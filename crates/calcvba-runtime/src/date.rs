//! Deterministic VBA date text parsing shared by compile-time constants and VM
//! coercion. Locale-sensitive numeric forms such as `2/3/2026` intentionally
//! return `None`; callers that need host-locale parsing must provide an explicit
//! host collation/date service.

pub fn parse_date_literal_serial_bits(text: &str) -> Option<u64> {
    let inner = text.strip_prefix('#')?.strip_suffix('#')?.trim();
    // A source literal is always month/day/year, never the host locale's order, and may end
    // in a time (`#6/15/2005 8:30:00 AM#`, `#13:45#`).
    let mut tokens: Vec<&str> = inner.split_whitespace().collect();
    let mut time = None;
    if let Some(last) = tokens.last().copied() {
        let ampm = last.eq_ignore_ascii_case("am") || last.eq_ignore_ascii_case("pm");
        if ampm && tokens.len() >= 2 && tokens[tokens.len() - 2].contains(':') {
            time = parse_time_of_day(tokens[tokens.len() - 2], Some(last));
            tokens.truncate(tokens.len() - 2);
        } else if last.contains(':') {
            time = parse_time_of_day(last, None);
            tokens.pop();
        }
        if time.is_none() && (ampm || last.contains(':')) {
            return None;
        }
    }
    let date_text = tokens.join(" ");
    let date_serial = if date_text.is_empty() {
        0.0
    } else if let Some(serial) = parse_month_day_year_slashes(&date_text) {
        serial
    } else {
        parse_date_text_serial(&date_text)?
    };
    Some((date_serial + time.unwrap_or(0.0)).to_bits())
}

/// `m/d/yyyy` with numeric parts, always in month-day-year order.
fn parse_month_day_year_slashes(text: &str) -> Option<f64> {
    let parts: Vec<&str> = text.split('/').collect();
    let [month, day, year] = parts.as_slice() else {
        return None;
    };
    let month = month.trim().parse::<i32>().ok()?;
    let day = day.trim().parse::<i32>().ok()?;
    let year = year.trim().parse::<i32>().ok()?;
    packed_date_to_ole_serial(year.saturating_mul(10_000) + month.saturating_mul(100) + day)
}

/// `h:mm[:ss]` with an optional AM/PM marker, as a fraction of a day.
fn parse_time_of_day(text: &str, ampm: Option<&str>) -> Option<f64> {
    let parts: Vec<&str> = text.split(':').collect();
    if !(2..=3).contains(&parts.len()) {
        return None;
    }
    let mut hour = parts[0].parse::<u32>().ok()?;
    let minute = parts[1].parse::<u32>().ok()?;
    let second = parts.get(2).map_or(Some(0), |s| s.parse::<u32>().ok())?;
    if minute > 59 || second > 59 {
        return None;
    }
    match ampm.map(str::to_ascii_lowercase).as_deref() {
        Some("am") => {
            if !(1..=12).contains(&hour) {
                return None;
            }
            hour %= 12;
        }
        Some("pm") => {
            if !(1..=12).contains(&hour) {
                return None;
            }
            hour = hour % 12 + 12;
        }
        _ => {
            if hour > 23 {
                return None;
            }
        }
    }
    Some(f64::from(hour * 3600 + minute * 60 + second) / 86_400.0)
}

pub fn parse_date_text_serial_bits(text: &str) -> Option<u64> {
    Some(parse_date_text_serial(text)?.to_bits())
}

pub fn parse_date_text_serial(text: &str) -> Option<f64> {
    let packed = parse_date_text_to_packed(text)?;
    packed_date_to_ole_serial(packed)
}

fn parse_date_text_to_packed(text: &str) -> Option<i32> {
    let normalized = text.trim().replace([',', '.', '-', '/'], " ");
    let parts: Vec<&str> = normalized.split_whitespace().collect();
    let packed = match parts.as_slice() {
        [year, month, day] if year.len() == 4 => {
            let year = year.parse::<i32>().ok()?;
            let month = parse_month_token(month).or_else(|| month.parse::<i32>().ok())?;
            let day = day.parse::<i32>().ok()?;
            year.saturating_mul(10_000) + month.saturating_mul(100) + day
        }
        [month, day, year] if parse_month_token(month).is_some() => {
            let month = parse_month_token(month)?;
            let day = day.parse::<i32>().ok()?;
            let year = year.parse::<i32>().ok()?;
            year.saturating_mul(10_000) + month.saturating_mul(100) + day
        }
        [day, month, year] if parse_month_token(month).is_some() => {
            let day = day.parse::<i32>().ok()?;
            let month = parse_month_token(month)?;
            let year = year.parse::<i32>().ok()?;
            year.saturating_mul(10_000) + month.saturating_mul(100) + day
        }
        [first, second, year] => {
            let first = first.parse::<i32>().ok()?;
            let second = second.parse::<i32>().ok()?;
            let year = year.parse::<i32>().ok()?;
            let (month, day) = match unambiguous_numeric_date_order(first, second)? {
                NumericDateOrder::MonthDayYear => (first, second),
                NumericDateOrder::DayMonthYear => (second, first),
            };
            year.saturating_mul(10_000) + month.saturating_mul(100) + day
        }
        _ => return None,
    };
    packed_date_components(packed)?;
    Some(packed)
}

enum NumericDateOrder {
    MonthDayYear,
    DayMonthYear,
}

fn unambiguous_numeric_date_order(first: i32, second: i32) -> Option<NumericDateOrder> {
    let first_can_be_month = (1..=12).contains(&first);
    let second_can_be_month = (1..=12).contains(&second);
    match (first_can_be_month, second_can_be_month) {
        // `2/3/2026` is locale-sensitive; `2/2/2026` has the same result in either
        // numeric order, so it remains deterministic.
        (true, true) if first == second => Some(NumericDateOrder::MonthDayYear),
        (true, true) => None,
        (true, false) => Some(NumericDateOrder::MonthDayYear),
        (false, true) => Some(NumericDateOrder::DayMonthYear),
        (false, false) => None,
    }
}

fn parse_month_token(text: &str) -> Option<i32> {
    match text.trim().to_ascii_lowercase().as_str() {
        "jan" | "january" => Some(1),
        "feb" | "february" => Some(2),
        "mar" | "march" => Some(3),
        "apr" | "april" => Some(4),
        "may" => Some(5),
        "jun" | "june" => Some(6),
        "jul" | "july" => Some(7),
        "aug" | "august" => Some(8),
        "sep" | "sept" | "september" => Some(9),
        "oct" | "october" => Some(10),
        "nov" | "november" => Some(11),
        "dec" | "december" => Some(12),
        _ => None,
    }
}

fn packed_date_components(packed: i32) -> Option<(i32, u32, u32)> {
    let year = packed / 10_000;
    let month = ((packed / 100) % 100) as u32;
    let day = (packed % 100) as u32;
    if !(1..=12).contains(&month) {
        return None;
    }
    let max_day = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if is_leap_year(year) => 29,
        2 => 28,
        _ => 0,
    };
    if max_day == 0 || day == 0 || day > max_day {
        return None;
    }
    Some((year, month, day))
}

fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

fn packed_date_to_ole_serial(packed: i32) -> Option<f64> {
    let (year, month, day) = packed_date_components(packed)?;
    Some((days_from_civil(year, month, day) + 25_569) as f64)
}

fn days_from_civil(year: i32, month: u32, day: u32) -> i64 {
    let year = i64::from(year) - i64::from((month <= 2) as i32);
    let era = if year >= 0 { year } else { year - 399 } / 400;
    let year_of_era = year - era * 400;
    let month_index = i64::from(month) + if month > 2 { -3 } else { 9 };
    let day_of_year = (153 * month_index + 2) / 5 + i64::from(day) - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era * 146_097 + day_of_era - 719_468
}

#[cfg(test)]
mod tests {
    use super::{parse_date_literal_serial_bits, parse_date_text_serial_bits};

    #[test]
    fn date_literals_accept_unambiguous_numeric_orders() {
        let expected = Some(46_081.0f64.to_bits());
        assert_eq!(parse_date_literal_serial_bits("#2026-02-28#"), expected);
        assert_eq!(parse_date_literal_serial_bits("#2/28/2026#"), expected);
        assert_eq!(parse_date_literal_serial_bits("#28/2/2026#"), expected);
    }

    #[test]
    fn date_literals_are_month_day_year_and_take_a_time() {
        // A source literal is never read in the host locale's order.
        let feb3 = parse_date_literal_serial_bits("#2/3/2026#");
        assert_eq!(feb3, Some(46_056.0f64.to_bits()));
        let near = |text: &str, want: f64| {
            let got = f64::from_bits(parse_date_literal_serial_bits(text).expect(text));
            assert!((got - want).abs() < 1e-9, "{text}: {got} vs {want}");
        };
        near("#6/15/2005 8:30:00 AM#", 38_518.0 + 8.5 / 24.0);
        near("#12/31/1999 11:59:59 PM#", 36_525.0 + 86_399.0 / 86_400.0);
        near("#13:30#", 13.5 / 24.0);
        assert_eq!(parse_date_literal_serial_bits("#1/1/2026 13:00 PM#"), None);
    }

    #[test]
    fn date_text_accepts_same_deterministic_orders_as_literals() {
        let expected = Some(46_081.0f64.to_bits());
        assert_eq!(parse_date_text_serial_bits("2026-02-28"), expected);
        assert_eq!(parse_date_text_serial_bits("February 28, 2026"), expected);
        assert_eq!(parse_date_text_serial_bits("28/2/2026"), expected);
        assert_eq!(parse_date_text_serial_bits("2/3/2026"), None);
    }
}
