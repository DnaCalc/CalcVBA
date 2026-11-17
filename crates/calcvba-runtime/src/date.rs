//! Deterministic VBA date text parsing shared by compile-time constants and VM
//! coercion. Locale-sensitive numeric forms such as `2/3/2026` intentionally
//! return `None`; callers that need host-locale parsing must provide an explicit
//! host collation/date service.

#[derive(Debug, Clone, PartialEq)]
enum LitTok {
    /// Digits and their text (the length distinguishes `99` from `1999`).
    Num(String),
    Word(String),
    Colon,
}

fn tokenize_literal(inner: &str) -> Option<Vec<LitTok>> {
    let mut out = Vec::new();
    let mut chars = inner.chars().peekable();
    while let Some(&c) = chars.peek() {
        if c.is_ascii_digit() {
            let mut n = String::new();
            while let Some(&d) = chars.peek().filter(|d| d.is_ascii_digit()) {
                n.push(d);
                chars.next();
            }
            out.push(LitTok::Num(n));
        } else if c.is_ascii_alphabetic() {
            let mut w = String::new();
            while let Some(&d) = chars.peek().filter(|d| d.is_ascii_alphabetic()) {
                w.push(d.to_ascii_lowercase());
                chars.next();
            }
            out.push(LitTok::Word(w));
        } else if c == ':' {
            out.push(LitTok::Colon);
            chars.next();
        } else if c.is_whitespace() || matches!(c, '/' | '-' | ',' | '.') {
            chars.next();
        } else {
            return None;
        }
    }
    Some(out)
}

/// A source date literal (MS-VBAL 3.3.3): month/day/year order whatever the host locale, with
/// optional time and AM/PM, month names, two-part dates and two-digit year windowing.
pub fn parse_date_literal_serial_bits(text: &str) -> Option<u64> {
    let inner = text.strip_prefix('#')?.strip_suffix('#')?.trim();
    let mut toks = tokenize_literal(inner)?;
    // AM/PM marker: am, pm, a or p.
    let mut ampm = None;
    if let Some(i) = toks
        .iter()
        .position(|t| matches!(t, LitTok::Word(w) if matches!(w.as_str(), "am" | "pm" | "a" | "p")))
    {
        let LitTok::Word(w) = toks.remove(i) else {
            return None;
        };
        ampm = Some(w.starts_with('p'));
        // The marker must follow the time or hour.
        if i == 0 || !matches!(toks[i - 1], LitTok::Num(_)) {
            return None;
        }
    }
    let nums = |t: &[LitTok]| t.iter().any(|x| matches!(x, LitTok::Colon));
    let time = if nums(&toks) {
        let colon = toks.iter().position(|t| *t == LitTok::Colon)?;
        if colon == 0 {
            return None;
        }
        let start = colon - 1;
        let mut end = colon;
        let mut fields = 1;
        while end + 1 < toks.len() && toks[end] == LitTok::Colon {
            if !matches!(toks[end + 1], LitTok::Num(_)) {
                return None;
            }
            end += 2;
            fields += 1;
        }
        if !(2..=3).contains(&fields) {
            return None;
        }
        let values: Vec<u32> = toks[start..end]
            .iter()
            .filter_map(|t| match t {
                LitTok::Num(n) => n.parse().ok(),
                _ => None,
            })
            .collect();
        toks.drain(start..end);
        Some(time_fraction(&values, ampm)?)
    } else if let Some(pm) = ampm {
        // `#3 PM#`: the number before the marker is the hour.
        let LitTok::Num(h) = toks.pop()? else {
            return None;
        };
        // Bare hour must be the last token before the removed marker; the marker sat after it.
        Some(time_fraction(&[h.parse().ok()?, 0], Some(pm))?)
    } else {
        None
    };
    let date_serial = if toks.is_empty() {
        0.0
    } else {
        literal_date_serial(&toks)?
    };
    Some((date_serial + time.unwrap_or(0.0)).to_bits())
}

/// Hour with optional PM flag: an hour above 12 ignores the marker, `0 PM` is noon.
fn time_fraction(values: &[u32], pm: Option<bool>) -> Option<f64> {
    let hour = *values.first()?;
    let minute = *values.get(1)?;
    let second = values.get(2).copied().unwrap_or(0);
    if hour > 23 || minute > 59 || second > 59 {
        return None;
    }
    let hour = match pm {
        Some(true) if hour < 12 => hour + 12,
        Some(true) => {
            if hour == 0 {
                12
            } else {
                hour
            }
        }
        Some(false) if hour == 12 => 0,
        _ => hour,
    };
    Some(f64::from(hour * 3600 + minute * 60 + second) / 86_400.0)
}

fn current_year() -> i32 {
    let secs = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map_or(0, |d| d.as_secs() as i64);
    let days = secs.div_euclid(86_400) + 719_468;
    let era = days.div_euclid(146_097);
    let doe = days - era * 146_097;
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let year = yoe + era * 400 + i64::from(mp >= 10);
    year as i32
}

/// A year token: two digits are windowed (00-29 -> 20xx, 30-99 -> 19xx); otherwise 100-9999.
fn literal_year(text: &str) -> Option<i32> {
    let v = text.parse::<i32>().ok()?;
    if text.len() <= 2 {
        Some(if v < 30 { 2000 + v } else { 1900 + v })
    } else if (100..=9999).contains(&v) {
        Some(v)
    } else {
        None
    }
}

fn literal_date_serial(toks: &[LitTok]) -> Option<f64> {
    #[derive(Clone)]
    enum D {
        N(String),
        M(i32),
    }
    let items: Vec<D> = toks
        .iter()
        .map(|t| match t {
            LitTok::Num(n) => Some(D::N(n.clone())),
            LitTok::Word(w) => parse_month_token(w).map(D::M),
            LitTok::Colon => None,
        })
        .collect::<Option<_>>()?;
    let num = |s: &str| s.parse::<i32>().ok();
    let year_like = |s: &str| s.len() >= 3 || num(s).is_some_and(|v| v > 31);
    let (year, month, day) = match items.as_slice() {
        [D::N(a), D::N(b)] => {
            if year_like(b) {
                (literal_year(b)?, num(a)?, 1)
            } else if year_like(a) {
                (literal_year(a)?, num(b)?, 1)
            } else {
                (current_year(), num(a)?, num(b)?)
            }
        }
        [D::M(m), D::N(n)] | [D::N(n), D::M(m)] => {
            if year_like(n) {
                (literal_year(n)?, *m, 1)
            } else {
                (current_year(), *m, num(n)?)
            }
        }
        [D::M(m), D::N(a), D::N(b)] => (literal_year(b)?, *m, num(a)?),
        [D::N(a), D::M(m), D::N(b)] => {
            if year_like(a) {
                (literal_year(a)?, *m, num(b)?)
            } else {
                (literal_year(b)?, *m, num(a)?)
            }
        }
        [D::N(a), D::N(b), D::M(m)] => {
            if year_like(a) {
                (literal_year(a)?, *m, num(b)?)
            } else {
                (literal_year(b)?, *m, num(a)?)
            }
        }
        [D::N(a), D::N(b), D::N(c)] => {
            let first = num(a)?;
            if year_like(a) {
                (literal_year(a)?, num(b)?, num(c)?)
            } else if (1..=12).contains(&first) {
                (literal_year(c)?, first, num(b)?)
            } else if c.len() >= 3 {
                (literal_year(c)?, num(b)?, first)
            } else {
                (literal_year(a)?, num(b)?, num(c)?)
            }
        }
        _ => return None,
    };
    if !(100..=9999).contains(&year) {
        return None;
    }
    packed_date_to_ole_serial(year * 10_000 + month.checked_mul(100)? + day)
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
        near("#1/1/2026 13:00 PM#", 46_023.0 + 13.0 / 24.0);
    }

    #[test]
    fn date_literal_forms_match_excel() {
        let near = |text: &str, want: f64| {
            let got = f64::from_bits(parse_date_literal_serial_bits(text).expect(text));
            assert!((got - want).abs() < 1e-9, "{text}: {got} vs {want}");
        };
        near("#3 PM#", 0.625);
        near("#3 P#", 0.625);
        near("#8:30AM#", 0.354_166_666_666_667);
        near("#8:30 a#", 0.354_166_666_666_667);
        near("#2-3-2026#", 46_056.0);
        near("#3/2020#", 43_891.0);
        near("#1/1/2026 0:00 PM#", 46_023.5);
        near("#1/1/2026 12:00 AM#", 46_023.0);
        near("#99/12/31#", 36_525.0);
        near("#1/1/99#", 36_161.0);
        near("#1/1/29#", 47_119.0);
        near("#Jan 2 2026#", 46_024.0);
        near("#2 January 2026#", 46_024.0);
        near("#January 2, 2026#", 46_024.0);
        near("#1/1/2026 8:30:15 PM#", 46_023.0 + 73_815.0 / 86_400.0);
        let year = super::current_year();
        let got = f64::from_bits(parse_date_literal_serial_bits("#1/15#").unwrap());
        let want = parse_date_literal_serial_bits(&format!("#1/15/{year}#")).unwrap();
        assert_eq!(got.to_bits(), want);
        assert_eq!(parse_date_literal_serial_bits("#1/1/99999#"), None);
        assert_eq!(parse_date_literal_serial_bits("#1/1/50 PM#"), None);
        assert_eq!(parse_date_literal_serial_bits("#2/30/2026#"), None);
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
