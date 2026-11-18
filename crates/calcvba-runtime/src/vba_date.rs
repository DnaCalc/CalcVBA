//! VBA `Date` serial ↔ civil-calendar decomposition and display formatting.
//!
//! A VBA `Date` is an OLE-Automation serial: the integer part is days since 1899-12-30 and the
//! fractional part is the time of day. The calendar is proleptic Gregorian (no Excel 1900
//! leap-year quirk). This is the single source of truth for the serial math — `calcvba-lib`'s
//! date functions re-import it, and the value-display formatters here (`variant_to_vba_string`
//! / `print_display_text` / `write_display_text`) use it so a `Date` renders as a formatted
//! date string, not its raw serial number.

/// Days from the Unix epoch (1970-01-01) back to the VBA epoch (1899-12-30).
pub const VBA_EPOCH_DAYS_FROM_UNIX: i64 = -25569;

/// Days since 1970-01-01 for a proleptic-Gregorian `(y, m, d)` (Howard Hinnant's algorithm).
/// Counts outside the i64 day carrier saturate; VBA-facing callers check Date bounds.
pub fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    days_from_civil_wide(y, m, d).clamp(i128::from(i64::MIN), i128::from(i64::MAX)) as i64
}

fn days_from_civil_wide(y: i64, m: i64, d: i64) -> i128 {
    let (y, m, d) = (i128::from(y), i128::from(m), i128::from(d));
    // A month outside 1-12 carries into the year (`DateSerial(2000, 15, 1)` is March 2001).
    let y = y + (m - 1).div_euclid(12);
    let m = (m - 1).rem_euclid(12) + 1;
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let doy = (153 * (if m > 2 { m - 3 } else { m + 9 }) + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe - 719468
}

/// Proleptic-Gregorian `(y, m, d)` for a days-since-1970-01-01 count (inverse of
/// [`days_from_civil`]).
pub fn civil_from_days(z: i64) -> (i64, i64, i64) {
    civil_from_days_wide(i128::from(z))
}

fn civil_from_days_wide(z: i128) -> (i64, i64, i64) {
    let z = z + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    // An i64 day count plus the epoch offset has a year well inside i64.
    ((if m <= 2 { y + 1 } else { y }) as i64, m as i64, d as i64)
}

/// The VBA `Date` serial for a `(y, m, d)`.
pub fn ymd_to_serial(y: i64, m: i64, d: i64) -> f64 {
    (days_from_civil_wide(y, m, d) - i128::from(VBA_EPOCH_DAYS_FROM_UNIX)) as f64
}

/// Proleptic-Gregorian leap year (divisible by 4, except centuries not divisible by 400).
pub fn is_leap_year(y: i64) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

/// The number of days in month `m` (1-12) of proleptic-Gregorian year `y`.
pub fn days_in_month(y: i64, m: i64) -> i64 {
    match m {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if is_leap_year(y) => 29,
        2 => 28,
        _ => 30,
    }
}

fn serial_day_and_seconds(serial: f64) -> (i64, i64) {
    let mut day = serial.trunc() as i64;
    let frac = (serial - serial.trunc()).abs();
    let mut total = (frac * 86_400.0).round() as i64;
    if total >= 86_400 {
        day = day.saturating_add(1);
        total -= 86_400;
    }
    (day, total)
}

/// The `(year, month, day)` of a VBA `Date` serial.
pub fn serial_to_ymd(serial: f64) -> (i64, i64, i64) {
    let (day, _) = serial_day_and_seconds(serial);
    civil_from_days_wide(i128::from(day) + i128::from(VBA_EPOCH_DAYS_FROM_UNIX))
}

/// The `(hour, minute, second)` of a VBA `Date` serial's time-of-day fraction.
pub fn serial_to_hms(serial: f64) -> (i64, i64, i64) {
    let (_, total) = serial_day_and_seconds(serial);
    (total / 3600, (total % 3600) / 60, total % 60)
}

/// VBA's long-time display: `h:mm:ss AM/PM` (12-hour, no leading zero on the hour).
fn format_long_time(h: i64, mi: i64, s: i64) -> String {
    let (h12, ampm) = match h {
        0 => (12, "AM"),
        1..=11 => (h, "AM"),
        12 => (12, "PM"),
        _ => (h - 12, "PM"),
    };
    format!("{h12}:{mi:02}:{s:02} {ampm}")
}

/// VBA's "General Date" display (en-US): a value with a real date and a time renders as
/// `M/D/YYYY h:mm:ss AM/PM`; a date with no time renders as `M/D/YYYY`; a value whose date part
/// is the epoch (1899-12-30, integer part 0) renders as the time only, `h:mm:ss AM/PM`.
pub fn format_general_date(serial: f64) -> String {
    let (y, m, d) = serial_to_ymd(serial);
    let (h, mi, s) = serial_to_hms(serial);
    let has_date = serial.trunc() != 0.0;
    let has_time = (h, mi, s) != (0, 0, 0);
    let date = format!("{m}/{d}/{y}");
    let time = format_long_time(h, mi, s);
    match (has_date, has_time) {
        (true, true) => format!("{date} {time}"),
        (true, false) => date,
        (false, _) => time,
    }
}

/// VBA `Write #` date literal: `#YYYY-MM-DD HH:MM:SS#` (or `#YYYY-MM-DD#` at midnight) — the
/// locale-independent, machine-readable form `Write` emits (distinct from "General Date").
pub fn format_write_date(serial: f64) -> String {
    let (y, m, d) = serial_to_ymd(serial);
    let (h, mi, s) = serial_to_hms(serial);
    if (h, mi, s) == (0, 0, 0) {
        format!("#{y:04}-{m:02}-{d:02}#")
    } else {
        format!("#{y:04}-{m:02}-{d:02} {h:02}:{mi:02}:{s:02}#")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn civil_calendar_round_trips_extreme_day_counts() {
        for day in [
            i64::MIN,
            i64::MIN + 1,
            -719_469,
            -1,
            0,
            i64::MAX - 1,
            i64::MAX,
        ] {
            let (year, month, date) = civil_from_days(day);
            assert!((1..=12).contains(&month));
            assert!((1..=days_in_month(year, month)).contains(&date));
            assert_eq!(days_from_civil(year, month, date), day);
        }
        assert_eq!(days_from_civil(i64::MAX, i64::MAX, i64::MAX), i64::MAX);
        assert_eq!(days_from_civil(i64::MIN, i64::MIN, i64::MIN), i64::MIN);
    }

    #[test]
    fn invalid_serial_display_cannot_overflow_calendar_math() {
        for serial in [
            f64::MIN,
            -1e30,
            1e30,
            f64::MAX,
            f64::NEG_INFINITY,
            f64::INFINITY,
            f64::NAN,
        ] {
            let (_, month, date) = serial_to_ymd(serial);
            assert!((1..=12).contains(&month));
            assert!((1..=31).contains(&date));
            let _ = format_general_date(serial);
            let _ = format_write_date(serial);
        }
        assert!(ymd_to_serial(i64::MAX, i64::MAX, i64::MAX).is_finite());
        assert!(ymd_to_serial(i64::MIN, i64::MIN, i64::MIN).is_finite());
    }

    #[test]
    fn general_date_date_only() {
        // 2020-01-15 (no time).
        let serial = ymd_to_serial(2020, 1, 15);
        assert_eq!(format_general_date(serial), "1/15/2020");
    }

    #[test]
    fn general_date_with_time() {
        // 2020-01-15 13:30:00 → 1:30:00 PM.
        let serial = ymd_to_serial(2020, 1, 15) + (13.0 * 3600.0 + 30.0 * 60.0) / 86_400.0;
        assert_eq!(format_general_date(serial), "1/15/2020 1:30:00 PM");
    }

    #[test]
    fn general_date_time_only_and_midnight() {
        // Epoch-date noon → time only.
        assert_eq!(format_general_date(0.5), "12:00:00 PM");
        // Exactly 0 (epoch midnight) → time only, 12:00:00 AM.
        assert_eq!(format_general_date(0.0), "12:00:00 AM");
    }

    #[test]
    fn negative_serial_uses_whole_number_date_part() {
        assert_eq!(format_general_date(-1.25), "12/29/1899 6:00:00 AM");
    }

    #[test]
    fn rounded_time_can_carry_to_next_date() {
        let almost_next_day = ymd_to_serial(2020, 1, 15) + 86_399.6 / 86_400.0;
        assert_eq!(format_general_date(almost_next_day), "1/16/2020");
    }

    #[test]
    fn write_date_form() {
        assert_eq!(
            format_write_date(ymd_to_serial(2020, 1, 15)),
            "#2020-01-15#"
        );
        let serial = ymd_to_serial(2020, 1, 15) + 0.5;
        assert_eq!(format_write_date(serial), "#2020-01-15 12:00:00#");
    }
}
