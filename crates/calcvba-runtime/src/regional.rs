//! Host regional settings for VBA date and number text.

use std::cell::RefCell;
use std::rc::Rc;

use crate::vba_date::{days_from_civil, serial_day_and_seconds, serial_to_hms, serial_to_ymd};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateOrder {
    MonthDayYear,
    DayMonthYear,
    YearMonthDay,
}

#[derive(Clone)]
pub struct RegionalSettings {
    pub date_order: DateOrder,
    pub short_date: String,
    pub long_date: String,
    pub long_time: String,
    pub medium_time: String,
    pub am: String,
    pub pm: String,
    pub months: [String; 12],
    pub abbreviated_months: [String; 12],
    pub weekdays: [String; 7],
    pub abbreviated_weekdays: [String; 7],
    pub decimal_separator: String,
    pub group_separator: String,
    pub currency_symbol: String,
    pub two_digit_year_max: i64,
}

impl RegionalSettings {
    pub fn en_us() -> Self {
        let months = [
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December",
        ];
        let weekdays = [
            "Sunday",
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
        ];
        Self {
            date_order: DateOrder::MonthDayYear,
            short_date: "M/d/yyyy".into(),
            long_date: "dddd, MMMM d, yyyy".into(),
            long_time: "h:mm:ss tt".into(),
            medium_time: "h:mm tt".into(),
            am: "AM".into(),
            pm: "PM".into(),
            months: months.map(str::to_owned),
            abbreviated_months: months.map(|name| name[..3].to_owned()),
            weekdays: weekdays.map(str::to_owned),
            abbreviated_weekdays: weekdays.map(|name| name[..3].to_owned()),
            decimal_separator: ".".into(),
            group_separator: ",".into(),
            currency_symbol: "$".into(),
            two_digit_year_max: 2029,
        }
    }

    pub fn en_za() -> Self {
        let mut settings = Self {
            date_order: DateOrder::YearMonthDay,
            short_date: "yyyy/MM/dd".into(),
            long_date: "d MMMM yyyy".into(),
            long_time: "HH:mm:ss".into(),
            medium_time: "hh:mm tt".into(),
            am: "am".into(),
            pm: "pm".into(),
            group_separator: "\u{a0}".into(),
            currency_symbol: "R".into(),
            ..Self::en_us()
        };
        settings.abbreviated_months[8] = "Sept".into();
        settings
    }

    fn by_name(name: &str) -> Option<Self> {
        let name = name.split(['.', '@']).next().unwrap_or(name);
        if name.eq_ignore_ascii_case("en-ZA") || name.eq_ignore_ascii_case("en_ZA") {
            Some(Self::en_za())
        } else if name.eq_ignore_ascii_case("en-US") || name.eq_ignore_ascii_case("en_US") {
            Some(Self::en_us())
        } else {
            None
        }
    }

    fn host() -> Self {
        if let Some(settings) = std::env::var("CALCVBA_LOCALE")
            .ok()
            .and_then(|name| Self::by_name(&name))
        {
            return settings;
        }
        #[cfg(windows)]
        {
            windows::settings()
        }
        #[cfg(not(windows))]
        {
            ["LC_ALL", "LC_TIME", "LANG"]
                .iter()
                .filter_map(|key| std::env::var(key).ok())
                .find(|name| !name.is_empty())
                .and_then(|name| Self::by_name(&name))
                .unwrap_or_else(Self::en_us)
        }
    }

    pub fn expand_year(&self, year: i64) -> i64 {
        if !(0..=99).contains(&year) {
            return year;
        }
        let century = self.two_digit_year_max.div_euclid(100) * 100;
        let candidate = century + year;
        if candidate > self.two_digit_year_max {
            candidate - 100
        } else {
            candidate
        }
    }

    pub fn designator(&self, hour: i64) -> &str {
        if hour < 12 { &self.am } else { &self.pm }
    }

    pub fn format_general_date(&self, serial: f64) -> String {
        let (day, seconds) = serial_day_and_seconds(serial);
        let date = || self.format_date_pattern(serial, &self.short_date, false);
        let time = || self.format_date_pattern(serial, &self.long_time, true);
        match (day != 0, seconds != 0) {
            (true, true) => format!("{} {}", date(), time()),
            (true, false) => date(),
            (false, _) => time(),
        }
    }

    /// Render a Windows regional date or time pattern.
    pub fn format_date_pattern(&self, serial: f64, pattern: &str, pad_year: bool) -> String {
        let (y, m, d) = serial_to_ymd(serial);
        let (h, minute, s) = serial_to_hms(serial);
        let month = (m - 1) as usize;
        let day = (days_from_civil(y, m, d).rem_euclid(7) + 4) % 7;
        let day = day as usize;
        let hour12 = (h + 11) % 12 + 1;
        let mut chars = pattern.chars().peekable();
        let mut output = String::new();
        while let Some(token) = chars.next() {
            if token == '\'' {
                output.extend(chars.by_ref().take_while(|&ch| ch != '\''));
                continue;
            }
            let mut width = 1;
            while chars.next_if_eq(&token).is_some() {
                width += 1;
            }
            let name = match (token, width) {
                ('M', 3) => Some(self.abbreviated_months[month].as_str()),
                ('M', 4..) => Some(self.months[month].as_str()),
                ('d', 3) => Some(self.abbreviated_weekdays[day].as_str()),
                ('d', 4..) => Some(self.weekdays[day].as_str()),
                ('t', 2..) => Some(self.designator(h)),
                _ => None,
            };
            if let Some(name) = name {
                output.push_str(name);
                continue;
            }
            let number = match token {
                'M' => m,
                'd' => d,
                'y' if width <= 2 => y.rem_euclid(100),
                'y' => y,
                'H' => h,
                'h' => hour12,
                'm' => minute,
                's' => s,
                't' => {
                    output.extend(self.designator(h).chars().take(1));
                    continue;
                }
                other => {
                    output.extend(std::iter::repeat_n(other, width));
                    continue;
                }
            };
            if token == 'y' && width >= 3 && !pad_year {
                width = 1;
            }
            use std::fmt::Write as _;
            let _ = write!(output, "{number:0width$}");
        }
        output
    }
}

thread_local! {
    static SETTINGS: RefCell<Rc<RegionalSettings>> =
        RefCell::new(Rc::new(RegionalSettings::host()));
}

pub fn current() -> Rc<RegionalSettings> {
    SETTINGS.with_borrow(Rc::clone)
}

/// Run with explicit regional settings on the current VBA thread.
pub fn with_settings<T>(settings: RegionalSettings, run: impl FnOnce() -> T) -> T {
    struct Restore(Rc<RegionalSettings>);
    impl Drop for Restore {
        fn drop(&mut self) {
            SETTINGS.with_borrow_mut(|current| *current = Rc::clone(&self.0));
        }
    }
    let _restore = Restore(SETTINGS.replace(Rc::new(settings)));
    run()
}

#[cfg(windows)]
mod windows {
    use super::*;
    use windows_sys::Win32::Globalization::*;

    fn read(kind: u32) -> Option<String> {
        // SAFETY: a null locale name selects the current user's settings and a
        // null output with length zero asks Windows for the buffer size.
        let len = unsafe { GetLocaleInfoEx(std::ptr::null(), kind, std::ptr::null_mut(), 0) };
        if len <= 0 {
            return None;
        }
        let mut buffer = vec![0u16; len as usize];
        // SAFETY: the buffer owns len initialized, writable UTF-16 units.
        let written = unsafe { GetLocaleInfoEx(std::ptr::null(), kind, buffer.as_mut_ptr(), len) };
        (written > 0).then(|| String::from_utf16_lossy(&buffer[..written as usize - 1]))
    }

    pub(super) fn settings() -> RegionalSettings {
        let mut settings = read(LOCALE_SNAME)
            .and_then(|name| RegionalSettings::by_name(&name))
            .unwrap_or_else(RegionalSettings::en_us);
        for (kind, field) in [
            (LOCALE_SSHORTDATE, &mut settings.short_date),
            (LOCALE_SLONGDATE, &mut settings.long_date),
            (LOCALE_STIMEFORMAT, &mut settings.long_time),
            (LOCALE_S1159, &mut settings.am),
            (LOCALE_S2359, &mut settings.pm),
            (LOCALE_SDECIMAL, &mut settings.decimal_separator),
            (LOCALE_STHOUSAND, &mut settings.group_separator),
            (LOCALE_SCURRENCY, &mut settings.currency_symbol),
        ] {
            if let Some(value) = read(kind) {
                *field = value;
            }
        }
        settings.date_order = match read(LOCALE_IDATE).as_deref() {
            Some("0") => DateOrder::MonthDayYear,
            Some("1") => DateOrder::DayMonthYear,
            Some("2") => DateOrder::YearMonthDay,
            _ => settings.date_order,
        };
        for (first, names) in [
            (LOCALE_SMONTHNAME1, &mut settings.months),
            (LOCALE_SABBREVMONTHNAME1, &mut settings.abbreviated_months),
        ] {
            for (index, name) in names.iter_mut().enumerate() {
                if let Some(value) = read(first + index as u32) {
                    *name = value;
                }
            }
        }
        for (first, names) in [
            (LOCALE_SDAYNAME1, &mut settings.weekdays),
            (LOCALE_SABBREVDAYNAME1, &mut settings.abbreviated_weekdays),
        ] {
            for (index, name) in names.iter_mut().enumerate() {
                let monday_based = (index + 6) % 7;
                if let Some(value) = read(first + monday_based as u32) {
                    *name = value;
                }
            }
        }
        let mut max = 0;
        // SAFETY: requested numeric output is a uniquely writable DWORD.
        let count = unsafe {
            GetCalendarInfoEx(
                std::ptr::null(),
                CAL_GREGORIAN,
                std::ptr::null(),
                CAL_ITWODIGITYEARMAX | CAL_RETURN_NUMBER,
                std::ptr::null_mut(),
                0,
                &mut max,
            )
        };
        if count > 0 {
            settings.two_digit_year_max = i64::from(max);
        }
        settings
    }
}
