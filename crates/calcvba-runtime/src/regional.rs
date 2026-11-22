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

/// Check an explicit VBA locale ID.
pub fn locale_id_supported(id: i32) -> bool {
    let Ok(id) = u32::try_from(id) else {
        return false;
    };
    if id == 0 {
        return true;
    }
    #[cfg(windows)]
    {
        windows::locale_id_supported(id)
    }
    #[cfg(not(windows))]
    {
        matches!(id, 0x0400 | 0x0800) || WINDOWS_LCIDS.binary_search(&id).is_ok()
    }
}

// MS-LCID §2.2 defines the permanent language IDs and alternate sorts.
// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-lcid/63d3d639-7fd2-4afb-abbe-0d5b5551eef8
// The Windows path asks the host about identifiers specific to that installation.
#[cfg(not(windows))]
#[rustfmt::skip]
const WINDOWS_LCIDS: &[u32] = &[
    0x0001, 0x0002, 0x0003, 0x0004, 0x0005, 0x0006, 0x0007, 0x0008,
    0x0009, 0x000a, 0x000b, 0x000c, 0x000d, 0x000e, 0x000f, 0x0010,
    0x0011, 0x0012, 0x0013, 0x0014, 0x0015, 0x0016, 0x0017, 0x0018,
    0x0019, 0x001a, 0x001b, 0x001c, 0x001d, 0x001e, 0x001f, 0x0020,
    0x0021, 0x0022, 0x0023, 0x0024, 0x0025, 0x0026, 0x0027, 0x0028,
    0x0029, 0x002a, 0x002b, 0x002c, 0x002d, 0x002e, 0x002f, 0x0030,
    0x0031, 0x0032, 0x0033, 0x0034, 0x0035, 0x0036, 0x0037, 0x0038,
    0x0039, 0x003a, 0x003b, 0x003c, 0x003e, 0x003f, 0x0040, 0x0041,
    0x0042, 0x0043, 0x0044, 0x0045, 0x0046, 0x0047, 0x0048, 0x0049,
    0x004a, 0x004b, 0x004c, 0x004d, 0x004e, 0x004f, 0x0050, 0x0051,
    0x0052, 0x0053, 0x0054, 0x0055, 0x0056, 0x0057, 0x0059, 0x005a,
    0x005b, 0x005c, 0x005d, 0x005e, 0x005f, 0x0060, 0x0061, 0x0062,
    0x0063, 0x0064, 0x0065, 0x0067, 0x0068, 0x006a, 0x006b, 0x006c,
    0x006d, 0x006e, 0x006f, 0x0070, 0x0072, 0x0073, 0x0074, 0x0075,
    0x0078, 0x007a, 0x007c, 0x007e, 0x007f, 0x0080, 0x0081, 0x0082,
    0x0083, 0x0084, 0x0085, 0x0086, 0x0087, 0x0088, 0x008c, 0x0091,
    0x0092, 0x0401, 0x0402, 0x0403, 0x0404, 0x0405, 0x0406, 0x0407,
    0x0408, 0x0409, 0x040a, 0x040b, 0x040c, 0x040d, 0x040e, 0x040f,
    0x0410, 0x0411, 0x0412, 0x0413, 0x0414, 0x0415, 0x0416, 0x0417,
    0x0418, 0x0419, 0x041a, 0x041b, 0x041c, 0x041d, 0x041e, 0x041f,
    0x0420, 0x0421, 0x0422, 0x0423, 0x0424, 0x0425, 0x0426, 0x0427,
    0x0428, 0x0429, 0x042a, 0x042b, 0x042c, 0x042d, 0x042e, 0x042f,
    0x0430, 0x0431, 0x0432, 0x0433, 0x0434, 0x0435, 0x0436, 0x0437,
    0x0438, 0x0439, 0x043a, 0x043b, 0x043d, 0x043e, 0x043f, 0x0440,
    0x0441, 0x0442, 0x0443, 0x0444, 0x0445, 0x0446, 0x0447, 0x0448,
    0x0449, 0x044a, 0x044b, 0x044c, 0x044d, 0x044e, 0x044f, 0x0450,
    0x0451, 0x0452, 0x0453, 0x0454, 0x0455, 0x0456, 0x0457, 0x045a,
    0x045b, 0x045c, 0x045d, 0x045e, 0x045f, 0x0460, 0x0461, 0x0462,
    0x0463, 0x0464, 0x0465, 0x0467, 0x0468, 0x046a, 0x046b, 0x046c,
    0x046d, 0x046e, 0x046f, 0x0470, 0x0471, 0x0472, 0x0473, 0x0474,
    0x0475, 0x0476, 0x0477, 0x0478, 0x047a, 0x047c, 0x047e, 0x0480,
    0x0481, 0x0482, 0x0483, 0x0484, 0x0485, 0x0487, 0x0488, 0x048c,
    0x0491, 0x0492, 0x0501, 0x05fe, 0x0801, 0x0803, 0x0804, 0x0807,
    0x0809, 0x080a, 0x080c, 0x0810, 0x0813, 0x0814, 0x0816, 0x0818,
    0x0819, 0x081a, 0x081d, 0x0820, 0x082e, 0x0832, 0x083b, 0x083c,
    0x083e, 0x0845, 0x0846, 0x0849, 0x0859, 0x085d, 0x085f, 0x0860,
    0x0861, 0x0867, 0x086b, 0x0873, 0x09ff, 0x0c00, 0x0c01, 0x0c04,
    0x0c07, 0x0c09, 0x0c0a, 0x0c0c, 0x0c1a, 0x0c3b, 0x0c50, 0x0c51,
    0x0c6b, 0x1000, 0x1001, 0x1004, 0x1007, 0x1009, 0x100a, 0x100c,
    0x101a, 0x103b, 0x105f, 0x1401, 0x1404, 0x1407, 0x1409, 0x140a,
    0x140c, 0x141a, 0x143b, 0x1801, 0x1809, 0x180a, 0x180c, 0x181a,
    0x183b, 0x1c01, 0x1c09, 0x1c0a, 0x1c0c, 0x1c1a, 0x1c3b, 0x2000,
    0x2001, 0x2009, 0x200a, 0x200c, 0x201a, 0x203b, 0x2400, 0x2401,
    0x240a, 0x240c, 0x241a, 0x243b, 0x2800, 0x2801, 0x2809, 0x280a,
    0x280c, 0x281a, 0x2c00, 0x2c01, 0x2c09, 0x2c0a, 0x2c0c, 0x2c1a,
    0x3000, 0x3001, 0x3009, 0x300a, 0x300c, 0x301a, 0x3400, 0x3401,
    0x3409, 0x340a, 0x340c, 0x3800, 0x3801, 0x380a, 0x380c, 0x3c00,
    0x3c01, 0x3c09, 0x3c0a, 0x3c0c, 0x4000, 0x4001, 0x4009, 0x400a,
    0x4400, 0x4409, 0x440a, 0x4800, 0x4809, 0x480a, 0x4c00, 0x4c09,
    0x4c0a, 0x500a, 0x540a, 0x5c0a, 0x641a, 0x681a, 0x6c1a, 0x701a,
    0x703b, 0x742c, 0x743b, 0x7804, 0x7814, 0x781a, 0x782c, 0x783b,
    0x7843, 0x7850, 0x785d, 0x785f, 0x7c04, 0x7c14, 0x7c1a, 0x7c28,
    0x7c2e, 0x7c3b, 0x7c43, 0x7c46, 0x7c50, 0x7c59, 0x7c5c, 0x7c5d,
    0x7c5f, 0x7c67, 0x7c68, 0x7c92, 0x0001_007f, 0x0001_0407, 0x0001_040e, 0x0001_0437,
    0x0002_0804, 0x0002_1004, 0x0002_1404, 0x0003_0404, 0x0004_0404, 0x0004_0411, 0x0004_0c04, 0x0004_1404,
    0x0005_0804, 0x0005_1004,
];

#[cfg(windows)]
mod windows {
    use super::*;
    use windows_sys::Win32::Globalization::*;

    pub(super) fn locale_id_supported(id: u32) -> bool {
        // SAFETY: a null output and length zero request only the required buffer size.
        let length = unsafe { GetLocaleInfoW(id, LOCALE_SNAME, std::ptr::null_mut(), 0) };
        length > 0
    }

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
