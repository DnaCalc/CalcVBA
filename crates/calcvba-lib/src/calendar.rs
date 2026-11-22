use calcvba_runtime::{days_from_civil, days_in_month, serial_to_ymd, ymd_to_serial};

#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum Calendar {
    Gregorian,
    Hijri,
}

impl Calendar {
    pub(crate) fn days(self, year: i64, month: i64, day: i64) -> i64 {
        if self == Self::Gregorian {
            return days_from_civil(year, month, day);
        }
        let month = i128::from(month) - 1;
        let year = i128::from(year) + month.div_euclid(12);
        let day = i128::from(day) - 1;
        (i128::from(days_from_civil(622, 7, 18))
            + hijri_year_start(year)
            + hijri_month_start(month.rem_euclid(12))
            + day)
            .clamp(i128::from(i64::MIN), i128::from(i64::MAX)) as i64
    }

    pub(crate) fn serial(self, year: i64, month: i64, day: i64) -> f64 {
        match self {
            Self::Gregorian => ymd_to_serial(year, month, day),
            Self::Hijri => {
                let epoch = days_from_civil(1899, 12, 30);
                (i128::from(self.days(year, month, day)) - i128::from(epoch)) as f64
            }
        }
    }

    pub(crate) fn parts(self, serial: f64) -> (i64, i64, i64) {
        let (year, month, day) = serial_to_ymd(serial);
        if self == Self::Gregorian {
            return (year, month, day);
        }
        let days =
            i128::from(days_from_civil(year, month, day)) - i128::from(days_from_civil(622, 7, 18));
        // Invert the 10,631 days of a 30-year Hijri cycle.
        let year = (30 * days + 10_646).div_euclid(10_631);
        let day_of_year = days - hijri_year_start(year);
        let month = (2 * day_of_year).div_euclid(59).min(11);
        let day = day_of_year - hijri_month_start(month) + 1;
        (year as i64, month as i64 + 1, day as i64)
    }

    pub(crate) fn month_days(self, year: i64, month: i64) -> i64 {
        match self {
            Self::Gregorian => days_in_month(year, month),
            Self::Hijri => self.days(year, month + 1, 1) - self.days(year, month, 1),
        }
    }
}

// Eleven years in each 30-year cycle add a day to the last month.
fn hijri_year_start(year: i128) -> i128 {
    354 * (year - 1) + (3 + 11 * year).div_euclid(30)
}

// Hijri months alternate between 30 and 29 days. The offset is zero-based.
fn hijri_month_start(month: i128) -> i128 {
    (59 * month + 1).div_euclid(2)
}
