//! VBA `Format(expr, format)` engine.
//!
//! Covers the documented surface:
//! - **Named numeric:** General Number, Currency, Fixed, Standard, Percent,
//!   Scientific; **Boolean:** Yes/No, True/False, On/Off.
//! - **Named date:** General/Long/Medium/Short Date and Long/Medium/Short Time.
//! - **Custom numeric masks:** `0` `#` `.` `,` `%`, `E+`/`E-` exponential, and
//!   `;`-separated positive;negative;zero sections.
//! - **Custom date/time masks:** `y`/`m`/`d`/`h`/`n`/`s` tokens (1, 2, and named
//!   widths), `AM/PM`, with the VBA rule that `m`/`mm` after an hour token means
//!   minutes.
//!
//! FIDELITY: string-format characters (`@ & < >`), quoted literal runs in masks,
//! and locale-specific currency symbols / separators are first-cut. Numeric
//! masks and fixed formats round their decimal digits half up.

use calcvba_runtime::coerce::{coerce_to, parse_vba_numeric_string};
use calcvba_runtime::date::parse_date_text_serial_bits;
use calcvba_runtime::variant::VarType;
use calcvba_runtime::{Decimal96, Variant, variant_to_vba_string};

use crate::pure::{day_of_week, days_from_civil, serial_to_ymd, week_of_year};

const MONTHS: [&str; 12] = [
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
const MONTHS_ABBR: [&str; 12] = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];
const WEEKDAYS: [&str; 7] = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
];
const WEEKDAYS_ABBR: [&str; 7] = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

#[cfg(test)]
fn apply(value: &Variant, mask: &str) -> String {
    apply_with(value, mask, 1, 1)
}

/// `Format(value, mask, firstdayofweek, firstweekofyear)`: the two week arguments steer the `w`
/// and `ww` date tokens (`vbSunday`/`vbFirstJan1` by default).
pub fn apply_with(value: &Variant, mask: &str, first_day: i32, first_week: i32) -> String {
    let weeks = Weeks {
        first_day,
        first_week,
    };
    if value.vtype() == VarType::Empty {
        return String::new();
    }
    let trimmed = mask.trim();
    if trimmed.is_empty() {
        return general_string(value);
    }
    // A Null formats as nothing, or as the mask's fourth section.
    if value.vtype() == VarType::Null {
        return trimmed
            .split(';')
            .nth(3)
            .map(literal_text)
            .unwrap_or_default();
    }
    if let Some(out) = named(value, trimmed) {
        return out;
    }
    if has_string_mask_chars(trimmed) {
        return format_string(&general_string(value), trimmed);
    }
    if value.vtype() == VarType::String {
        // A number mask formats numeric text and returns other text unchanged.
        let numeric = value
            .as_bstr()
            .and_then(|b| parse_vba_numeric_string(&b.as_str()))
            .is_some();
        let date_text = is_date_mask(trimmed)
            && value
                .as_bstr()
                .and_then(|b| parse_date_text_serial_bits(&b.as_str()))
                .is_some();
        if !numeric && !date_text {
            return general_string(value);
        }
    }
    if is_date_mask(trimmed) {
        format_date(value, split_sections(trimmed)[0], weeks)
    } else {
        format_numeric(value, trimmed)
    }
}

/// [`apply_with`] as UTF-16 code units: where `Format` returns or rearranges the value's own
/// text (no mask, a string mask, text a number mask leaves alone) the text's units are kept,
/// unpaired surrogates included, as VBA keeps them.
pub fn apply_units(value: &Variant, mask: &str, first_day: i32, first_week: i32) -> Vec<u16> {
    let trimmed = mask.trim();
    let own_text = || variant_to_vba_string(value).map(|text| text.payload_units().to_vec());
    if !matches!(value.vtype(), VarType::Empty | VarType::Null) {
        if trimmed.is_empty() {
            if let Ok(units) = own_text() {
                return units;
            }
        } else if named(value, trimmed).is_none() && has_string_mask_chars(trimmed) {
            if let Ok(units) = own_text() {
                return format_string_units(&units, trimmed);
            }
        } else if value.vtype() == VarType::String
            && named(value, trimmed).is_none()
            && !is_date_or_number_text(value)
            && let Ok(units) = own_text()
        {
            return units;
        }
    }
    apply_with(value, mask, first_day, first_week)
        .encode_utf16()
        .collect()
}

/// [`format_string`] over code units: each character of the text (a surrogate pair, or one
/// unpaired surrogate) fills one placeholder.
fn format_string_units(units: &[u16], mask: &str) -> Vec<u16> {
    let text: Vec<u16> = format_string(&with_stand_ins(units), mask)
        .encode_utf16()
        .collect();
    restore_unpaired(&text)
}

/// Each unpaired surrogate stands in as a supplementary private-use character, so text that
/// holds one can go through `str`-based code; [`restore_unpaired`] puts them back.
const STAND_IN: u32 = 0xF_0000;

/// The units as text, an unpaired surrogate as its stand-in.
pub fn with_stand_ins(units: &[u16]) -> String {
    char::decode_utf16(units.iter().copied())
        .map(|item| match item {
            Ok(c) => c,
            Err(unpaired) => char::from_u32(STAND_IN + u32::from(unpaired.unpaired_surrogate()))
                .expect("supplementary private-use character"),
        })
        .collect()
}

/// The units with each stand-in back as its unpaired surrogate (units already unpaired are kept).
pub fn restore_unpaired(units: &[u16]) -> Vec<u16> {
    let mut out = Vec::with_capacity(units.len());
    for item in char::decode_utf16(units.iter().copied()) {
        match item {
            Ok(c) if (STAND_IN + 0xD800..STAND_IN + 0xE000).contains(&u32::from(c)) => {
                out.push((u32::from(c) - STAND_IN) as u16);
            }
            Ok(c) => {
                let mut buf = [0u16; 2];
                out.extend_from_slice(c.encode_utf16(&mut buf));
            }
            Err(unpaired) => out.push(unpaired.unpaired_surrogate()),
        }
    }
    out
}

/// The week-numbering settings of a `Format` call.
#[derive(Clone, Copy)]
struct Weeks {
    first_day: i32,
    first_week: i32,
}

impl Weeks {
    const DEFAULT: Weeks = Weeks {
        first_day: 1,
        first_week: 1,
    };
}

/// True when the mask has a string-format character outside quotes.
fn has_string_mask_chars(mask: &str) -> bool {
    let mut quoted = false;
    for c in mask.chars() {
        match c {
            '"' => quoted = !quoted,
            '@' | '&' | '<' | '>' if !quoted => return true,
            _ => {}
        }
    }
    false
}

/// A mask section's literal text: quotes dropped, a backslash taking the next character as is.
fn literal_text(section: &str) -> String {
    let mut out = String::new();
    let mut chars = section.chars();
    let mut quoted = false;
    while let Some(c) = chars.next() {
        match c {
            '"' => quoted = !quoted,
            '\\' if !quoted => out.extend(chars.next()),
            _ => out.push(c),
        }
    }
    out
}

/// `Format` of text with a string mask: `@` is a character or a space, `&` a character or
/// nothing. Padding goes before the source unless `!` requests left alignment; `>`
/// upper-cases and `<` lower-cases. Surplus source characters follow the mask.
/// The second section is for empty text.
fn format_string(text: &str, mask: &str) -> String {
    let sections: Vec<&str> = mask.split(';').collect();
    let section = if text.is_empty() && sections.len() >= 2 {
        sections[1]
    } else {
        sections[0]
    };
    let mut upper = false;
    let mut lower = false;
    let mut left = false;
    // Pattern items: Some(true) = `@`, Some(false) = `&`, None carries a literal.
    let mut items: Vec<(Option<bool>, char)> = Vec::new();
    let mut chars = section.chars();
    let mut quoted = false;
    while let Some(c) = chars.next() {
        match c {
            '"' => quoted = !quoted,
            _ if quoted => items.push((None, c)),
            '\\' => {
                if let Some(next) = chars.next() {
                    items.push((None, next));
                }
            }
            '@' => items.push((Some(true), ' ')),
            '&' => items.push((Some(false), ' ')),
            '!' => left = true,
            '>' => upper = true,
            '<' => lower = true,
            other => items.push((None, other)),
        }
    }
    let mut source = text.chars();
    let placeholders = items.iter().filter(|(kind, _)| kind.is_some()).count();
    let mut padding = if left {
        0
    } else {
        placeholders.saturating_sub(text.chars().count())
    };
    let mut out = String::new();
    for (kind, literal) in items {
        match kind {
            None => out.push(literal),
            Some(pad_space) => {
                if padding > 0 {
                    padding -= 1;
                    if pad_space {
                        out.push(' ');
                    }
                } else if let Some(c) = source.next() {
                    out.push(c);
                } else if pad_space {
                    out.push(' ');
                }
            }
        }
    }
    out.extend(source);
    if upper {
        out = out.to_uppercase();
    } else if lower {
        out = out.to_lowercase();
    }
    out
}

fn general_string(value: &Variant) -> String {
    variant_to_vba_string(value)
        .map(|b| b.as_str())
        .unwrap_or_default()
}

fn num(value: &Variant) -> f64 {
    // A String operand parses per VBA's lenient numeric rules ("123.5" -> 123.5)
    // before defaulting to 0 — Format's inputs are pervasively strings (cell text,
    // control `.Value`, DB fields), and coerce_to(String, Double) has no arm, so a
    // numeric string used to silently format as 0.
    if value.vtype() == VarType::String {
        return value
            .as_bstr()
            .and_then(|b| parse_vba_numeric_string(&b.as_str()))
            .unwrap_or(0.0);
    }
    coerce_to(value, VarType::Double)
        .ok()
        .and_then(|v| v.as_f64())
        .unwrap_or(0.0)
}

/// True when text parses as a number or a date.
pub fn is_date_or_number_text(value: &Variant) -> bool {
    value.as_bstr().is_some_and(|b| {
        let t = b.as_str();
        parse_vba_numeric_string(&t).is_some() || parse_date_text_serial_bits(&t).is_some()
    })
}

fn serial_of(value: &Variant) -> f64 {
    if let Some(serial) = value.as_date_f64() {
        return serial;
    }
    // A date-like String under a date mask (Format("2020-05-01", "yyyy")) parses
    // to a serial; a numeric string falls through to `num`.
    if value.vtype() == VarType::String
        && let Some(b) = value.as_bstr()
        && let Some(bits) = parse_date_text_serial_bits(&b.as_str())
    {
        return f64::from_bits(bits);
    }
    num(value)
}

// ── Named formats ─────────────────────────────────────────────────────────────

fn named(value: &Variant, mask: &str) -> Option<String> {
    match mask.to_ascii_lowercase().as_str() {
        "general number" => Some(general_number(num(value))),
        "currency" => Some(format_numeric(value, "$#,##0.00")),
        "fixed" => Some(format_numeric(value, "0.00")),
        "standard" => Some(format_numeric(value, "#,##0.00")),
        "percent" => Some(format_numeric(value, "0.00%")),
        "scientific" => Some(format_numeric(value, "0.00E+00")),
        "yes/no" => Some(bool_word(value, "Yes", "No")),
        "true/false" => Some(bool_word(value, "True", "False")),
        "on/off" => Some(bool_word(value, "On", "Off")),
        "general date" => Some(format_date(value, "general date", Weeks::DEFAULT)),
        "long date" => Some(format_date(value, "dddd, mmmm d, yyyy", Weeks::DEFAULT)),
        "medium date" => Some(format_date(value, "dd-mmm-yy", Weeks::DEFAULT)),
        "short date" => Some(format_date(value, "m/d/yyyy", Weeks::DEFAULT)),
        "long time" => Some(format_date(value, "h:nn:ss AM/PM", Weeks::DEFAULT)),
        "medium time" => Some(format_date(value, "h:nn AM/PM", Weeks::DEFAULT)),
        "short time" => Some(format_date(value, "hh:nn", Weeks::DEFAULT)),
        _ => None,
    }
}

fn bool_word(value: &Variant, yes: &str, no: &str) -> String {
    if num(value) != 0.0 { yes } else { no }.to_string()
}

/// Deterministic `FormatNumber` helper. VBA delegates omitted/default settings to
/// regional options; vm3 uses this module's invariant separators instead.
pub fn format_number(
    value: &Variant,
    decimals: usize,
    include_leading_digit: bool,
    parens_for_negative: bool,
    group_digits: bool,
) -> String {
    format_signed_fixed(
        num(value),
        decimals,
        include_leading_digit,
        parens_for_negative,
        group_digits,
        "",
        "",
    )
}

/// Deterministic `FormatCurrency` helper using the same `$`/`,`/`.` convention
/// already used by the named `Format(..., "Currency")` path.
pub fn format_currency(
    value: &Variant,
    decimals: usize,
    include_leading_digit: bool,
    parens_for_negative: bool,
    group_digits: bool,
) -> String {
    format_signed_fixed(
        num(value),
        decimals,
        include_leading_digit,
        parens_for_negative,
        group_digits,
        "$",
        "",
    )
}

/// Deterministic `FormatPercent` helper: multiply by 100, append `%`.
pub fn format_percent(
    value: &Variant,
    decimals: usize,
    include_leading_digit: bool,
    parens_for_negative: bool,
    group_digits: bool,
) -> String {
    format_signed_fixed(
        num(value) * 100.0,
        decimals,
        include_leading_digit,
        parens_for_negative,
        group_digits,
        "",
        "%",
    )
}

/// Deterministic `FormatDateTime` helper. `named_format` is one of the
/// `vbGeneralDate`/`vbLongDate`/`vbShortDate`/`vbLongTime`/`vbShortTime`
/// constants; regional settings are bounded to this module's existing masks.
pub fn format_date_time(value: &Variant, named_format: i32) -> Option<String> {
    Some(match named_format {
        0 => format_date(value, "general date", Weeks::DEFAULT),
        1 => format_date(value, "dddd, mmmm d, yyyy", Weeks::DEFAULT),
        2 => format_date(value, "m/d/yyyy", Weeks::DEFAULT),
        3 => format_date(value, "h:nn:ss AM/PM", Weeks::DEFAULT),
        4 => format_date(value, "hh:nn", Weeks::DEFAULT),
        _ => return None,
    })
}

fn format_signed_fixed(
    value: f64,
    decimals: usize,
    include_leading_digit: bool,
    parens_for_negative: bool,
    group_digits: bool,
    prefix: &str,
    suffix: &str,
) -> String {
    let negative = value < 0.0;
    let number = format_fixed_abs(value.abs(), decimals, include_leading_digit, group_digits);
    let body = format!("{prefix}{number}{suffix}");
    if negative {
        if parens_for_negative {
            format!("({body})")
        } else {
            format!("-{body}")
        }
    } else {
        body
    }
}

fn format_fixed_abs(
    value: f64,
    decimals: usize,
    include_leading_digit: bool,
    group_digits: bool,
) -> String {
    if !value.is_finite() {
        return value.to_string();
    }
    let (integer, fraction) = Dec::from_f64(value).to_fixed(decimals);
    let mut int_s = if integer.is_empty() {
        "0".to_string()
    } else {
        integer.into_iter().map(digit_char).collect()
    };
    if !include_leading_digit && decimals > 0 && int_s == "0" {
        int_s.clear();
    } else if group_digits {
        int_s = group_thousands(&int_s);
    }
    if decimals > 0 {
        int_s.push('.');
        int_s.extend(fraction.into_iter().map(digit_char));
    }
    int_s
}

/// VBA "General Number": up to 15 significant digits, no grouping, trailing zeros trimmed,
/// `E+nn` form for very large or small magnitudes.
fn general_number(n: f64) -> String {
    if n == 0.0 || !n.is_finite() {
        return "0".to_string();
    }
    let dec = Dec::from_f64(n);
    let mut digits = dec.digits.clone();
    while digits.len() > 1 && digits.last() == Some(&0) {
        digits.pop();
    }
    let e = dec.exp10;
    let sign = if n < 0.0 { "-" } else { "" };
    let text = |r: std::ops::Range<usize>| -> String {
        digits[r].iter().map(|d| digit_char(*d)).collect()
    };
    if !(-6..15).contains(&e) {
        let mut out = format!("{sign}{}", text(0..1));
        if digits.len() > 1 {
            out.push('.');
            out.push_str(&text(1..digits.len()));
        }
        let esign = if e < 0 { '-' } else { '+' };
        return format!("{out}E{esign}{:02}", e.abs());
    }
    if e >= 0 {
        let int_len = e as usize + 1;
        let mut all = digits.clone();
        if all.len() < int_len {
            all.resize(int_len, 0);
        }
        let int: String = all[..int_len].iter().map(|d| digit_char(*d)).collect();
        if all.len() > int_len {
            let frac: String = all[int_len..].iter().map(|d| digit_char(*d)).collect();
            format!("{sign}{int}.{frac}")
        } else {
            format!("{sign}{int}")
        }
    } else {
        let zeros = "0".repeat((-e - 1) as usize);
        format!("{sign}0.{zeros}{}", text(0..digits.len()))
    }
}

// ── Numeric masks ───────────────────────────────────────────────────────────

/// True when the mask holds a date/time token outside quotes, escapes and `[...]`.
fn is_date_mask(mask: &str) -> bool {
    let chars: Vec<char> = mask.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        match chars[i] {
            '"' => {
                i += 1;
                while i < chars.len() && chars[i] != '"' {
                    i += 1;
                }
            }
            '\\' => i += 1,
            '[' => {
                while i < chars.len() && chars[i] != ']' {
                    i += 1;
                }
            }
            c if matches!(
                c.to_ascii_lowercase(),
                'y' | 'm' | 'd' | 'h' | 'n' | 's' | 'w' | 'q'
            ) =>
            {
                return true;
            }
            _ => {
                let rest = &chars[i..];
                if starts_ci(rest, "a/p") || starts_ci(rest, "am/pm") || starts_ci(rest, "ampm") {
                    return true;
                }
            }
        }
        i += 1;
    }
    false
}

/// Splits a mask at `;` outside quotes and escapes.
fn split_sections(mask: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut start = 0;
    let mut quoted = false;
    let mut escaped = false;
    for (i, c) in mask.char_indices() {
        if escaped {
            escaped = false;
            continue;
        }
        match c {
            '\\' if !quoted => escaped = true,
            '"' => quoted = !quoted,
            ';' if !quoted => {
                out.push(&mask[start..i]);
                start = i + 1;
            }
            _ => {}
        }
    }
    out.push(&mask[start..]);
    out
}

/// A number as up to 15 significant decimal digits (what Excel works from), `d.ddd × 10^exp10`.
struct Dec {
    digits: Vec<u8>,
    exp10: i32,
}

impl Dec {
    fn from_decimal(value: Decimal96) -> Self {
        let magnitude = value.magnitude_u128();
        let digits: Vec<_> = magnitude
            .to_string()
            .bytes()
            .map(|digit| digit - b'0')
            .collect();
        let exp10 = if magnitude == 0 {
            0
        } else {
            digits.len() as i32 - 1 - i32::from(value.scale())
        };
        Self { digits, exp10 }
    }

    fn from_f64(v: f64) -> Dec {
        if v == 0.0 || !v.is_finite() {
            return Dec {
                digits: vec![0],
                exp10: 0,
            };
        }
        let s = format!("{:.14e}", v.abs());
        let (mant, exp) = s.split_once('e').unwrap_or((&s, "0"));
        Dec {
            digits: mant
                .bytes()
                .filter(u8::is_ascii_digit)
                .map(|b| b - b'0')
                .collect(),
            exp10: exp.parse().unwrap_or(0),
        }
    }

    /// Integer digits (no leading zeros) and exactly `frac` fraction digits, rounded half up.
    fn to_fixed(&self, frac: usize) -> (Vec<u8>, Vec<u8>) {
        let mut all = self.digits.clone();
        let mut ipos = self.exp10 + 1;
        if ipos <= 0 {
            let mut z = vec![0u8; (-ipos) as usize];
            z.extend(all);
            all = z;
            ipos = 0;
        } else if ipos as usize > all.len() {
            all.resize(ipos as usize, 0);
        }
        let mut ipos = ipos as usize;
        let keep = ipos + frac;
        if all.len() > keep {
            let up = all[keep] >= 5;
            all.truncate(keep);
            if up {
                let mut i = keep;
                loop {
                    if i == 0 {
                        all.insert(0, 1);
                        ipos += 1;
                        break;
                    }
                    i -= 1;
                    if all[i] == 9 {
                        all[i] = 0;
                    } else {
                        all[i] += 1;
                        break;
                    }
                }
            }
        } else {
            all.resize(keep, 0);
        }
        let frac_digits = all[ipos..].to_vec();
        let int_digits: Vec<u8> = all[..ipos]
            .iter()
            .copied()
            .skip_while(|d| *d == 0)
            .collect();
        (int_digits, frac_digits)
    }
}

#[derive(Clone, Debug)]
enum Tok {
    Digit(char),
    Dot,
    Comma,
    Percent,
    Exp { ch: char, plus: bool, digits: usize },
    Lit(String),
}

fn tokenize(section: &str) -> Vec<Tok> {
    let chars: Vec<char> = section.chars().collect();
    let mut toks = Vec::new();
    let mut seen_digit = false;
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        match c {
            '"' => {
                let mut s = String::new();
                i += 1;
                while i < chars.len() && chars[i] != '"' {
                    s.push(chars[i]);
                    i += 1;
                }
                toks.push(Tok::Lit(s));
            }
            '\\' => {
                i += 1;
                if let Some(n) = chars.get(i) {
                    toks.push(Tok::Lit(n.to_string()));
                }
            }
            '0' | '#' => {
                seen_digit = true;
                toks.push(Tok::Digit(c));
            }
            '.' => toks.push(Tok::Dot),
            ',' => toks.push(Tok::Comma),
            '%' => toks.push(Tok::Percent),
            'e' | 'E' if seen_digit && matches!(chars.get(i + 1), Some('+') | Some('-')) => {
                let plus = chars[i + 1] == '+';
                let mut digits = 0;
                let mut j = i + 2;
                while j < chars.len() && matches!(chars[j], '0' | '#') {
                    digits += 1;
                    j += 1;
                }
                toks.push(Tok::Exp {
                    ch: c,
                    plus,
                    digits,
                });
                i = j - 1;
            }
            other => toks.push(Tok::Lit(other.to_string())),
        }
        i += 1;
    }
    toks
}

fn tok_text(tok: &Tok) -> String {
    match tok {
        Tok::Lit(s) => s.clone(),
        Tok::Percent => "%".to_string(),
        Tok::Dot => ".".to_string(),
        Tok::Comma => ",".to_string(),
        Tok::Digit(c) => c.to_string(),
        Tok::Exp { .. } => String::new(),
    }
}

struct Rendered {
    text: String,
    zero: bool,
}

fn format_numeric(value: &Variant, mask: &str) -> String {
    let (decimal, negative, zero) = if let Some(value) = value.as_decimal96()
        && value.scale() <= 28
    {
        let decimal = Dec::from_decimal(value);
        let zero = decimal.digits == [0];
        (decimal, value.is_negative() && !zero, zero)
    } else {
        let n = num(value);
        (Dec::from_f64(n), n < 0.0, n == 0.0)
    };
    let sections = split_sections(mask);
    let (section, force_minus) = if negative {
        if sections.len() >= 2 && !sections[1].is_empty() {
            (sections[1], false) // the negative section carries its own sign
        } else {
            (sections[0], true)
        }
    } else if zero && sections.len() >= 3 {
        (sections[2], false)
    } else {
        (sections[0], false)
    };
    let r = render_section(section, decimal);
    if force_minus && !r.zero {
        format!("-{}", r.text)
    } else {
        r.text
    }
}

fn digit_char(d: u8) -> char {
    (b'0' + d) as char
}

/// Lays the integer digits `d` over the placeholder tokens (right to left, the leftmost
/// placeholder taking any surplus) and the fraction digits `f` over the fraction placeholders.
fn layout_number(
    before: &[Tok],
    after: &[Tok],
    d: &[u8],
    f: &[u8],
    has_dot: bool,
    grouping: bool,
) -> String {
    let leftmost = before.iter().position(|t| matches!(t, Tok::Digit(_)));
    let mut rev: Vec<String> = Vec::new();
    let mut pos = 0usize;
    let push_digit = |rev: &mut Vec<String>, ch: char, pos: usize| {
        if grouping && pos > 0 && pos.is_multiple_of(3) {
            rev.push(GROUP_SEP.to_string());
        }
        rev.push(ch.to_string());
    };
    for (idx, tok) in before.iter().enumerate().rev() {
        match tok {
            Tok::Digit(c) => {
                if Some(idx) == leftmost {
                    if pos < d.len() {
                        while pos < d.len() {
                            push_digit(&mut rev, digit_char(d[d.len() - 1 - pos]), pos);
                            pos += 1;
                        }
                    } else {
                        if *c == '0' {
                            push_digit(&mut rev, '0', pos);
                        }
                        pos += 1;
                    }
                } else if pos < d.len() {
                    push_digit(&mut rev, digit_char(d[d.len() - 1 - pos]), pos);
                    pos += 1;
                } else {
                    if *c == '0' {
                        push_digit(&mut rev, '0', pos);
                    }
                    pos += 1;
                }
            }
            other => rev.push(tok_text(other)),
        }
    }
    let mut out: String = rev.into_iter().rev().collect();
    if leftmost.is_none() && !d.is_empty() {
        out = d.iter().map(|x| digit_char(*x)).collect::<String>() + &out;
    }
    if has_dot {
        out.push('.');
    }
    let slots: Vec<usize> = after
        .iter()
        .enumerate()
        .filter(|(_, t)| matches!(t, Tok::Digit(_)))
        .map(|(i, _)| i)
        .collect();
    let mut keep = slots.len();
    while keep > 0 {
        let i = slots[keep - 1];
        if matches!(after[i], Tok::Digit('#')) && f.get(keep - 1).copied().unwrap_or(0) == 0 {
            keep -= 1;
        } else {
            break;
        }
    }
    let mut k = 0;
    for tok in after {
        match tok {
            Tok::Digit(_) => {
                if k < keep {
                    out.push(digit_char(f.get(k).copied().unwrap_or(0)));
                }
                k += 1;
            }
            other => out.push_str(&tok_text(other)),
        }
    }
    out
}

fn main_bounds(toks: &[Tok]) -> (Option<usize>, usize, Option<usize>) {
    let exp_at = toks.iter().position(|t| matches!(t, Tok::Exp { .. }));
    let main_end = exp_at.unwrap_or(toks.len());
    let dot_at = toks[..main_end].iter().position(|t| matches!(t, Tok::Dot));
    (exp_at, main_end, dot_at)
}

fn render_section(section: &str, mut dec: Dec) -> Rendered {
    let mut toks = tokenize(section);
    let Some(first_digit) = toks.iter().position(|t| matches!(t, Tok::Digit(_))) else {
        return Rendered {
            text: literal_text(section),
            zero: true,
        };
    };
    let (_, main_end, dot_at) = main_bounds(&toks);
    let before_end = dot_at.unwrap_or(main_end);

    // Percent signs after the first placeholder scale by 100; commas become grouping (between
    // placeholders), scaling by 1000 (trailing the integer part) or nothing (after the point).
    let mut shift = 0i32;
    let mut grouping = false;
    for (i, t) in toks.iter().enumerate() {
        if matches!(t, Tok::Percent) && i > first_digit {
            shift += 2;
        }
        if matches!(t, Tok::Comma) && i < before_end && i > first_digit {
            if toks[i + 1..before_end]
                .iter()
                .any(|x| matches!(x, Tok::Digit(_)))
            {
                grouping = true;
            } else {
                shift -= 3;
            }
        }
    }
    toks.retain(|t| !matches!(t, Tok::Comma));
    let (exp_at, main_end, dot_at) = main_bounds(&toks);
    let before = &toks[..dot_at.unwrap_or(main_end)];
    let after: &[Tok] = match dot_at {
        Some(d) => &toks[d + 1..main_end],
        None => &[],
    };
    let frac_places = after.iter().filter(|t| matches!(t, Tok::Digit(_))).count();

    let is_zero_value = dec.digits.iter().all(|d| *d == 0);
    if !is_zero_value {
        dec.exp10 += shift;
    }

    if let Some(e) = exp_at {
        let Tok::Exp { ch, plus, digits } = toks[e].clone() else {
            unreachable!()
        };
        let int_places = before
            .iter()
            .filter(|t| matches!(t, Tok::Digit(_)))
            .count()
            .max(1);
        let m = Dec {
            digits: dec.digits.clone(),
            exp10: int_places as i32 - 1,
        };
        let mut exp = if is_zero_value {
            0
        } else {
            dec.exp10 - (int_places as i32 - 1)
        };
        let (mut d, f) = m.to_fixed(frac_places);
        if d.len() > int_places {
            d.truncate(int_places);
            exp += 1;
        }
        let mut text = layout_number(before, after, &d, &f, dot_at.is_some(), false);
        text.push(ch);
        if exp < 0 {
            text.push('-');
        } else if plus {
            text.push('+');
        }
        text.push_str(&format!("{:0>w$}", exp.abs(), w = digits.max(1)));
        for t in &toks[e + 1..] {
            text.push_str(&tok_text(t));
        }
        return Rendered {
            text,
            zero: is_zero_value,
        };
    }

    let (d, f) = dec.to_fixed(frac_places);
    let zero = d.is_empty() && f.iter().all(|x| *x == 0);
    let mut text = layout_number(before, after, &d, &f, dot_at.is_some(), grouping);
    for t in &toks[main_end..] {
        text.push_str(&tok_text(t));
    }
    Rendered { text, zero }
}

/// The separator between digit groups.
const GROUP_SEP: char = ',';

fn group_thousands(digits: &str) -> String {
    let n = digits.len();
    let mut out = String::with_capacity(n + n / 3);
    for (i, ch) in digits.chars().enumerate() {
        if i > 0 && (n - i).is_multiple_of(3) {
            out.push(GROUP_SEP);
        }
        out.push(ch);
    }
    out
}

// ── Date / time masks ─────────────────────────────────────────────────────────

struct DateParts {
    year: i64,
    month: i64,
    day: i64,
    hour: i64,
    minute: i64,
    second: i64,
    dow: usize, // 0 = Sunday … 6 = Saturday
}

fn decompose(value: &Variant) -> DateParts {
    let serial = serial_of(value);
    let (year, month, day) = serial_to_ymd(serial);
    let frac = serial - serial.floor();
    let total = ((frac * 86_400.0).round() as i64).rem_euclid(86_400);
    let dow = day_of_week(year as i32, month as u32, day as u32).rem_euclid(7) as usize;
    DateParts {
        year,
        month,
        day,
        hour: total / 3600,
        minute: (total % 3600) / 60,
        second: total % 60,
        dow,
    }
}

fn format_date(value: &Variant, mask: &str, weeks: Weeks) -> String {
    let p = decompose(value);
    // "General Date" routes here with a non-token mask; render date + present
    // time components like VBA's default.
    if mask.eq_ignore_ascii_case("general date") {
        let serial = serial_of(value);
        let has_time = serial.fract() != 0.0;
        let has_date = serial.trunc() != 0.0;
        return match (has_date, has_time) {
            (true, true) => format!(
                "{} {}",
                format_date(value, "m/d/yyyy", weeks),
                format_date(value, "h:nn:ss AM/PM", weeks)
            ),
            (false, true) => format_date(value, "h:nn:ss AM/PM", weeks),
            _ => format_date(value, "m/d/yyyy", weeks),
        };
    }

    let lower = mask.to_ascii_lowercase();
    let use_12h = lower.contains("am/pm") || lower.contains("ampm") || lower.contains("a/p");
    let chars: Vec<char> = mask.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    // VBA: `m`/`mm` mean minutes when they follow an hour token (across literal
    // separators), month otherwise.
    let mut after_hour = false;
    while i < chars.len() {
        let rest = &chars[i..];
        match chars[i] {
            '"' => {
                i += 1;
                while i < chars.len() && chars[i] != '"' {
                    out.push(chars[i]);
                    i += 1;
                }
                i += 1;
                continue;
            }
            '\\' => {
                if let Some(c) = chars.get(i + 1) {
                    out.push(*c);
                }
                i += 2;
                continue;
            }
            '[' => {
                while i < chars.len() && chars[i] != ']' {
                    i += 1;
                }
                i += 1;
                continue;
            }
            _ => {}
        }
        let lower_case = chars[i].is_lowercase();
        let (text, advance, hour_token, calendar_token) = if starts_ci(rest, "yyyy") {
            (format!("{:04}", p.year), 4, false, true)
        } else if starts_ci(rest, "yy") {
            (format!("{:02}", p.year.rem_euclid(100)), 2, false, true)
        } else if starts_ci(rest, "y") {
            let day_of_year =
                days_from_civil(p.year, p.month, p.day) - days_from_civil(p.year, 1, 1) + 1;
            (day_of_year.to_string(), 1, false, true)
        } else if starts_ci(rest, "ww") {
            let week = week_of_year(p.year, p.month, p.day, weeks.first_day, weeks.first_week);
            (week.to_string(), 2, false, true)
        } else if starts_ci(rest, "w") {
            let weekday = (p.dow as i32 - (weeks.first_day - 1)).rem_euclid(7) + 1;
            (weekday.to_string(), 1, false, true)
        } else if starts_ci(rest, "q") {
            (((p.month - 1) / 3 + 1).to_string(), 1, false, true)
        } else if starts_ci(rest, "mmmm") {
            (month_name(p.month, false), 4, false, true)
        } else if starts_ci(rest, "mmm") {
            (month_name(p.month, true), 3, false, true)
        } else if starts_ci(rest, "mm") {
            let v = if after_hour { p.minute } else { p.month };
            (format!("{v:02}"), 2, false, true)
        } else if starts_ci(rest, "m") {
            let v = if after_hour { p.minute } else { p.month };
            (v.to_string(), 1, false, true)
        } else if starts_ci(rest, "dddd") {
            (WEEKDAYS[p.dow].to_string(), 4, false, true)
        } else if starts_ci(rest, "ddd") {
            (WEEKDAYS_ABBR[p.dow].to_string(), 3, false, true)
        } else if starts_ci(rest, "dd") {
            (format!("{:02}", p.day), 2, false, true)
        } else if starts_ci(rest, "d") {
            (p.day.to_string(), 1, false, true)
        } else if starts_ci(rest, "hh") {
            (
                format!("{:02}", hour_display(p.hour, use_12h)),
                2,
                true,
                true,
            )
        } else if starts_ci(rest, "h") {
            (hour_display(p.hour, use_12h).to_string(), 1, true, true)
        } else if starts_ci(rest, "nn") {
            (format!("{:02}", p.minute), 2, false, true)
        } else if starts_ci(rest, "n") {
            (p.minute.to_string(), 1, false, true)
        } else if starts_ci(rest, "ss") {
            (format!("{:02}", p.second), 2, false, true)
        } else if starts_ci(rest, "s") {
            (p.second.to_string(), 1, false, true)
        } else if starts_ci(rest, "am/pm") {
            (
                if p.hour < 12 { "AM" } else { "PM" }.to_string(),
                5,
                false,
                true,
            )
        } else if starts_ci(rest, "ampm") {
            (
                if p.hour < 12 { "AM" } else { "PM" }.to_string(),
                4,
                false,
                true,
            )
        } else if starts_ci(rest, "a/p") {
            (
                if p.hour < 12 { "A" } else { "P" }.to_string(),
                3,
                false,
                true,
            )
        } else {
            (chars[i].to_string(), 1, false, false)
        };
        if lower_case && (text == "AM" || text == "PM" || text == "A" || text == "P") {
            out.push_str(&text.to_lowercase());
        } else {
            out.push_str(&text);
        }
        i += advance;
        if calendar_token {
            after_hour = hour_token;
        }
        // literal separators leave `after_hour` unchanged so "hh:mm" reads mm as
        // minutes.
    }
    out
}

fn month_name(month: i64, abbr: bool) -> String {
    let idx = (month.clamp(1, 12) - 1) as usize;
    if abbr { MONTHS_ABBR[idx] } else { MONTHS[idx] }.to_string()
}

fn hour_display(hour24: i64, use_12h: bool) -> i64 {
    if use_12h {
        let h = hour24 % 12;
        if h == 0 { 12 } else { h }
    } else {
        hour24
    }
}

fn starts_ci(rest: &[char], token: &str) -> bool {
    let t: Vec<char> = token.chars().collect();
    rest.len() >= t.len()
        && rest[..t.len()]
            .iter()
            .zip(&t)
            .all(|(a, b)| a.eq_ignore_ascii_case(b))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fmt(n: f64, mask: &str) -> String {
        apply(&Variant::from_f64(n), mask)
    }
    fn fmt_date(serial: f64, mask: &str) -> String {
        apply(&Variant::from_date_f64(serial), mask)
    }

    #[test]
    fn numeric_masks() {
        assert_eq!(fmt(1234.5, "#,##0.00"), "1,234.50");
        assert_eq!(fmt(-1234.5, "#,##0.00"), "-1,234.50");
        assert_eq!(fmt(1234.567, "0.00"), "1234.57");
        assert_eq!(fmt(5.0, "00000"), "00005");
        assert_eq!(fmt(0.5, "0%"), "50%");
        assert_eq!(fmt(0.5, "0.0%"), "50.0%");
        assert_eq!(fmt(12345.0, "0.00E+00"), "1.23E+04");
    }

    #[test]
    fn numeric_mask_preserves_decimal96_digits() {
        // Excel's Format3.37 capture uses nonbreaking spaces for grouping.
        // Use the current format profile's separator to check the same digits.
        let value = Variant::from_decimal96(calcvba_runtime::Decimal96::from_parts(
            u32::MAX,
            u32::MAX,
            u32::MAX,
            0,
            false,
        ));
        let expected =
            "79\u{a0}228\u{a0}162\u{a0}514\u{a0}264\u{a0}337\u{a0}593\u{a0}543\u{a0}950\u{a0}335"
                .replace('\u{a0}', &GROUP_SEP.to_string());
        assert_eq!(apply(&value, "#,##0"), expected);
    }

    #[test]
    fn named_formats() {
        assert_eq!(fmt(1234.5, "Currency"), "$1,234.50");
        assert_eq!(fmt(1234.5, "Standard"), "1,234.50");
        assert_eq!(fmt(0.5, "Percent"), "50.00%");
        assert_eq!(fmt(0.0, "Yes/No"), "No");
        assert_eq!(fmt(1.0, "True/False"), "True");
        assert_eq!(fmt(42.0, "General Number"), "42");
    }

    #[test]
    fn date_masks() {
        // Serial 0 = 1899-12-30 (the VBA Date epoch).
        assert_eq!(fmt_date(0.0, "yyyy-mm-dd"), "1899-12-30");
        // Half a day = noon.
        assert_eq!(fmt_date(0.5, "hh:nn:ss"), "12:00:00");
        assert_eq!(fmt_date(0.5, "h:nn AM/PM"), "12:00 PM");
        // `mm` after an hour token means minutes, across the `:` separator.
        assert_eq!(fmt_date(0.5, "hh:mm"), "12:00");
    }
}
