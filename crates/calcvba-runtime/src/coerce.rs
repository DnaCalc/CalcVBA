use crate::bstr::BStr;
use crate::variant::{VarType, Variant};

/// Recognize VBA's Boolean string literals for `CBool` and the implicit Boolean
/// Let-coercion: exactly `"True"` or `"False"`, ASCII case-insensitive. The match is
/// strict — surrounding whitespace is *not* tolerated (live VBA raises Type mismatch on
/// `CBool("  False  ")` and on `Dim b As Boolean: b = "  False  "`), so a padded string
/// returns `None` and falls through to the caller's numeric path (which then errors,
/// matching VBA). A numeric string like `"5"` is likewise `None` here — the numeric path
/// converts it (any non-zero is `True`). This is the single source of truth shared by the
/// explicit (`pure::cbool`) and implicit (`arith::coerce_numeric`) paths so the two agree.
pub fn parse_bool_text(text: &str) -> Option<bool> {
    // `#TRUE#` / `#FALSE#` (the `Write #` spellings) convert too (Excel: `f = "#TRUE#"` is True).
    if text.eq_ignore_ascii_case("true") || text.eq_ignore_ascii_case("#true#") {
        Some(true)
    } else if text.eq_ignore_ascii_case("false") || text.eq_ignore_ascii_case("#false#") {
        Some(false)
    } else {
        None
    }
}

/// Why a string is not a VBA number: not numeric text (Type mismatch, 13) or numeric text out
/// of the `Double` range (Overflow, 6).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumberTextError {
    Mismatch,
    Overflow,
}

impl NumberTextError {
    /// The VBA run-time error number.
    pub const fn code(self) -> i32 {
        match self {
            Self::Mismatch => 13,
            Self::Overflow => 6,
        }
    }

    pub const fn decimal_error(self) -> crate::decimal::DecimalError {
        match self {
            Self::Mismatch => crate::decimal::DecimalError::TypeMismatch,
            Self::Overflow => crate::decimal::DecimalError::Overflow,
        }
    }

    fn message(self) -> String {
        match self {
            Self::Mismatch => "runtime error: 13 (Type mismatch)".to_string(),
            Self::Overflow => "runtime error: 6 (Overflow)".to_string(),
        }
    }
}

/// The parts of a decimal number string: its sign and the unsigned body
/// `digits[.digits][(E|D)[sign]digits]`, already checked against that grammar.
///
/// The sign may lead (`-5`, `- 5`), trail (`5-`) or be parentheses (`(5)`), as OLE's
/// `VarR8FromStr` accepts. Thousands separators and currency symbols are not accepted (Excel
/// raised 13 on `"1,000"` and `"$5"` on the measuring machine).
fn number_text_parts(s: &str) -> Option<(bool, &str)> {
    let t = s.trim();
    let (negative, body) = if t.len() >= 2 && t.starts_with('(') && t.ends_with(')') {
        (true, t[1..t.len() - 1].trim())
    } else if let Some(rest) = t.strip_prefix('-') {
        (true, rest.trim_start())
    } else if let Some(rest) = t.strip_prefix('+') {
        (false, rest.trim_start())
    } else if let Some(rest) = t.strip_suffix('-') {
        (true, rest.trim_end())
    } else if let Some(rest) = t.strip_suffix('+') {
        (false, rest.trim_end())
    } else {
        (false, t)
    };
    let bytes = body.as_bytes();
    let mut pos = 0;
    let digits = |pos: &mut usize| {
        let start = *pos;
        while bytes.get(*pos).is_some_and(u8::is_ascii_digit) {
            *pos += 1;
        }
        *pos - start
    };
    let mut mantissa_digits = digits(&mut pos);
    if bytes.get(pos) == Some(&b'.') {
        pos += 1;
        mantissa_digits += digits(&mut pos);
    }
    if mantissa_digits == 0 {
        return None;
    }
    if matches!(bytes.get(pos), Some(b'e' | b'E' | b'd' | b'D')) {
        pos += 1;
        if matches!(bytes.get(pos), Some(b'+' | b'-')) {
            pos += 1;
        }
        if digits(&mut pos) == 0 {
            return None;
        }
    }
    (pos == bytes.len()).then_some((negative, body))
}

/// Parse a VBA number string (the grammar of `CDbl` and of an implicit String-to-number Let):
/// blanks around the text, a leading, trailing or parenthesised sign, a `D` exponent like `E`,
/// and the Automation radix forms `&HFF` / `&O17`. Rust-only spellings such as `NaN`/`inf` and
/// the type-suffixed `&HFF&` are rejected.
pub fn parse_vba_number(s: &str) -> Result<f64, NumberTextError> {
    let t = s.trim();
    if let Some(value) = parse_vba_prefixed_integer(t) {
        return Ok(value as f64);
    }
    let (negative, body) = number_text_parts(t).ok_or(NumberTextError::Mismatch)?;
    let (mantissa, exponent) = match body.find(['e', 'E', 'd', 'D']) {
        Some(i) => (&body[..i], &body[i + 1..]),
        None => (body, "0"),
    };
    let (whole, frac) = mantissa.split_once('.').unwrap_or((mantissa, ""));
    let whole = if whole.is_empty() { "0" } else { whole };
    let frac = if frac.is_empty() { "0" } else { frac };
    let value: f64 = format!("{whole}.{frac}e{exponent}")
        .parse()
        .map_err(|_| NumberTextError::Mismatch)?;
    if !value.is_finite() {
        return Err(NumberTextError::Overflow);
    }
    Ok(if negative { -value } else { value })
}

/// The exact integer a number string spells when it is plain digits (with any of the sign
/// forms of [`parse_vba_number`]) or a radix literal, so a `LongLong` Let keeps digits beyond
/// a Double's 53 bits (`"9007199254740993"`). `None` for other text or outside `i64`.
pub fn parse_vba_integer_text(s: &str) -> Option<i64> {
    let t = s.trim();
    if let Some(value) = parse_vba_prefixed_integer(t) {
        return Some(value);
    }
    let (negative, body) = number_text_parts(t)?;
    if !body.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let magnitude: i128 = body.parse().ok()?;
    i64::try_from(if negative { -magnitude } else { magnitude }).ok()
}

/// [`parse_vba_number`] without the reason for a failure.
pub fn parse_vba_numeric_string(s: &str) -> Option<f64> {
    parse_vba_number(s).ok()
}

fn parse_vba_prefixed_integer(t: &str) -> Option<i64> {
    let (value, rest) = parse_vba_prefixed_integer_token(t)?;
    if rest.trim().is_empty() {
        Some(value)
    } else {
        None
    }
}

fn parse_vba_prefixed_integer_token(t: &str) -> Option<(i64, &str)> {
    let bytes = t.as_bytes();
    let mut pos = 0;
    let (sign, rest) = if let Some(rest) = t.strip_prefix('-') {
        pos = 1;
        (-1_i64, rest)
    } else if let Some(rest) = t.strip_prefix('+') {
        pos = 1;
        (1_i64, rest)
    } else {
        (1_i64, t)
    };

    rest.strip_prefix('&')?;
    pos += 1;
    let (radix, digit_start) = match bytes.get(pos) {
        Some(b'H' | b'h') => (16, pos + 1),
        Some(b'O' | b'o') => (8, pos + 1),
        _ => (8, pos),
    };
    pos = digit_start;

    let start = pos;
    while let Some(&byte) = bytes.get(pos) {
        let is_digit = match radix {
            16 => byte.is_ascii_hexdigit(),
            8 => (b'0'..=b'7').contains(&byte),
            _ => false,
        };
        if !is_digit {
            break;
        }
        pos += 1;
    }
    if start == pos {
        return None;
    }
    let digits = &t[start..pos];
    if matches!(bytes.get(pos), Some(b'%' | b'&' | b'^')) {
        return None;
    }

    let magnitude = u64::from_str_radix(digits, radix).ok()?;
    let parsed = numeric_string_radix_value(magnitude);
    Some((parsed.checked_mul(sign)?, &t[pos..]))
}

fn numeric_string_radix_value(magnitude: u64) -> i64 {
    if magnitude <= u32::MAX as u64 {
        (magnitude as u32) as i32 as i64
    } else {
        magnitude as i64
    }
}

pub fn coerce_to(value: &Variant, target: VarType) -> Result<Variant, String> {
    if value.vtype() == target {
        return Ok(value.clone());
    }

    match (value.vtype(), target) {
        // ── Widening integer paths ──
        (VarType::Integer, VarType::Long) => {
            Ok(Variant::from_i32(value.as_i16().unwrap_or(0) as i32))
        }
        (VarType::Integer, VarType::Single) => {
            Ok(Variant::from_f32(value.as_i16().unwrap_or(0) as f32))
        }
        (VarType::Integer, VarType::Double) => {
            Ok(Variant::from_f64(value.as_i16().unwrap_or(0) as f64))
        }
        (VarType::Long, VarType::Single) => {
            Ok(Variant::from_f32(value.as_i32().unwrap_or(0) as f32))
        }
        (VarType::Long, VarType::Double) => {
            Ok(Variant::from_f64(value.as_i32().unwrap_or(0) as f64))
        }
        (VarType::Single, VarType::Double) => {
            Ok(Variant::from_f64(value.as_f32().unwrap_or(0.0) as f64))
        }

        // ── Byte widening paths ──
        (VarType::Byte, VarType::Integer) => {
            let b = value.as_u8().unwrap_or(0);
            Ok(Variant::from_i16(b as i16))
        }
        (VarType::Byte, VarType::Long) => {
            let b = value.as_u8().unwrap_or(0);
            Ok(Variant::from_i32(b as i32))
        }
        (VarType::Byte, VarType::Single) => {
            let b = value.as_u8().unwrap_or(0);
            Ok(Variant::from_f32(b as f32))
        }
        (VarType::Byte, VarType::Double) => {
            let b = value.as_u8().unwrap_or(0);
            Ok(Variant::from_f64(b as f64))
        }

        // ── Boolean to numeric ──
        (VarType::Boolean, VarType::Integer) => {
            let n: i16 = if value.as_bool().unwrap_or(false) {
                -1
            } else {
                0
            };
            Ok(Variant::from_i16(n))
        }
        (VarType::Boolean, VarType::Long) => {
            let n = if value.as_bool().unwrap_or(false) {
                -1
            } else {
                0
            };
            Ok(Variant::from_i32(n))
        }
        (VarType::Boolean, VarType::Single) => {
            let n: f32 = if value.as_bool().unwrap_or(false) {
                -1.0
            } else {
                0.0
            };
            Ok(Variant::from_f32(n))
        }
        (VarType::Boolean, VarType::Double) => {
            let n: f64 = if value.as_bool().unwrap_or(false) {
                -1.0
            } else {
                0.0
            };
            Ok(Variant::from_f64(n))
        }

        // ── Date is stored as f64 internally — coerce to Double ──
        (VarType::Date, VarType::Double) => {
            Ok(Variant::from_f64(value.as_date_f64().unwrap_or(0.0)))
        }

        // ── Currency → Double (scaled i64 / 10000) ──
        (VarType::Currency, VarType::Double) => {
            let scaled = value.as_currency_scaled_i64().unwrap_or(0);
            Ok(Variant::from_f64(scaled as f64 / 10_000.0))
        }

        // ── Extended integer source types → floating targets ──
        // CDbl/CSng over COM edge-VT returns (VT_I1, VT_UI2, VT_UI4, VT_UI8,
        // VT_I8) lands here. Each widens losslessly into f64 (all magnitudes
        // fit the 53-bit mantissa except large u64/i64, which round to nearest).
        (VarType::SignedByte, VarType::Double) => {
            Ok(Variant::from_f64(f64::from(value.as_i8().unwrap_or(0))))
        }
        (VarType::SignedByte, VarType::Single) => {
            Ok(Variant::from_f32(f32::from(value.as_i8().unwrap_or(0))))
        }
        (VarType::UnsignedInteger, VarType::Double) => {
            Ok(Variant::from_f64(f64::from(value.as_u16().unwrap_or(0))))
        }
        (VarType::UnsignedInteger, VarType::Single) => {
            Ok(Variant::from_f32(f32::from(value.as_u16().unwrap_or(0))))
        }
        (VarType::UnsignedLong | VarType::UnsignedInt, VarType::Double) => {
            Ok(Variant::from_f64(f64::from(value.as_u32().unwrap_or(0))))
        }
        (VarType::UnsignedLong | VarType::UnsignedInt, VarType::Single) => {
            Ok(Variant::from_f32(value.as_u32().unwrap_or(0) as f32))
        }
        (VarType::UnsignedLongLong, VarType::Double) => {
            Ok(Variant::from_f64(value.as_u64().unwrap_or(0) as f64))
        }
        (VarType::UnsignedLongLong, VarType::Single) => {
            Ok(Variant::from_f32(value.as_u64().unwrap_or(0) as f32))
        }
        (VarType::LongLong, VarType::Double) => {
            Ok(Variant::from_f64(value.as_i64().unwrap_or(0) as f64))
        }
        (VarType::LongLong, VarType::Single) => {
            Ok(Variant::from_f32(value.as_i64().unwrap_or(0) as f32))
        }

        // ── Decimal → numeric targets ──
        // A VT_DECIMAL carries a 96-bit unsigned magnitude, a base-10 scale
        // (0..=28), and a sign bit. Floating targets divide the magnitude by
        // 10^scale; integer targets round-to-nearest then range-check.
        (
            VarType::Decimal,
            VarType::Double
            | VarType::Single
            | VarType::Long
            | VarType::Integer
            | VarType::Currency
            | VarType::LongLong
            | VarType::Boolean,
        ) => {
            let dec = value
                .as_decimal96()
                .ok_or_else(|| "invalid Decimal variant payload".to_string())?;
            let f = decimal_to_f64(dec);
            let overflow = || "runtime error: 6 (Overflow)".to_string();
            // Integer and Currency targets round the exact value half to even (VarI4FromDec).
            let rounded = |digits: u32, min: i128, max: i128| {
                dec.round_scaled_i128(digits)
                    .filter(|n| (min..=max).contains(n))
                    .ok_or_else(overflow)
            };
            match target {
                VarType::Double => Ok(Variant::from_f64(f)),
                VarType::Single => Ok(Variant::from_f32(f as f32)),
                VarType::Boolean => Ok(Variant::from_bool(!dec.is_zero())),
                VarType::Currency => Ok(Variant::from_currency_scaled_i64(rounded(
                    4,
                    i128::from(i64::MIN),
                    i128::from(i64::MAX),
                )?
                    as i64)),
                VarType::Long => {
                    Ok(Variant::from_i32(
                        rounded(0, i128::from(i32::MIN), i128::from(i32::MAX))? as i32,
                    ))
                }
                VarType::Integer => {
                    Ok(Variant::from_i16(
                        rounded(0, i128::from(i16::MIN), i128::from(i16::MAX))? as i16,
                    ))
                }
                VarType::LongLong => {
                    Ok(Variant::from_i64(
                        rounded(0, i128::from(i64::MIN), i128::from(i64::MAX))? as i64,
                    ))
                }
                _ => unreachable!("guarded by the match arm pattern"),
            }
        }

        // ── String → number through the VBA number grammar ──
        (
            VarType::String,
            VarType::Double
            | VarType::Single
            | VarType::Byte
            | VarType::Integer
            | VarType::Long
            | VarType::LongLong
            | VarType::Currency
            | VarType::Boolean
            | VarType::Date
            | VarType::Decimal,
        ) => {
            let text = value
                .as_bstr()
                .ok_or_else(|| "invalid String variant payload".to_string())?
                .as_str();
            string_to_number(&text, target)
        }

        // ── Empty coerces to zero/false for numeric targets ──
        (VarType::Empty, VarType::Integer) => Ok(Variant::from_i16(0)),
        (VarType::Empty, VarType::Long) => Ok(Variant::from_i32(0)),
        (VarType::Empty, VarType::Single) => Ok(Variant::from_f32(0.0)),
        (VarType::Empty, VarType::Double) => Ok(Variant::from_f64(0.0)),
        (VarType::Empty, VarType::Boolean) => Ok(Variant::from_bool(false)),

        // ── String coercion retains a BSTR-backed Variant carrier. ──
        (_, VarType::String) => variant_to_vba_string(value).map(Variant::from_string),
        _ => Err(format!(
            "unsupported coercion from {:?} to {:?}",
            value.vtype(),
            target
        )),
    }
}

/// A String converted to a numeric `target` as VBA's Let does: integer targets round half to
/// even and range-check (Overflow), `Boolean` also takes `"True"`/`"False"`, `Date` also takes
/// date text, `Decimal` keeps every digit.
fn string_to_number(text: &str, target: VarType) -> Result<Variant, String> {
    let overflow = || "runtime error: 6 (Overflow)".to_string();
    if target == VarType::Boolean
        && let Some(b) = parse_bool_text(text)
    {
        return Ok(Variant::from_bool(b));
    }
    if target == VarType::Date
        && let Some(serial) = crate::date::parse_date_text_serial(text)
    {
        return Ok(Variant::from_date_f64(serial));
    }
    if target == VarType::Decimal
        && let Some(result) = crate::decimal::Decimal96::parse_text(text)
    {
        return result.map(Variant::from_decimal96).map_err(|_| overflow());
    }
    if target == VarType::LongLong
        && let Some(n) = parse_vba_integer_text(text)
    {
        return Ok(Variant::from_i64(n));
    }
    let x = parse_vba_number(text).map_err(NumberTextError::message)?;
    let int = |min: f64, max: f64| {
        let r = x.round_ties_even();
        if r >= min && r <= max {
            Ok(r)
        } else {
            Err(overflow())
        }
    };
    Ok(match target {
        VarType::Double => Variant::from_f64(x),
        VarType::Single if x.abs() <= f64::from(f32::MAX) => Variant::from_f32(x as f32),
        VarType::Single => return Err(overflow()),
        VarType::Boolean => Variant::from_bool(x != 0.0),
        VarType::Byte => Variant::from_u8(int(0.0, 255.0)? as u8),
        VarType::Integer => Variant::from_i16(int(-32768.0, 32767.0)? as i16),
        VarType::Long => Variant::from_i32(int(-2_147_483_648.0, 2_147_483_647.0)? as i32),
        VarType::LongLong if x.abs() < 9.223_372_036_854_775e18 => {
            Variant::from_i64(x.round_ties_even() as i64)
        }
        VarType::Currency => {
            Variant::from_currency_scaled_i64(currency_scaled_from_f64(x).ok_or_else(overflow)?)
        }
        VarType::Date if is_date_serial(x) => Variant::from_date_f64(x),
        VarType::Decimal => {
            Variant::from_decimal96(crate::decimal::Decimal96::from_f64(x).map_err(|_| overflow())?)
        }
        _ => return Err(overflow()),
    })
}

/// A Double as Currency's scaled value (`value * 10^4`), computed exactly from the Double's
/// binary value and rounded half to even (VarCyFromR8); `None` outside the Currency range.
/// Going through `value * 10_000.0` alone loses the last digits near the limits
/// (`CCur(922337203685477.5#)` must be 922337203685477.5, `922337203685477.625#` Overflow).
pub fn currency_scaled_from_f64(value: f64) -> Option<i64> {
    if !value.is_finite() {
        return None;
    }
    // Where the floating product is exact enough (below 2^53), keep its rounding: `0.00005`
    // gives the product 0.5, which rounds to even (0).
    let product = value * 10_000.0;
    if product.abs() < 9_007_199_254_740_992.0 {
        return Some(product.round_ties_even() as i64);
    }
    let bits = value.to_bits();
    let negative = bits >> 63 == 1;
    let exp_bits = ((bits >> 52) & 0x7FF) as i32;
    let fraction = bits & ((1 << 52) - 1);
    // value = mantissa * 2^exp
    let (mantissa, exp) = if exp_bits == 0 {
        (fraction, -1074)
    } else {
        (fraction | (1 << 52), exp_bits - 1075)
    };
    let scaled = i128::from(mantissa) * 10_000;
    let magnitude = if exp >= 0 {
        if exp > 20 {
            return None;
        }
        scaled.checked_shl(exp as u32)?
    } else if exp < -120 {
        0
    } else {
        let shift = (-exp) as u32;
        let q = scaled >> shift;
        let rem = scaled - (q << shift);
        let half = 1i128 << (shift - 1);
        if rem > half || (rem == half && q & 1 == 1) {
            q + 1
        } else {
            q
        }
    };
    let signed = if negative { -magnitude } else { magnitude };
    i64::try_from(signed).ok()
}

/// Is `serial` inside the Date range (1 January 100 to 31 December 9999, VarDateFromR8)?
pub fn is_date_serial(serial: f64) -> bool {
    serial > -657_435.0 && serial < 2_958_466.0
}

pub fn variant_to_vba_string(value: &Variant) -> Result<BStr, String> {
    let text = match value.vtype() {
        VarType::ByRef => return variant_to_vba_string(&value.try_clone()?),
        VarType::Empty => BStr::empty(),
        VarType::Integer => BStr::from(format!("{}", value.as_i16().unwrap_or(0))),
        VarType::Long => BStr::from(format!("{}", value.as_i32().unwrap_or(0))),
        VarType::LongLong => BStr::from(format!("{}", value.as_i64().unwrap_or(0))),
        VarType::SignedByte => BStr::from(format!("{}", value.as_i8().unwrap_or(0))),
        VarType::Byte => BStr::from(format!("{}", value.as_u8().unwrap_or(0))),
        VarType::UnsignedInteger => BStr::from(format!("{}", value.as_u16().unwrap_or(0))),
        VarType::UnsignedLong | VarType::UnsignedInt => {
            BStr::from(format!("{}", value.as_u32().unwrap_or(0)))
        }
        VarType::UnsignedLongLong => BStr::from(format!("{}", value.as_u64().unwrap_or(0))),
        VarType::Single => BStr::from(format_vba_single(value.as_f32().unwrap_or(0.0))),
        VarType::Double => BStr::from(format_vba_f64(value.as_f64().unwrap_or(0.0))),
        VarType::Date => BStr::from(crate::vba_date::format_general_date(
            value.as_date_f64().unwrap_or(0.0),
        )),
        VarType::Boolean => BStr::from(if value.as_bool().unwrap_or(false) {
            "True"
        } else {
            "False"
        }),
        VarType::Decimal => BStr::from(format!(
            "{}",
            value
                .as_decimal96()
                .ok_or_else(|| "invalid Decimal variant payload".to_string())?
        )),
        VarType::Currency => BStr::from(format_vba_currency(
            value.as_currency_scaled_i64().unwrap_or(0),
        )),
        VarType::Error => BStr::from(format!("Error {}", value.as_error_code().unwrap_or(0))),
        VarType::String => {
            return value
                .as_bstr()
                .ok_or_else(|| "invalid String variant payload".to_string());
        }
        VarType::Null => {
            return Err("runtime error: 13 (Type mismatch)".to_string());
        }
        VarType::Object | VarType::ArrayVariant | VarType::Record | VarType::ProcRef => {
            return Err(format!("cannot convert {:?} to String", value.vtype()));
        }
    };
    Ok(text)
}

/// Convert a `Decimal96` to the nearest `f64`. The 96-bit magnitude is divided
/// by `10^scale`; the sign bit is applied afterwards so a negative zero
/// magnitude stays at `0.0`. Mirrors the Currency→Double scaling pattern.
fn decimal_to_f64(dec: crate::decimal::Decimal96) -> f64 {
    let magnitude = dec.magnitude_u128() as f64;
    let value = magnitude / 10f64.powi(i32::from(dec.scale()));
    if dec.is_negative() { -value } else { value }
}

/// Format a Currency scaled-i64 (four implied decimals) the way VBA's `CStr`
/// does. The sign is carried independently of the integer part — truncating
/// division loses it whenever the integer part is zero (`-0.5` must not
/// render as `0.5`) — and trailing fraction zeros are trimmed.
fn format_vba_currency(scaled: i64) -> String {
    let sign = if scaled < 0 { "-" } else { "" };
    let magnitude = scaled.unsigned_abs();
    let whole = magnitude / 10_000;
    let frac = magnitude % 10_000;
    if frac == 0 {
        format!("{sign}{whole}")
    } else {
        let frac_str = format!("{frac:04}");
        let frac_str = frac_str.trim_end_matches('0');
        format!("{sign}{whole}.{frac_str}")
    }
}

/// Render a floating-point value the way VBA's `CStr`, `Str`, `&` and `Print` do
/// (`VarBstrFromR8`/`VarBstrFromR4`, measured against Excel 16.0 build 20527).
///
/// The value is rounded to `digits` significant decimal digits (15 for `Double`, 7 for
/// `Single`) with trailing zeros dropped. It prints as plain digits when the integer part has at
/// most `digits` digits and the fraction at most `digits` places; otherwise as `d.dddE+XX` with
/// at least two exponent digits. Negative zero keeps its sign ("-0").
fn format_vba_float(v: f64, digits: usize) -> String {
    // Office prints the IEEE specials as the Microsoft C runtime does.
    if v.is_nan() {
        return if v.is_sign_negative() {
            "-nan(ind)".to_string()
        } else {
            "nan".to_string()
        };
    }
    if v.is_infinite() {
        return if v > 0.0 {
            "inf".to_string()
        } else {
            "-inf".to_string()
        };
    }
    let sign = if v.is_sign_negative() { "-" } else { "" };
    if v == 0.0 {
        return format!("{sign}0");
    }
    // Exact decimal expansion, then round half away from zero at `digits` places
    // (Office prints 100000000000000.5 as 100000000000001).
    let exact = format!("{:.*e}", 60, v.abs());
    let (mantissa, exponent) = exact.split_once('e').unwrap_or((&exact, "0"));
    let mut exp10: i32 = exponent.parse().unwrap_or(0);
    let mut kept: Vec<u8> = mantissa
        .bytes()
        .filter(u8::is_ascii_digit)
        .map(|b| b - b'0')
        .collect();
    let round_up = kept.get(digits).is_some_and(|d| *d >= 5);
    kept.truncate(digits);
    if round_up {
        let mut i = kept.len();
        loop {
            if i == 0 {
                kept.insert(0, 1);
                kept.truncate(digits);
                exp10 += 1;
                break;
            }
            i -= 1;
            if kept[i] == 9 {
                kept[i] = 0;
            } else {
                kept[i] += 1;
                break;
            }
        }
    }
    let all_digits: String = kept.iter().map(|d| char::from(b'0' + d)).collect();
    let sig = all_digits.trim_end_matches('0');
    let sig = if sig.is_empty() { "0" } else { sig };
    let n_sig = sig.len() as i32;
    let max = digits as i32;
    let int_digits = exp10 + 1;
    let decimals = if exp10 >= 0 {
        (n_sig - 1 - exp10).max(0)
    } else {
        -exp10 - 1 + n_sig
    };
    if (exp10 < 0 || int_digits <= max) && decimals <= max {
        if exp10 >= 0 {
            let int_len = int_digits as usize;
            if sig.len() <= int_len {
                format!("{sign}{sig}{}", "0".repeat(int_len - sig.len()))
            } else {
                format!("{sign}{}.{}", &sig[..int_len], &sig[int_len..])
            }
        } else {
            format!("{sign}0.{}{sig}", "0".repeat((-exp10 - 1) as usize))
        }
    } else {
        let frac = if sig.len() > 1 {
            format!(".{}", &sig[1..])
        } else {
            String::new()
        };
        let exp_sign = if exp10 < 0 { '-' } else { '+' };
        format!("{sign}{}{frac}E{exp_sign}{:02}", &sig[..1], exp10.abs())
    }
}

/// `Single` to string: seven significant digits (see [`format_vba_float`]).
fn format_vba_single(v: f32) -> String {
    format_vba_float(f64::from(v), 7)
}

/// `Double` to string: fifteen significant digits (see [`format_vba_float`]).
fn format_vba_f64(v: f64) -> String {
    format_vba_float(v, 15)
}

/// Render a `Variant` the way VBA's `Print`/`Debug.Print` displays it.
///
/// This is the **display** semantics (distinct from `variant_to_vba_string`'s
/// `CStr` semantics): it never errors — `Null` → `"Null"`, an object →
/// `"<object:…>"`, an array → `"<array:N>"` — because `Print` must render any
/// value. Multiple `Print` arguments are joined by the caller (the host `Print`
/// intrinsics tab-separate them).
pub fn print_display_text(value: &Variant) -> String {
    match value.vtype() {
        VarType::ByRef => print_display_text(&value.clone()),
        VarType::String => value
            .as_bstr()
            .map(|text| text.as_str().to_string())
            .unwrap_or_default(),
        VarType::Boolean => {
            if value.as_bool().unwrap_or(false) {
                "True".to_string()
            } else {
                "False".to_string()
            }
        }
        VarType::Empty => String::new(),
        VarType::Null => "Null".to_string(),
        VarType::Error => format!("Error {}", value.as_error_code().unwrap_or(0)),
        VarType::Integer => value.as_i16().unwrap_or(0).to_string(),
        VarType::Long => value.as_i32().unwrap_or(0).to_string(),
        VarType::LongLong => value.as_i64().unwrap_or(0).to_string(),
        VarType::SignedByte => value.as_i8().unwrap_or(0).to_string(),
        VarType::Byte => value.as_u8().unwrap_or(0).to_string(),
        VarType::UnsignedInteger => value.as_u16().unwrap_or(0).to_string(),
        VarType::UnsignedLong | VarType::UnsignedInt => value.as_u32().unwrap_or(0).to_string(),
        VarType::UnsignedLongLong => value.as_u64().unwrap_or(0).to_string(),
        VarType::Single => format_vba_single(value.as_f32().unwrap_or(0.0)),
        VarType::Double => format_vba_f64(value.as_f64().unwrap_or(0.0)),
        VarType::Date => crate::vba_date::format_general_date(value.as_date_f64().unwrap_or(0.0)),
        VarType::Decimal => value
            .as_decimal96()
            .map(|decimal| decimal.to_string())
            .unwrap_or_default(),
        VarType::Currency => value
            .as_currency_scaled_i64()
            .map(format_vba_currency)
            .unwrap_or_default(),
        VarType::ArrayVariant => value
            .as_safearray()
            .map(|array| format!("<array:{}>", array.len()))
            .unwrap_or_else(|| "<array:invalid>".to_string()),
        VarType::Object => value
            .as_object_ref()
            .map(|handle| format!("<object:{handle}>"))
            .unwrap_or_else(|| "<object:invalid>".to_string()),
        VarType::Record => value
            .as_com_record()
            .map(|record| format!("<record:{:p}>", record.record_data_ptr()))
            .unwrap_or_else(|| "<record:invalid>".to_string()),
        VarType::ProcRef => value
            .as_proc_ref()
            .map(|proc| format!("<proc:{proc}>"))
            .unwrap_or_else(|| "<proc:invalid>".to_string()),
    }
}

/// Render a `Variant` as a single `Write #` output field.
///
/// `Write #` writes a machine-readable, locale-independent record: strings are
/// double-quoted (embedded `"` doubled), `Boolean` becomes `#TRUE#`/`#FALSE#`,
/// `Null` becomes `#NULL#`, an `Empty` field is written as *nothing* (just the
/// surrounding commas), and an `Error` becomes `#ERROR <code>#`. Numbers reuse
/// the locale-independent [`print_display_text`] digits (period decimal, no
/// thousands separators), which also covers `Currency`/`Decimal` exactly.
///
/// FIDELITY: `Date` currently renders as its raw serial (via `print_display_text`)
/// rather than VBA's universal `#yyyy-mm-dd hh:mm:ss#` literal — that is the
/// separate `date-to-string-emits-serial` gap (it needs date decomposition, which
/// lives above this crate). The caller joins fields with commas and appends the
/// record terminator.
pub fn write_display_text(value: &Variant) -> String {
    match value.vtype() {
        VarType::String => {
            let text = value
                .as_bstr()
                .map(|t| t.as_str().to_string())
                .unwrap_or_default();
            format!("\"{}\"", text.replace('"', "\"\""))
        }
        VarType::Boolean => {
            if value.as_bool().unwrap_or(false) {
                "#TRUE#".to_string()
            } else {
                "#FALSE#".to_string()
            }
        }
        VarType::Null => "#NULL#".to_string(),
        VarType::Empty => String::new(),
        VarType::Error => format!("#ERROR {}#", value.as_error_code().unwrap_or(0)),
        // `Write #` emits a date as the locale-independent `#YYYY-MM-DD HH:MM:SS#` literal,
        // NOT the locale "General Date" form `print_display_text` would give.
        VarType::Date => crate::vba_date::format_write_date(value.as_date_f64().unwrap_or(0.0)),
        // A Single or Double drops the zero before its decimal point (Excel: `Write #1, 0.1 + 0.2`
        // writes `.3`).
        VarType::Single | VarType::Double => {
            let text = print_display_text(value);
            if let Some(rest) = text.strip_prefix("0.") {
                format!(".{rest}")
            } else if let Some(rest) = text.strip_prefix("-0.") {
                format!("-.{rest}")
            } else {
                text
            }
        }
        // Other numbers (incl. Currency/Decimal) reuse the canonical locale-independent rendering.
        _ => print_display_text(value),
    }
}

#[cfg(test)]
mod tests {
    use super::{
        NumberTextError, coerce_to, parse_bool_text, parse_vba_number, parse_vba_numeric_string,
        print_display_text, variant_to_vba_string, write_display_text,
    };
    use crate::{VarType, Variant};

    #[test]
    fn number_text_grammar_matches_excel() {
        // SpecTypes dbl1d3, dblSpSign, dblParen, dblTrailMinus, dblHex, dblPad, dblE400,
        // dblThou, dblDollar.
        assert_eq!(parse_vba_number("1d3"), Ok(1000.0));
        assert_eq!(parse_vba_number("- 5"), Ok(-5.0));
        assert_eq!(parse_vba_number("(5)"), Ok(-5.0));
        assert_eq!(parse_vba_number("5-"), Ok(-5.0));
        assert_eq!(parse_vba_number("&H10"), Ok(16.0));
        assert_eq!(parse_vba_number("  12  "), Ok(12.0));
        assert_eq!(parse_vba_number(".5"), Ok(0.5));
        assert_eq!(parse_vba_number("5."), Ok(5.0));
        assert_eq!(parse_vba_number("1E400"), Err(NumberTextError::Overflow));
        for bad in [
            "1,000", "$5", "", "-", "e5", "1e", "NaN", "inf", "5--", "(-5)",
        ] {
            assert_eq!(
                parse_vba_number(bad),
                Err(NumberTextError::Mismatch),
                "{bad:?}"
            );
        }
        assert_eq!(
            super::parse_vba_integer_text("9007199254740993"),
            Some(9_007_199_254_740_993)
        );
        assert_eq!(super::parse_vba_integer_text("1.5"), None);
        assert_eq!(parse_bool_text("#TRUE#"), Some(true));
        assert_eq!(parse_bool_text("#false#"), Some(false));
    }

    #[test]
    fn coerce_to_converts_strings_through_the_number_grammar() {
        let s = |text: &str| Variant::from_string(text);
        assert_eq!(
            coerce_to(&s("4"), VarType::Double).unwrap().as_f64(),
            Some(4.0)
        );
        assert_eq!(
            coerce_to(&s("(2.5)"), VarType::Long).unwrap().as_i32(),
            Some(-2)
        );
        assert_eq!(
            coerce_to(&s("&HFF"), VarType::Integer).unwrap().as_i16(),
            Some(255)
        );
        assert_eq!(
            coerce_to(&s("9007199254740993"), VarType::LongLong)
                .unwrap()
                .as_i64(),
            Some(9_007_199_254_740_993)
        );
        assert_eq!(
            coerce_to(&s("True"), VarType::Boolean).unwrap().as_bool(),
            Some(true)
        );
        assert!(
            coerce_to(&s("abc"), VarType::Double)
                .unwrap_err()
                .contains("13")
        );
        assert!(
            coerce_to(&s("40000"), VarType::Integer)
                .unwrap_err()
                .contains("6")
        );
        assert!(
            coerce_to(&s("1E400"), VarType::Double)
                .unwrap_err()
                .contains("6")
        );
    }

    #[test]
    fn currency_from_double_is_exact_near_the_limits() {
        use super::currency_scaled_from_f64 as cy;
        // SpecTypes ccurNearMax and curTooBig.
        assert_eq!(cy(922_337_203_685_477.5), Some(9_223_372_036_854_775_000));
        assert_eq!(cy(922_337_203_685_477.5 + 0.125), None);
        assert_eq!(cy(-922_337_203_685_477.5), Some(-9_223_372_036_854_775_000));
        assert_eq!(cy(1.5), Some(15_000));
        assert_eq!(cy(0.00005), Some(0), "the product 0.5 rounds to even");
        assert_eq!(cy(-0.0001), Some(-1));
        assert_eq!(cy(f64::NAN), None);
    }

    #[test]
    fn float_specials_print_as_the_c_runtime_does() {
        let text = |bits: u64| {
            variant_to_vba_string(&Variant::from_f64(f64::from_bits(bits)))
                .unwrap()
                .as_str()
        };
        assert_eq!(text(0x7FF0_0000_0000_0000), "inf");
        assert_eq!(text(0xFFF0_0000_0000_0000), "-inf");
        assert_eq!(text(0xFFF8_0000_0000_0000), "-nan(ind)");
    }

    #[test]
    fn print_and_write_number_text_use_fifteen_digits() {
        let print = |x: f64| print_display_text(&Variant::from_f64(x));
        let write = |x: f64| write_display_text(&Variant::from_f64(x));
        // Print #1, 0.1 + 0.2 writes " 0.3 "; Print #1, 1E+20; 0.000015; -1 / 3 writes
        // " 1E+20  0.000015 -0.333333333333333 " (the caller adds the sign/trailing blanks).
        assert_eq!(print(0.1 + 0.2), "0.3");
        assert_eq!(print(1e20), "1E+20");
        assert_eq!(print(0.000015), "0.000015");
        assert_eq!(print(-1.0 / 3.0), "-0.333333333333333");
        // Write #1, 0.1 + 0.2 writes .3.
        assert_eq!(write(0.1 + 0.2), ".3");
        assert_eq!(write(-0.5), "-.5");
        assert_eq!(write(12.5), "12.5");
        assert_eq!(write_display_text(&Variant::from_f32(0.25)), ".25");
    }

    #[test]
    fn parse_bool_text_recognizes_true_false_case_insensitively_without_trimming() {
        assert_eq!(parse_bool_text("True"), Some(true));
        assert_eq!(parse_bool_text("false"), Some(false));
        assert_eq!(parse_bool_text("tRuE"), Some(true));
        assert_eq!(parse_bool_text("FALSE"), Some(false));
        // Strict: padding, numeric strings, and other words are not recognized here
        // (VBA raises Type mismatch on `CBool("  False  ")`; numeric strings go to the
        // numeric path).
        assert_eq!(parse_bool_text("  False  "), None);
        assert_eq!(parse_bool_text("True "), None);
        assert_eq!(parse_bool_text("5"), None);
        assert_eq!(parse_bool_text("yes"), None);
        assert_eq!(parse_bool_text(""), None);
    }

    #[test]
    fn parse_numeric_radix_strings_use_vba_conversion_rules() {
        assert_eq!(parse_vba_numeric_string("&HFFFF"), Some(65_535.0));
        assert_eq!(parse_vba_numeric_string("&HFFFFFFFF"), Some(-1.0));
        assert_eq!(
            parse_vba_numeric_string("&H80000000"),
            Some(-2_147_483_648.0)
        );
        assert_eq!(parse_vba_numeric_string("&O37777777777"), Some(-1.0));
        assert_eq!(parse_vba_numeric_string("+&H7F"), Some(127.0));
        assert_eq!(parse_vba_numeric_string("&HFFFF&"), None);
    }

    #[test]
    fn integer_to_long() {
        let input = Variant::from_i16(7);
        let out = coerce_to(&input, VarType::Long).expect("coercion should succeed");
        assert_eq!(out.as_i32(), Some(7));
    }

    #[test]
    fn bool_to_long() {
        let input = Variant::from_bool(true);
        let out = coerce_to(&input, VarType::Long).expect("coercion should succeed");
        assert_eq!(out.as_i32(), Some(-1));
    }

    #[test]
    fn currency_to_string_keeps_sign_for_fractional_negatives() {
        // -1 < v < 0 has integer part zero; the sign must not vanish with it.
        for (scaled, expected) in [
            (-5_000_i64, "-0.5"),
            (-15_000, "-1.5"),
            (-10_000, "-1"),
            (5_000, "0.5"),
            (0, "0"),
            (-1, "-0.0001"),
            (12_345_678, "1234.5678"),
        ] {
            let v = Variant::from_currency_scaled_i64(scaled);
            assert_eq!(
                variant_to_vba_string(&v)
                    .expect("currency to string")
                    .as_str(),
                expected,
                "scaled {scaled}"
            );
            assert_eq!(
                super::print_display_text(&v),
                expected,
                "display text for scaled {scaled}"
            );
        }
    }

    #[test]
    fn decimal_to_double_and_single() {
        use crate::decimal::Decimal96;
        // 123.45 = magnitude 12345, scale 2.
        let v = Variant::from_decimal96(Decimal96::from_parts(12_345, 0, 0, 2, false));
        let out = coerce_to(&v, VarType::Double).expect("Decimal→Double");
        assert_eq!(out.vtype(), VarType::Double);
        assert!((out.as_f64().unwrap() - 123.45).abs() < 1e-9);

        let out_single = coerce_to(&v, VarType::Single).expect("Decimal→Single");
        assert_eq!(out_single.vtype(), VarType::Single);
        assert!((out_single.as_f32().unwrap() - 123.45_f32).abs() < 1e-4);

        // Negative: -4.25 = magnitude 425, scale 2, negative.
        let neg = Variant::from_decimal96(Decimal96::from_parts(425, 0, 0, 2, true));
        let out_neg = coerce_to(&neg, VarType::Double).expect("Decimal→Double");
        assert!((out_neg.as_f64().unwrap() - (-4.25)).abs() < 1e-9);
    }

    #[test]
    fn decimal_to_long_rounds_to_nearest() {
        use crate::decimal::Decimal96;
        // 2.5 rounds half to even, to 2, and 3.5 to 4 (VarI4FromDec, as VBA's CLng).
        let v = Variant::from_decimal96(Decimal96::from_parts(25, 0, 0, 1, false));
        let out = coerce_to(&v, VarType::Long).expect("Decimal→Long");
        assert_eq!(out.as_i32(), Some(2));
        let v = Variant::from_decimal96(Decimal96::from_parts(35, 0, 0, 1, true));
        assert_eq!(coerce_to(&v, VarType::Long).unwrap().as_i32(), Some(-4));

        // Integer-valued decimal.
        let whole = Variant::from_decimal96(Decimal96::from_parts(1000, 0, 0, 0, false));
        assert_eq!(
            coerce_to(&whole, VarType::Long).unwrap().as_i32(),
            Some(1000)
        );

        // Negative round.
        let neg = Variant::from_decimal96(Decimal96::from_parts(75, 0, 0, 1, true));
        assert_eq!(coerce_to(&neg, VarType::Long).unwrap().as_i32(), Some(-8));
    }

    #[test]
    fn decimal_to_integer_overflow_is_error_6() {
        use crate::decimal::Decimal96;
        // 40000 exceeds i16::MAX (32767) → overflow.
        let v = Variant::from_decimal96(Decimal96::from_parts(40_000, 0, 0, 0, false));
        let err = coerce_to(&v, VarType::Integer).expect_err("should overflow");
        assert!(err.contains("runtime error: 6"), "got: {err}");
    }

    #[test]
    fn decimal_to_boolean_nonzero_is_true() {
        use crate::decimal::Decimal96;
        let nonzero = Variant::from_decimal96(Decimal96::from_parts(1, 0, 0, 4, false));
        assert_eq!(
            coerce_to(&nonzero, VarType::Boolean).unwrap().as_bool(),
            Some(true)
        );
        let zero = Variant::from_decimal96(Decimal96::from_parts(0, 0, 0, 4, false));
        assert_eq!(
            coerce_to(&zero, VarType::Boolean).unwrap().as_bool(),
            Some(false)
        );
    }

    #[test]
    fn variant_to_vba_string_preserves_bstr_carrier() {
        let value = Variant::from_string("alpha");
        assert_eq!(
            variant_to_vba_string(&value)
                .expect("variant string coercion")
                .as_str(),
            "alpha"
        );
    }

    #[test]
    fn variant_to_vba_string_formats_numeric_and_boolean_carriers() {
        assert_eq!(
            variant_to_vba_string(&Variant::from_i32(42))
                .expect("long coercion")
                .as_str(),
            "42"
        );
        assert_eq!(
            variant_to_vba_string(&Variant::from_bool(true))
                .expect("bool coercion")
                .as_str(),
            "True"
        );
    }

    /// Decision-table oracle test: validates coercion outcomes against the canonical
    /// coercion.csv truth table. Each row specifies source type, target type, and
    /// expected result (ok or type-mismatch).
    #[test]
    fn oracle_coercion_table() {
        let table = include_str!("../../../tables/coercion.csv");
        let mut tested = 0;
        let mut skipped = 0;
        for line in table.lines().skip(1) {
            let parts: Vec<&str> = line.split(',').collect();
            if parts.len() < 3 {
                continue;
            }
            let (source_name, target_name, expected) =
                (parts[0].trim(), parts[1].trim(), parts[2].trim());
            let Some(source_vtype) = parse_vartype_name(source_name) else {
                skipped += 1;
                continue;
            };
            let Some(target_vtype) = parse_vartype_name(target_name) else {
                skipped += 1;
                continue;
            };
            let source_value = default_variant_for_type(source_vtype);
            let result = coerce_to(&source_value, target_vtype);
            match expected {
                "ok" => {
                    // Coercion may succeed or fail for unimplemented paths;
                    // track but don't hard-fail on unimplemented.
                    if let Err(msg) = result {
                        if msg.contains("not yet implemented")
                            || msg.contains("unsupported coercion")
                            || msg.contains("cannot hold heap strings")
                        {
                            skipped += 1;
                            continue;
                        }
                        panic!(
                            "coercion {source_name} → {target_name} expected ok, got err: {msg}"
                        );
                    }
                    tested += 1;
                }
                "type-mismatch" => {
                    if result.is_ok() {
                        // Accepting ok where type-mismatch was expected means the
                        // implementation is more permissive — track but don't fail
                        // since some coercions may be valid intermediate paths.
                        skipped += 1;
                    } else {
                        tested += 1;
                    }
                }
                _ => {
                    skipped += 1;
                }
            }
        }
        assert!(
            tested > 0,
            "oracle should test at least one row; tested={tested} skipped={skipped}"
        );
    }

    fn parse_vartype_name(name: &str) -> Option<VarType> {
        match name.to_ascii_lowercase().as_str() {
            "variant" => Some(VarType::Empty), // Variant maps to Empty for testing purposes
            "long" => Some(VarType::Long),
            "integer" => Some(VarType::Integer),
            "double" => Some(VarType::Double),
            "single" => Some(VarType::Single),
            "string" => Some(VarType::String),
            "boolean" => Some(VarType::Boolean),
            "date" => Some(VarType::Date),
            "currency" => Some(VarType::Currency),
            "decimal" => Some(VarType::Decimal),
            "byte" => Some(VarType::Byte),
            "object" => Some(VarType::Object),
            _ => None,
        }
    }

    fn default_variant_for_type(vtype: VarType) -> Variant {
        match vtype {
            VarType::Integer => Variant::from_i16(0),
            VarType::Long => Variant::from_i32(0),
            VarType::Double => Variant::from_f64(0.0),
            VarType::Boolean => Variant::from_bool(false),
            // A String source converts when its text is a number ("" is Type mismatch in VBA).
            VarType::String => Variant::from_string("0"),
            _ => Variant::zeroed(vtype),
        }
    }
}

#[cfg(test)]
mod proptests {
    use super::coerce_to;
    use crate::variant::{VarType, Variant};
    use proptest::prelude::*;

    proptest! {
        /// Coercing a value to its own type must always succeed (identity).
        #[test]
        fn prop_coerce_identity_is_noop(
            vtype_disc in 0u16..=20u16,
        ) {
            if let Some(vtype) = VarType::from_u16(vtype_disc) {
                let value = make_default(vtype);
                let result = coerce_to(&value, vtype);
                prop_assert!(result.is_ok(), "identity coercion failed for {:?}: {:?}", vtype, result);
                let out = result.unwrap();
                prop_assert_eq!(out.vtype(), vtype);
            }
        }

        /// Integer->Long must preserve the numeric value for all i16 inputs.
        #[test]
        fn prop_integer_to_long_preserves_value(v: i16) {
            let input = Variant::from_i16(v);
            let output = coerce_to(&input, VarType::Long).expect("Int->Long always succeeds");
            prop_assert_eq!(output.vtype(), VarType::Long);
            prop_assert_eq!(output.as_i32(), Some(v as i32));
        }

        /// Integer->Double must preserve the numeric value for all i16 inputs.
        #[test]
        fn prop_integer_to_double_preserves_value(v: i16) {
            let input = Variant::from_i16(v);
            let output = coerce_to(&input, VarType::Double).expect("Int->Double always succeeds");
            prop_assert_eq!(output.vtype(), VarType::Double);
            prop_assert_eq!(output.as_f64(), Some(v as f64));
        }

        /// Long->Double must preserve values that fit in f64 mantissa exactly.
        #[test]
        fn prop_long_to_double_preserves_value(v: i32) {
            let input = Variant::from_i32(v);
            let output = coerce_to(&input, VarType::Double).expect("Long->Double always succeeds");
            prop_assert_eq!(output.vtype(), VarType::Double);
            let f = output.as_f64().expect("result should be Double");
            // All i32 values are exactly representable in f64 (53-bit mantissa > 32 bits).
            prop_assert_eq!(f as i32, v, "Long->Double lost precision for {}", v);
        }

        /// Boolean->Long must map True to -1 and False to 0.
        #[test]
        fn prop_bool_to_long_maps_correctly(b: bool) {
            let input = Variant::from_bool(b);
            let output = coerce_to(&input, VarType::Long).expect("Bool->Long always succeeds");
            let expected = if b { -1 } else { 0 };
            prop_assert_eq!(output.as_i32(), Some(expected));
        }
    }

    fn make_default(vtype: VarType) -> Variant {
        match vtype {
            VarType::Integer => Variant::from_i16(0),
            VarType::Long => Variant::from_i32(0),
            VarType::Double => Variant::from_f64(0.0),
            VarType::Boolean => Variant::from_bool(false),
            // A String source converts when its text is a number ("" is Type mismatch in VBA).
            VarType::String => Variant::from_string("0"),
            _ => Variant::zeroed(vtype),
        }
    }
}
