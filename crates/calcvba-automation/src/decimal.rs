#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Decimal96 {
    pub lo: u32,
    pub mid: u32,
    pub hi: u32,
    pub scale_sign: u16,
}

impl Decimal96 {
    pub const SIGN_MASK: u16 = 0x8000;

    pub const fn from_parts(lo: u32, mid: u32, hi: u32, scale: u8, negative: bool) -> Self {
        let sign = if negative { Self::SIGN_MASK } else { 0 };
        Self {
            lo,
            mid,
            hi,
            scale_sign: sign | (scale as u16),
        }
    }

    pub const fn from_scale_sign(lo: u32, mid: u32, hi: u32, scale_sign: u16) -> Self {
        Self {
            lo,
            mid,
            hi,
            scale_sign,
        }
    }

    pub const fn scale(self) -> u8 {
        (self.scale_sign & 0x00FF) as u8
    }

    pub const fn is_negative(self) -> bool {
        (self.scale_sign & Self::SIGN_MASK) != 0
    }

    pub const fn magnitude_u128(self) -> u128 {
        (self.lo as u128) | ((self.mid as u128) << 32) | ((self.hi as u128) << 64)
    }
}

impl core::fmt::Display for Decimal96 {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        let scale = u32::from(self.scale());
        if scale > 28 {
            return write!(
                f,
                "<decimal:lo={};mid={};hi={};scale_sign={}>",
                self.lo, self.mid, self.hi, self.scale_sign
            );
        }
        let magnitude = self.magnitude_u128();
        if magnitude == 0 {
            return write!(f, "0");
        }
        let divisor = 10u128.pow(scale);
        let whole = magnitude / divisor;
        let fractional = magnitude % divisor;
        if self.is_negative() {
            write!(f, "-")?;
        }
        if scale == 0 || fractional == 0 {
            return write!(f, "{whole}");
        }
        let mut fractional_text = format!("{fractional:0width$}", width = scale as usize);
        while fractional_text.ends_with('0') {
            fractional_text.pop();
        }
        write!(f, "{whole}.{fractional_text}")
    }
}

// ── Exact Decimal arithmetic (MS-OAUT VarDecAdd/VarDecSub/VarDecMul/VarDecDiv) ──
//
// A Decimal is a 96-bit unsigned magnitude, a power-of-ten scale 0..=28 and a sign. Every
// operation computes its result exactly in a wide integer, then reduces the scale (rounding half
// to even, as Windows does) until the magnitude fits in 96 bits; a result that does not fit at
// scale 0 is Overflow.

/// Why a Decimal operation failed (VBA errors 6, 11 and 13).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecimalError {
    Overflow,
    DivideByZero,
    TypeMismatch,
}

impl DecimalError {
    /// The VBA run-time error number.
    pub const fn code(self) -> i32 {
        match self {
            Self::Overflow => 6,
            Self::DivideByZero => 11,
            Self::TypeMismatch => 13,
        }
    }
}

const MAX_SCALE: u32 = 28;
const LIMBS: usize = 12;

/// A 384-bit unsigned integer: wide enough for a 96-bit magnitude times 10^56.
#[derive(Clone, Copy, PartialEq, Eq)]
struct Wide([u32; LIMBS]);

impl Wide {
    const ZERO: Self = Self([0; LIMBS]);

    fn from_u128(v: u128) -> Self {
        let mut w = Self::ZERO;
        for (i, limb) in w.0.iter_mut().take(4).enumerate() {
            *limb = (v >> (32 * i)) as u32;
        }
        w
    }

    fn is_zero(&self) -> bool {
        self.0.iter().all(|&l| l == 0)
    }

    /// The value when it fits in 96 bits.
    fn to_u96(self) -> Option<u128> {
        if self.0[3..].iter().any(|&l| l != 0) {
            return None;
        }
        Some(u128::from(self.0[0]) | (u128::from(self.0[1]) << 32) | (u128::from(self.0[2]) << 64))
    }

    fn mul_small(mut self, m: u32) -> Self {
        let mut carry = 0u64;
        for limb in &mut self.0 {
            let p = u64::from(*limb) * u64::from(m) + carry;
            *limb = p as u32;
            carry = p >> 32;
        }
        debug_assert_eq!(carry, 0, "Wide overflow");
        self
    }

    fn mul_pow10(mut self, n: u32) -> Self {
        for _ in 0..n {
            self = self.mul_small(10);
        }
        self
    }

    fn pow10(n: u32) -> Self {
        Self::from_u128(1).mul_pow10(n)
    }

    fn mul(&self, other: &Self) -> Self {
        let mut out = [0u64; LIMBS * 2];
        for (i, &a) in self.0.iter().enumerate() {
            if a == 0 {
                continue;
            }
            let mut carry = 0u64;
            for (j, &b) in other.0.iter().enumerate() {
                let t = out[i + j] + u64::from(a) * u64::from(b) + carry;
                out[i + j] = t & 0xFFFF_FFFF;
                carry = t >> 32;
            }
            out[i + LIMBS] += carry;
        }
        let mut w = Self::ZERO;
        for (limb, &v) in w.0.iter_mut().zip(out.iter()) {
            *limb = v as u32;
        }
        w
    }

    fn add(&self, other: &Self) -> Self {
        let mut w = Self::ZERO;
        let mut carry = 0u64;
        for i in 0..LIMBS {
            let t = u64::from(self.0[i]) + u64::from(other.0[i]) + carry;
            w.0[i] = t as u32;
            carry = t >> 32;
        }
        w
    }

    /// `self - other`, for `self >= other`.
    fn sub(&self, other: &Self) -> Self {
        let mut w = Self::ZERO;
        let mut borrow = 0i64;
        for i in 0..LIMBS {
            let mut t = i64::from(self.0[i]) - i64::from(other.0[i]) - borrow;
            borrow = 0;
            if t < 0 {
                t += 1 << 32;
                borrow = 1;
            }
            w.0[i] = t as u32;
        }
        w
    }

    fn bit(&self, i: usize) -> bool {
        (self.0[i / 32] >> (i % 32)) & 1 == 1
    }

    fn shl1(mut self) -> Self {
        let mut carry = 0;
        for limb in &mut self.0 {
            let next = *limb >> 31;
            *limb = (*limb << 1) | carry;
            carry = next;
        }
        self
    }

    /// Long division by shift-and-subtract; `d` must be non-zero.
    fn divrem(&self, d: &Self) -> (Self, Self) {
        let mut q = Self::ZERO;
        let mut r = Self::ZERO;
        for i in (0..LIMBS * 32).rev() {
            r = r.shl1();
            if self.bit(i) {
                r.0[0] |= 1;
            }
            if r >= *d {
                r = r.sub(d);
                q.0[i / 32] |= 1 << (i % 32);
            }
        }
        (q, r)
    }

    /// `round_half_even(self / d)`.
    fn div_round(&self, d: &Self) -> Self {
        let (q, r) = self.divrem(d);
        let twice = r.shl1();
        let round_up = match twice.cmp(d) {
            core::cmp::Ordering::Greater => true,
            core::cmp::Ordering::Equal => q.0[0] & 1 == 1,
            core::cmp::Ordering::Less => false,
        };
        if round_up {
            q.add(&Self::from_u128(1))
        } else {
            q
        }
    }
}

impl PartialOrd for Wide {
    fn partial_cmp(&self, other: &Self) -> Option<core::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

impl Ord for Wide {
    fn cmp(&self, other: &Self) -> core::cmp::Ordering {
        self.0.iter().rev().cmp(other.0.iter().rev())
    }
}

/// Build a Decimal from an exact wide magnitude at `scale` (which may exceed 28), dropping
/// digits with round-half-even until the magnitude fits 96 bits and the scale is at most 28.
fn finish(mag: Wide, scale: u32, negative: bool) -> Result<Decimal96, DecimalError> {
    let first = scale.saturating_sub(MAX_SCALE);
    for drop in first..=scale {
        let reduced = if drop == 0 {
            mag
        } else {
            mag.div_round(&Wide::pow10(drop))
        };
        if let Some(m) = reduced.to_u96() {
            return Ok(Decimal96::from_u96(m, (scale - drop) as u8, negative));
        }
    }
    Err(DecimalError::Overflow)
}

impl Decimal96 {
    /// A Decimal from a magnitude below 2^96, a scale 0..=28 and a sign.
    pub const fn from_u96(magnitude: u128, scale: u8, negative: bool) -> Self {
        Self::from_parts(
            magnitude as u32,
            (magnitude >> 32) as u32,
            (magnitude >> 64) as u32,
            scale,
            negative,
        )
    }

    pub fn from_i64(value: i64) -> Self {
        Self::from_u96(u128::from(value.unsigned_abs()), 0, value < 0)
    }

    pub fn is_zero(self) -> bool {
        self.magnitude_u128() == 0
    }

    /// The value with its sign flipped (VarDecNeg).
    pub fn negated(self) -> Self {
        Self::from_scale_sign(
            self.lo,
            self.mid,
            self.hi,
            self.scale_sign ^ Self::SIGN_MASK,
        )
    }

    fn wide(self) -> Wide {
        Wide::from_u128(self.magnitude_u128())
    }

    fn signed_negative(self) -> bool {
        self.is_negative() && !self.is_zero()
    }

    /// Exact `self + other` (VarDecAdd).
    pub fn checked_add(self, other: Self) -> Result<Self, DecimalError> {
        let scale = u32::from(self.scale().max(other.scale()));
        let a = self.wide().mul_pow10(scale - u32::from(self.scale()));
        let b = other.wide().mul_pow10(scale - u32::from(other.scale()));
        let (an, bn) = (self.signed_negative(), other.signed_negative());
        if an == bn {
            return finish(a.add(&b), scale, an);
        }
        if a >= b {
            finish(a.sub(&b), scale, an && a != b)
        } else {
            finish(b.sub(&a), scale, bn)
        }
    }

    /// Exact `self - other` (VarDecSub).
    pub fn checked_sub(self, other: Self) -> Result<Self, DecimalError> {
        self.checked_add(other.negated())
    }

    /// Exact `self * other` (VarDecMul).
    pub fn checked_mul(self, other: Self) -> Result<Self, DecimalError> {
        let mag = self.wide().mul(&other.wide());
        let negative = (self.signed_negative() != other.signed_negative()) && !mag.is_zero();
        finish(
            mag,
            u32::from(self.scale()) + u32::from(other.scale()),
            negative,
        )
    }

    /// `self / other` to as many digits as fit (VarDecDiv): the largest result scale (at most
    /// 28) whose rounded quotient fits 96 bits, then trailing zeros removed.
    pub fn checked_div(self, other: Self) -> Result<Self, DecimalError> {
        if other.is_zero() {
            return Err(DecimalError::DivideByZero);
        }
        if self.is_zero() {
            return Ok(Self::default());
        }
        let negative = self.signed_negative() != other.signed_negative();
        let (sa, sb) = (u32::from(self.scale()), u32::from(other.scale()));
        let den = other.wide().mul_pow10(sa);
        for t in (0..=MAX_SCALE).rev() {
            let q = self.wide().mul_pow10(t + sb).div_round(&den);
            if let Some(m) = q.to_u96() {
                return Ok(Self::from_u96(m, t as u8, negative).trimmed());
            }
        }
        Err(DecimalError::Overflow)
    }

    /// The same value with trailing fractional zeros removed.
    pub fn trimmed(self) -> Self {
        let mut mag = self.magnitude_u128();
        let mut scale = self.scale();
        while scale > 0 && mag.is_multiple_of(10) {
            mag /= 10;
            scale -= 1;
        }
        Self::from_u96(mag, scale, self.is_negative())
    }

    /// The value times `10^digits`, rounded half to even to an integer (exact). `digits` 0
    /// gives the integer conversion VBA uses for `CLng(CDec(2.5))` (2); 4 gives Currency's
    /// scaled value.
    pub fn round_scaled_i128(self, digits: u32) -> Option<i128> {
        let scale = u32::from(self.scale());
        let mag = if digits >= scale {
            self.wide().mul_pow10(digits - scale)
        } else {
            self.wide().div_round(&Wide::pow10(scale - digits))
        };
        let m = mag.to_u96()? as i128;
        Some(if self.is_negative() { -m } else { m })
    }

    /// The nearest `f64`.
    pub fn to_f64(self) -> f64 {
        let value = self.magnitude_u128() as f64 / 10f64.powi(i32::from(self.scale()));
        if self.is_negative() { -value } else { value }
    }

    /// A Decimal from a Double, keeping its 15 significant digits (VarDecFromR8).
    pub fn from_f64(value: f64) -> Result<Self, DecimalError> {
        Self::from_float_digits(value, 15)
    }

    /// A Decimal from a Single, keeping its 7 significant digits (VarDecFromR4).
    pub fn from_f32(value: f32) -> Result<Self, DecimalError> {
        Self::from_float_digits(f64::from(value), 7)
    }

    fn from_float_digits(value: f64, digits: usize) -> Result<Self, DecimalError> {
        if !value.is_finite() {
            return Err(DecimalError::Overflow);
        }
        if value == 0.0 {
            return Ok(Self::default());
        }
        let text = format!("{:.*e}", digits - 1, value.abs());
        let (mantissa, exponent) = text.split_once('e').unwrap_or((&text, "0"));
        let exponent: i32 = exponent.parse().unwrap_or(0);
        let digits_text: String = mantissa.chars().filter(char::is_ascii_digit).collect();
        let m: u128 = digits_text.parse().unwrap_or(0);
        let shift = exponent - (digits as i32 - 1);
        let negative = value < 0.0;
        let out = if shift >= 0 {
            finish(Wide::from_u128(m).mul_pow10(shift as u32), 0, negative)?
        } else {
            finish(Wide::from_u128(m), (-shift) as u32, negative)?
        };
        Ok(out.trimmed())
    }

    /// Parse plain decimal text (`[+|-]digits[.digits][E[+|-]digits]`, surrounding blanks
    /// allowed) exactly, rounding digits beyond the 28th decimal place. `None` when the text
    /// is not of that form; `Some(Err(Overflow))` when it does not fit.
    pub fn parse_text(text: &str) -> Option<Result<Self, DecimalError>> {
        let t = text.trim();
        let (negative, rest) = match t.as_bytes().first()? {
            b'-' => (true, &t[1..]),
            b'+' => (false, &t[1..]),
            _ => (false, t),
        };
        let (mantissa, exponent) = match rest.find(['e', 'E']) {
            Some(i) => (&rest[..i], rest[i + 1..].parse::<i32>().ok()?),
            None => (rest, 0),
        };
        let (whole, frac) = mantissa.split_once('.').unwrap_or((mantissa, ""));
        let digit_count = whole.len() + frac.len();
        if digit_count == 0
            || digit_count > 60
            || !whole
                .bytes()
                .chain(frac.bytes())
                .all(|b| b.is_ascii_digit())
        {
            return None;
        }
        let mut mag = Wide::ZERO;
        for b in whole.bytes().chain(frac.bytes()) {
            mag = mag
                .mul_small(10)
                .add(&Wide::from_u128(u128::from(b - b'0')));
        }
        let scale = frac.len() as i64 - i64::from(exponent);
        Some(if scale >= 0 {
            if scale > 88 {
                // At most 60 digits: the value is below 10^-28 and rounds to zero.
                Ok(Self::default())
            } else {
                finish(mag, scale as u32, negative && !mag.is_zero())
            }
        } else if -scale > 40 {
            if mag.is_zero() {
                Ok(Self::default())
            } else {
                Err(DecimalError::Overflow)
            }
        } else {
            finish(
                mag.mul_pow10((-scale) as u32),
                0,
                negative && !mag.is_zero(),
            )
        })
    }
}

#[cfg(test)]
mod tests {
    use super::{Decimal96, DecimalError};

    fn dec(text: &str) -> Decimal96 {
        Decimal96::parse_text(text)
            .expect("decimal text")
            .expect("in range")
    }

    #[test]
    fn decimal_arithmetic_matches_excel() {
        // SpecOps decAdd, decMul, decDiv3, decDiv0, decOverflow, decSubPrec.
        assert_eq!(dec("1").checked_add(dec("1")).unwrap().to_string(), "2");
        assert_eq!(dec("0.1").checked_mul(dec("3")).unwrap().to_string(), "0.3");
        assert_eq!(
            dec("1").checked_div(dec("3")).unwrap().to_string(),
            "0.3333333333333333333333333333"
        );
        assert_eq!(
            dec("1").checked_div(dec("0")),
            Err(DecimalError::DivideByZero)
        );
        assert_eq!(
            dec("79228162514264337593543950335").checked_add(dec("1")),
            Err(DecimalError::Overflow)
        );
        assert_eq!(
            dec("1.0000000000000000000000000001")
                .checked_sub(dec("1"))
                .unwrap()
                .to_string(),
            "0.0000000000000000000000000001"
        );
    }

    #[test]
    fn decimal_arithmetic_rounds_half_to_even_when_reducing_scale() {
        let tenth = dec("0.1");
        assert_eq!(
            dec("0.0000000000000000000000000015")
                .checked_mul(tenth)
                .unwrap()
                .to_string(),
            "0.0000000000000000000000000002"
        );
        assert_eq!(
            dec("0.0000000000000000000000000025")
                .checked_mul(tenth)
                .unwrap()
                .to_string(),
            "0.0000000000000000000000000002"
        );
        assert_eq!(
            dec("79228162514264337593543950334")
                .checked_add(dec("0.5"))
                .unwrap()
                .to_string(),
            "79228162514264337593543950334"
        );
        assert_eq!(
            dec("79228162514264337593543950335").checked_add(dec("0.5")),
            Err(DecimalError::Overflow)
        );
        assert_eq!(
            dec("2").checked_div(dec("3")).unwrap().to_string(),
            "0.6666666666666666666666666667"
        );
        assert_eq!(
            dec("10").checked_div(dec("3")).unwrap().to_string(),
            "3.3333333333333333333333333333"
        );
        let six_halves = dec("6").checked_div(dec("2")).unwrap();
        assert_eq!(
            (six_halves.to_string().as_str(), six_halves.scale()),
            ("3", 0)
        );
    }

    #[test]
    fn decimal_signs_and_conversions() {
        let zero = dec("-1.5").checked_add(dec("1.5")).unwrap();
        assert!(zero.is_zero() && !zero.is_negative());
        assert_eq!(dec("-1.5").checked_mul(dec("-2")).unwrap().to_string(), "3");
        assert_eq!(dec("1.5").negated().to_string(), "-1.5");
        assert_eq!(dec("2.5").round_scaled_i128(0), Some(2));
        assert_eq!(dec("3.5").round_scaled_i128(0), Some(4));
        assert_eq!(dec("-2.5").round_scaled_i128(0), Some(-2));
        assert_eq!(
            dec("922337203685477.5807").round_scaled_i128(4),
            Some(9_223_372_036_854_775_807)
        );
        assert_eq!(Decimal96::from_f64(0.1).unwrap().to_string(), "0.1");
        assert_eq!(
            Decimal96::from_f64(1e20).unwrap().to_string(),
            "100000000000000000000"
        );
        assert_eq!(Decimal96::from_f64(1e30), Err(DecimalError::Overflow));
        assert_eq!(Decimal96::from_i64(-7).to_string(), "-7");
    }

    #[test]
    fn decimal96_preserves_exact_parts() {
        let value = Decimal96::from_parts(0x5566_7788, 0x1122_3344, 0x99AA_BBCC, 4, true);
        assert_eq!(value.lo, 0x5566_7788);
        assert_eq!(value.mid, 0x1122_3344);
        assert_eq!(value.hi, 0x99AA_BBCC);
        assert_eq!(value.scale(), 4);
        assert!(value.is_negative());
        assert_eq!(
            Decimal96::from_scale_sign(value.lo, value.mid, value.hi, value.scale_sign),
            value
        );
    }

    #[test]
    fn decimal96_display_trims_fractional_trailing_zeroes() {
        assert_eq!(
            Decimal96::from_parts(1_250_000, 0, 0, 4, false).to_string(),
            "125"
        );
        assert_eq!(
            Decimal96::from_parts(123_450, 0, 0, 3, false).to_string(),
            "123.45"
        );
        assert_eq!(
            Decimal96::from_parts(42_500, 0, 0, 4, true).to_string(),
            "-4.25"
        );
    }
}

#[cfg(test)]
mod proptests {
    use super::Decimal96;
    use proptest::prelude::*;

    proptest! {
        #[test]
        fn prop_decimal96_from_parts_accessor_roundtrip(
            lo: u32,
            mid: u32,
            hi: u32,
            scale in 0u8..=28u8,
            negative: bool,
        ) {
            let d = Decimal96::from_parts(lo, mid, hi, scale, negative);
            prop_assert_eq!(d.lo, lo);
            prop_assert_eq!(d.mid, mid);
            prop_assert_eq!(d.hi, hi);
            prop_assert_eq!(d.scale(), scale);
            prop_assert_eq!(d.is_negative(), negative);
        }
    }
}
