// The trigonometric kernels and argument reducers are adapted from the SSE2
// path in AMD's win-libm, commit 6121d02aef0c775947d5dc655d7fc4ccfe2144ab.
// https://github.com/amd/win-libm/tree/6121d02aef0c775947d5dc655d7fc4ccfe2144ab
//
// MIT License
//
// Copyright (c) 2002-2019 Advanced Micro Devices, Inc.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this Software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.

// VBA follows the unfused evaluation order in these kernels. Regrouping the
// polynomials or combining multiply and add into one operation changes bits.

const SINE: [f64; 6] = [
    f64::from_bits(0xbfc5_5555_5555_5555),
    f64::from_bits(0x3f81_1111_1111_0bb3),
    f64::from_bits(0xbf2a_01a0_19e8_3e5c),
    f64::from_bits(0x3ec7_1de3_796c_de01),
    f64::from_bits(0xbe5a_e600_b42f_dfa7),
    f64::from_bits(0x3de5_e0b2_f9a4_3bb8),
];

const COSINE: [f64; 6] = [
    f64::from_bits(0x3fa5_5555_5555_5555),
    f64::from_bits(0xbf56_c16c_16c1_6967),
    f64::from_bits(0x3efa_01a0_19f4_ec91),
    f64::from_bits(0xbe92_7e4f_a17f_667b),
    f64::from_bits(0x3e21_eeb6_9038_2eec),
    f64::from_bits(0xbda9_07db_4725_8aa7),
];

const TWO_TO_NEG_13: f64 = f64::from_bits(0x3f20_0000_0000_0000);

pub fn sin(x: f64) -> f64 {
    if x == 0.0 {
        return x;
    }
    if !x.is_finite() {
        return f64::from_bits(x.to_bits() | 0x0008_0000_0000_0000);
    }
    // The direct interval includes the rounded pi/4: its sine gives the
    // rounding recorded for Tan(0.7853981633974483) in P_Tan.10.
    if x.abs() <= std::f64::consts::FRAC_PI_4 {
        return x + sine_correction(x);
    }
    let (r, tail, quadrant) = reduce(x.abs());
    let result = if quadrant & 1 == 0 {
        reduced_sine(r, tail)
    } else {
        reduced_cosine(r, tail)
    };
    if (quadrant >= 2) != x.is_sign_negative() {
        0.0 - result
    } else {
        result
    }
}

pub fn cos(x: f64) -> f64 {
    if !x.is_finite() {
        return f64::from_bits(x.to_bits() | 0x0008_0000_0000_0000);
    }
    let x = x.abs();
    if x < std::f64::consts::FRAC_PI_4 {
        return small_cosine(x);
    }
    let (r, tail, quadrant) = reduce(x);
    let result = if quadrant & 1 == 0 {
        reduced_cosine(r, tail)
    } else {
        reduced_sine(r, tail)
    };
    if quadrant == 1 || quadrant == 2 {
        0.0 - result
    } else {
        result
    }
}

pub fn tan(x: f64) -> f64 {
    sin(x) / cos(x)
}

fn sine_correction(x: f64) -> f64 {
    let [s1, s2, s3, s4, s5, s6] = SINE;
    let x2 = x * x;
    let x4 = x2 * x2;
    let x6 = x4 * x2;
    let low = s1 + x2 * (s2 + x2 * s3);
    let high = x6 * (s4 + x2 * (s5 + x2 * s6));
    (x * x2) * (low + high)
}

fn reduced_sine(x: f64, tail: f64) -> f64 {
    // Near a multiple of pi/2, VBA uses the leading remainder with a cubic
    // correction. Adding the low remainder changes the last bit in some cases.
    if x.abs() < TWO_TO_NEG_13 {
        return x - (x * (x * x)) * (1.0 / 6.0);
    }
    let half_square = 0.5 * (x * x);
    (tail + (sine_correction(x) - half_square * tail)) + x
}

fn reduced_cosine(x: f64, tail: f64) -> f64 {
    let [c1, c2, c3, c4, c5, c6] = COSINE;
    let x2 = x * x;
    let x4 = x2 * x2;
    let x6 = x4 * x2;
    let low = c1 + x2 * (c2 + x2 * c3);
    let high = x6 * (c4 + x2 * (c5 + x2 * c6));
    let half_square = 0.5 * x2;
    let neg_rounded_leading = half_square - 1.0;
    let rounding = (1.0 + neg_rounded_leading) - half_square;
    (x4 * (low + high) + (rounding - x * tail)) - neg_rounded_leading
}

fn small_cosine(x: f64) -> f64 {
    const TWO_TO_NEG_27: f64 = f64::from_bits(0x3e40_0000_0000_0000);
    if x < TWO_TO_NEG_27 {
        return 1.0;
    }
    let x2 = x * x;
    if x < TWO_TO_NEG_13 {
        return 1.0 - 0.5 * x2;
    }
    let [c1, c2, c3, c4, c5, c6] = COSINE;
    let x4 = x2 * x2;
    let x8 = x4 * x4;
    let correction = (x4 * (c1 + x2 * c2) + x8 * (c3 + x2 * c4)) + (x4 * x8) * (c5 + x2 * c6);
    let neg_half_square = -0.5 * x2;
    let leading = 1.0 + neg_half_square;
    let rounding = (1.0 - leading) + neg_half_square;
    (rounding + correction) + leading
}

fn reduce(x: f64) -> (f64, f64, u32) {
    if x >= 500_000.0 {
        return reduce_large(x);
    }
    const PI_2_LEAD: f64 = f64::from_bits(0x3ff9_21fb_5440_0000);
    const PI_2_TAIL: f64 = f64::from_bits(0x3dd0_b461_1a62_6331);
    const PI_2_TAIL_LEAD: f64 = f64::from_bits(0x3dd0_b461_1a60_0000);
    const PI_2_TAIL_TAIL: f64 = f64::from_bits(0x3ba3_198a_2e03_7073);
    let multiple = (x * std::f64::consts::FRAC_2_PI + 0.5) as u32;
    let n = f64::from(multiple);
    let mut lead = x - n * PI_2_LEAD;
    let mut tail = n * PI_2_TAIL;
    let mut r = lead - tail;
    // A near multiple of pi/2 cancels enough bits to need another split.
    let exponent = (x.to_bits() >> 52) as i32;
    let r_exponent = ((r.abs().to_bits()) >> 52) as i32;
    if exponent - r_exponent > 15 {
        let old_lead = lead;
        let tail_lead = n * PI_2_TAIL_LEAD;
        lead = old_lead - tail_lead;
        tail = n * PI_2_TAIL_TAIL - ((old_lead - lead) - tail_lead);
        r = lead - tail;
    }
    (r, (lead - r) - tail, multiple & 3)
}

// Little-endian fixed-point digits of 2/pi, with an upper guard word. The
// large reducer selects the 24-byte slice that brackets x * 2/pi's radix point.
const TWO_OVER_PI: [u8; 158] = [
    224, 241, 27, 193, 12, 88, 33, 116, 53, 126, 196, 126, 237, 175, 169, 75, 74, 41, 222, 231, 28,
    244, 236, 197, 151, 175, 31, 235, 158, 212, 181, 168, 127, 121, 154, 253, 24, 61, 221, 38, 44,
    159, 60, 251, 217, 180, 125, 180, 41, 104, 45, 70, 188, 188, 63, 96, 22, 120, 255, 95, 226,
    127, 236, 160, 228, 247, 46, 126, 17, 114, 210, 231, 76, 13, 230, 88, 71, 230, 4, 249, 125,
    209, 154, 192, 113, 166, 19, 18, 237, 186, 212, 215, 8, 162, 251, 156, 166, 196, 114, 172, 119,
    248, 115, 72, 70, 39, 168, 187, 36, 25, 128, 75, 55, 9, 233, 184, 145, 220, 134, 21, 239, 122,
    175, 142, 69, 249, 7, 65, 14, 241, 100, 86, 138, 109, 3, 119, 211, 212, 71, 95, 157, 240, 167,
    84, 16, 57, 185, 13, 230, 139, 2, 0, 0, 0, 0, 0, 0, 0,
];

fn reduce_large(x: f64) -> (f64, f64, u32) {
    const MANTISSA_MASK: u64 = (1 << 52) - 1;
    let bits = x.to_bits();
    let exponent = (bits >> 52) as u32 - 1023;
    let mantissa = u128::from((bits & MANTISSA_MASK) | (1 << 52));
    let offset = 134 - (exponent / 8) as usize;
    let word = |i| u64::from_le_bytes(TWO_OVER_PI[i..i + 8].try_into().unwrap());
    let lo_product = u128::from(word(offset)) * mantissa;
    let mid_product = u128::from(word(offset + 8)) * mantissa + (lo_product >> 64);
    let mut lo = lo_product as u64;
    let mut mid = mid_product as u64;
    let mut hi = word(offset + 16)
        .wrapping_mul(mantissa as u64)
        .wrapping_add((mid_product >> 64) as u64);
    let fractional_bits = 54 - exponent % 8;
    let half = (hi >> (fractional_bits - 1)) & 1;
    let quadrant = (((hi >> fractional_bits) + half) & 3) as u32;
    // When rounding up, the remainder is negative. Complementing the product
    // gives its magnitude to the accuracy needed by the floating-point splits.
    if half != 0 {
        hi = !hi;
        mid = !mid;
        lo = !lo;
    }
    hi &= (1_u64 << fractional_bits) - 1;

    // Scale the fixed-point remainder and split it into two truncated doubles.
    let mut binary_scale = -(fractional_bits as i32);
    if hi == 0 {
        (hi, mid, lo) = (mid, lo, 0);
        binary_scale -= 64;
    }
    if hi == 0 {
        (hi, mid, lo) = (mid, 0, 0);
        binary_scale -= 64;
    }
    if hi == 0 {
        return (0.0, 0.0, quadrant);
    }
    let highest = 63 - hi.leading_zeros();
    let (head_bits, tail_bits) = if highest < 52 {
        let shift = 52 - highest;
        (
            (hi << shift) | (mid >> (64 - shift)),
            (mid << shift) | (lo >> (64 - shift)),
        )
    } else if highest > 52 {
        let shift = highest - 52;
        (hi >> shift, (mid >> shift) | (hi << (64 - shift)))
    } else {
        (hi, mid)
    };
    let sign = half << 63;
    let biased_exponent = (binary_scale + highest as i32 + 1023) as u64;
    let head = f64::from_bits(sign | (biased_exponent << 52) | (head_bits & MANTISSA_MASK));
    let tail = if tail_bits == 0 {
        0.0
    } else {
        let shift = tail_bits.leading_zeros();
        let exponent = biased_exponent - u64::from(53 + shift);
        let significand = (tail_bits << shift) >> 11;
        f64::from_bits(sign | (exponent << 52) | (significand & MANTISSA_MASK))
    };

    const PI_2_PART1: f64 = f64::from_bits(0x3ff9_21fb_5000_0000);
    const PI_2_PART2: f64 = f64::from_bits(0x3e51_10b4_6000_0000);
    const PI_2_PART3: f64 = f64::from_bits(0x3c91_a626_3314_5c06);
    let head_high = f64::from_bits(head.to_bits() & !((1 << 27) - 1));
    let head_low = head - head_high;
    let product = std::f64::consts::FRAC_PI_2 * head;
    let product_tail = (((PI_2_PART1 * head_high - product) + PI_2_PART1 * head_low)
        + PI_2_PART2 * head_high)
        + PI_2_PART2 * head_low;
    let product_tail = product_tail + (PI_2_PART3 * head + std::f64::consts::FRAC_PI_2 * tail);
    let r = product + product_tail;
    (r, (product - r) + product_tail, quadrant)
}
