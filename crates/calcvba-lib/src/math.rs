// The logarithm and arctangent kernels are adapted from AMD win-libm,
// commit 6121d02aef0c775947d5dc655d7fc4ccfe2144ab.
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

/// Arctangent with the unfused evaluation order used by VBA.
pub fn atan(x: f64) -> f64 {
    use std::f64::consts::{FRAC_PI_2, FRAC_PI_4};

    if x.is_nan() {
        return f64::from_bits(x.to_bits() | 0x0008_0000_0000_0000);
    }
    let magnitude = x.abs();
    if magnitude > 72_057_594_037_927_936.0 {
        return FRAC_PI_2.copysign(x);
    }

    // Reduce to [-7/16, 7/16] and keep the offset split into head and tail.
    let (reduced, head, tail) = if magnitude > 2.4375 {
        (
            -1.0 / magnitude,
            FRAC_PI_2,
            f64::from_bits(0x3c91_a626_3314_5c06),
        )
    } else if magnitude > 1.1875 {
        (
            (magnitude - 1.5) / (1.0 + 1.5 * magnitude),
            f64::from_bits(0x3fef_730b_d281_f69b),
            f64::from_bits(0x3c70_0788_7af0_cbbc),
        )
    } else if magnitude > 0.6875 {
        (
            (magnitude - 1.0) / (1.0 + magnitude),
            FRAC_PI_4,
            f64::from_bits(0x3c81_a626_3314_5c06),
        )
    } else if magnitude > 0.4375 {
        (
            (2.0 * magnitude - 1.0) / (2.0 + magnitude),
            f64::from_bits(0x3fdd_ac67_0561_bb4f),
            f64::from_bits(0x3c7a_2b7f_222f_65e0),
        )
    } else {
        (magnitude, 0.0, 0.0)
    };

    // Remez rational approximation of the correction to the reduced argument.
    let s = reduced * reduced;
    let numerator = 0.268_297_920_532_545_9
        + s * (0.447_677_206_805_497_5
            + s * (0.220_638_780_716_667_43
                + s * (0.030_445_591_950_485_303 + s * 0.000_142_316_903_342_317_78)));
    let denominator = 0.804_893_761_597_637_7
        + s * (1.825_967_877_375_070_7
            + s * (1.412_542_599_319_589_3
                + s * (0.424_602_594_203_847_1 + s * 0.038_952_587_394_474_22)));
    let correction = reduced * s * numerator / denominator;
    let value = head - ((correction - tail) - reduced);
    if x.is_sign_negative() { -value } else { value }
}

/// Natural logarithm with the unfused evaluation order used by VBA.
pub fn ln(x: f64) -> f64 {
    if x.is_nan() {
        return f64::from_bits(x.to_bits() | 0x0008_0000_0000_0000);
    }
    if x == 0.0 {
        return f64::NEG_INFINITY;
    }
    if x < 0.0 {
        return f64::from_bits(0xfff8_0000_0000_0000);
    }
    if x.is_infinite() {
        return x;
    }

    let delta = x - 1.0;
    if delta.abs() < 0.0625 {
        return near_one(delta);
    }

    // Write x as 2^(exponent+1) * y, with 0.5 <= y < 1.0.
    const MANTISSA: u64 = (1 << 52) - 1;
    const HALF: u64 = 0.5_f64.to_bits();
    let (normalized, bias) = if x < f64::MIN_POSITIVE {
        (x * 4_503_599_627_370_496.0, 1075)
    } else {
        (x, 1023)
    };
    let bits = normalized.to_bits();
    let exponent = f64::from((bits >> 52) as i32 - bias);
    let mantissa = bits & MANTISSA;
    let y = f64::from_bits(HALF | mantissa);

    // Round y to the nearest 1/512, using the ninth fractional bit to break
    // ties upward. The corresponding value of log(2 * f) has a split table.
    let index = ((mantissa >> 44) + ((mantissa >> 43) & 1)) as usize;
    let f = f64::from_bits(HALF | ((index as u64) << 44));
    let r = (f - y) * RECIPROCAL[index];
    let r2 = r * r;
    let r4 = r2 * r2;
    let low = ((1.0 / 3.0) * r + 0.5) * r2 + r;
    let high = (((1.0 / 6.0) * r + 0.2) * r + 0.25) * r4;
    let poly = low + high;

    const LN2_LEAD: f64 = f64::from_bits(0x3fe6_2e42_e000_0000);
    const LN2_TAIL: f64 = f64::from_bits(0x3e6e_fa39_ef35_793c);
    let (lead, tail) = LOG_SPLIT[index];
    let lead = f64::from_bits(lead) + exponent * LN2_LEAD;
    let tail = f64::from_bits(tail) + (exponent * LN2_TAIL - poly);
    lead + tail
}

fn near_one(delta: f64) -> f64 {
    const CA1: f64 = f64::from_bits(0x3fb5_5555_5555_54e6);
    const CA2: f64 = f64::from_bits(0x3f89_9999_99ba_c6d4);
    const CA3: f64 = f64::from_bits(0x3f62_4923_07f1_519f);
    const CA4: f64 = f64::from_bits(0x3f3c_8034_c85d_fff0);
    let half = delta / (2.0 + delta);
    let correction = delta * half;
    let u = half + half;
    let u2 = u * u;
    let u3 = u2 * u;
    let u7 = (u3 * u3) * u;
    let low = (CA1 + CA2 * u2) * u3;
    let high = (CA3 + CA4 * u2) * u7;
    delta + ((low + high) - correction)
}

const RECIPROCAL: [f64; 257] = {
    let mut reciprocal = [0.0; 257];
    let mut index = 0;
    while index < reciprocal.len() {
        reciprocal[index] = 512.0 / (256 + index) as f64;
        index += 1;
    }
    reciprocal
};

// Head and tail bits for log(1 + index / 256), from
// amd/win-libm/log_256_lead_tail_table.asm.
const LOG_SPLIT: [(u64, u64); 257] = [
    (0x0000_0000_0000_0000, 0x0000_0000_0000_0000),
    (0x3f6f_f00a_a000_0000, 0x3db5_885e_0250_435a),
    (0x3f7f_e02a_6000_0000, 0x3de6_20cf_11f8_6ed2),
    (0x3f87_dc47_5000_0000, 0x3dff_0214_edba_4a25),
    (0x3f8f_c0a8_b000_0000, 0x3dbf_807c_79f3_db4e),
    (0x3f93_cea4_4000_0000, 0x3dea_352b_a779_a52b),
    (0x3f97_b91b_0000_0000, 0x3dff_56c4_6aa4_9fd5),
    (0x3f9b_9fc0_2000_0000, 0x3dfe_be46_5fef_5196),
    (0x3f9f_829b_0000_0000, 0x3e0c_f066_0099_f1f8),
    (0x3fa1_b0d9_8000_0000, 0x3e12_47b2_ff85_945d),
    (0x3fa3_9e87_b000_0000, 0x3e13_fd7a_bf52_02b6),
    (0x3fa5_8a5b_a000_0000, 0x3e1f_91c9_a918_d51e),
    (0x3fa7_7458_f000_0000, 0x3e08_cb73_f118_d3ca),
    (0x3fa9_5c83_0000_0000, 0x3e1d_91c7_d6fa_d074),
    (0x3fab_42dd_7000_0000, 0x3de1_971b_ec28_d14c),
    (0x3fad_276b_8000_0000, 0x3e15_b616_a423_c78a),
    (0x3faf_0a30_c000_0000, 0x3da1_62a6_617c_c971),
    (0x3fb0_7598_3000_0000, 0x3e16_6391_c4c0_6d29),
    (0x3fb1_6536_e000_0000, 0x3e2d_46f5_c1d0_c4b8),
    (0x3fb2_53f6_2000_0000, 0x3e2e_1428_2df1_f6d3),
    (0x3fb3_41d7_9000_0000, 0x3e18_6f47_424a_660d),
    (0x3fb4_2edc_b000_0000, 0x3e2d_4c8d_e077_753e),
    (0x3fb5_1b07_3000_0000, 0x3e2e_0c30_7ed2_4f1c),
    (0x3fb6_0658_a000_0000, 0x3e22_6ea1_8763_bdd3),
    (0x3fb6_f0d2_8000_0000, 0x3e25_cad6_9737_c933),
    (0x3fb7_da76_6000_0000, 0x3e2a_f625_9908_8901),
    (0x3fb8_c345_d000_0000, 0x3e18_c66c_83d6_b2d0),
    (0x3fb9_ab42_4000_0000, 0x3e18_80ce_b36f_b30f),
    (0x3fba_926d_3000_0000, 0x3e24_95aa_c6ca_17a4),
    (0x3fbb_78c8_2000_0000, 0x3e27_61db_4210_878c),
    (0x3fbc_5e54_8000_0000, 0x3e2e_b78e_862b_ac2f),
    (0x3fbd_4313_d000_0000, 0x3e19_b2cd_7579_0dd9),
    (0x3fbe_2707_6000_0000, 0x3e2c_55e5_cbd3_d50f),
    (0x3fbf_0a30_c000_0000, 0x3db1_62a6_617c_c971),
    (0x3fbf_ec91_3000_0000, 0x3dfd_beab_aaa2_e519),
    (0x3fc0_6715_1000_0000, 0x3e16_52cb_7150_c647),
    (0x3fc0_d77e_7000_0000, 0x3e39_a11c_b2cd_2ee2),
    (0x3fc1_4785_8000_0000, 0x3e21_9d0a_b1a2_8813),
    (0x3fc1_b72a_d000_0000, 0x3e24_bd9e_80a4_1811),
    (0x3fc2_266f_1000_0000, 0x3e32_14b5_96fa_a3df),
    (0x3fc2_9552_f000_0000, 0x3e30_3fea_4698_0bb8),
    (0x3fc3_03d7_1000_0000, 0x3e31_c8ff_a5fd_28c7),
    (0x3fc3_71fc_2000_0000, 0x3dce_8f74_3bcd_96c5),
    (0x3fc3_dfc2_b000_0000, 0x3dfd_98c5_3953_15c6),
    (0x3fc4_4d2b_6000_0000, 0x3e39_96fa_3ccf_a7b2),
    (0x3fc4_ba36_f000_0000, 0x3e1c_d2af_2ad1_3037),
    (0x3fc5_26e5_e000_0000, 0x3e1d_0da1_bd17_200e),
    (0x3fc5_9338_d000_0000, 0x3e33_3041_0ba6_8b75),
    (0x3fc5_ff30_7000_0000, 0x3df4_f27a_790e_7c41),
    (0x3fc6_6acd_4000_0000, 0x3e13_956a_86f6_ff1b),
    (0x3fc6_d60f_e000_0000, 0x3e2c_6748_7235_51d9),
    (0x3fc7_40f8_f000_0000, 0x3e25_00de_9326_cdfc),
    (0x3fc7_ab89_0000_0000, 0x3e10_86c8_48df_1b59),
    (0x3fc8_15c0_a000_0000, 0x3e04_357e_ad68_36ff),
    (0x3fc8_7fa0_6000_0000, 0x3e24_8324_4240_8024),
    (0x3fc8_e928_d000_0000, 0x3e3d_10da_8154_b13d),
    (0x3fc9_525a_9000_0000, 0x3e39_e8ad_68ec_8260),
    (0x3fc9_bb36_2000_0000, 0x3e3c_fbf7_06ab_af18),
    (0x3fca_23bc_1000_0000, 0x3e3f_c56a_c632_6e23),
    (0x3fca_8bec_f000_0000, 0x3e39_105e_3185_cf21),
    (0x3fca_f3c9_4000_0000, 0x3e3d_017f_e5b1_9cc0),
    (0x3fcb_5b51_9000_0000, 0x3e3d_1f6b_48dd_13fe),
    (0x3fcb_c286_7000_0000, 0x3e20_b633_58a7_e73a),
    (0x3fcc_2968_5000_0000, 0x3e26_3063_028c_211c),
    (0x3fcc_8ff7_c000_0000, 0x3e2e_6a68_86b0_9760),
    (0x3fcc_f635_4000_0000, 0x3e3c_138b_b891_cd03),
    (0x3fcd_5c21_6000_0000, 0x3e36_9f77_22b7_221a),
    (0x3fcd_c1bc_a000_0000, 0x3df5_7d8f_ac1a_628c),
    (0x3fce_2707_6000_0000, 0x3e3c_55e5_cbd3_d50f),
    (0x3fce_8c02_5000_0000, 0x3e15_52d2_ff48_fe2e),
    (0x3fce_f0ad_c000_0000, 0x3e37_b8b2_6ca4_31bc),
    (0x3fcf_550a_5000_0000, 0x3e29_2dec_dc1c_5f6d),
    (0x3fcf_b918_6000_0000, 0x3e3a_bc7c_551a_aa8c),
    (0x3fd0_0e6c_4000_0000, 0x3e36_b540_731a_354b),
    (0x3fd0_4025_9000_0000, 0x3e32_d341_036b_89ef),
    (0x3fd0_71b8_5000_0000, 0x3e4f_9ab2_1a3a_2e0f),
    (0x3fd0_a324_e000_0000, 0x3e23_9c87_1afb_9fbd),
    (0x3fd0_d46b_5000_0000, 0x3e3e_6add_2c81_f640),
    (0x3fd1_058b_f000_0000, 0x3e43_5c95_aa31_3f41),
    (0x3fd1_3687_0000_0000, 0x3e24_9d45_82f6_cc53),
    (0x3fd1_675c_a000_0000, 0x3e47_574c_1c07_398f),
    (0x3fd1_980d_2000_0000, 0x3e4b_a846_dece_9e8d),
    (0x3fd1_c898_c000_0000, 0x3e16_999f_afbc_68e7),
    (0x3fd1_f8ff_9000_0000, 0x3e4c_9145_e51b_0103),
    (0x3fd2_2941_f000_0000, 0x3e47_9ef2_cb44_850a),
    (0x3fd2_5960_1000_0000, 0x3e0b_eec7_3de1_1275),
    (0x3fd2_895a_1000_0000, 0x3e2e_f435_1af5_a498),
    (0x3fd2_b930_3000_0000, 0x3e45_713a_493b_4a50),
    (0x3fd2_e8e2_b000_0000, 0x3e45_c23a_6138_5992),
    (0x3fd3_1871_c000_0000, 0x3e42_a883_09f5_7299),
    (0x3fd3_47dd_9000_0000, 0x3e45_30fa_a9ac_8ace),
    (0x3fd3_7726_6000_0000, 0x3e25_fec2_d792_a758),
    (0x3fd3_a64c_5000_0000, 0x3e35_a517_a71c_bcd7),
    (0x3fd3_d54f_a000_0000, 0x3e37_07dc_3e1c_d9a3),
    (0x3fd4_0430_8000_0000, 0x3e3a_1a9f_8ef4_3049),
    (0x3fd4_32ef_2000_0000, 0x3e44_09d0_276b_3674),
    (0x3fd4_618b_c000_0000, 0x3e20_e2f6_13e8_5bd9),
    (0x3fd4_9006_8000_0000, 0x3df0_0274_3300_1e5f),
    (0x3fd4_be5f_9000_0000, 0x3e35_dde2_836d_3265),
    (0x3fd4_ec97_3000_0000, 0x3e23_0013_4d7a_af04),
    (0x3fd5_1aad_8000_0000, 0x3e3c_b7e0_b427_24f5),
    (0x3fd5_48a2_c000_0000, 0x3e2d_6e93_167e_6308),
    (0x3fd5_7677_1000_0000, 0x3e3d_1569_b152_6adb),
    (0x3fd5_a42a_b000_0000, 0x3e0e_99fc_338a_1a41),
    (0x3fd5_d1bd_b000_0000, 0x3e4e_b013_94a1_1b1c),
    (0x3fd5_ff30_7000_0000, 0x3e04_f27a_790e_7c41),
    (0x3fd6_2c82_f000_0000, 0x3e25_ce3c_a97b_7af9),
    (0x3fd6_59b5_7000_0000, 0x3e28_1f0f_940e_d857),
    (0x3fd6_86c8_1000_0000, 0x3e4d_3629_5d88_857c),
    (0x3fd6_b3bb_2000_0000, 0x3e21_aca1_ec4a_f526),
    (0x3fd6_e08e_a000_0000, 0x3e44_5743_c718_2726),
    (0x3fd7_0d42_e000_0000, 0x3e23_c491_aead_337e),
    (0x3fd7_39d7_f000_0000, 0x3e3a_ef40_1a73_8931),
    (0x3fd7_664e_1000_0000, 0x3e21_cede_7609_2a29),
    (0x3fd7_92a5_5000_0000, 0x3e4f_ba8f_44f8_2bb4),
    (0x3fd7_bede_0000_0000, 0x3e44_6f5f_7f3c_3e1a),
    (0x3fd7_eaf8_3000_0000, 0x3e47_055f_86c9_674b),
    (0x3fd8_16f4_1000_0000, 0x3e4b_41a9_2b6b_6e1a),
    (0x3fd8_42d1_d000_0000, 0x3e44_3d16_2e92_7628),
    (0x3fd8_6e91_9000_0000, 0x3e44_6617_4013_f9b1),
    (0x3fd8_9a33_8000_0000, 0x3e3b_0509_6ad6_9c62),
    (0x3fd8_c5b7_c000_0000, 0x3e40_b169_150f_aa58),
    (0x3fd8_f11e_8000_0000, 0x3e3c_d98b_1df8_5da7),
    (0x3fd9_1c67_e000_0000, 0x3e46_8b50_7b0f_8fa8),
    (0x3fd9_4794_1000_0000, 0x3e48_422d_f574_99ba),
    (0x3fd9_72a3_4000_0000, 0x3e11_3515_8697_0274),
    (0x3fd9_9d95_8000_0000, 0x3e11_7e08_acba_92ee),
    (0x3fd9_c86b_0000_0000, 0x3e26_e043_14dd_0229),
    (0x3fd9_f323_e000_0000, 0x3e49_7f30_97e5_6d1a),
    (0x3fda_1dc0_6000_0000, 0x3e33_56e6_5590_1286),
    (0x3fda_4840_9000_0000, 0x3e0c_b761_457f_94d6),
    (0x3fda_72a4_9000_0000, 0x3e39_af67_a85a_9dac),
    (0x3fda_9cec_9000_0000, 0x3e45_3410_931a_909f),
    (0x3fda_c718_c000_0000, 0x3e22_c587_2060_58f5),
    (0x3fda_f129_3000_0000, 0x3e22_3bc3_5889_9c22),
    (0x3fdb_1b1e_0000_0000, 0x3e4d_7bf8_b6d2_23cb),
    (0x3fdb_44f7_7000_0000, 0x3e47_991e_c519_7ddb),
    (0x3fdb_6eb5_9000_0000, 0x3e4a_79e6_bb3a_9219),
    (0x3fdb_9858_9000_0000, 0x3e3a_4c43_ed66_3ec5),
    (0x3fdb_c1e0_8000_0000, 0x3e46_1b5a_1484_f438),
    (0x3fdb_eb4d_9000_0000, 0x3e4b_4e36_f7ef_0c3a),
    (0x3fdc_149f_f000_0000, 0x3e11_5f02_6acd_0d1b),
    (0x3fdc_3dd7_a000_0000, 0x3e3f_36b5_35ce_cf05),
    (0x3fdc_66f4_e000_0000, 0x3e2f_fb7f_bf3e_b5c6),
    (0x3fdc_8ff7_c000_0000, 0x3e3e_6a68_86b0_9760),
    (0x3fdc_b8e0_7000_0000, 0x3e31_35eb_27f5_bbc3),
    (0x3fdc_e1af_0000_0000, 0x3e47_0be7_d6f6_fa57),
    (0x3fdd_0a63_a000_0000, 0x3e4c_e43c_c84a_b338),
    (0x3fdd_32fe_7000_0000, 0x3e4c_01d7_aac3_bd91),
    (0x3fdd_5b7f_9000_0000, 0x3e45_c58d_0796_1060),
    (0x3fdd_83e7_2000_0000, 0x3e36_28bc_f941_456e),
    (0x3fdd_ac35_3000_0000, 0x3e4c_58b2_a846_1cd2),
    (0x3fdd_d46a_0000_0000, 0x3e33_0712_82fb_989a),
    (0x3fdd_fc85_9000_0000, 0x3e42_0dab_6a80_f09c),
    (0x3fde_2488_1000_0000, 0x3e44_f8d8_4c39_7b1e),
    (0x3fde_4c71_a000_0000, 0x3e40_d0ee_0859_9e48),
    (0x3fde_7442_6000_0000, 0x3e1d_6878_7e37_da36),
    (0x3fde_9bfa_6000_0000, 0x3e36_6187_d591_bafc),
    (0x3fde_c399_d000_0000, 0x3e22_3466_00ba_e772),
    (0x3fde_eb20_c000_0000, 0x3e39_0377_d0d6_1b8e),
    (0x3fdf_128f_5000_0000, 0x3e4f_5e0d_d966_b907),
    (0x3fdf_39e5_b000_0000, 0x3e49_023c_b79a_00e2),
    (0x3fdf_6123_f000_0000, 0x3e44_e051_58c2_8ad8),
    (0x3fdf_884a_3000_0000, 0x3e3b_fa7b_08b1_8ae4),
    (0x3fdf_af58_8000_0000, 0x3e4e_f1e6_3db3_5f67),
    (0x3fdf_d64f_2000_0000, 0x3e0e_c2ae_3949_3d4f),
    (0x3fdf_fd2e_0000_0000, 0x3e40_afe9_30ab_2fa0),
    (0x3fe0_11fa_b000_0000, 0x3e22_5ff8_a181_0dd4),
    (0x3fe0_2552_a000_0000, 0x3e46_9743_fb1a_71a5),
    (0x3fe0_389e_e000_0000, 0x3e5f_9cc6_7678_5571),
    (0x3fe0_4bdf_9000_0000, 0x3e5b_524d_a4cb_f982),
    (0x3fe0_5f14_b000_0000, 0x3e5a_4c8b_3815_35b8),
    (0x3fe0_723e_5000_0000, 0x3e58_39be_809c_af2c),
    (0x3fe0_855c_8000_0000, 0x3e50_968a_1cb8_2c13),
    (0x3fe0_986f_4000_0000, 0x3e5e_ae6a_4172_3fb5),
    (0x3fe0_ab76_b000_0000, 0x3e5d_9c29_a380_a4db),
    (0x3fe0_be72_e000_0000, 0x3e40_94aa_0ada_625e),
    (0x3fe0_d163_c000_0000, 0x3e59_73ad_6fc1_08ca),
    (0x3fe0_e449_8000_0000, 0x3e47_4732_2fdb_ab97),
    (0x3fe0_f724_1000_0000, 0x3e59_3692_fa9d_4221),
    (0x3fe1_09f3_9000_0000, 0x3e5c_5a99_2dfb_c7d9),
    (0x3fe1_1cb8_1000_0000, 0x3e4e_1f33_e102_387a),
    (0x3fe1_2f71_9000_0000, 0x3e46_4fbe_f14c_048c),
    (0x3fe1_4220_2000_0000, 0x3e44_90f5_13ca_5e3b),
    (0x3fe1_54c3_d000_0000, 0x3e37_a6af_4d4c_799d),
    (0x3fe1_675c_a000_0000, 0x3e57_574c_1c07_398f),
    (0x3fe1_79ea_b000_0000, 0x3e57_b133_417f_8c1c),
    (0x3fe1_8c6e_0000_0000, 0x3e5f_eb9e_0c17_6514),
    (0x3fe1_9ee6_b000_0000, 0x3e41_9f25_bb31_72f7),
    (0x3fe1_b154_b000_0000, 0x3e45_f68a_7bbf_b852),
    (0x3fe1_c3b8_1000_0000, 0x3e5e_e278_4979_29f1),
    (0x3fe1_d610_f000_0000, 0x3e5c_cee0_0610_9d58),
    (0x3fe1_e85f_5000_0000, 0x3e5c_e081_a07b_d8b3),
    (0x3fe1_faa3_4000_0000, 0x3e57_0e12_9818_17b8),
    (0x3fe2_0cdc_d000_0000, 0x3e29_2ab6_d935_03d0),
    (0x3fe2_1f0b_f000_0000, 0x3e58_cb7d_d7c3_b61e),
    (0x3fe2_3130_d000_0000, 0x3e4e_fafd_0a0b_78da),
    (0x3fe2_434b_6000_0000, 0x3e5e_9072_67c4_288e),
    (0x3fe2_555b_c000_0000, 0x3e5d_31ef_9678_0875),
    (0x3fe2_6762_0000_0000, 0x3e23_430d_fcd2_ad50),
    (0x3fe2_795e_1000_0000, 0x3e34_4d88_d75b_c1f9),
    (0x3fe2_8b50_0000_0000, 0x3e5b_ec0f_055e_04fc),
    (0x3fe2_9d37_f000_0000, 0x3e5d_8561_1590_b9ad),
    (0x3fe2_af15_f000_0000, 0x3df3_2056_8e58_3229),
    (0x3fe2_c0e9_e000_0000, 0x3e5a_891d_1772_f538),
    (0x3fe2_d2b4_0000_0000, 0x3e22_edc9_dabb_a74d),
    (0x3fe2_e474_3000_0000, 0x3e4b_9009_a101_5086),
    (0x3fe2_f62a_9000_0000, 0x3e52_a12a_8c5b_1a19),
    (0x3fe3_07d7_3000_0000, 0x3e3a_7885_f0fd_ac85),
    (0x3fe3_197a_0000_0000, 0x3e5f_4ffc_d43a_c691),
    (0x3fe3_2b13_3000_0000, 0x3e52_243a_e264_0aad),
    (0x3fe3_3ca2_b000_0000, 0x3e54_6513_2990_35d3),
    (0x3fe3_4e28_9000_0000, 0x3e5b_39c3_a62d_d725),
    (0x3fe3_5fa4_e000_0000, 0x3e5b_a6dd_4004_9f51),
    (0x3fe3_7117_b000_0000, 0x3e45_1d1e_d717_7409),
    (0x3fe3_8280_f000_0000, 0x3e5c_b0f2_fd7f_5216),
    (0x3fe3_93e0_d000_0000, 0x3e3a_b150_cd4e_2213),
    (0x3fe3_a537_3000_0000, 0x3e5c_fd7b_f319_3844),
    (0x3fe3_b684_4000_0000, 0x3e53_fff8_455f_1dbd),
    (0x3fe3_c7c7_f000_0000, 0x3e5f_ee64_0b90_5fc9),
    (0x3fe3_d902_6000_0000, 0x3e54_e2ad_f548_084c),
    (0x3fe3_ea33_9000_0000, 0x3e3b_597a_dc1e_cdd2),
    (0x3fe3_fb5b_8000_0000, 0x3e43_45bd_096d_3a75),
    (0x3fe4_0c7a_4000_0000, 0x3e51_01b9_d245_3c8b),
    (0x3fe4_1d8f_e000_0000, 0x3e50_8ce5_5cc8_c979),
    (0x3fe4_2e9c_6000_0000, 0x3e5b_bf01_7e59_5f71),
    (0x3fe4_3f9f_e000_0000, 0x3e37_ce73_3bd3_93dc),
    (0x3fe4_509a_5000_0000, 0x3e23_3bb0_a503_f8a1),
    (0x3fe4_618b_c000_0000, 0x3e30_e2f6_13e8_5bd9),
    (0x3fe4_7274_3000_0000, 0x3e5e_6755_5a63_5b3c),
    (0x3fe4_8353_d000_0000, 0x3e2e_a88d_f73d_5e8b),
    (0x3fe4_942a_8000_0000, 0x3e3d_17e0_3bda_18a8),
    (0x3fe4_a4f8_5000_0000, 0x3e5b_607d_7604_4f7e),
    (0x3fe4_b5bd_6000_0000, 0x3e52_adc4_e71b_c2fc),
    (0x3fe4_c679_a000_0000, 0x3e5f_99dc_7362_d1d9),
    (0x3fe4_d72d_3000_0000, 0x3e54_73fa_008e_6a6a),
    (0x3fe4_e7d8_1000_0000, 0x3e2b_75bb_09cb_0985),
    (0x3fe4_f87a_3000_0000, 0x3e5e_a04d_d10b_9aba),
    (0x3fe5_0913_c000_0000, 0x3e58_02d0_d697_9674),
    (0x3fe5_19a4_c000_0000, 0x3e17_4688_ccd9_9094),
    (0x3fe5_2a2d_2000_0000, 0x3e49_6f16_abb9_df22),
    (0x3fe5_3aad_0000_0000, 0x3e46_e66d_f2aa_374f),
    (0x3fe5_4b24_6000_0000, 0x3e4e_6652_5ea4_550a),
    (0x3fe5_5b93_5000_0000, 0x3e42_d02f_34f2_0cbd),
    (0x3fe5_6bf9_d000_0000, 0x3e46_cfce_6504_7188),
    (0x3fe5_7c57_f000_0000, 0x3e39_b78c_842d_58b8),
    (0x3fe5_8cad_b000_0000, 0x3e47_35e6_24c2_4bc9),
    (0x3fe5_9cfb_2000_0000, 0x3e47_eba1_f7dd_1adf),
    (0x3fe5_ad40_4000_0000, 0x3e58_6b3e_59f6_5355),
    (0x3fe5_bd7d_3000_0000, 0x3e1c_e38e_637f_1b4d),
    (0x3fe5_cdb1_d000_0000, 0x3e58_d82e_c919_edc7),
    (0x3fe5_ddde_5000_0000, 0x3e4c_5264_8ddc_fa37),
    (0x3fe5_ee02_a000_0000, 0x3e52_482c_eae1_ac12),
    (0x3fe5_fe1e_d000_0000, 0x3e55_a312_311a_ba4f),
    (0x3fe6_0e32_f000_0000, 0x3e41_1e23_6329_f225),
    (0x3fe6_1e3e_f000_0000, 0x3e5b_48c8_cd2f_246c),
    (0x3fe6_2e42_e000_0000, 0x3e6e_fa39_ef35_793c),
];
