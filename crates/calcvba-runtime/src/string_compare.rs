//! VBA string ordering shared by execution backends and StrComp.
use std::cmp::Ordering;

pub fn compare_utf16(left: &[u16], right: &[u16], text: bool) -> Ordering {
    if !text {
        return left.cmp(right);
    }
    #[cfg(all(target_os = "windows", not(miri)))]
    if let (Ok(left_len), Ok(right_len)) = (i32::try_from(left.len()), i32::try_from(right.len())) {
        // SAFETY: both buffers remain live for their explicit UTF-16 lengths.
        let result = unsafe {
            windows_sys::Win32::Globalization::CompareStringW(
                windows_sys::Win32::Globalization::LOCALE_USER_DEFAULT,
                windows_sys::Win32::Globalization::NORM_IGNORECASE,
                left.as_ptr(),
                left_len,
                right.as_ptr(),
                right_len,
            )
        };
        match result {
            1 => return Ordering::Less,
            2 => return Ordering::Equal,
            3 => return Ordering::Greater,
            _ => {}
        }
    }
    word_sort_key(left).cmp(&word_sort_key(right))
}

#[derive(Default, PartialEq, Eq, PartialOrd, Ord)]
struct WordSortKey {
    primary: Vec<u32>,
    accents: Vec<u8>,
    widths: Vec<u8>,
    punctuation: Vec<(usize, u8)>,
}

// Use the default Latin weights on the portable path. Other characters
// retain their folded UTF-16 order after the shared Latin range.
fn word_sort_key(units: &[u16]) -> WordSortKey {
    let mut key = WordSortKey::default();
    for unit in fold(units) {
        let weight = match unit {
            0..=0xff => LATIN1_WEIGHTS[usize::from(unit)],
            0x0300..=0x036f => COMBINING_WEIGHTS[usize::from(unit - 0x0300)],
            _ => {
                key.primary.push(0x10000 + u32::from(unit));
                key.accents.push(2);
                key.widths.push(2);
                continue;
            }
        };
        let [script, primary, accent, case] = weight.to_be_bytes();
        match script {
            0 => {}
            1 => {
                if let Some(previous) = key.accents.last_mut() {
                    *previous = previous.saturating_add(accent);
                } else {
                    key.accents.push(accent);
                }
            }
            6 => key.punctuation.push((key.primary.len(), primary)),
            _ => {
                key.primary
                    .push(u32::from(u16::from_be_bytes([script, primary])));
                key.accents.push(accent);
                key.widths.push(case & !0x10);
            }
        }
    }
    key
}

// The default Latin-1 and Combining Diacritical Marks weights are from the
// Windows 11 Sorting Weight Table, published by Microsoft.
// https://www.microsoft.com/en-us/download/details.aspx?id=10921
// Text comparison uses separate primary, diacritic and punctuation keys
// as described in MS-UCODEREF 3.1.5.2.4.

#[rustfmt::skip]
const LATIN1_WEIGHTS: [u32; 256] = [
    0x0000_0000, 0x0603_0202, 0x0604_0202, 0x0605_0202, 0x0606_0202, 0x0607_0202, 0x0608_0202, 0x0609_0202,
    0x060a_0202, 0x0705_0202, 0x0706_0202, 0x0707_0202, 0x0708_0202, 0x0709_0202, 0x060b_0202, 0x060c_0202,
    0x060d_0202, 0x060e_0202, 0x060f_0202, 0x0610_0202, 0x0611_0202, 0x0612_0202, 0x0613_0202, 0x0614_0202,
    0x0615_0202, 0x0616_0202, 0x0617_0202, 0x0618_0202, 0x0619_0202, 0x061a_0202, 0x061b_0202, 0x061c_0202,
    0x0702_0202, 0x071c_0202, 0x071d_0202, 0x071f_0202, 0x0721_0202, 0x0723_0202, 0x0725_0202, 0x0680_0202,
    0x0727_0202, 0x072a_0202, 0x072d_0202, 0x0803_0202, 0x072f_0202, 0x0682_0202, 0x0733_0202, 0x0735_0202,
    0x0d03_0202, 0x0d1a_0202, 0x0d1c_0202, 0x0d1e_0202, 0x0d20_0202, 0x0d22_0202, 0x0d24_0202, 0x0d26_0202,
    0x0d28_0202, 0x0d2a_0202, 0x0737_0202, 0x073a_0202, 0x080e_0202, 0x0812_0202, 0x0814_0202, 0x073c_0202,
    0x073e_0202, 0x0e02_0212, 0x0e09_0212, 0x0e0a_0212, 0x0e1a_0212, 0x0e21_0212, 0x0e23_0212, 0x0e25_0212,
    0x0e2c_0212, 0x0e32_0212, 0x0e35_0212, 0x0e36_0212, 0x0e48_0212, 0x0e51_0212, 0x0e70_0212, 0x0e7c_0212,
    0x0e7e_0212, 0x0e89_0212, 0x0e8a_0212, 0x0e91_0212, 0x0e99_0212, 0x0e9f_0212, 0x0ea2_0212, 0x0ea4_0212,
    0x0ea6_0212, 0x0ea7_0212, 0x0ea9_0212, 0x073f_0202, 0x0741_0202, 0x0742_0202, 0x0743_0202, 0x0744_0202,
    0x0748_0202, 0x0e02_0202, 0x0e09_0202, 0x0e0a_0202, 0x0e1a_0202, 0x0e21_0202, 0x0e23_0202, 0x0e25_0202,
    0x0e2c_0202, 0x0e32_0202, 0x0e35_0202, 0x0e36_0202, 0x0e48_0202, 0x0e51_0202, 0x0e70_0202, 0x0e7c_0202,
    0x0e7e_0202, 0x0e89_0202, 0x0e8a_0202, 0x0e91_0202, 0x0e99_0202, 0x0e9f_0202, 0x0ea2_0202, 0x0ea4_0202,
    0x0ea6_0202, 0x0ea7_0202, 0x0ea9_0202, 0x074a_0202, 0x074c_0202, 0x074e_0202, 0x0750_0202, 0x061d_0202,
    0x0cfa_1d02, 0x0cfa_1e02, 0x0cfa_1f02, 0x0cfa_2002, 0x0cfa_2102, 0x0cfa_2202, 0x0cfa_2302, 0x0cfa_2402,
    0x0cfa_2502, 0x0cfa_2602, 0x0cfa_2702, 0x0cfa_2802, 0x0cfa_2902, 0x0cfa_2a02, 0x0cfa_2b02, 0x0cfa_2c02,
    0x0cfa_2d02, 0x0cfa_2e02, 0x0cfa_2f02, 0x0cfa_3002, 0x0cfa_3102, 0x0cfa_3202, 0x0cfa_3302, 0x0cfa_3402,
    0x0cfa_3502, 0x0cfa_3602, 0x0cfa_3702, 0x0cfa_3802, 0x0cfa_3902, 0x0cfa_3a02, 0x0cfa_3b02, 0x0cfa_3c02,
    0x0704_0202, 0x0751_0202, 0x0797_0202, 0x0798_0202, 0x0799_0202, 0x079a_0202, 0x0752_0202, 0x0a06_0202,
    0x0753_0202, 0x0a07_0202, 0x0e02_030e, 0x0818_0202, 0x0a08_0202, 0x0000_0000, 0x0a09_0202, 0x0754_0202,
    0x0a0a_0202, 0x0817_0202, 0x0d1c_020e, 0x0d1e_020e, 0x0755_0202, 0x0a0b_0202, 0x0a0c_0202, 0x0a0d_0202,
    0x0756_0202, 0x0d1a_020e, 0x0e7c_030e, 0x081a_0202, 0x0d0d_0202, 0x0d11_0202, 0x0d15_0202, 0x0757_0202,
    0x0e02_0f12, 0x0e02_0e12, 0x0e02_1212, 0x0e02_1912, 0x0e02_1312, 0x0e02_1a12, 0x0000_0000, 0x0e0a_1c12,
    0x0e21_0f12, 0x0e21_0e12, 0x0e21_1212, 0x0e21_1312, 0x0e32_0f12, 0x0e32_0e12, 0x0e32_1212, 0x0e32_1312,
    0x0e1a_6812, 0x0e70_1912, 0x0e7c_0f12, 0x0e7c_0e12, 0x0e7c_1212, 0x0e7c_1912, 0x0e7c_1312, 0x081c_0202,
    0x0e7c_2112, 0x0e9f_0f12, 0x0e9f_0e12, 0x0e9f_1212, 0x0e9f_1312, 0x0ea7_0e12, 0x0000_0000, 0x0000_0000,
    0x0e02_0f02, 0x0e02_0e02, 0x0e02_1202, 0x0e02_1902, 0x0e02_1302, 0x0e02_1a02, 0x0000_0000, 0x0e0a_1c02,
    0x0e21_0f02, 0x0e21_0e02, 0x0e21_1202, 0x0e21_1302, 0x0e32_0f02, 0x0e32_0e02, 0x0e32_1202, 0x0e32_1302,
    0x0e1a_6802, 0x0e70_1902, 0x0e7c_0f02, 0x0e7c_0e02, 0x0e7c_1202, 0x0e7c_1902, 0x0e7c_1302, 0x081d_0202,
    0x0e7c_2102, 0x0e9f_0f02, 0x0e9f_0e02, 0x0e9f_1202, 0x0e9f_1302, 0x0ea7_0e02, 0x0000_0000, 0x0ea7_1302,
];

#[rustfmt::skip]
const COMBINING_WEIGHTS: [u32; 112] = [
    0x0100_0d00, 0x0100_0c00, 0x0100_1000, 0x0100_1700, 0x0100_1500, 0x0100_3e00, 0x0100_1300, 0x0100_0e00,
    0x0100_1100, 0x0100_3f00, 0x0100_1800, 0x0100_1b00, 0x0100_1200, 0x0100_0300, 0x0100_4100, 0x0100_4200,
    0x0100_4300, 0x0100_4400, 0x0100_4500, 0x0100_4600, 0x0100_4700, 0x0100_4800, 0x0100_4900, 0x0100_4a00,
    0x0100_4b00, 0x0100_4c00, 0x0100_4d00, 0x0100_4e00, 0x0100_4f00, 0x0100_5000, 0x0100_5100, 0x0100_5200,
    0x0100_5300, 0x0100_5400, 0x0100_5500, 0x0100_5600, 0x0100_5700, 0x0100_5800, 0x0100_5900, 0x0100_1a00,
    0x0100_1900, 0x0100_5a00, 0x0100_5b00, 0x0100_5c00, 0x0100_5d00, 0x0100_5e00, 0x0100_5f00, 0x0100_6000,
    0x0100_6100, 0x0100_6200, 0x0100_6300, 0x0100_6400, 0x0100_6500, 0x0100_1c00, 0x0100_1c00, 0x0100_1d00,
    0x0100_1e00, 0x0100_6700, 0x0100_6800, 0x0100_6900, 0x0100_6a00, 0x0100_6b00, 0x0100_6c00, 0x0100_6d00,
    0x0100_0d00, 0x0100_0c00, 0x0100_8d00, 0x0100_4600, 0x0100_1d00, 0x0100_9000, 0x0100_7400, 0x0100_7500,
    0x0100_7600, 0x0100_5c00, 0x0100_5d00, 0x0100_5e00, 0x0100_5f00, 0x0100_6000, 0x0100_6100, 0x0000_0000,
    0x0100_5400, 0x0100_5500, 0x0100_5600, 0x0100_5700, 0x0100_5800, 0x0100_5900, 0x0100_5a00, 0x0100_5b00,
    0x0100_1e00, 0x0100_1f00, 0x0100_2000, 0x0100_2100, 0x0100_2800, 0x0100_1400, 0x0100_1500, 0x0100_1500,
    0x0100_1700, 0x0100_4400, 0x0100_2700, 0x0e02_0302, 0x0e21_0302, 0x0e32_0302, 0x0e7c_0302, 0x0e9f_0302,
    0x0e0a_0302, 0x0e1a_0302, 0x0e2c_0302, 0x0e51_0302, 0x0e8a_0302, 0x0e99_0302, 0x0ea2_0302, 0x0ea6_0302,
];

fn fold(units: &[u16]) -> Vec<u16> {
    let mut result = Vec::with_capacity(units.len());
    for decoded in char::decode_utf16(units.iter().copied()) {
        match decoded {
            Ok(value) => {
                for lower in value.to_lowercase() {
                    // Linguistic text compare expands these letters (`CompareString`).
                    let expansion = match lower {
                        '\u{e6}' => Some("ae"),
                        '\u{153}' => Some("oe"),
                        '\u{df}' => Some("ss"),
                        _ => None,
                    };
                    match expansion {
                        Some(text) => result.extend(text.encode_utf16()),
                        None => result.extend_from_slice(lower.encode_utf16(&mut [0; 2])),
                    }
                }
            }
            Err(value) => result.push(value.unpaired_surrogate()),
        }
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn unicode_text_and_binary_keep_their_distinct_rules() {
        assert_eq!(compare_utf16(&[0xc5], &[0xe5], true), Ordering::Equal);
        assert_eq!(compare_utf16(&[0xc5], &[0xe5], false), Ordering::Less);
        assert_eq!(compare_utf16(&[0xd800], &[0xdc00], false), Ordering::Less);
        assert_eq!(
            compare_utf16(&[65, 0, 66], &[97, 0, 98], true),
            Ordering::Equal
        );
    }
}
