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
    fold(left).cmp(&fold(right))
}

fn fold(units: &[u16]) -> Vec<u16> {
    let mut result = Vec::with_capacity(units.len());
    for decoded in char::decode_utf16(units.iter().copied()) {
        match decoded {
            Ok(value) => {
                for lower in value.to_lowercase() {
                    result.extend_from_slice(lower.encode_utf16(&mut [0; 2]));
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
