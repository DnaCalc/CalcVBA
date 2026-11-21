//! ANSI (system code page) ↔ Unicode conversion for byte-exact VBA string I/O.
//!
//! VBA 7.1 stores strings on disk, marshals them across `Declare` 'A' boundaries,
//! and computes `Chr`/`Asc` against **ANSI** — the active system code page
//! (`CP_ACP`), via `WideCharToMultiByte`/`MultiByteToWideChar`. We do exactly that
//! on Windows, so the codec is faithful to whatever ACP the host machine runs.
//!
//! On non-Windows targets (where VBA itself does not run — these builds exist only
//! for cross-platform development and CI) we use a built-in **Windows-1252** codec,
//! the predominant Western ANSI code page, so the dev lane is correct for Latin
//! text. DBCS / non-1252 system code pages are only reproduced on real Windows.
//!
//! This lives in `calcvba-runtime` (the value substrate at the bottom of the stack)
//! so both the `Declare` ANSI string marshalling (`calcvba-com`/`calcvba-hal`) and the
//! `Chr`/`Asc` built-ins (`calcvba-lib`) share one codec rather than duplicating a
//! hardcoded CP-1252 table.

/// Encode a Unicode string to ANSI (system code page) bytes. No terminator.
pub fn ansi_encode(text: &str) -> Vec<u8> {
    #[cfg(target_os = "windows")]
    {
        windows_codec::encode(text)
    }
    #[cfg(not(target_os = "windows"))]
    {
        cp1252::encode(text)
    }
}

/// Decode ANSI (system code page) bytes to a Unicode string.
pub fn ansi_decode(bytes: &[u8]) -> String {
    #[cfg(target_os = "windows")]
    {
        windows_codec::decode(bytes)
    }
    #[cfg(not(target_os = "windows"))]
    {
        cp1252::decode(bytes)
    }
}

/// Encode UTF-16 units (a VBA String, unpaired surrogates included) to ANSI bytes, best fit,
/// as VBA converts a `Declare` String argument. No terminator.
pub fn ansi_encode_units(units: &[u16]) -> Vec<u8> {
    #[cfg(target_os = "windows")]
    {
        windows_codec::encode_units(units)
    }
    #[cfg(not(target_os = "windows"))]
    {
        char::decode_utf16(units.iter().copied())
            .map(|c| cp1252::encode_char(c.unwrap_or('?')))
            .collect()
    }
}

/// Decode ANSI bytes to UTF-16 units at their full length (embedded NULs kept), as VBA
/// converts a `Declare` String buffer back.
pub fn ansi_decode_units(bytes: &[u8]) -> Vec<u16> {
    #[cfg(target_os = "windows")]
    {
        windows_codec::decode_units(bytes)
    }
    #[cfg(not(target_os = "windows"))]
    {
        cp1252::decode(bytes).encode_utf16().collect()
    }
}

/// Decode a VBA `Chr` argument in the system ANSI code page. Two-byte codes
/// use a packed lead and trail byte on DBCS hosts.
pub fn ansi_decode_character(code: i32) -> Option<u16> {
    if let Ok(byte) = u8::try_from(code) {
        return ansi_decode_units(&[byte]).first().copied();
    }
    #[cfg(target_os = "windows")]
    {
        if !(-32768..=65535).contains(&code) {
            return None;
        }
        windows_codec::decode_dbcs_character((code as u16).to_be_bytes())
    }
    #[cfg(not(target_os = "windows"))]
    {
        None
    }
}

#[cfg(target_os = "windows")]
mod windows_codec {
    const CP_ACP: u32 = 0;
    const MB_ERR_INVALID_CHARS: u32 = 8;

    // SAFETY: these declarations transcribe the documented Win32 ABI signatures of
    // kernel32's WideCharToMultiByte/MultiByteToWideChar exactly.
    unsafe extern "system" {
        fn IsDBCSLeadByteEx(code_page: u32, byte: u8) -> i32;
        fn WideCharToMultiByte(
            code_page: u32,
            flags: u32,
            wide: *const u16,
            wide_len: i32,
            multi_byte: *mut u8,
            multi_byte_len: i32,
            default_char: *const u8,
            used_default: *mut i32,
        ) -> i32;
        fn MultiByteToWideChar(
            code_page: u32,
            flags: u32,
            multi_byte: *const u8,
            multi_byte_len: i32,
            wide: *mut u16,
            wide_len: i32,
        ) -> i32;
    }

    pub fn decode_dbcs_character(bytes: [u8; 2]) -> Option<u16> {
        // SAFETY: CP_ACP selects the Windows ANSI code page and the byte is passed by value.
        if unsafe { IsDBCSLeadByteEx(CP_ACP, bytes[0]) } == 0 {
            return None;
        }
        let mut unit = 0;
        // SAFETY: the two input bytes and one output unit are live for this call.
        let written = unsafe {
            MultiByteToWideChar(
                CP_ACP,
                MB_ERR_INVALID_CHARS,
                bytes.as_ptr(),
                2,
                &mut unit,
                1,
            )
        };
        (written == 1).then_some(unit)
    }

    pub fn encode(text: &str) -> Vec<u8> {
        if text.is_empty() {
            return Vec::new();
        }
        let wide: Vec<u16> = text.encode_utf16().collect();
        let out = encode_units(&wide);
        if out.is_empty() {
            return text.as_bytes().to_vec();
        }
        out
    }

    pub fn encode_units(wide: &[u16]) -> Vec<u8> {
        if wide.is_empty() {
            return Vec::new();
        }
        // SAFETY: `wide` is valid for `wide.len()` UTF-16 units for the duration of
        // the call, and a null output buffer with length 0 is the documented
        // size-query mode of WideCharToMultiByte (nothing is written).
        let needed = unsafe {
            WideCharToMultiByte(
                CP_ACP,
                0,
                wide.as_ptr(),
                wide.len() as i32,
                std::ptr::null_mut(),
                0,
                std::ptr::null(),
                std::ptr::null_mut(),
            )
        };
        if needed <= 0 {
            return Vec::new();
        }
        let mut out = vec![0u8; needed as usize];
        // SAFETY: `wide` is valid for `wide.len()` UTF-16 units and `out` was
        // allocated just above with exactly `needed` bytes, so the output window the
        // converter may write is in bounds.
        let written = unsafe {
            WideCharToMultiByte(
                CP_ACP,
                0,
                wide.as_ptr(),
                wide.len() as i32,
                out.as_mut_ptr(),
                needed,
                std::ptr::null(),
                std::ptr::null_mut(),
            )
        };
        out.truncate(written.max(0) as usize);
        out
    }

    pub fn decode(bytes: &[u8]) -> String {
        if bytes.is_empty() {
            return String::new();
        }
        let wide = decode_units(bytes);
        if wide.is_empty() {
            return String::from_utf8_lossy(bytes).into_owned();
        }
        String::from_utf16_lossy(&wide)
    }

    pub fn decode_units(bytes: &[u8]) -> Vec<u16> {
        if bytes.is_empty() {
            return Vec::new();
        }
        // SAFETY: `bytes` is valid for `bytes.len()` bytes for the duration of the
        // call, and a null output buffer with length 0 is the documented size-query
        // mode of MultiByteToWideChar (nothing is written).
        let needed = unsafe {
            MultiByteToWideChar(
                CP_ACP,
                0,
                bytes.as_ptr(),
                bytes.len() as i32,
                std::ptr::null_mut(),
                0,
            )
        };
        if needed <= 0 {
            return Vec::new();
        }
        let mut wide = vec![0u16; needed as usize];
        // SAFETY: `bytes` is valid for `bytes.len()` bytes and `wide` was allocated
        // just above with exactly `needed` u16 slots, so the output window is in
        // bounds; the API writes at most `needed` units, keeping the
        // `[..written]` slice below in bounds too.
        let written = unsafe {
            MultiByteToWideChar(
                CP_ACP,
                0,
                bytes.as_ptr(),
                bytes.len() as i32,
                wide.as_mut_ptr(),
                needed,
            )
        };
        wide.truncate(written.max(0) as usize);
        wide
    }
}

#[cfg(not(target_os = "windows"))]
mod cp1252 {
    /// Windows-1252 mapping for the 0x80..=0x9F range (the rest of 0x00..=0xFF is
    /// Latin-1 identity). Windows maps the five unassigned positions to their C1 control codes.
    const HIGH: [char; 32] = [
        '\u{20AC}', '\u{0081}', '\u{201A}', '\u{0192}', '\u{201E}', '\u{2026}', '\u{2020}',
        '\u{2021}', '\u{02C6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{008D}',
        '\u{017D}', '\u{008F}', '\u{0090}', '\u{2018}', '\u{2019}', '\u{201C}', '\u{201D}',
        '\u{2022}', '\u{2013}', '\u{2014}', '\u{02DC}', '\u{2122}', '\u{0161}', '\u{203A}',
        '\u{0153}', '\u{009D}', '\u{017E}', '\u{0178}',
    ];

    pub fn encode(text: &str) -> Vec<u8> {
        text.chars().map(encode_char).collect()
    }

    pub fn encode_char(c: char) -> u8 {
        let u = c as u32;
        if u < 0x80 || (0xA0..=0xFF).contains(&u) {
            u as u8
        } else if let Some(i) = HIGH.iter().position(|&h| h == c) {
            0x80 + i as u8
        } else {
            b'?'
        }
    }

    pub fn decode(bytes: &[u8]) -> String {
        bytes
            .iter()
            .map(|&b| {
                if !(0x80..0xA0).contains(&b) {
                    b as char
                } else {
                    HIGH[(b - 0x80) as usize]
                }
            })
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ascii_round_trips_and_is_one_byte_per_char() {
        let bytes = ansi_encode("Hello");
        assert_eq!(bytes, b"Hello");
        assert_eq!(ansi_decode(&bytes), "Hello");
    }

    #[test]
    fn units_keep_embedded_nul_and_length() {
        let units: Vec<u16> = "a\0b".encode_utf16().collect();
        let bytes = ansi_encode_units(&units);
        assert_eq!(bytes, b"a\0b");
        assert_eq!(ansi_decode_units(&bytes), units);
    }

    // On non-Windows the CP-1252 fallback maps é (U+00E9) to the single byte 0xE9.
    #[cfg(not(target_os = "windows"))]
    #[test]
    fn latin1_high_char_is_single_ansi_byte() {
        let bytes = ansi_encode("é");
        assert_eq!(bytes, vec![0xE9]);
        assert_eq!(ansi_decode(&bytes), "é");
    }
}
