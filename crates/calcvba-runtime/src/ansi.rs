//! ANSI (system or project code page) ↔ Unicode conversion for byte-exact VBA string I/O.
//!
//! VBA 7.1 stores strings on disk, marshals them across `Declare` 'A' boundaries,
//! and computes `Chr`/`Asc`/`String`/`StrConv` against the active ANSI code page.
//!
//! When a project or host configures an explicit code page via [`enter_codepage`] /
//! [`with_codepage`], that code page is used deterministically on every platform.
//! Otherwise, Windows uses the host's `CP_ACP` via `WideCharToMultiByte` /
//! `MultiByteToWideChar` and non-Windows hosts default to Windows-1252.

use std::cell::Cell;

/// Default ANSI code page when a project or host does not configure one off Windows.
pub const DEFAULT_CODEPAGE: u16 = 1252;

thread_local! {
    static ACTIVE_CODEPAGE: Cell<Option<u16>> = const { Cell::new(None) };
}

/// True if `codepage` is a supported single-byte or multibyte ANSI code page.
pub fn is_supported_codepage(codepage: u16) -> bool {
    codepage == DEFAULT_CODEPAGE
        || codepage::to_encoding(codepage)
            .is_some_and(|enc| enc != encoding_rs::UTF_16LE && enc != encoding_rs::UTF_16BE)
}

/// Currently active project/host code-page override on this thread (`None` = host default).
pub fn active_codepage() -> Option<u16> {
    ACTIVE_CODEPAGE.with(Cell::get)
}

/// RAII guard restoring the previous thread-local code page on drop.
pub struct CodepageGuard {
    previous: Option<u16>,
}

impl Drop for CodepageGuard {
    fn drop(&mut self) {
        ACTIVE_CODEPAGE.with(|cell| cell.set(self.previous));
    }
}

/// Activate `codepage` on the owning VBA thread until the returned guard is dropped.
/// Passing `None` keeps the platform default (`CP_ACP` on Windows, 1252 elsewhere).
pub fn enter_codepage(codepage: Option<u16>) -> CodepageGuard {
    let previous = ACTIVE_CODEPAGE.with(|cell| cell.replace(codepage));
    CodepageGuard { previous }
}

/// Run `f` with `codepage` active on the owning VBA thread.
pub fn with_codepage<T>(codepage: Option<u16>, f: impl FnOnce() -> T) -> T {
    let _guard = enter_codepage(codepage);
    f()
}

/// True if `byte` is a DBCS lead byte in the active ANSI code page.
pub fn is_dbcs_lead_byte(byte: u8) -> bool {
    if let Some(cp) = active_codepage() {
        return is_dbcs_lead_byte_for(cp, byte);
    }
    #[cfg(target_os = "windows")]
    {
        windows_codec::is_dbcs_lead_byte(byte)
    }
    #[cfg(not(target_os = "windows"))]
    {
        false
    }
}

fn is_dbcs_lead_byte_for(codepage: u16, byte: u8) -> bool {
    match codepage {
        // Shift-JIS (932)
        932 => (0x81..=0x9F).contains(&byte) || (0xE0..=0xFC).contains(&byte),
        // GBK (936), EUC-KR/UHC (949), Big5 (950)
        936 | 949 | 950 => (0x81..=0xFE).contains(&byte),
        _ => false,
    }
}

/// Encode a Unicode string to ANSI bytes in the active code page. No terminator.
pub fn ansi_encode(text: &str) -> Vec<u8> {
    if let Some(cp) = active_codepage() {
        return if cp == DEFAULT_CODEPAGE {
            cp1252::encode(text)
        } else {
            configured::encode(text, cp)
        };
    }
    #[cfg(target_os = "windows")]
    {
        windows_codec::encode(text)
    }
    #[cfg(not(target_os = "windows"))]
    {
        cp1252::encode(text)
    }
}

/// Decode ANSI bytes in the active code page to a Unicode string.
pub fn ansi_decode(bytes: &[u8]) -> String {
    if let Some(cp) = active_codepage() {
        return if cp == DEFAULT_CODEPAGE {
            cp1252::decode(bytes)
        } else {
            configured::decode(bytes, cp)
        };
    }
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
    if let Some(cp) = active_codepage() {
        return if cp == DEFAULT_CODEPAGE {
            cp1252::encode_units(units)
        } else {
            configured::encode_units(units, cp)
        };
    }
    #[cfg(target_os = "windows")]
    {
        windows_codec::encode_units(units)
    }
    #[cfg(not(target_os = "windows"))]
    {
        cp1252::encode_units(units)
    }
}

/// Decode ANSI bytes to UTF-16 units at their full length (embedded NULs kept), as VBA
/// converts a `Declare` String buffer back.
pub fn ansi_decode_units(bytes: &[u8]) -> Vec<u16> {
    if let Some(cp) = active_codepage() {
        return if cp == DEFAULT_CODEPAGE {
            cp1252::decode(bytes).encode_utf16().collect()
        } else {
            configured::decode(bytes, cp).encode_utf16().collect()
        };
    }
    #[cfg(target_os = "windows")]
    {
        windows_codec::decode_units(bytes)
    }
    #[cfg(not(target_os = "windows"))]
    {
        cp1252::decode(bytes).encode_utf16().collect()
    }
}

/// Decode a VBA `Chr` argument in the active ANSI code page. Two-byte codes
/// use a packed lead and trail byte on DBCS code pages.
pub fn ansi_decode_character(code: i32) -> Option<u16> {
    if let Ok(byte) = u8::try_from(code) {
        return ansi_decode_units(&[byte]).first().copied();
    }
    if !(-32768..=65535).contains(&code) {
        return None;
    }
    if let Some(cp) = active_codepage() {
        return configured::decode_dbcs_character(cp, (code as u16).to_be_bytes());
    }
    #[cfg(target_os = "windows")]
    {
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

    pub fn is_dbcs_lead_byte(byte: u8) -> bool {
        // SAFETY: CP_ACP selects the Windows ANSI code page and the byte is passed by value.
        unsafe { IsDBCSLeadByteEx(CP_ACP, byte) != 0 }
    }

    pub fn decode_dbcs_character(bytes: [u8; 2]) -> Option<u16> {
        if !is_dbcs_lead_byte(bytes[0]) {
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

mod configured {
    use super::{cp1252, is_dbcs_lead_byte_for};

    fn encoding_for(cp: u16) -> Option<&'static encoding_rs::Encoding> {
        codepage::to_encoding(cp)
            .filter(|enc| *enc != encoding_rs::UTF_16LE && *enc != encoding_rs::UTF_16BE)
    }

    fn encode_char(encoding: &'static encoding_rs::Encoding, c: char, out: &mut Vec<u8>) {
        if c.is_ascii() {
            out.push(c as u8);
            return;
        }
        let mut utf8 = [0u8; 4];
        let text = c.encode_utf8(&mut utf8);
        let (encoded, _, had_errors) = encoding.encode(text);
        if had_errors {
            out.push(b'?');
        } else {
            out.extend_from_slice(&encoded);
        }
    }

    pub fn encode(text: &str, cp: u16) -> Vec<u8> {
        let Some(encoding) = encoding_for(cp) else {
            return cp1252::encode(text);
        };
        let mut out = Vec::with_capacity(text.len());
        for ch in text.chars() {
            encode_char(encoding, ch, &mut out);
        }
        out
    }

    pub fn encode_units(units: &[u16], cp: u16) -> Vec<u8> {
        let Some(encoding) = encoding_for(cp) else {
            return cp1252::encode_units(units);
        };
        let mut out = Vec::with_capacity(units.len());
        for ch in char::decode_utf16(units.iter().copied()) {
            match ch {
                Ok(c) => encode_char(encoding, c, &mut out),
                Err(_) => out.push(b'?'),
            }
        }
        out
    }

    pub fn decode(bytes: &[u8], cp: u16) -> String {
        let Some(encoding) = encoding_for(cp) else {
            return cp1252::decode(bytes);
        };
        encoding.decode_without_bom_handling(bytes).0.into_owned()
    }

    pub fn decode_dbcs_character(cp: u16, bytes: [u8; 2]) -> Option<u16> {
        if !is_dbcs_lead_byte_for(cp, bytes[0]) {
            return None;
        }
        let encoding = encoding_for(cp)?;
        let decoded = encoding.decode_without_bom_handling_and_without_replacement(&bytes)?;
        let mut units = decoded.encode_utf16();
        let first = units.next()?;
        units.next().is_none().then_some(first)
    }
}

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

    pub fn encode_units(units: &[u16]) -> Vec<u8> {
        char::decode_utf16(units.iter().copied())
            .map(|c| encode_char(c.unwrap_or('?')))
            .collect()
    }

    pub fn encode_char(c: char) -> u8 {
        let u = c as u32;
        if u < 0x80 || (0xA0..=0xFF).contains(&u) {
            u as u8
        } else if let Some(i) = HIGH.iter().position(|&h| h == c) {
            0x80 + i as u8
        } else {
            // https://www.unicode.org/Public/MAPPINGS/VENDORS/MICSFT/WindowsBestFit/bestfit1252.txt
            match c {
                '\u{0130}' => b'I',
                '\u{3000}' => b' ',
                _ => b'?',
            }
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

    #[test]
    fn configured_codepages_round_trip_and_restore_previous() {
        assert!(is_supported_codepage(1252));
        assert!(is_supported_codepage(1251));
        assert!(is_supported_codepage(1250));
        assert!(is_supported_codepage(932));
        assert!(!is_supported_codepage(65000));
        assert!(!is_supported_codepage(1200));

        assert_eq!(active_codepage(), None);
        with_codepage(Some(1251), || {
            assert_eq!(active_codepage(), Some(1251));
            assert_eq!(ansi_encode("Привет"), b"\xcf\xf0\xe8\xe2\xe5\xf2");
            assert_eq!(ansi_decode(b"\xcf\xf0\xe8\xe2\xe5\xf2"), "Привет");
            // Unrepresentable characters become '?' rather than HTML entities.
            assert_eq!(ansi_encode("漢"), b"?");
            // Nested override restores 1251.
            with_codepage(Some(1250), || {
                assert_eq!(ansi_encode("Łž"), b"\xa3\x9e");
                assert_eq!(ansi_decode(b"\xa3\x9e"), "Łž");
            });
            assert_eq!(active_codepage(), Some(1251));
        });
        assert_eq!(active_codepage(), None);
    }

    #[test]
    fn configured_dbcs_codepage_decodes_packed_chr_and_lead_bytes() {
        with_codepage(Some(932), || {
            assert!(is_dbcs_lead_byte(0x82));
            assert!(!is_dbcs_lead_byte(b'A'));
            assert_eq!(ansi_encode("あ"), b"\x82\xa0");
            assert_eq!(ansi_decode(b"\x82\xa0"), "あ");
            assert_eq!(ansi_decode_character(0x82A0), Some('あ' as u16));
            // Signed negative i16 packing for 0x82A0 (-32096).
            assert_eq!(ansi_decode_character(-32096), Some('あ' as u16));
        });
    }
}
