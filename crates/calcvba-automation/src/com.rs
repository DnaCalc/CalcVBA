//! COM interface layouts with the Windows ABI, usable on every platform.
//!
//! Off Windows these are CalcVBA's portable stand-ins: the same vtable shapes and calling
//! convention (`extern "system"` is the C convention there), so a CalcVBA object can sit in a
//! `VARIANT`'s `punkVal`/`pdispVal` and a CalcVBA record type in `pRecInfo` on every platform.
//! Interface identifiers travel as `REFIID` (`*const Guid`), never by value.

use core::ffi::c_void;

/// `GUID` / `IID` / `CLSID`.
#[repr(C)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct Guid {
    pub data1: u32,
    pub data2: u16,
    pub data3: u16,
    pub data4: [u8; 8],
}

impl Guid {
    pub const fn new(data1: u32, data2: u16, data3: u16, data4: [u8; 8]) -> Self {
        Self {
            data1,
            data2,
            data3,
            data4,
        }
    }

    /// Parse `XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX`, with or without braces.
    pub fn parse(input: &str) -> Option<Self> {
        fn hex<T: TryFrom<u32>>(text: &str, len: usize) -> Option<T> {
            if text.len() != len {
                return None;
            }
            u32::from_str_radix(text, 16).ok()?.try_into().ok()
        }
        let trimmed = input.trim();
        let body = trimmed
            .strip_prefix('{')
            .and_then(|inner| inner.strip_suffix('}'))
            .unwrap_or(trimmed);
        let mut parts = body.split('-');
        let data1 = hex::<u32>(parts.next()?, 8)?;
        let data2 = hex::<u16>(parts.next()?, 4)?;
        let data3 = hex::<u16>(parts.next()?, 4)?;
        let data4_hi = parts.next()?;
        let data4_lo = parts.next()?;
        if parts.next().is_some() || data4_hi.len() != 4 || data4_lo.len() != 12 {
            return None;
        }
        let mut data4 = [0u8; 8];
        data4[0] = hex::<u8>(data4_hi.get(0..2)?, 2)?;
        data4[1] = hex::<u8>(data4_hi.get(2..4)?, 2)?;
        for (index, byte) in data4[2..].iter_mut().enumerate() {
            *byte = hex::<u8>(data4_lo.get(index * 2..index * 2 + 2)?, 2)?;
        }
        Some(Self::new(data1, data2, data3, data4))
    }
}

/// Upper-case canonical form without braces.
impl core::fmt::Display for Guid {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        let d = &self.data4;
        write!(
            f,
            "{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}",
            self.data1, self.data2, self.data3, d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7]
        )
    }
}

/// An interface pointer: its first word is the vtable.
#[repr(C)]
pub struct IUnknown {
    pub vtbl: *const IUnknownVtbl,
}

/// `HRESULT`.
pub type HResult = i32;

/// `IID_IUnknown` {00000000-0000-0000-C000-000000000046}.
pub const IID_IUNKNOWN: Guid = Guid::new(0, 0, 0, [0xC0, 0, 0, 0, 0, 0, 0, 0x46]);
/// `IID_IDispatch` {00020400-0000-0000-C000-000000000046}.
pub const IID_IDISPATCH: Guid = Guid::new(0x0002_0400, 0, 0, [0xC0, 0, 0, 0, 0, 0, 0, 0x46]);
/// `IID_IRecordInfo` {0000002F-0000-0000-C000-000000000046}.
pub const IID_IRECORDINFO: Guid = Guid::new(0x0000_002F, 0, 0, [0xC0, 0, 0, 0, 0, 0, 0, 0x46]);

/// `IUnknown`'s vtable.
#[repr(C)]
pub struct IUnknownVtbl {
    pub query_interface: unsafe extern "system" fn(
        this: *mut c_void,
        riid: *const Guid,
        object: *mut *mut c_void,
    ) -> HResult,
    pub add_ref: unsafe extern "system" fn(this: *mut c_void) -> u32,
    pub release: unsafe extern "system" fn(this: *mut c_void) -> u32,
}

/// `IDispatch`'s vtable. Names and type info pointers are left untyped here; the marshalling
/// layer owns their meaning.
#[repr(C)]
pub struct IDispatchVtbl {
    pub unknown: IUnknownVtbl,
    pub get_type_info_count:
        unsafe extern "system" fn(this: *mut c_void, count: *mut u32) -> HResult,
    pub get_type_info: unsafe extern "system" fn(
        this: *mut c_void,
        index: u32,
        lcid: u32,
        type_info: *mut *mut c_void,
    ) -> HResult,
    pub get_ids_of_names: unsafe extern "system" fn(
        this: *mut c_void,
        riid: *const Guid,
        names: *const *const u16,
        count: u32,
        lcid: u32,
        dispids: *mut i32,
    ) -> HResult,
    #[allow(clippy::type_complexity)]
    pub invoke: unsafe extern "system" fn(
        this: *mut c_void,
        dispid: i32,
        riid: *const Guid,
        lcid: u32,
        flags: u16,
        params: *mut c_void,
        result: *mut c_void,
        exception: *mut c_void,
        arg_error: *mut u32,
    ) -> HResult,
}

/// `IRecordInfo`'s vtable, in the header's slot order. Record data pointers are untyped; field
/// values are `VARIANT*` (`*mut c_void` here, typed by the value crate that owns `Variant`).
#[repr(C)]
pub struct IRecordInfoVtbl {
    pub unknown: IUnknownVtbl,
    pub record_init: unsafe extern "system" fn(this: *mut c_void, record: *mut c_void) -> HResult,
    pub record_clear: unsafe extern "system" fn(this: *mut c_void, record: *mut c_void) -> HResult,
    pub record_copy: unsafe extern "system" fn(
        this: *mut c_void,
        existing: *mut c_void,
        new: *mut c_void,
    ) -> HResult,
    pub get_guid: unsafe extern "system" fn(this: *mut c_void, guid: *mut Guid) -> HResult,
    pub get_name: unsafe extern "system" fn(this: *mut c_void, name: *mut *mut u16) -> HResult,
    pub get_size: unsafe extern "system" fn(this: *mut c_void, size: *mut u32) -> HResult,
    pub get_type_info:
        unsafe extern "system" fn(this: *mut c_void, type_info: *mut *mut c_void) -> HResult,
    pub get_field: unsafe extern "system" fn(
        this: *mut c_void,
        data: *mut c_void,
        name: *const u16,
        field: *mut c_void,
    ) -> HResult,
    pub get_field_no_copy: unsafe extern "system" fn(
        this: *mut c_void,
        data: *mut c_void,
        name: *const u16,
        field: *mut c_void,
        data_carray: *mut *mut c_void,
    ) -> HResult,
    pub put_field: unsafe extern "system" fn(
        this: *mut c_void,
        flags: u32,
        data: *mut c_void,
        name: *const u16,
        field: *mut c_void,
    ) -> HResult,
    pub put_field_no_copy: unsafe extern "system" fn(
        this: *mut c_void,
        flags: u32,
        data: *mut c_void,
        name: *const u16,
        field: *mut c_void,
    ) -> HResult,
    pub get_field_names: unsafe extern "system" fn(
        this: *mut c_void,
        count: *mut u32,
        names: *mut *mut u16,
    ) -> HResult,
    pub is_matching_type: unsafe extern "system" fn(this: *mut c_void, other: *mut c_void) -> i32,
    pub record_create: unsafe extern "system" fn(this: *mut c_void) -> *mut c_void,
    pub record_create_copy: unsafe extern "system" fn(
        this: *mut c_void,
        source: *mut c_void,
        dest: *mut *mut c_void,
    ) -> HResult,
    pub record_destroy:
        unsafe extern "system" fn(this: *mut c_void, record: *mut c_void) -> HResult,
}

#[cfg(test)]
mod tests {
    use super::*;
    use core::mem::size_of;

    const P: usize = size_of::<*const ()>();

    #[test]
    fn vtable_slot_counts() {
        assert_eq!(size_of::<IUnknownVtbl>(), 3 * P);
        assert_eq!(size_of::<IDispatchVtbl>(), 7 * P);
        // IUnknown (3) + RecordInit .. RecordDestroy (16).
        assert_eq!(size_of::<IRecordInfoVtbl>(), 19 * P);
    }

    #[test]
    fn guid_parse_and_display_round_trip() {
        let parsed = Guid::parse("{00020400-0000-0000-c000-000000000046}").unwrap();
        assert_eq!(parsed, IID_IDISPATCH);
        assert_eq!(
            Guid::parse("00000000-0000-0000-C000-000000000046"),
            Some(IID_IUNKNOWN)
        );
        assert!(Guid::parse("00020400-0000-0000-C000").is_none());
        assert!(Guid::parse("not-a-guid").is_none());
        assert!(Guid::parse("00020400-0000-0000-C0é0-000000000046").is_none());
        struct Buf([u8; 36], usize);
        impl core::fmt::Write for Buf {
            fn write_str(&mut self, s: &str) -> core::fmt::Result {
                self.0[self.1..self.1 + s.len()].copy_from_slice(s.as_bytes());
                self.1 += s.len();
                Ok(())
            }
        }
        let mut buf = Buf([0; 36], 0);
        core::fmt::write(&mut buf, format_args!("{IID_IDISPATCH}")).unwrap();
        assert_eq!(&buf.0, b"00020400-0000-0000-C000-000000000046");
    }

    #[test]
    #[cfg(windows)]
    fn guid_matches_windows() {
        use core::mem::{align_of, offset_of};
        use windows_sys::core::GUID;
        assert_eq!(size_of::<Guid>(), size_of::<GUID>());
        assert_eq!(align_of::<Guid>(), align_of::<GUID>());
        assert_eq!(offset_of!(Guid, data2), offset_of!(GUID, data2));
        assert_eq!(offset_of!(Guid, data4), offset_of!(GUID, data4));
    }
}
