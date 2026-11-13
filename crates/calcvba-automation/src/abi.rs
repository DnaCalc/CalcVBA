//! Byte layout of the Automation `VARIANT` on the build target.
//!
//! ```text
//! offset 0       vt (for VT_DECIMAL: DECIMAL.wReserved)
//! offset 2, 4, 6 wReserved1..3 (for VT_DECIMAL: scale, sign, Hi32)
//! offset 8       union: llVal, dblVal, bstrVal, punkVal, parray, byref, ...; DECIMAL.Lo64;
//!                BRECORD.pvRecord
//! offset 8 + P   BRECORD.pRecInfo (P = pointer size)
//! ```

use core::mem::size_of;

const POINTER: usize = size_of::<*const ()>();

/// Offset of the `vt` type tag.
pub const VT_OFFSET: usize = 0;
/// Offset of the value union (`llVal`, `bstrVal`, `parray`, `byref`, `pvRecord`, ...).
pub const PAYLOAD_OFFSET: usize = 8;
/// Offset of `BRECORD.pRecInfo`, the second pointer of a `VT_RECORD` payload.
pub const RECINFO_OFFSET: usize = PAYLOAD_OFFSET + POINTER;
/// Size of the value union: eight bytes for 64-bit scalars, or `BRECORD`'s two pointers.
pub const PAYLOAD_SIZE: usize = if 2 * POINTER > 8 { 2 * POINTER } else { 8 };
/// Size of a `VARIANT`: 24 bytes on 64-bit targets, 16 where pointers are 32-bit.
pub const VARIANT_SIZE: usize = PAYLOAD_OFFSET + PAYLOAD_SIZE;
/// Alignment of a `VARIANT` (its 64-bit members).
pub const VARIANT_ALIGN: usize = 8;

/// `VT_BYREF`: the payload is a pointer to storage of the remaining type.
pub const VT_BYREF: u16 = 0x4000;
/// `VT_ARRAY`: the payload is a `SAFEARRAY*` of the remaining element type.
pub const VT_ARRAY: u16 = 0x2000;
/// Mask of the base type bits under `VT_ARRAY` / `VT_BYREF`.
pub const VT_TYPEMASK: u16 = 0x0fff;

const _: () = assert!(VARIANT_SIZE == 8 + 2 * POINTER || (POINTER == 4 && VARIANT_SIZE == 16));
const _: () = assert!(VARIANT_SIZE.is_multiple_of(VARIANT_ALIGN));

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    #[cfg(target_pointer_width = "64")]
    fn sixty_four_bit_layout() {
        assert_eq!(VARIANT_SIZE, 24);
        assert_eq!(PAYLOAD_SIZE, 16);
        assert_eq!(RECINFO_OFFSET, 16);
    }

    #[test]
    #[cfg(target_pointer_width = "32")]
    fn thirty_two_bit_layout() {
        assert_eq!(VARIANT_SIZE, 16);
        assert_eq!(PAYLOAD_SIZE, 8);
        assert_eq!(RECINFO_OFFSET, 12);
    }

    /// Every constant against the platform header's own definition.
    #[test]
    #[cfg(windows)]
    fn matches_windows_variant() {
        use core::mem::{align_of, offset_of};
        use windows_sys::Win32::System::Variant::{VARIANT, VARIANT_0_0};
        assert_eq!(size_of::<VARIANT>(), VARIANT_SIZE);
        assert_eq!(align_of::<VARIANT>(), VARIANT_ALIGN);
        assert_eq!(offset_of!(VARIANT_0_0, vt), VT_OFFSET);
        assert_eq!(offset_of!(VARIANT_0_0, Anonymous), PAYLOAD_OFFSET);
        // BRECORD is the union arm that sets the payload size.
        let brecord = size_of::<windows_sys::Win32::System::Variant::VARIANT_0_0_0_0>();
        assert_eq!(brecord, 2 * POINTER);
        assert_eq!(PAYLOAD_OFFSET + POINTER, RECINFO_OFFSET);
        // DECIMAL overlays the whole header and the first eight payload bytes.
        assert_eq!(size_of::<windows_sys::Win32::Foundation::DECIMAL>(), 16);
        assert_eq!(VT_BYREF, windows_sys::Win32::System::Variant::VT_BYREF);
        assert_eq!(VT_ARRAY, windows_sys::Win32::System::Variant::VT_ARRAY);
        assert_eq!(
            VT_TYPEMASK,
            windows_sys::Win32::System::Variant::VT_TYPEMASK
        );
    }
}
