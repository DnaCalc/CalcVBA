//! Portable Automation cells used at the Linux native boundary.
use calcvba_automation::variant::VariantCore;
use calcvba_runtime::{ObjectRef, Variant, bstr::BStr, object_ref::RawRuntimeIUnknown};

pub(super) fn boundary(work: impl FnOnce() -> Result<(), i32>) -> i32 {
    let saved = super::thread::last_error();
    let status = match std::panic::catch_unwind(std::panic::AssertUnwindSafe(work)) {
        Ok(Ok(())) => 0,
        Ok(Err(code)) => code,
        Err(_) => 7,
    };
    super::thread::set_last_error(saved);
    status
}

fn width(kind: u32) -> Result<usize, i32> {
    match kind {
        17 => Ok(1),
        2 | 11 => Ok(2),
        3 | 4 => Ok(4),
        5..=8 | 20 => Ok(8),
        _ => Err(13),
    }
}

pub(super) unsafe fn read(
    _context: *mut calcvba_rt_abi::native_export::Context,
    cell: *const u8,
    kind: u32,
) -> Result<Variant, i32> {
    // SAFETY: the caller retains a live cell of the exact admitted kind.
    unsafe {
        if kind == 12 {
            return Variant::from_trusted_cell(&*cell.cast::<VariantCore>()).map_err(|_| 13);
        }
        if kind == 9 {
            let pointer = cell.cast::<*mut RawRuntimeIUnknown>().read();
            return Ok(ObjectRef::from_raw_iunknown_addref(pointer)
                .map_or_else(Variant::nothing, Variant::from_object_ref));
        }
        if kind == 8 {
            return BStr::clone_from_raw_bstr(cell.cast::<*mut u16>().read())
                .map(Variant::from_string)
                .map_err(|_| 7);
        }
        let size = width(kind)?;
        let mut bytes = Variant::empty().to_wire_bytes();
        bytes[..2].copy_from_slice(&(kind as u16).to_le_bytes());
        std::ptr::copy_nonoverlapping(cell, bytes.as_mut_ptr().add(8), size);
        Variant::from_wire_bytes(bytes).map_err(|_| 13)
    }
}

pub(super) unsafe fn write(
    _context: *mut calcvba_rt_abi::native_export::Context,
    cell: *mut u8,
    kind: u32,
    value: &Variant,
) -> Result<(), i32> {
    if cell.is_null() || kind == 0 {
        return Ok(());
    }
    // SAFETY: the caller owns the initialized writable cell of the exact kind.
    unsafe {
        if kind == 9 {
            if value.vtype() != calcvba_runtime::VarType::Object {
                return Err(13);
            }
            let object = value
                .as_object_ref()
                .map_or(std::ptr::null_mut(), |o| o.into_raw_iunknown());
            let previous = cell.cast::<*mut RawRuntimeIUnknown>().replace(object);
            drop(ObjectRef::from_raw_iunknown_owned(previous));
            return Ok(());
        }
        if kind == 12 {
            *cell.cast::<Variant>() = value.try_clone().map_err(|_| 7)?;
            return Ok(());
        }
        let converted;
        let value = if u32::from(value.core().raw_vt()) == kind {
            value
        } else {
            converted = calcvba_eval::arith::coerce_for_byref_target(value, kind as u16)
                .map_err(|error| error.code)?;
            &converted
        };
        if kind == 8 {
            let string = value.try_as_bstr().map_err(|_| 7)?.ok_or(13)?;
            let previous = cell.cast::<*mut u16>().replace(string.into_raw_bstr());
            drop(BStr::from_raw_bstr(previous));
        } else {
            let size = width(kind)?;
            std::ptr::copy_nonoverlapping(value.to_wire_bytes().as_ptr().add(8), cell, size);
        }
        Ok(())
    }
}

pub(super) unsafe fn copy(dst: *mut u8, src: *const u8, to: u32, from: u32) -> Result<(), i32> {
    // SAFETY: read produces a separate owner before the destination is changed.
    unsafe {
        let value = read(std::ptr::null_mut(), src, from)?;
        write(std::ptr::null_mut(), dst, to, &value)
    }
}

/// # Safety
/// The source and destination are initialized cells of their respective kinds.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_value_copy(dst: *mut u8, src: *const u8, to: u32, from: u32) -> i32 {
    // SAFETY: forwards the initialized cell contract.
    boundary(|| unsafe { copy(dst, src, to, from) })
}

/// # Safety
/// The caller owns this initialized writable cell of kind.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_value_clear(cell: *mut u8, kind: u32) -> i32 {
    boundary(|| {
        // SAFETY: transfer and release the owner in this cell exactly once.
        unsafe {
            match kind {
                8 => drop(BStr::from_raw_bstr(
                    cell.cast::<*mut u16>().replace(std::ptr::null_mut()),
                )),
                9 => drop(ObjectRef::from_raw_iunknown_owned(
                    cell.cast::<*mut RawRuntimeIUnknown>()
                        .replace(std::ptr::null_mut()),
                )),
                12 => *cell.cast::<Variant>() = Variant::empty(),
                _ => {}
            }
        }
        Ok(())
    })
}

/// # Safety
/// The input is null or a terminated UTF-16 string. The output is a live BSTR cell.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_string_import(src: *const u8, encoding: u32, out: *mut u8) -> i32 {
    boundary(|| {
        if encoding != 1 {
            return Err(5);
        }
        let mut units = &[][..];
        if !src.is_null() {
            let src = src.cast::<u16>();
            let mut length = 0;
            // SAFETY: the foreign caller supplies a terminated UTF-16 string.
            unsafe {
                while src.add(length).read() != 0 {
                    length += 1;
                }
                units = std::slice::from_raw_parts(src, length);
            }
        }
        // SAFETY: out is an initialized BSTR destination.
        unsafe {
            let value = Variant::from_utf16_units(units);
            write(std::ptr::null_mut(), out, 8, &value)
        }
    })
}

/// # Safety
/// These are initialized BSTR cells. The result may alias either input.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_string_concat(out: *mut u8, a: *const u8, b: *const u8) -> i32 {
    boundary(|| {
        // SAFETY: obtain independent copies before updating out.
        unsafe {
            let a = BStr::clone_from_raw_bstr(a.cast::<*mut u16>().read()).map_err(|_| 7)?;
            let b = BStr::clone_from_raw_bstr(b.cast::<*mut u16>().read()).map_err(|_| 7)?;
            let units = [a.payload_units(), b.payload_units()].concat();
            write(
                std::ptr::null_mut(),
                out,
                8,
                &Variant::from_utf16_units(&units),
            )
        }
    })
}

/// # Safety
/// Input is a live BSTR cell, out is a writable Long.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_string_len(input: *const *mut u16, out: *mut i32) -> i32 {
    boundary(|| {
        // SAFETY: both cells are live for this call.
        unsafe {
            let value = BStr::clone_from_raw_bstr(input.read()).map_err(|_| 7)?;
            out.write(i32::try_from(value.payload_units().len()).map_err(|_| 6)?);
        }
        Ok(())
    })
}

/// # Safety
/// A Variant input is live. The output is a writable VBA Long.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_value_type(out: *mut i32, src: *const VariantCore, kind: u32) -> i32 {
    // SAFETY: the generated caller retains both exact typed cells for this call.
    unsafe {
        let mut vt = if kind == 12 {
            (*src).raw_vt()
        } else {
            kind as u16
        } & !0x4000;
        if vt & 0xfff == 13 {
            vt = (vt & 0xf000) | 9;
        }
        out.write(i32::from(vt));
    }
    0
}

/// # Safety
/// The caller owns a BSTR returned by this library, or passes null.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_free_bstr(value: *mut u16) {
    // SAFETY: release the exact allocation whose ownership the caller transferred.
    unsafe { drop(BStr::from_raw_bstr(value)) };
}
