//! Individual native Declare calls. No VBA instruction or procedure dispatch.
use calcvba_com::windows_ffi_bridge::{self as ffi, FfiArg, FfiReturnType, OwnedLibrary};
use calcvba_runtime::bstr::BStr;
use std::{
    ffi::{CStr, c_void},
    ptr,
};
use windows_sys::Win32::Foundation::{
    FreeLibrary, GetLastError, SetLastError, SysFreeString, SysStringByteLen,
};

const BYREF: u32 = 0x10000;
const WRITE_STRING: u32 = 0x20000;

// The compiler allocates two zeroed pointer words per Declare descriptor.
// Only accepted calls on the owning thread reach this state.
unsafe fn invoke(
    state: *mut usize,
    metadata: *const u32,
    cells: *const *mut u8,
    out: *mut u8,
    last_error: &mut u32,
) -> Result<(), i32> {
    // SAFETY: all spans and cell kinds are compiler-generated, retained image
    // data; their pointees are live for this synchronous foreign call.
    unsafe {
        let result_kind = metadata.read();
        let count = metadata.add(1).read() as usize;
        let library = CStr::from_ptr(
            metadata
                .cast::<u8>()
                .add(metadata.add(2).read() as usize)
                .cast(),
        )
        .to_str()
        .map_err(|_| 53)?;
        let symbol = CStr::from_ptr(
            metadata
                .cast::<u8>()
                .add(metadata.add(3).read() as usize)
                .cast(),
        )
        .to_str()
        .map_err(|_| 453)?;
        let mut address = state.read();
        if address == 0 {
            let owner = OwnedLibrary::load(library).map_err(|_| 53)?;
            address = ffi::get_proc_address(owner.handle(), symbol).map_err(|_| 453)?;
            state.add(1).write(owner.handle());
            state.write(address);
            std::mem::forget(owner);
        }
        let mut args = Vec::with_capacity(count);
        let mut strings = Vec::new();
        for i in 0..count {
            let flags = metadata.add(4 + i).read();
            let kind = flags & 0xffff;
            let cell = cells.add(i).read();
            let arg = if flags & BYREF != 0 {
                FfiArg::Pointer(cell.cast())
            } else {
                match kind {
                    17 => FfiArg::Byte(cell.read()),
                    2 => FfiArg::Integer(cell.cast::<i16>().read()),
                    3 => FfiArg::Long(cell.cast::<i32>().read()),
                    11 => FfiArg::Boolean(cell.cast::<i16>().read()),
                    20 | 6 => FfiArg::LongLong(cell.cast::<i64>().read()),
                    4 => FfiArg::Single(cell.cast::<f32>().read()),
                    5 | 7 => FfiArg::Double(cell.cast::<f64>().read()),
                    8 => {
                        let source = cell.cast::<*mut u16>().read();
                        let ansi = super::values::bstr_to_ansi(source)?;
                        let len = SysStringByteLen(ansi) as usize;
                        let mut buffer = if len == 0 {
                            Vec::new()
                        } else {
                            std::slice::from_raw_parts(ansi.cast::<u8>(), len).to_vec()
                        };
                        SysFreeString(ansi);
                        buffer.push(0);
                        let original = buffer.clone();
                        let address = buffer.as_mut_ptr();
                        strings.push((cell, buffer, original, flags & WRITE_STRING != 0));
                        FfiArg::Pointer(address.cast::<c_void>())
                    }
                    _ => return Err(13),
                }
            };
            args.push(arg);
        }
        let result_type = match result_kind {
            0 => FfiReturnType::Void,
            17 => FfiReturnType::Byte,
            2 => FfiReturnType::Integer,
            3 => FfiReturnType::Long,
            11 => FfiReturnType::Boolean,
            20 | 6 => FfiReturnType::LongLong,
            4 => FfiReturnType::Single,
            5 | 7 => FfiReturnType::Double,
            _ => return Err(13),
        };
        let (result, error) =
            ffi::invoke_stdcall_with_last_error(address, &args, result_type).map_err(|_| 5)?;
        *last_error = error;
        for (cell, buffer, original, writable) in strings {
            if writable && buffer != original {
                let units = ffi::utf16_from_ansi(&buffer[..buffer.len() - 1]);
                let value = BStr::from_utf16_units(&units)
                    .map_err(|_| 7)?
                    .into_raw_bstr();
                SysFreeString(cell.cast::<*mut u16>().read());
                cell.cast::<*mut u16>().write(value);
            }
        }
        if result_kind != 0 {
            let width = match result_kind {
                17 => 1,
                2 | 11 => 2,
                3 | 4 => 4,
                _ => 8,
            };
            ptr::copy_nonoverlapping((&result as *const i64).cast::<u8>(), out, width);
        }
        Ok(())
    }
}

/// # Safety
/// The compiler supplies the retained descriptor, two-word state, argument cell
/// vector and private result cell, after the public owning-thread check.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_declare(
    state: *mut usize,
    metadata: *const u32,
    cells: *const *mut u8,
    out: *mut u8,
) -> i32 {
    // SAFETY: these OS calls access only the current thread's last-error slot.
    let mut error = unsafe { GetLastError() };
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        // SAFETY: forwarded compiler-proven carrier and state contract above.
        unsafe { invoke(state, metadata, cells, out, &mut error) }
    }));
    let status = match result {
        Ok(Ok(())) => 0,
        Ok(Err(code)) => code,
        Err(_) => 7,
    };
    // SAFETY: argument/CIF cleanup completed before restoring the call's error.
    unsafe { SetLastError(error) };
    status
}

/// # Safety
/// state is this image's two-word import state, on its owning thread at shutdown.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_declare_clear(state: *mut usize) -> i32 {
    // SAFETY: take only our recorded LoadLibrary reference and clear it first.
    unsafe {
        let error = GetLastError();
        let module = state.add(1).read();
        state.write(0);
        state.add(1).write(0);
        if module != 0 {
            FreeLibrary(module as *mut c_void);
        }
        SetLastError(error);
    }
    0
}
