//! Retained native error text, including dynamic Err.Raise arguments.
use calcvba_rt_abi::native_export::Context;
use calcvba_runtime::bstr::BStr;

pub(super) struct ErrorText {
    pub source: BStr,
    pub description: BStr,
    pub help_file: BStr,
    pub help_context: i32,
}

pub(super) unsafe fn clear(context: *mut Context) {
    // SAFETY: the owning VBA thread alone owns this allocation.
    unsafe {
        let owner = std::mem::replace(&mut (*context).error_owner, std::ptr::null_mut());
        if !owner.is_null() {
            drop(Box::from_raw(owner.cast::<ErrorText>()));
        }
    }
}

/// # Safety
/// The context belongs to the calling VBA thread; text arguments are borrowed
/// BSTRs retained for this call. Flags distinguish omitted and empty arguments.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_raise_dynamic(
    context: *mut Context,
    number: i32,
    source: *const u8,
    source_kind: u32,
    description: *const u8,
    description_kind: u32,
    help_file: *const u8,
    help_file_kind: u32,
    help_context: *const u8,
    help_context_kind: u32,
    flags: u32,
) -> i32 {
    // SAFETY: this reads only calling-thread Win32 state.
    let saved = unsafe { windows_sys::Win32::Foundation::GetLastError() };
    let result = std::panic::catch_unwind(|| {
        // SAFETY: forwarded live BSTR and unique context contract above.
        unsafe {
            unsafe fn text(cell: *const u8, kind: u32, present: bool) -> Result<BStr, i32> {
                if !present {
                    return Ok(BStr::empty());
                }
                let mut value: *mut u16 = std::ptr::null_mut();
                // SAFETY: convert into a private empty BSTR before retaining the error.
                unsafe {
                    super::values::copy((&mut value as *mut *mut u16).cast(), cell, 8, kind)?;
                    let result = BStr::clone_from_raw_bstr(value).map_err(|_| 7);
                    windows_sys::Win32::Foundation::SysFreeString(value);
                    result
                }
            }
            let mut number_context = 0i32;
            if flags & 8 != 0 {
                super::values::copy(
                    (&mut number_context as *mut i32).cast(),
                    help_context,
                    3,
                    help_context_kind,
                )?;
            }
            let text = ErrorText {
                source: text(source, source_kind, flags & 1 != 0)?,
                description: text(description, description_kind, flags & 2 != 0)?,
                help_file: text(help_file, help_file_kind, flags & 4 != 0)?,
                help_context: number_context,
            };
            Ok::<_, i32>(super::services::raise_explicit(
                context, number, &text, flags,
            ))
        }
    });
    // SAFETY: all temporary conversions are gone before restoring LastDllError.
    unsafe { windows_sys::Win32::Foundation::SetLastError(saved) };
    match result {
        Ok(Ok(number)) => number,
        Ok(Err(number)) => number,
        Err(_) => 7,
    }
}
