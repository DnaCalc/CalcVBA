//! The thread's `IErrorInfo` (`GetErrorInfo` / `SetErrorInfo`) as a COM call's rich error, on
//! every Windows target.
#![cfg(target_os = "windows")]

use crate::windows_invoke::{ComInvokeExceptionInfo, bstr_to_string_and_free};
use std::ffi::c_void;

// ── IErrorInfo + Get/SetErrorInfo ──
//
// windows-sys 0.59 exposes GetErrorInfo/SetErrorInfo as free functions over an
// untyped `*mut c_void` IErrorInfo; it does not expose the IErrorInfo vtable, so
// we declare the slots we read here (the four accessors after IUnknown).

#[repr(C)]
struct RawIErrorInfoVtbl {
    query_interface: unsafe extern "system" fn(
        this: *mut c_void,
        iid: *const windows_sys::core::GUID,
        ppv: *mut *mut c_void,
    ) -> i32,
    add_ref: unsafe extern "system" fn(this: *mut c_void) -> u32,
    release: unsafe extern "system" fn(this: *mut c_void) -> u32,
    get_guid:
        unsafe extern "system" fn(this: *mut c_void, pguid: *mut windows_sys::core::GUID) -> i32,
    get_source:
        unsafe extern "system" fn(this: *mut c_void, pbstr: *mut windows_sys::core::BSTR) -> i32,
    get_description:
        unsafe extern "system" fn(this: *mut c_void, pbstr: *mut windows_sys::core::BSTR) -> i32,
    get_help_file:
        unsafe extern "system" fn(this: *mut c_void, pbstr: *mut windows_sys::core::BSTR) -> i32,
    get_help_context: unsafe extern "system" fn(this: *mut c_void, pdw: *mut u32) -> i32,
}

#[repr(C)]
struct RawIErrorInfo {
    vtbl: *const RawIErrorInfoVtbl,
}

// SAFETY: oleaut32 exports transcribed with the stdcall `system` ABI.
// GetErrorInfo retrieves (and clears) the thread's pending IErrorInfo into an
// out-pointer; SetErrorInfo installs (or clears, with a null pointer) it. Each
// call site documents the pointer invariants it upholds.
unsafe extern "system" {
    fn GetErrorInfo(dwreserved: u32, pperrinfo: *mut *mut c_void) -> i32;
    fn SetErrorInfo(dwreserved: u32, perrinfo: *mut c_void) -> i32;
}

/// Clear any pending thread error info so a stale `IErrorInfo` from an earlier
/// call cannot be mis-attributed to this one (mirrors what the COM runtime does
/// before a server entry point).
pub(crate) fn clear_thread_error_info() {
    // SAFETY: SetErrorInfo(0, NULL) is the documented form that clears the
    // thread's current error object; it dereferences no pointer.
    unsafe {
        let _ = SetErrorInfo(0, std::ptr::null_mut());
    }
}

/// Retrieve the rich error the failing call installed via `SetErrorInfo`, if
/// any, and project it into the shared [`ComInvokeExceptionInfo`]. The
/// `IErrorInfo` is Released here; its Source/Description/HelpFile BSTRs are
/// transferred to us and freed via [`bstr_to_string_and_free`].
pub(crate) fn take_error_info() -> Option<ComInvokeExceptionInfo> {
    let mut perrinfo: *mut c_void = std::ptr::null_mut();
    // SAFETY: `perrinfo` is a valid out-pointer; GetErrorInfo writes either null
    // (no pending error, returns S_FALSE) or a single owned IErrorInfo reference.
    let hr = unsafe { GetErrorInfo(0, &mut perrinfo) };
    if hr < 0 || perrinfo.is_null() {
        return None;
    }
    // SAFETY: `perrinfo` is a non-null IErrorInfo carrying one reference we now
    // own; its first field is its vtable (RawIErrorInfo prefix). We call the
    // accessors, then Release exactly once below, and do not use the pointer
    // afterward.
    let info = unsafe {
        let err = perrinfo.cast::<RawIErrorInfo>();
        let vtbl = &*(*err).vtbl;

        let mut source_bstr: windows_sys::core::BSTR = std::ptr::null_mut();
        let source = if (vtbl.get_source)(perrinfo, &mut source_bstr) >= 0 {
            bstr_to_string_and_free(source_bstr)
        } else {
            None
        };

        let mut description_bstr: windows_sys::core::BSTR = std::ptr::null_mut();
        let description = if (vtbl.get_description)(perrinfo, &mut description_bstr) >= 0 {
            bstr_to_string_and_free(description_bstr)
        } else {
            None
        };

        let mut help_file_bstr: windows_sys::core::BSTR = std::ptr::null_mut();
        let help_file = if (vtbl.get_help_file)(perrinfo, &mut help_file_bstr) >= 0 {
            bstr_to_string_and_free(help_file_bstr)
        } else {
            None
        };

        let mut help_context: u32 = 0;
        let help_context =
            if (vtbl.get_help_context)(perrinfo, &mut help_context) >= 0 && help_context != 0 {
                Some(help_context)
            } else {
                None
            };

        (vtbl.release)(perrinfo);

        ComInvokeExceptionInfo {
            source,
            description,
            help_file,
            help_context,
            scode: None,
            wcode: None,
        }
    };

    // An IErrorInfo with nothing readable is no richer than the bare HRESULT.
    if info.source.is_none()
        && info.description.is_none()
        && info.help_file.is_none()
        && info.help_context.is_none()
    {
        None
    } else {
        Some(info)
    }
}
