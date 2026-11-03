//! Linked once by CalcVBA developers, then finished and shipped. These exports
//! retain helpers through linking; the finisher hides them in the final image.
//! The runtime has no loader, dispatcher, interpreter or JIT dependency.

#[unsafe(no_mangle)]
pub extern "C" fn cvb_identity_i32(value: i32) -> i32 {
    calcvba_runtime::Variant::from_i32(value)
        .as_i32()
        .unwrap_or(0)
}

/// # Safety
/// out is a live aligned, writable, nonaliasing f64 result for this call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_abs_f64(
    _context: *mut std::ffi::c_void,
    value: f64,
    out: *mut f64,
) -> i32 {
    // SAFETY: the generated caller supplies its private result stack slot.
    unsafe {
        out.write(calcvba_lib::abs_f64(value));
    }
    0
}
/// # Safety
/// out is a live aligned, writable, nonaliasing f64 result for this call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_sqr_f64(
    _context: *mut std::ffi::c_void,
    value: f64,
    out: *mut f64,
) -> i32 {
    match calcvba_lib::sqr_f64(value) {
        Ok(value) => {
            // SAFETY: the generated caller supplies its private result slot.
            unsafe {
                out.write(value);
            }
            0
        }
        Err(error) => error.code,
    }
}

/// # Safety
/// out is a live aligned, writable, nonaliasing i32 result for this call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_long_from_f64(
    _context: *mut std::ffi::c_void,
    value: f64,
    out: *mut i32,
) -> i32 {
    match calcvba_eval::typed::long_from_f64(value) {
        Ok(value) => {
            // SAFETY: the generated caller supplies its private result slot.
            unsafe { out.write(value) };
            0
        }
        Err(error) => error.code,
    }
}

const fn abi_hash(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf29ce484222325u64;
    let mut at = 0;
    while at < bytes.len() {
        if bytes[at] != b'\r' {
            hash = (hash ^ bytes[at] as u64).wrapping_mul(0x100000001b3);
        }
        at += 1;
    }
    hash
}
#[unsafe(export_name = "cvb_template_abi")]
pub static CVB_TEMPLATE_ABI: u64 =
    abi_hash(include_bytes!("../../../assets/native/helpers-v1.txt"));

#[cfg(windows)]
mod thread {
    use std::sync::atomic::{AtomicPtr, Ordering};
    use windows_sys::Win32::{
        Foundation::{
            CloseHandle, DUPLICATE_SAME_ACCESS, DuplicateHandle, GetLastError, SetLastError,
            WAIT_TIMEOUT,
        },
        System::{
            LibraryLoader::DisableThreadLibraryCalls,
            SystemServices::{DLL_PROCESS_ATTACH, DLL_PROCESS_DETACH},
            Threading::{
                GetCurrentProcess, GetCurrentThread, GetCurrentThreadId, GetThreadId,
                WaitForSingleObject,
            },
        },
    };
    static OWNER: AtomicPtr<std::ffi::c_void> = AtomicPtr::new(std::ptr::null_mut());

    #[unsafe(no_mangle)]
    pub extern "C" fn cvb_enter_scalar() -> i32 {
        // SAFETY: OS pseudo handles are valid in this process. The retained
        // real handle identifies the original thread even after termination.
        unsafe {
            let mut owner = OWNER.load(Ordering::Acquire);
            if owner.is_null() {
                let mut candidate = std::ptr::null_mut();
                let process = GetCurrentProcess();
                if DuplicateHandle(
                    process,
                    GetCurrentThread(),
                    process,
                    &mut candidate,
                    0,
                    0,
                    DUPLICATE_SAME_ACCESS,
                ) == 0
                {
                    return GetLastError().max(1) as i32;
                }
                match OWNER.compare_exchange(
                    std::ptr::null_mut(),
                    candidate,
                    Ordering::AcqRel,
                    Ordering::Acquire,
                ) {
                    Ok(_) => owner = candidate,
                    Err(existing) => {
                        CloseHandle(candidate);
                        owner = existing;
                    }
                }
            }
            // A zero timeout only tests thread state; it never waits. Do not
            // rely on an ID remaining unique after its thread terminates.
            if WaitForSingleObject(owner, 0) == WAIT_TIMEOUT
                && GetThreadId(owner) == GetCurrentThreadId()
            {
                0
            } else {
                SetLastError(1444);
                1444
            }
        }
    }
    #[unsafe(no_mangle)]
    unsafe extern "system" fn DllMain(
        module: windows_sys::Win32::Foundation::HMODULE,
        reason: u32,
        _reserved: *mut std::ffi::c_void,
    ) -> i32 {
        // SAFETY: the loader serializes attach/detach. No VBA code, allocation,
        // image loading or callbacks execute here; only kernel handle cleanup.
        unsafe {
            if reason == DLL_PROCESS_ATTACH {
                DisableThreadLibraryCalls(module);
            }
            if reason == DLL_PROCESS_DETACH {
                let owner = OWNER.swap(std::ptr::null_mut(), Ordering::AcqRel);
                if !owner.is_null() {
                    CloseHandle(owner);
                }
            }
        }
        1
    }
}
