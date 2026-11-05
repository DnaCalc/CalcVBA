//! Linked once by CalcVBA developers, then finished and shipped. These exports
//! retain helpers through linking; the finisher hides them in the final image.
//! The runtime has no loader, dispatcher, interpreter or JIT dependency.
#[cfg(windows)]
mod errors;
#[cfg(windows)]
mod foreign;
#[cfg(windows)]
mod services;
#[cfg(windows)]
mod values;

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
    use calcvba_rt_abi::native_export::{Context, WRONG_THREAD};
    use std::sync::atomic::{AtomicI32, AtomicPtr, Ordering};
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
    static LAST_ERROR: AtomicI32 = AtomicI32::new(0);

    #[unsafe(no_mangle)]
    pub extern "C" fn cvb_enter_scalar() -> i32 {
        // SAFETY: OS pseudo handles are valid in this process. The retained
        // real handle identifies the original thread even after termination.
        unsafe {
            let saved_error = GetLastError();
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
                SetLastError(saved_error);
                0
            } else {
                LAST_ERROR.store(WRONG_THREAD, Ordering::Release);
                SetLastError(WRONG_THREAD as u32);
                WRONG_THREAD
            }
        }
    }

    /// # Safety
    /// context belongs to this DLL and the accepted owning thread; source is
    /// immutable UTF-8 image data that remains live until the DLL is unloaded.
    #[unsafe(no_mangle)]
    pub unsafe extern "C" fn cvb_record_error(
        context: *mut Context,
        code: i32,
        source: *const u8,
        len: usize,
    ) -> i32 {
        // SAFETY: error-owner destruction must not change Err.LastDllError.
        let saved = unsafe { GetLastError() };
        // SAFETY: emitted wrappers call only after the thread check succeeds.
        unsafe {
            let explicit = (*context).error_number == code && code != 0;
            if !explicit || (*context).custom_error & 0x100 == 0 {
                super::errors::clear(context);
            }
            if !explicit || (*context).custom_error & 1 == 0 {
                (*context).error_source = source;
                (*context).error_source_len = len;
            }
            if !explicit {
                (*context).custom_error = 0;
            }
            (*context).error_number = code;
        }
        LAST_ERROR.store(code, Ordering::Release);
        // SAFETY: all owner cleanup has completed on this calling thread.
        unsafe { SetLastError(saved) };
        calcvba_rt_abi::native_export::hresult(code)
    }

    /// # Safety
    /// context is uniquely owned by this calling VBA thread. Text spans immutable
    /// UTF-8 image data; flags distinguishes omitted text from explicit empty text.
    #[unsafe(no_mangle)]
    pub unsafe extern "C" fn cvb_raise(
        context: *mut Context,
        number: i32,
        source: *const u8,
        source_len: usize,
        description: *const u8,
        description_len: usize,
        flags: u32,
    ) -> i32 {
        let number = if number == 0 { 5 } else { number };
        // SAFETY: generated code retains this unique context and both constants.
        unsafe {
            super::errors::clear(context);
            (*context).error_number = number;
            (*context).error_source = source;
            (*context).error_source_len = source_len;
            (*context).error_description = description;
            (*context).error_description_len = description_len;
            (*context).custom_error = flags;
        }
        number
    }

    #[unsafe(no_mangle)]
    pub extern "C" fn cvb_last_dll_error() -> i32 {
        // SAFETY: GetLastError reads only calling-thread state.
        unsafe { GetLastError() as i32 }
    }

    /// # Safety
    /// context is this image's state; nonnull outputs are writable BSTR* cells.
    /// Returned BSTRs belong to the caller (SysFreeString).
    #[unsafe(no_mangle)]
    pub unsafe extern "C" fn cvb_get_last_error(
        context: *const Context,
        source: *mut *mut u16,
        description: *mut *mut u16,
    ) -> i32 {
        // SAFETY: GetLastError reads only the current thread state.
        let saved_error = unsafe { GetLastError() };
        let wrong = cvb_enter_scalar() != 0;
        if wrong {
            return WRONG_THREAD;
        }
        let code = LAST_ERROR.load(Ordering::Acquire);
        let text = if wrong || code == WRONG_THREAD {
            "CalcVBA native export"
        } else if code == 0 {
            ""
        } else {
            // SAFETY: only the owner reads its source. Other threads only set
            // the atomic code and never access or mutate these fields.
            unsafe { image_text((*context).error_source, (*context).error_source_len) }
        };
        // SAFETY: the owner alone accesses context fields; foreign calls returned above.
        let explicit_description =
            unsafe { code == (*context).error_number && (*context).custom_error & 2 != 0 };
        let message = if code == 0 {
            ""
        } else if code == WRONG_THREAD {
            "Native VBA calls must run on the owning thread"
        } else if explicit_description {
            // SAFETY: explicit text is retained immutable UTF-8 image data.
            unsafe {
                image_text(
                    (*context).error_description,
                    (*context).error_description_len,
                )
            }
        } else {
            calcvba_runtime::vba_error::default_error_message(code)
        };
        // SAFETY: caller provided live optional output cells, never inputs.
        unsafe {
            if !source.is_null() {
                let value = if code == (*context).error_number
                    && (*context).custom_error & 0x101 == 0x101
                {
                    (&*(*context).error_owner.cast::<super::errors::ErrorText>())
                        .source
                        .clone()
                } else {
                    calcvba_runtime::bstr::BStr::from(text)
                };
                source.write(value.into_raw_bstr());
            }
            if !description.is_null() {
                let value = if code == (*context).error_number
                    && (*context).custom_error & 0x102 == 0x102
                {
                    (&*(*context).error_owner.cast::<super::errors::ErrorText>())
                        .description
                        .clone()
                } else {
                    calcvba_runtime::bstr::BStr::from(message)
                };
                description.write(value.into_raw_bstr());
            }
            SetLastError(if wrong {
                WRONG_THREAD as u32
            } else {
                saved_error
            });
        }
        code
    }

    /// # Safety
    /// The context belongs to this image. Nonnull outputs are writable BSTR*/Long*.
    #[unsafe(no_mangle)]
    pub unsafe extern "C" fn cvb_get_error_help(
        context: *const Context,
        help_file: *mut *mut u16,
        help_context: *mut i32,
    ) -> i32 {
        // SAFETY: read only the current thread's Win32 error state.
        let saved = unsafe { GetLastError() };
        if cvb_enter_scalar() != 0 {
            return WRONG_THREAD;
        }
        let code = LAST_ERROR.load(Ordering::Acquire);
        // SAFETY: only the accepted owning thread reads retained error owners.
        unsafe {
            let text = if code != 0
                && code == (*context).error_number
                && (*context).custom_error & 0x100 != 0
            {
                Some(&*(*context).error_owner.cast::<super::errors::ErrorText>())
            } else {
                None
            };
            if !help_file.is_null() {
                help_file.write(
                    text.map(|t| t.help_file.clone())
                        .unwrap_or_else(calcvba_runtime::bstr::BStr::empty)
                        .into_raw_bstr(),
                );
            }
            if !help_context.is_null() {
                help_context.write(text.map_or(0, |t| t.help_context));
            }
            SetLastError(saved);
        }
        code
    }

    unsafe fn image_text<'a>(pointer: *const u8, len: usize) -> &'a str {
        if len == 0 {
            return "";
        }
        // SAFETY: caller supplies a retained immutable UTF-8 span of len bytes.
        unsafe { std::str::from_utf8_unchecked(std::slice::from_raw_parts(pointer, len)) }
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
