//! The oleaut32 and ole32 calls the Automation values need, with the platform's signatures over
//! raw pointers: the real functions on Windows, and the same contracts and memory layout
//! elsewhere (where this crate is the only allocator of BSTRs, SAFEARRAYs and task memory).
#![allow(clippy::missing_safety_doc)]

#[cfg(not(target_os = "windows"))]
use crate::{Variant, safe_array::SafeArray};
use core::ffi::c_void;

const S_OK: i32 = 0;
const E_INVALIDARG: i32 = 0x8007_0057_u32 as i32;
#[cfg(not(target_os = "windows"))]
const E_OUTOFMEMORY: i32 = 0x8007_000E_u32 as i32;

/// `SysAllocStringByteLen`: a BSTR holding `len` bytes copied from `source` (zeroed when
/// `source` is null).
pub unsafe fn sys_alloc_string_byte_len(source: *const u8, len: u32) -> *mut u16 {
    let bytes = if source.is_null() {
        vec![0u8; len as usize]
    } else {
        // SAFETY: the caller supplies `len` readable bytes at `source`.
        unsafe { core::slice::from_raw_parts(source, len as usize) }.to_vec()
    };
    crate::bstr::BStr::from_bytes(&bytes)
        .map(crate::bstr::BStr::into_raw_bstr)
        .unwrap_or(core::ptr::null_mut())
}

/// `SysFreeString` (null is allowed).
pub unsafe fn sys_free_string(bstr: *mut u16) {
    // SAFETY: the caller transfers a BSTR from this allocator; the owner frees it.
    drop(unsafe { crate::bstr::BStr::from_raw_bstr(bstr) });
}

/// `SysStringByteLen` (0 for null).
pub unsafe fn sys_string_byte_len(bstr: *const u16) -> u32 {
    if bstr.is_null() {
        return 0;
    }
    // SAFETY: a live BSTR's byte length is the u32 just before its payload.
    unsafe { bstr.cast::<u8>().sub(4).cast::<u32>().read_unaligned() }
}

/// Header of a task-memory block off Windows: its size, so `co_task_mem_free` needs none.
#[cfg(not(target_os = "windows"))]
const TASK_HEADER: usize = 16;

/// `CoTaskMemAlloc`: `size` bytes (uninitialised), freed with [`co_task_mem_free`].
pub unsafe fn co_task_mem_alloc(size: usize) -> *mut c_void {
    #[cfg(target_os = "windows")]
    {
        // SAFETY: plain allocation.
        unsafe { windows_sys::Win32::System::Com::CoTaskMemAlloc(size) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        let Ok(layout) = std::alloc::Layout::from_size_align(TASK_HEADER + size, 16) else {
            return core::ptr::null_mut();
        };
        // SAFETY: a non-zero layout.
        let base = unsafe { std::alloc::alloc(layout) };
        if base.is_null() {
            return core::ptr::null_mut();
        }
        // SAFETY: the header is the first TASK_HEADER bytes of the block.
        unsafe {
            base.cast::<usize>().write(size);
            base.add(TASK_HEADER).cast()
        }
    }
}

/// `CoTaskMemFree` (null is allowed).
pub unsafe fn co_task_mem_free(memory: *mut c_void) {
    if memory.is_null() {
        return;
    }
    #[cfg(target_os = "windows")]
    {
        // SAFETY: the caller transfers a CoTaskMemAlloc block.
        unsafe { windows_sys::Win32::System::Com::CoTaskMemFree(memory) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: a block from `co_task_mem_alloc`, whose header precedes it.
        unsafe {
            let base = memory.cast::<u8>().sub(TASK_HEADER);
            let size = base.cast::<usize>().read();
            let layout = std::alloc::Layout::from_size_align_unchecked(TASK_HEADER + size, 16);
            std::alloc::dealloc(base, layout);
        }
    }
}

/// `VariantClear`: releases what the cell owns and leaves it `VT_EMPTY`.
pub unsafe fn variant_clear(cell: *mut c_void) -> i32 {
    if cell.is_null() {
        return E_INVALIDARG;
    }
    #[cfg(target_os = "windows")]
    {
        // SAFETY: a live VARIANT.
        unsafe { windows_sys::Win32::System::Variant::VariantClear(cell.cast()) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: a live cell this crate's types own; replacing it drops the old value.
        unsafe { cell.cast::<Variant>().replace(Variant::empty()) };
        S_OK
    }
}

/// `VariantCopyInd`: a copy of `source` (references read through) into the empty cell `dst`.
pub unsafe fn variant_copy_ind(dst: *mut c_void, source: *const c_void) -> i32 {
    if dst.is_null() || source.is_null() {
        return E_INVALIDARG;
    }
    #[cfg(target_os = "windows")]
    {
        // SAFETY: live VARIANTs.
        unsafe { windows_sys::Win32::System::Variant::VariantCopyInd(dst.cast(), source.cast()) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: `source` is a live cell; cloning reads a reference's target.
        match unsafe { &*source.cast::<Variant>() }.try_clone() {
            Ok(copy) => {
                // SAFETY: `dst` is an empty (owning nothing) writable cell.
                unsafe { dst.cast::<Variant>().write(copy) };
                S_OK
            }
            Err(_) => E_OUTOFMEMORY,
        }
    }
}

/// `SafeArrayCopy`: a deep copy of the descriptor `array` in `*out` (null for null).
pub unsafe fn safe_array_copy(array: *mut c_void, out: *mut *mut c_void) -> i32 {
    if out.is_null() {
        return E_INVALIDARG;
    }
    #[cfg(target_os = "windows")]
    {
        // SAFETY: a live descriptor and writable out slot.
        unsafe { windows_sys::Win32::System::Ole::SafeArrayCopy(array.cast(), out.cast()) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: a live descriptor the caller keeps.
        match unsafe { SafeArray::try_clone_from_raw_safearray(array, None) } {
            Ok(copy) => {
                // SAFETY: writable out slot.
                unsafe {
                    out.write(copy.map_or(core::ptr::null_mut(), SafeArray::into_raw_safearray_ptr))
                };
                S_OK
            }
            Err(_) => E_OUTOFMEMORY,
        }
    }
}

/// `SafeArrayDestroy` (null is allowed).
pub unsafe fn safe_array_destroy(array: *mut c_void) -> i32 {
    if array.is_null() {
        return S_OK;
    }
    #[cfg(target_os = "windows")]
    {
        // SAFETY: the caller transfers a live descriptor.
        unsafe { windows_sys::Win32::System::Ole::SafeArrayDestroy(array.cast()) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: the caller transfers a live descriptor this crate laid out.
        match unsafe { SafeArray::from_raw_safearray_owned(array, None) } {
            Some(owner) => {
                drop(owner);
                S_OK
            }
            None => E_INVALIDARG,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Variant;

    #[test]
    fn task_memory_and_bstrs_round_trip() {
        // SAFETY: allocations are used within their sizes and freed once.
        unsafe {
            let block = co_task_mem_alloc(24).cast::<u8>();
            assert!(!block.is_null());
            block.write_bytes(7, 24);
            co_task_mem_free(block.cast());
            let text = sys_alloc_string_byte_len(b"abc".as_ptr(), 3);
            assert_eq!(sys_string_byte_len(text), 3);
            sys_free_string(text);
        }
    }

    #[test]
    fn variant_copy_and_clear() {
        let source = Variant::from_string("x");
        let mut copy = Variant::empty();
        // SAFETY: `copy` is empty and owns nothing; both cells are live.
        unsafe {
            let dst = (&raw mut copy).cast::<c_void>();
            core::mem::forget(core::ptr::read(dst.cast::<Variant>()));
            assert_eq!(variant_copy_ind(dst, (&raw const source).cast()), S_OK);
        }
        assert_eq!(copy, source);
        // SAFETY: a live cell.
        assert_eq!(unsafe { variant_clear((&raw mut copy).cast()) }, S_OK);
        assert_eq!(copy, Variant::empty());
    }
}
