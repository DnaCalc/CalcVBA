//! The oleaut32 and ole32 calls the Automation values need, with the platform's signatures over
//! raw pointers: the real functions on Windows, and the same contracts and memory layout
//! elsewhere (where this crate is the only allocator of BSTRs, SAFEARRAYs and task memory).
#![allow(clippy::missing_safety_doc)]

#[cfg(not(target_os = "windows"))]
use crate::{Variant, safe_array::SafeArray};
use core::ffi::c_void;

const S_OK: i32 = 0;
#[cfg(not(target_os = "windows"))]
const S_FALSE: i32 = 1;
#[cfg(not(target_os = "windows"))]
const E_POINTER: i32 = 0x8000_4003_u32 as i32;
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
        // SAFETY: a live cell; handle `VT_ARRAY | VT_RECORD` (Automation `FADF_RECORD`
        // arrays created via `safe_array_create_ex`) through `safe_array_destroy`.
        unsafe {
            let vt = cell.cast::<u16>().read();
            if vt == (crate::abi::VT_ARRAY | crate::safe_array::VT_RECORD_VALUE) {
                let parray = cell
                    .cast::<u8>()
                    .add(crate::abi::PAYLOAD_OFFSET)
                    .cast::<*mut c_void>()
                    .read();
                let hr = safe_array_destroy(parray);
                if hr >= 0 {
                    core::ptr::write_bytes(cell.cast::<u8>(), 0, crate::abi::VARIANT_SIZE);
                }
                return hr;
            }
            cell.cast::<Variant>().replace(Variant::empty());
        }
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
        // SAFETY: `source` is a live cell; handle `VT_ARRAY | VT_RECORD` via `safe_array_copy`.
        unsafe {
            let vt = source.cast::<u16>().read();
            if vt == (crate::abi::VT_ARRAY | crate::safe_array::VT_RECORD_VALUE) {
                let parray = source
                    .cast::<u8>()
                    .add(crate::abi::PAYLOAD_OFFSET)
                    .cast::<*mut c_void>()
                    .read();
                let mut copy = core::ptr::null_mut();
                let hr = safe_array_copy(parray, &mut copy);
                if hr < 0 {
                    return hr;
                }
                core::ptr::write_bytes(dst.cast::<u8>(), 0, crate::abi::VARIANT_SIZE);
                dst.cast::<u16>().write(vt);
                dst.cast::<u8>()
                    .add(crate::abi::PAYLOAD_OFFSET)
                    .cast::<*mut c_void>()
                    .write(copy);
                return S_OK;
            }
        }
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

/// `SafeArrayCreateEx`: allocates a descriptor + data (for `VT_RECORD`, `extra` is `IRecordInfo*`).
pub unsafe fn safe_array_create_ex(
    vt: u16,
    dims: u32,
    bounds: *const crate::safe_array::SafeArrayBound,
    extra: *mut c_void,
) -> *mut c_void {
    #[cfg(target_os = "windows")]
    {
        // SAFETY: `SafeArrayBound` has `#[repr(C)] { count: u32, lower: i32 }`, layout-identical
        // to Win32 `SAFEARRAYBOUND { cElements: u32, lLbound: i32 }`.
        const _: () = assert!(
            core::mem::size_of::<crate::safe_array::SafeArrayBound>()
                == core::mem::size_of::<windows_sys::Win32::System::Com::SAFEARRAYBOUND>()
        );
        // SAFETY: the caller passes `dims` valid bounds and an `extra` matching `vt`, as
        // SafeArrayCreateEx requires; the bound layout is asserted above.
        unsafe {
            windows_sys::Win32::System::Ole::SafeArrayCreateEx(vt, dims, bounds.cast(), extra)
                .cast()
        }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: forwarded from caller.
        unsafe { crate::safe_array::record_oleaut::create_ex(vt, dims, bounds, extra) }
    }
}

/// `SafeArrayPutElement`: copies one element (for `FADF_RECORD`, via `IRecordInfo::RecordCopy`).
pub unsafe fn safe_array_put_element(
    array: *mut c_void,
    indices: *const i32,
    value: *const c_void,
) -> i32 {
    #[cfg(target_os = "windows")]
    {
        // SAFETY: forwarded from caller.
        unsafe {
            windows_sys::Win32::System::Ole::SafeArrayPutElement(array.cast(), indices, value)
        }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: forwarded from caller.
        unsafe { crate::safe_array::record_oleaut::put_element(array, indices, value) }
    }
}

/// `SafeArrayPtrOfIndex`: computes the element address for 1..=dims indices (dim 1 first).
pub unsafe fn safe_array_ptr_of_index(
    array: *mut c_void,
    indices: *const i32,
    out: *mut *mut c_void,
) -> i32 {
    #[cfg(target_os = "windows")]
    {
        // SAFETY: forwarded from caller.
        unsafe { windows_sys::Win32::System::Ole::SafeArrayPtrOfIndex(array.cast(), indices, out) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: forwarded from caller.
        unsafe { crate::safe_array::record_oleaut::ptr_of_index(array, indices, out) }
    }
}

/// `SafeArrayGetDim`: returns the array's rank (`cDims`).
pub unsafe fn safe_array_get_dim(array: *const c_void) -> u32 {
    #[cfg(target_os = "windows")]
    {
        // SAFETY: forwarded from caller.
        unsafe { windows_sys::Win32::System::Ole::SafeArrayGetDim(array.cast()) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: forwarded from caller.
        unsafe { crate::safe_array::record_oleaut::get_dim(array) }
    }
}

/// `SafeArrayGetElemsize`: returns the array's element size in bytes (`cbElements`).
pub unsafe fn safe_array_get_elemsize(array: *const c_void) -> u32 {
    #[cfg(target_os = "windows")]
    {
        // SAFETY: forwarded from caller.
        unsafe { windows_sys::Win32::System::Ole::SafeArrayGetElemsize(array.cast()) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: forwarded from caller.
        unsafe { crate::safe_array::record_oleaut::get_elemsize(array) }
    }
}

/// `SafeArrayGetLBound`: reads the 1-based dimension `dim`'s lower bound.
pub unsafe fn safe_array_get_lbound(array: *const c_void, dim: u32, out: *mut i32) -> i32 {
    #[cfg(target_os = "windows")]
    {
        // SAFETY: forwarded from caller.
        unsafe { windows_sys::Win32::System::Ole::SafeArrayGetLBound(array.cast(), dim, out) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: forwarded from caller.
        unsafe { crate::safe_array::record_oleaut::get_lbound(array, dim, out) }
    }
}

/// `SafeArrayGetUBound`: reads the 1-based dimension `dim`'s upper bound.
pub unsafe fn safe_array_get_ubound(array: *const c_void, dim: u32, out: *mut i32) -> i32 {
    #[cfg(target_os = "windows")]
    {
        // SAFETY: forwarded from caller.
        unsafe { windows_sys::Win32::System::Ole::SafeArrayGetUBound(array.cast(), dim, out) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: forwarded from caller.
        unsafe { crate::safe_array::record_oleaut::get_ubound(array, dim, out) }
    }
}

/// `SafeArrayGetRecordInfo`: retrieves an owned `IRecordInfo*` from an `FADF_RECORD` descriptor.
pub unsafe fn safe_array_get_record_info(array: *const c_void, out: *mut *mut c_void) -> i32 {
    #[cfg(target_os = "windows")]
    {
        // SAFETY: forwarded from caller.
        unsafe { windows_sys::Win32::System::Ole::SafeArrayGetRecordInfo(array.cast(), out.cast()) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: forwarded from caller.
        unsafe { crate::safe_array::record_oleaut::get_record_info(array, out) }
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
        // SAFETY: live descriptor.
        if unsafe { crate::safe_array::record_oleaut::is_record_array(array) } {
            // SAFETY: handle `FADF_RECORD` descriptors through `IRecordInfo::RecordCopy`.
            return unsafe { crate::safe_array::record_oleaut::copy(array, out) };
        }
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

/// `SafeArrayGetIID`: reads the 16-byte IID from an `FADF_HAVEIID` descriptor.
pub unsafe fn safe_array_get_iid(array: *const c_void, out: *mut crate::com::Guid) -> i32 {
    if array.is_null() || out.is_null() {
        return E_INVALIDARG;
    }
    #[cfg(target_os = "windows")]
    {
        // SAFETY: `array` is a live SAFEARRAY and `out` is a writable GUID.
        unsafe { windows_sys::Win32::System::Ole::SafeArrayGetIID(array.cast(), out.cast()) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: `array` is a live descriptor this crate laid out.
        let Some(view) = (unsafe { SafeArray::borrow_raw(array.cast_mut(), None) }) else {
            return E_INVALIDARG;
        };
        let Some(iid) = view.iid() else {
            return E_INVALIDARG;
        };
        // SAFETY: `out` is a writable GUID.
        unsafe { out.write(crate::com::Guid::from_le_bytes(iid)) };
        S_OK
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
        // SAFETY: live descriptor.
        if unsafe { crate::safe_array::record_oleaut::is_record_array(array) } {
            // SAFETY: handle `FADF_RECORD` descriptors through `IRecordInfo::RecordClear`/`Release`.
            return unsafe { crate::safe_array::record_oleaut::destroy(array) };
        }
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

#[cfg(not(target_os = "windows"))]
struct ThreadErrorInfoSlot {
    ptr: *mut c_void,
}

#[cfg(not(target_os = "windows"))]
impl ThreadErrorInfoSlot {
    const fn empty() -> Self {
        Self {
            ptr: core::ptr::null_mut(),
        }
    }

    unsafe fn replace(&mut self, next: *mut c_void) {
        let prev = core::mem::replace(&mut self.ptr, next);
        if !prev.is_null() {
            // SAFETY: `prev` is a COM `IErrorInfo*` retained by `set_error_info`.
            unsafe {
                let vtbl = (*prev.cast::<crate::com::IUnknown>()).vtbl;
                ((*vtbl).release)(prev);
            }
        }
    }

    fn take(&mut self) -> *mut c_void {
        core::mem::replace(&mut self.ptr, core::ptr::null_mut())
    }
}

#[cfg(not(target_os = "windows"))]
impl Drop for ThreadErrorInfoSlot {
    fn drop(&mut self) {
        // SAFETY: release any remaining `IErrorInfo*` when the thread exits.
        unsafe {
            self.replace(core::ptr::null_mut());
        }
    }
}

#[cfg(not(target_os = "windows"))]
thread_local! {
    static THREAD_ERROR_INFO: std::cell::RefCell<ThreadErrorInfoSlot> =
        const { std::cell::RefCell::new(ThreadErrorInfoSlot::empty()) };
}

/// `SetErrorInfo`: stores (or clears, when `err_info` is null) the current thread's `IErrorInfo*`.
pub unsafe fn set_error_info(reserved: u32, err_info: *mut c_void) -> i32 {
    #[cfg(target_os = "windows")]
    {
        // SAFETY: forwarded to COM `SetErrorInfo`.
        unsafe { windows_sys::Win32::System::Com::SetErrorInfo(reserved, err_info) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        if reserved != 0 {
            return E_INVALIDARG;
        }
        if !err_info.is_null() {
            // SAFETY: `err_info` is a live COM `IErrorInfo*`; retain one reference for the slot.
            unsafe {
                let vtbl = (*err_info.cast::<crate::com::IUnknown>()).vtbl;
                ((*vtbl).add_ref)(err_info);
            }
        }
        THREAD_ERROR_INFO.with(|slot| {
            // SAFETY: replaces and releases any previous retained `IErrorInfo*`.
            unsafe {
                slot.borrow_mut().replace(err_info);
            }
        });
        S_OK
    }
}

/// `GetErrorInfo`: retrieves and clears the current thread's `IErrorInfo*`.
/// Returns `S_OK` when an error object was present, or `S_FALSE` when none was set.
pub unsafe fn get_error_info(reserved: u32, out: *mut *mut c_void) -> i32 {
    #[cfg(target_os = "windows")]
    {
        // SAFETY: forwarded to COM `GetErrorInfo`.
        unsafe { windows_sys::Win32::System::Com::GetErrorInfo(reserved, out) }
    }
    #[cfg(not(target_os = "windows"))]
    {
        if reserved != 0 {
            return E_INVALIDARG;
        }
        if out.is_null() {
            return E_POINTER;
        }
        let info = THREAD_ERROR_INFO.with(|slot| slot.borrow_mut().take());
        // SAFETY: `out` is non-null writable out-slot.
        unsafe {
            out.write(info);
        }
        if info.is_null() { S_FALSE } else { S_OK }
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

    #[test]
    fn safe_array_get_iid_reads_custom_iid_and_rejects_non_iid_arrays() {
        let collection_iid = crate::com::IID_VBA_COLLECTION;
        let arr = crate::safe_array::SafeArray::from_typed_variants(
            crate::safe_array::VT_DISPATCH_VALUE,
            vec![Variant::nothing()],
        )
        .unwrap()
        .with_iid(collection_iid.to_le_bytes());
        let mut out = crate::com::Guid::default();
        // SAFETY: live descriptor and writable GUID.
        let hr = unsafe { safe_array_get_iid(arr.raw_safearray_ptr().cast(), &raw mut out) };
        assert_eq!(hr, S_OK);
        assert_eq!(out, collection_iid);

        let longs = crate::safe_array::SafeArray::from_typed_variants(
            crate::safe_array::VT_I4_VALUE,
            vec![Variant::from_i32(1)],
        )
        .unwrap();
        // SAFETY: live descriptor and writable GUID; non-FADF_HAVEIID returns E_INVALIDARG.
        let hr = unsafe { safe_array_get_iid(longs.raw_safearray_ptr().cast(), &raw mut out) };
        assert_ne!(hr, S_OK);
    }

    #[test]
    fn record_safearray_round_trips_nd_bounds_copy_and_destroys_cleanly() {
        // Record layout: 16 bytes: offset 0 = i32, offset 8 = BSTR (pointer-sized on both 32/64).
        let bstr_offset: u32 = 8;
        let size: u32 = 8 + core::mem::size_of::<*mut u16>() as u32;
        let descriptor: [u32; 8] = [
            size,
            1, // 1 owning field
            bstr_offset,
            8, // VT_BSTR
            0x1234_5678,
            0x9abc_def0,
            0x4556_4143,
            0x5245_434f,
        ];
        let before = crate::current_thread_live_handle_counts();
        // SAFETY: descriptor is well-formed and all pointers stay valid for each oleaut call.
        unsafe {
            let info = crate::automation_record::create_record_info(descriptor.as_ptr());
            assert!(!info.is_null());
            let bounds = [
                crate::safe_array::SafeArrayBound { count: 2, lower: 1 },
                crate::safe_array::SafeArrayBound { count: 3, lower: 5 },
            ];
            let psa =
                safe_array_create_ex(crate::safe_array::VT_RECORD_VALUE, 2, bounds.as_ptr(), info);
            assert!(!psa.is_null());
            crate::automation_record::release_record_info(info);

            assert_eq!(safe_array_get_dim(psa), 2);
            assert_eq!(safe_array_get_elemsize(psa), size);
            let mut lb = 0;
            let mut ub = 0;
            assert_eq!(safe_array_get_lbound(psa, 1, &mut lb), S_OK);
            assert_eq!(safe_array_get_ubound(psa, 1, &mut ub), S_OK);
            assert_eq!((lb, ub), (1, 2));
            assert_eq!(safe_array_get_lbound(psa, 2, &mut lb), S_OK);
            assert_eq!(safe_array_get_ubound(psa, 2, &mut ub), S_OK);
            assert_eq!((lb, ub), (5, 7));

            let mut queried_info = core::ptr::null_mut();
            assert_eq!(safe_array_get_record_info(psa, &mut queried_info), S_OK);
            assert_eq!(queried_info, info);
            crate::automation_record::release_record_info(queried_info);

            // Populate element (2, 6).
            let mut buf = vec![0u64; (size as usize).div_ceil(8)];
            let base = buf.as_mut_ptr().cast::<u8>();
            base.cast::<i32>().write(42);
            let text_units: Vec<u16> = "hi".encode_utf16().collect();
            let bstr = sys_alloc_string_byte_len(
                text_units.as_ptr().cast(),
                (text_units.len() * 2) as u32,
            );
            base.add(bstr_offset as usize)
                .cast::<*mut u16>()
                .write(bstr);
            let idx = [2i32, 6i32];
            assert_eq!(safe_array_put_element(psa, idx.as_ptr(), base.cast()), S_OK);
            sys_free_string(bstr);

            // Deep copy through safe_array_copy and verify element (2, 6).
            let mut copy = core::ptr::null_mut();
            assert_eq!(safe_array_copy(psa, &mut copy), S_OK);
            assert!(!copy.is_null());
            let mut elem = core::ptr::null_mut();
            assert_eq!(safe_array_ptr_of_index(copy, idx.as_ptr(), &mut elem), S_OK);
            assert_eq!(elem.cast::<i32>().read(), 42);
            let copied_bstr = elem
                .cast::<u8>()
                .add(bstr_offset as usize)
                .cast::<*mut u16>()
                .read();
            assert_eq!(sys_string_byte_len(copied_bstr), 4);

            assert_eq!(safe_array_destroy(copy), S_OK);
            assert_eq!(safe_array_destroy(psa), S_OK);
        }
        assert_eq!(
            before.balance_to(crate::current_thread_live_handle_counts()),
            Default::default()
        );
    }

    #[test]
    fn set_and_get_error_info_round_trips_and_clears() {
        use crate::com::{
            Guid, HResult, IErrorInfoVtbl, IID_IERRORINFO, IID_IUNKNOWN, IUnknownVtbl,
        };
        use core::sync::atomic::{AtomicU32, Ordering};

        #[repr(C)]
        struct DummyErr {
            vtbl: *const IErrorInfoVtbl,
            refs: AtomicU32,
        }

        unsafe extern "system" fn qi(
            this: *mut c_void,
            riid: *const Guid,
            out: *mut *mut c_void,
        ) -> HResult {
            if out.is_null() || riid.is_null() {
                return 0x8000_4003u32 as i32;
            }
            // SAFETY: non-null `out` and `riid` from caller.
            unsafe {
                *out = core::ptr::null_mut();
                let iid = *riid;
                if iid == IID_IUNKNOWN || iid == IID_IERRORINFO {
                    add_ref(this);
                    *out = this;
                    S_OK
                } else {
                    0x8000_4002u32 as i32
                }
            }
        }
        unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
            // SAFETY: live `DummyErr`.
            unsafe {
                (*this.cast::<DummyErr>())
                    .refs
                    .fetch_add(1, Ordering::AcqRel)
                    + 1
            }
        }
        unsafe extern "system" fn release(this: *mut c_void) -> u32 {
            // SAFETY: live `DummyErr`.
            unsafe {
                let prev = (*this.cast::<DummyErr>())
                    .refs
                    .fetch_sub(1, Ordering::AcqRel);
                let rem = prev - 1;
                if rem == 0 {
                    drop(Box::from_raw(this.cast::<DummyErr>()));
                }
                rem
            }
        }
        unsafe extern "system" fn stub_guid(_: *mut c_void, _: *mut Guid) -> HResult {
            S_OK
        }
        unsafe extern "system" fn stub_bstr(_: *mut c_void, _: *mut *mut u16) -> HResult {
            S_OK
        }
        unsafe extern "system" fn stub_ctx(_: *mut c_void, _: *mut u32) -> HResult {
            S_OK
        }
        static VTBL: IErrorInfoVtbl = IErrorInfoVtbl {
            unknown: IUnknownVtbl {
                query_interface: qi,
                add_ref,
                release,
            },
            get_guid: stub_guid,
            get_source: stub_bstr,
            get_description: stub_bstr,
            get_help_file: stub_bstr,
            get_help_context: stub_ctx,
        };

        // SAFETY: exercises portable `set_error_info` / `get_error_info` contract.
        unsafe {
            assert_eq!(set_error_info(0, core::ptr::null_mut()), S_OK);
            let mut out = core::ptr::null_mut();
            assert_eq!(get_error_info(0, &mut out), 1);
            assert!(out.is_null());

            let err = Box::into_raw(Box::new(DummyErr {
                vtbl: &VTBL,
                refs: AtomicU32::new(1),
            }))
            .cast::<c_void>();
            assert_eq!(set_error_info(0, err), S_OK);
            assert_eq!(release(err), 1);

            assert_eq!(get_error_info(0, &mut out), S_OK);
            assert_eq!(out, err);
            assert_eq!(release(out), 0);

            let mut after = core::ptr::null_mut();
            assert_eq!(get_error_info(0, &mut after), 1);
            assert!(after.is_null());
        }
    }
}
