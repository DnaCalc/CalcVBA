//! Automation record owners, registration-free `IRecordInfo`, and nominal
//! `SAFEARRAY(VT_RECORD)` marshalling on every platform (`oleaut32` on Windows,
//! `crate::oleaut` elsewhere).
use crate::abi::{PAYLOAD_OFFSET, VARIANT_SIZE, VT_ARRAY};
use crate::com::{
    DISP_E_TYPEMISMATCH, E_INVALIDARG, E_NOINTERFACE, E_NOTIMPL, E_POINTER, Guid as GUID, HResult,
    IID_IRECORDINFO, IID_IUNKNOWN, S_OK,
};
use crate::oleaut::{
    co_task_mem_alloc as CoTaskMemAlloc, co_task_mem_free as CoTaskMemFree,
    safe_array_copy as SafeArrayCopy, safe_array_copy_data as SafeArrayCopyData,
    safe_array_create_ex as SafeArrayCreateEx, safe_array_destroy as SafeArrayDestroy,
    safe_array_features as SafeArrayFeatures, safe_array_get_dim as SafeArrayGetDim,
    safe_array_get_elemsize as SafeArrayGetElemsize, safe_array_get_lbound as SafeArrayGetLBound,
    safe_array_get_record_info as SafeArrayGetRecordInfo,
    safe_array_get_ubound as SafeArrayGetUBound,
    safe_array_is_record_array as SafeArrayIsRecordArray,
    safe_array_ptr_of_index as SafeArrayPtrOfIndex, safe_array_put_element as SafeArrayPutElement,
    sys_alloc_string_byte_len as SysAllocStringByteLen, sys_free_string as SysFreeString,
    sys_string_byte_len as SysStringByteLen, variant_clear as VariantClear,
    variant_copy_ind as VariantCopyInd,
};
use crate::safe_array::{
    FADF_FIXEDSIZE_VALUE, FADF_STATIC_VALUE, MAX_SAFEARRAY_RANK, SafeArray, SafeArrayBound,
    VT_RECORD_VALUE, VT_VARIANT_VALUE,
};
use crate::{Variant, VariantCore, VbaRecord, VbaRecordFieldKind, VbaRecordLayout};
use std::{
    ffi::c_void,
    ptr,
    sync::{
        Arc,
        atomic::{AtomicU32, Ordering},
    },
};

/// A `VARIANT`, by address only.
type VariantCell = c_void;
/// A `SAFEARRAY` descriptor, by address only.
type SafeArrayDescriptor = c_void;

const E_OUTOFMEMORY: i32 = 0x8007_000eu32 as i32;
const VT_BYREF: u16 = 0x4000;
const VT_ARRAY_RECORD: u16 = VT_ARRAY | VT_RECORD_VALUE;
const VT_ARRAY_VARIANT: u16 = VT_ARRAY | VT_VARIANT_VALUE;

/// `FADF_STATIC | FADF_FIXEDSIZE`: storage that cannot be resized or replaced.
pub const FADF_FIXED_STORAGE: u16 = FADF_STATIC_VALUE | FADF_FIXEDSIZE_VALUE;

/// AddRef through the interface's own IUnknown (null is Nothing).
unsafe fn interface_add_ref(pointer: *mut c_void) {
    if !pointer.is_null() {
        // SAFETY: a live interface pointer; slot 1 of every COM vtable is AddRef.
        unsafe {
            let table = *pointer.cast::<*const unsafe extern "system" fn(*mut c_void) -> u32>();
            (*table.add(1))(pointer);
        }
    }
}

/// Release through the interface's own IUnknown (null is Nothing).
unsafe fn interface_release(pointer: *mut c_void) {
    if !pointer.is_null() {
        // SAFETY: a live interface pointer owning one reference; slot 2 is Release.
        unsafe {
            let table = *pointer.cast::<*const unsafe extern "system" fn(*mut c_void) -> u32>();
            (*table.add(2))(pointer);
        }
    }
}

unsafe fn shape<'a>(descriptor: *const u32) -> (usize, &'a [u32]) {
    // SAFETY: the caller supplies a compiler-generated or retained descriptor.
    unsafe {
        (
            descriptor.read() as usize,
            std::slice::from_raw_parts(descriptor.add(2), descriptor.add(1).read() as usize * 2),
        )
    }
}

/// # Safety
/// `cell` is initialized native record storage matching the retained descriptor.
pub unsafe fn clear_record(cell: *mut u8, descriptor: *const u32) -> i32 {
    if cell.is_null() || descriptor.is_null() {
        return E_INVALIDARG;
    }
    // SAFETY: each owner offset and width was checked by the compiler or layout builder.
    unsafe {
        let (size, fields) = shape(descriptor);
        for field in fields.as_chunks::<2>().0 {
            let cell = cell.add(field[0] as usize);
            match field[1] {
                8 => {
                    let value = cell.cast::<*mut u16>().replace(ptr::null_mut());
                    SysFreeString(value);
                }
                12 => {
                    let hr = VariantClear(cell.cast());
                    if hr < 0 {
                        return hr;
                    }
                }
                9 | 13 => {
                    let value = cell.cast::<*mut c_void>().replace(ptr::null_mut());
                    interface_release(value);
                }
                kind if kind & 0x2000 != 0 => {
                    let value = cell
                        .cast::<*mut SafeArrayDescriptor>()
                        .replace(ptr::null_mut());
                    if !value.is_null() {
                        let hr = SafeArrayDestroy(value);
                        if hr < 0 {
                            return hr;
                        }
                    }
                }
                _ => return E_INVALIDARG,
            }
        }
        ptr::write_bytes(cell, 0, size);
    }
    S_OK
}

/// # Safety
/// Both records are initialized and have the exact descriptor shape. Aliasing
/// is permitted. All destination owners remain valid on a failed clone.
pub unsafe fn copy_record(dst: *mut u8, src: *const u8, descriptor: *const u32) -> i32 {
    if dst.is_null() || src.is_null() || descriptor.is_null() {
        return E_INVALIDARG;
    }
    // SAFETY: stage every owner before replacing destination storage.
    unsafe {
        let (size, fields) = shape(descriptor);
        if fields.is_empty() {
            ptr::copy(src, dst, size);
            return S_OK;
        }
        let mut words = vec![0u64; size.div_ceil(8)];
        let stage = words.as_mut_ptr().cast::<u8>();
        ptr::copy_nonoverlapping(src, stage, size);
        for field in fields.as_chunks::<2>().0 {
            ptr::write_bytes(
                stage.add(field[0] as usize),
                0,
                if field[1] == 12 {
                    VARIANT_SIZE
                } else {
                    std::mem::size_of::<*mut c_void>()
                },
            );
        }
        for field in fields.as_chunks::<2>().0 {
            let offset = field[0] as usize;
            let hr = if field[1] == 12 {
                VariantCopyInd(stage.add(offset).cast(), src.add(offset).cast())
            } else if matches!(field[1], 9 | 13) {
                let value = src.add(offset).cast::<*mut c_void>().read();
                interface_add_ref(value);
                stage.add(offset).cast::<*mut c_void>().write(value);
                S_OK
            } else if field[1] & 0x2000 != 0 {
                let value = src.add(offset).cast::<*mut SafeArrayDescriptor>().read();
                let mut copy = ptr::null_mut();
                let hr = if value.is_null() {
                    S_OK
                } else {
                    SafeArrayCopy(value, &mut copy)
                };
                stage
                    .add(offset)
                    .cast::<*mut SafeArrayDescriptor>()
                    .write(copy);
                hr
            } else if field[1] == 8 {
                let source = src.add(offset).cast::<*mut u16>().read();
                if source.is_null() {
                    stage.add(offset).cast::<*mut u16>().write(ptr::null_mut());
                    S_OK
                } else {
                    let len = SysStringByteLen(source);
                    let value = SysAllocStringByteLen(source.cast(), len);
                    if value.is_null() {
                        E_OUTOFMEMORY
                    } else {
                        stage.add(offset).cast::<*mut u16>().write(value);
                        S_OK
                    }
                }
            } else {
                E_INVALIDARG
            };
            if hr < 0 {
                clear_record(stage, descriptor);
                return hr;
            }
        }
        let hr = clear_record(dst, descriptor);
        if hr < 0 {
            clear_record(stage, descriptor);
            return hr;
        }
        ptr::copy_nonoverlapping(stage, dst, size);
    }
    S_OK
}

#[repr(C)]
struct RecordVtable {
    query_interface: unsafe extern "system" fn(
        this: *mut c_void,
        iid: *const GUID,
        ppv: *mut *mut c_void,
    ) -> i32,
    add_ref: unsafe extern "system" fn(this: *mut c_void) -> u32,
    release: unsafe extern "system" fn(this: *mut c_void) -> u32,
    record_init: unsafe extern "system" fn(this: *mut c_void, pvnew: *mut c_void) -> i32,
    record_clear: unsafe extern "system" fn(this: *mut c_void, pvexisting: *const c_void) -> i32,
    record_copy: unsafe extern "system" fn(
        this: *mut c_void,
        pvexisting: *const c_void,
        pvnew: *mut c_void,
    ) -> i32,
    get_guid: unsafe extern "system" fn(this: *mut c_void, pguid: *mut GUID) -> i32,
    get_name: unsafe extern "system" fn(this: *mut c_void, pbstrname: *mut *mut u16) -> i32,
    get_size: unsafe extern "system" fn(this: *mut c_void, pcbsize: *mut u32) -> i32,
    get_type_info:
        unsafe extern "system" fn(this: *mut c_void, pptypeinfo: *mut *mut c_void) -> i32,
    get_field: unsafe extern "system" fn(
        this: *mut c_void,
        pvdata: *const c_void,
        szfieldname: *const u16,
        pvarfield: *mut VariantCell,
    ) -> i32,
    get_field_no_copy: unsafe extern "system" fn(
        this: *mut c_void,
        pvdata: *const c_void,
        szfieldname: *const u16,
        pvarfield: *mut VariantCell,
        ppvdatacarray: *mut *mut c_void,
    ) -> i32,
    put_field: unsafe extern "system" fn(
        this: *mut c_void,
        wflags: u32,
        pvdata: *mut c_void,
        szfieldname: *const u16,
        pvarfield: *const VariantCell,
    ) -> i32,
    put_field_no_copy: unsafe extern "system" fn(
        this: *mut c_void,
        wflags: u32,
        pvdata: *mut c_void,
        szfieldname: *const u16,
        pvarfield: *const VariantCell,
    ) -> i32,
    get_field_names: unsafe extern "system" fn(
        this: *mut c_void,
        pcnames: *mut u32,
        rgbstrnames: *mut *mut u16,
    ) -> i32,
    is_matching_type: unsafe extern "system" fn(this: *mut c_void, precordinfo: *mut c_void) -> i32,
    record_create: unsafe extern "system" fn(this: *mut c_void) -> *mut c_void,
    record_create_copy: unsafe extern "system" fn(
        this: *mut c_void,
        pvsource: *const c_void,
        ppvdest: *mut *mut c_void,
    ) -> i32,
    record_destroy: unsafe extern "system" fn(this: *mut c_void, pvrecord: *const c_void) -> i32,
}

#[repr(C)]
struct RecordInfo {
    vtable: &'static RecordVtable,
    references: AtomicU32,
    descriptor: Box<[u32]>,
    guid: GUID,
    name: Box<str>,
    layout: Option<Arc<VbaRecordLayout>>,
}

fn protect(work: impl FnOnce() -> i32) -> i32 {
    std::panic::catch_unwind(std::panic::AssertUnwindSafe(work)).unwrap_or(E_OUTOFMEMORY)
}

unsafe extern "system" fn query(this: *mut c_void, iid: *const GUID, out: *mut *mut c_void) -> i32 {
    if out.is_null() {
        return E_POINTER;
    }
    // SAFETY: `out` is non-null writable output cell.
    unsafe { out.write(ptr::null_mut()) };
    if this.is_null() || iid.is_null() {
        return E_INVALIDARG;
    }
    // SAFETY: COM supplies a live object and readable GUID.
    unsafe {
        let iid = *iid;
        if iid == IID_IUNKNOWN || iid == IID_IRECORDINFO {
            addref(this);
            out.write(this);
            S_OK
        } else {
            E_NOINTERFACE
        }
    }
}

unsafe extern "system" fn addref(this: *mut c_void) -> u32 {
    // SAFETY: a COM owner keeps this allocation live during AddRef.
    unsafe {
        (*this.cast::<RecordInfo>())
            .references
            .fetch_add(1, Ordering::Relaxed)
            + 1
    }
}

unsafe extern "system" fn release(this: *mut c_void) -> u32 {
    // SAFETY: consume one retained COM reference, dropping only its last owner.
    unsafe {
        let info = this.cast::<RecordInfo>();
        let count = (*info).references.fetch_sub(1, Ordering::AcqRel) - 1;
        if count == 0 {
            drop(Box::from_raw(info));
        }
        count
    }
}

unsafe extern "system" fn initialize(this: *mut c_void, out: *mut c_void) -> i32 {
    if this.is_null() || out.is_null() {
        return E_INVALIDARG;
    }
    // SAFETY: the caller allocated GetSize bytes for this record type.
    unsafe {
        ptr::write_bytes(
            out.cast::<u8>(),
            0,
            (*this.cast::<RecordInfo>()).descriptor[0] as usize,
        );
    }
    S_OK
}

unsafe extern "system" fn clear(this: *mut c_void, record: *const c_void) -> i32 {
    if this.is_null() || record.is_null() {
        return E_INVALIDARG;
    }
    protect(|| {
        // SAFETY: initialized record storage belongs to this record descriptor.
        unsafe {
            clear_record(
                record.cast_mut().cast(),
                (*this.cast::<RecordInfo>()).descriptor.as_ptr(),
            )
        }
    })
}

unsafe extern "system" fn copy(this: *mut c_void, source: *const c_void, out: *mut c_void) -> i32 {
    if this.is_null() || source.is_null() || out.is_null() {
        return E_INVALIDARG;
    }
    protect(|| {
        // SAFETY: caller supplies initialized same-type source and destination.
        unsafe {
            copy_record(
                out.cast(),
                source.cast(),
                (*this.cast::<RecordInfo>()).descriptor.as_ptr(),
            )
        }
    })
}

unsafe extern "system" fn guid(this: *mut c_void, out: *mut GUID) -> i32 {
    if this.is_null() || out.is_null() {
        return E_INVALIDARG;
    }
    // SAFETY: live record info and writable GUID output.
    unsafe {
        out.write((*this.cast::<RecordInfo>()).guid);
    }
    S_OK
}

unsafe extern "system" fn name(this: *mut c_void, out: *mut *mut u16) -> i32 {
    if this.is_null() || out.is_null() {
        return E_INVALIDARG;
    }
    // SAFETY: live record info and writable BSTR output.
    unsafe {
        let text = &(*this.cast::<RecordInfo>()).name;
        let units: Vec<u16> = text.encode_utf16().collect();
        let src = if units.is_empty() {
            ptr::null()
        } else {
            units.as_ptr().cast::<u8>()
        };
        let value = SysAllocStringByteLen(src, (units.len() * 2) as u32);
        out.write(value);
        if value.is_null() { E_OUTOFMEMORY } else { S_OK }
    }
}

unsafe extern "system" fn size(this: *mut c_void, out: *mut u32) -> i32 {
    if this.is_null() || out.is_null() {
        return E_INVALIDARG;
    }
    // SAFETY: live info and writable ULONG output.
    unsafe {
        out.write((*this.cast::<RecordInfo>()).descriptor[0]);
    }
    S_OK
}

unsafe extern "system" fn type_info(_this: *mut c_void, out: *mut *mut c_void) -> i32 {
    if out.is_null() {
        return E_INVALIDARG;
    }
    // SAFETY: no registered type library is created or consulted.
    unsafe {
        out.write(ptr::null_mut());
    }
    E_NOTIMPL
}

unsafe extern "system" fn get_field(
    _this: *mut c_void,
    _data: *const c_void,
    _name: *const u16,
    _out: *mut VariantCell,
) -> i32 {
    E_NOTIMPL
}

unsafe extern "system" fn get_field_no_copy(
    _this: *mut c_void,
    _data: *const c_void,
    _name: *const u16,
    _out: *mut VariantCell,
    _array: *mut *mut c_void,
) -> i32 {
    E_NOTIMPL
}

unsafe extern "system" fn put_field(
    _this: *mut c_void,
    _flags: u32,
    _data: *mut c_void,
    _name: *const u16,
    _value: *const VariantCell,
) -> i32 {
    E_NOTIMPL
}

unsafe extern "system" fn field_names(
    _this: *mut c_void,
    count: *mut u32,
    _names: *mut *mut u16,
) -> i32 {
    if count.is_null() {
        return E_INVALIDARG;
    }
    // SAFETY: there are no reflective field names in the owning ABI descriptor.
    unsafe {
        count.write(0);
    }
    S_OK
}

unsafe extern "system" fn matches(this: *mut c_void, other: *mut c_void) -> i32 {
    if this.is_null() || other.is_null() {
        return 0;
    }
    if ptr::eq(this, other) {
        return 1;
    }
    // SAFETY: COM supplies another live IRecordInfo. Query its public GetGuid and GetSize.
    unsafe {
        let table = *other.cast::<*const RecordVtable>();
        let mut other_guid = GUID::default();
        if ((*table).get_guid)(other, &mut other_guid) < 0 {
            return 0;
        }
        let own = &*this.cast::<RecordInfo>();
        if own.guid != other_guid {
            return 0;
        }
        let mut other_size = 0u32;
        if ((*table).get_size)(other, &mut other_size) < 0 {
            return 0;
        }
        i32::from(own.descriptor[0] == other_size)
    }
}

unsafe extern "system" fn create(this: *mut c_void) -> *mut c_void {
    if this.is_null() {
        return ptr::null_mut();
    }
    // SAFETY: allocate exactly the advertised size and initialize all owners.
    unsafe {
        let out = CoTaskMemAlloc((*this.cast::<RecordInfo>()).descriptor[0] as usize);
        if !out.is_null() {
            initialize(this, out);
        }
        out
    }
}

unsafe extern "system" fn create_copy(
    this: *mut c_void,
    source: *const c_void,
    out: *mut *mut c_void,
) -> i32 {
    if out.is_null() {
        return E_POINTER;
    }
    // SAFETY: writable output cell.
    unsafe { out.write(ptr::null_mut()) };
    if this.is_null() || source.is_null() {
        return E_INVALIDARG;
    }
    // SAFETY: allocate/clone before publishing ownership to the caller.
    unsafe {
        let value = create(this);
        if value.is_null() {
            return E_OUTOFMEMORY;
        }
        let hr = copy(this, source, value);
        if hr < 0 {
            destroy(this, value);
            return hr;
        }
        out.write(value);
        S_OK
    }
}

unsafe extern "system" fn destroy(this: *mut c_void, record: *const c_void) -> i32 {
    if this.is_null() || record.is_null() {
        return E_INVALIDARG;
    }
    // SAFETY: RecordDestroy consumes this one RecordCreate allocation.
    unsafe {
        let hr = clear(this, record);
        if hr >= 0 {
            CoTaskMemFree(record.cast_mut());
        }
        hr
    }
}

static TABLE: RecordVtable = RecordVtable {
    query_interface: query,
    add_ref: addref,
    release,
    record_init: initialize,
    record_clear: clear,
    record_copy: copy,
    get_guid: guid,
    get_name: name,
    get_size: size,
    get_type_info: type_info,
    get_field,
    get_field_no_copy,
    put_field,
    put_field_no_copy: put_field,
    get_field_names: field_names,
    is_matching_type: matches,
    record_create: create,
    record_create_copy: create_copy,
    record_destroy: destroy,
};

/// # Safety
/// `descriptor` spans size/count/owner pairs followed by sixteen identity bytes.
/// Every owner offset/width is valid for that native record type.
/// The returned `IRecordInfo` owns its metadata; release it with [`release_record_info`].
pub unsafe fn create_record_info(descriptor: *const u32) -> *mut c_void {
    // SAFETY: copy all metadata before exposing the COM owner.
    unsafe {
        let count = descriptor.add(1).read() as usize * 2 + 2;
        let fields = std::slice::from_raw_parts(descriptor, count)
            .to_vec()
            .into_boxed_slice();
        let guid = descriptor.add(count).cast::<GUID>().read_unaligned();
        Box::into_raw(Box::new(RecordInfo {
            vtable: &TABLE,
            references: AtomicU32::new(1),
            descriptor: fields,
            guid,
            name: "CalcVBA native record".into(),
            layout: None,
        }))
        .cast()
    }
}

/// Build an owned COM `IRecordInfo` for a CalcVBA [`VbaRecordLayout`], carrying its nominal
/// `GetName`/`GetGuid`/`GetSize` and owning-field table (`RecordCopy`/`RecordClear`).
///
/// The returned pointer holds one reference; release it with [`release_record_info`].
pub fn create_record_info_for_layout(
    layout: &Arc<VbaRecordLayout>,
    fallback_name: &str,
    fallback_guid: GUID,
) -> Result<*mut c_void, HResult> {
    let size = u32::try_from(layout.size()).map_err(|_| DISP_E_TYPEMISMATCH)?;
    let owners = layout
        .automation_owners()
        .map_err(|_| DISP_E_TYPEMISMATCH)?;
    let owner_count = u32::try_from(owners.len()).map_err(|_| DISP_E_TYPEMISMATCH)?;
    let mut descriptor = Vec::with_capacity(2 + owners.len() * 2);
    descriptor.push(size);
    descriptor.push(owner_count);
    for (offset, owner_vt) in owners {
        descriptor.push(offset);
        descriptor.push(owner_vt);
    }
    let record_name = layout.type_name().unwrap_or(fallback_name);
    let record_guid = layout.type_guid().unwrap_or(fallback_guid);
    Ok(Box::into_raw(Box::new(RecordInfo {
        vtable: &TABLE,
        references: AtomicU32::new(1),
        descriptor: descriptor.into_boxed_slice(),
        guid: record_guid,
        name: record_name.into(),
        layout: Some(Arc::clone(layout)),
    }))
    .cast())
}

/// Retain one COM reference on a live `IRecordInfo*` (no-op for null).
///
/// # Safety
/// `pointer` must be null or a live `IRecordInfo*`.
pub unsafe fn add_ref_record_info(pointer: *mut c_void) {
    if !pointer.is_null() {
        // SAFETY: every `IRecordInfo` starts with `IUnknown`.
        unsafe {
            let table = *pointer.cast::<*const RecordVtable>();
            ((*table).add_ref)(pointer);
        }
    }
}

/// # Safety
/// `pointer` is null or a retained live `IRecordInfo`; consumes one reference.
pub unsafe fn release_record_info(pointer: *mut c_void) {
    if !pointer.is_null() {
        // SAFETY: every IRecordInfo starts with the standard COM reference slots.
        unsafe {
            let table = *pointer.cast::<*const RecordVtable>();
            ((*table).release)(pointer);
        }
    }
}

/// Query `IRecordInfo::GetSize`.
///
/// # Safety
/// `pointer` must be a live `IRecordInfo*`.
pub unsafe fn record_info_get_size(pointer: *mut c_void) -> Result<u32, HResult> {
    if pointer.is_null() {
        return Err(E_POINTER);
    }
    let mut out = 0u32;
    // SAFETY: live `IRecordInfo*` and writable `out`.
    let hr = unsafe {
        let table = *pointer.cast::<*const RecordVtable>();
        ((*table).get_size)(pointer, &mut out)
    };
    if hr < 0 { Err(hr) } else { Ok(out) }
}

/// Query `IRecordInfo::GetGuid`.
///
/// # Safety
/// `pointer` must be a live `IRecordInfo*`.
pub unsafe fn record_info_get_guid(pointer: *mut c_void) -> Result<GUID, HResult> {
    if pointer.is_null() {
        return Err(E_POINTER);
    }
    let mut out = GUID::default();
    // SAFETY: live `IRecordInfo*` and writable `out`.
    let hr = unsafe {
        let table = *pointer.cast::<*const RecordVtable>();
        ((*table).get_guid)(pointer, &mut out)
    };
    if hr < 0 { Err(hr) } else { Ok(out) }
}

/// Query `IRecordInfo::GetName` as a Rust `String` (frees the returned `BSTR`).
///
/// # Safety
/// `pointer` must be a live `IRecordInfo*`.
pub unsafe fn record_info_get_name(pointer: *mut c_void) -> Result<String, HResult> {
    if pointer.is_null() {
        return Err(E_POINTER);
    }
    let mut raw: *mut u16 = ptr::null_mut();
    // SAFETY: live `IRecordInfo*` and writable BSTR out-slot.
    let hr = unsafe {
        let table = *pointer.cast::<*const RecordVtable>();
        ((*table).get_name)(pointer, &mut raw)
    };
    if hr < 0 {
        return Err(hr);
    }
    if raw.is_null() {
        return Ok(String::new());
    }
    // SAFETY: `raw` is a fresh BSTR owned by us from `GetName`; read then free once.
    let text = unsafe {
        let byte_len = SysStringByteLen(raw) as usize;
        let units = std::slice::from_raw_parts(raw, byte_len / 2);
        let s = String::from_utf16_lossy(units);
        SysFreeString(raw);
        s
    };
    Ok(text)
}

/// Query `IRecordInfo::IsMatchingType`.
///
/// # Safety
/// `left` and `right` must be null or live `IRecordInfo*`.
pub unsafe fn record_info_is_matching_type(left: *mut c_void, right: *mut c_void) -> bool {
    if left.is_null() || right.is_null() {
        return false;
    }
    if ptr::eq(left, right) {
        return true;
    }
    // SAFETY: both pointers are live `IRecordInfo*`.
    unsafe {
        let table = *left.cast::<*const RecordVtable>();
        ((*table).is_matching_type)(left, right) != 0
    }
}

/// Recover the embedded [`VbaRecordLayout`] if `pointer` is a CalcVBA-created `IRecordInfo`.
///
/// # Safety
/// `pointer` must be null or a live COM `IRecordInfo*` whose first field is its vtable pointer.
pub unsafe fn layout_from_record_info(pointer: *mut c_void) -> Option<Arc<VbaRecordLayout>> {
    if pointer.is_null() {
        return None;
    }
    // SAFETY: every COM interface starts with its vtable pointer. Only if that pointer equals our
    // static `TABLE` address do we cast to `RecordInfo`.
    unsafe {
        let vtbl = *pointer.cast::<*const c_void>();
        if ptr::eq(vtbl, (&raw const TABLE).cast::<c_void>()) {
            (*pointer.cast::<RecordInfo>()).layout.clone()
        } else {
            None
        }
    }
}

/// RAII owner for one `IRecordInfo*` reference.
struct RecordInfoGuard(*mut c_void);

impl RecordInfoGuard {
    fn as_ptr(&self) -> *mut c_void {
        self.0
    }
}

impl Drop for RecordInfoGuard {
    fn drop(&mut self) {
        // SAFETY: owns one COM reference on `self.0`.
        unsafe { release_record_info(self.0) };
    }
}

/// RAII owner for a freshly allocated Automation `SAFEARRAY*` until transferred to a caller.
struct SafeArrayRollbackGuard(*mut c_void);

impl SafeArrayRollbackGuard {
    fn into_raw(mut self) -> *mut c_void {
        std::mem::replace(&mut self.0, ptr::null_mut())
    }
}

impl Drop for SafeArrayRollbackGuard {
    fn drop(&mut self) {
        if !self.0.is_null() {
            // SAFETY: owns this `SAFEARRAY*` until `into_raw` disarms it.
            let _ = unsafe { SafeArrayDestroy(self.0) };
        }
    }
}

/// RAII owner for one staged Automation record buffer initialized with `IRecordInfo::RecordInit`.
struct StagedAutomationRecord {
    info: *mut c_void,
    words: Vec<u64>,
}

impl StagedAutomationRecord {
    unsafe fn new(info: *mut c_void, size: usize) -> Result<Self, HResult> {
        let mut words = vec![0u64; size.div_ceil(8)];
        // SAFETY: `info` is a live `IRecordInfo*` and `words` covers `size` bytes.
        let hr = unsafe {
            let table = *info.cast::<*const RecordVtable>();
            ((*table).record_init)(info, words.as_mut_ptr().cast())
        };
        if hr < 0 {
            return Err(hr);
        }
        Ok(Self { info, words })
    }

    fn as_mut_ptr(&mut self) -> *mut u8 {
        self.words.as_mut_ptr().cast()
    }

    fn as_ptr(&self) -> *const u8 {
        self.words.as_ptr().cast()
    }
}

impl Drop for StagedAutomationRecord {
    fn drop(&mut self) {
        // SAFETY: `self.words` is an initialized record for `self.info`.
        unsafe {
            let table = *self.info.cast::<*const RecordVtable>();
            let _ = ((*table).record_clear)(self.info, self.words.as_ptr().cast());
        }
    }
}

/// Read dimensions and bounds (`dim 1..=cDims`) of any live Automation `SAFEARRAY`.
///
/// # Safety
/// `psa` must be a live non-null `SAFEARRAY*`.
pub unsafe fn read_automation_safearray_bounds(
    psa: *const c_void,
) -> Result<Vec<SafeArrayBound>, HResult> {
    if psa.is_null() {
        return Err(E_INVALIDARG);
    }
    // SAFETY: `psa` is a live descriptor.
    let dims = unsafe { SafeArrayGetDim(psa) };
    if dims == 0 || dims as usize > MAX_SAFEARRAY_RANK {
        return Err(DISP_E_TYPEMISMATCH);
    }
    let mut bounds = Vec::with_capacity(dims as usize);
    for dim in 1..=dims {
        let mut lb = 0i32;
        let mut ub = 0i32;
        // SAFETY: `psa` is live and `1 <= dim <= dims`.
        let hr_lb = unsafe { SafeArrayGetLBound(psa, dim, &mut lb) };
        if hr_lb < 0 {
            return Err(hr_lb);
        }
        // SAFETY: `psa` is live and `1 <= dim <= dims`.
        let hr_ub = unsafe { SafeArrayGetUBound(psa, dim, &mut ub) };
        if hr_ub < 0 {
            return Err(hr_ub);
        }
        let span = i64::from(ub) - i64::from(lb) + 1;
        let count = u32::try_from(span).map_err(|_| DISP_E_TYPEMISMATCH)?;
        bounds.push(SafeArrayBound { count, lower: lb });
    }
    Ok(bounds)
}

/// Walk all elements of an N-dimensional array in VBA/SAFEARRAY storage order (dim 1 fastest),
/// invoking `f(flat_index, &indices)` with 1-based dimension coordinates (`indices[0]` = dim 1).
fn for_each_nd_index(
    bounds: &[SafeArrayBound],
    mut f: impl FnMut(usize, &[i32]) -> Result<(), HResult>,
) -> Result<(), HResult> {
    if bounds.is_empty() || bounds.iter().any(|b| b.count == 0) {
        return Ok(());
    }
    let mut total = 1usize;
    for bound in bounds {
        total = total
            .checked_mul(bound.count as usize)
            .ok_or(DISP_E_TYPEMISMATCH)?;
    }
    let mut coords: Vec<i32> = bounds.iter().map(|b| b.lower).collect();
    for flat in 0..total {
        f(flat, &coords)?;
        for (dim, bound) in bounds.iter().enumerate() {
            let next_offset = (i64::from(coords[dim]) - i64::from(bound.lower)) + 1;
            if next_offset < i64::from(bound.count) {
                coords[dim] = (i64::from(bound.lower) + next_offset) as i32;
                break;
            }
            coords[dim] = bound.lower;
        }
    }
    Ok(())
}

/// Validate that a non-null `FADF_RECORD` `SAFEARRAY` matches `expected` nominally BEFORE reading
/// any element data. Returns an owned [`RecordInfoGuard`] on success.
unsafe fn validate_record_safearray_against_layout(
    psa: *const c_void,
    expected: &Arc<VbaRecordLayout>,
) -> Result<RecordInfoGuard, HResult> {
    // SAFETY: `psa` is a live non-null descriptor.
    if !unsafe { SafeArrayIsRecordArray(psa) } {
        return Err(DISP_E_TYPEMISMATCH);
    }
    let mut raw_info = ptr::null_mut();
    // SAFETY: `psa` is a live `FADF_RECORD` descriptor.
    let hr = unsafe { SafeArrayGetRecordInfo(psa, &mut raw_info) };
    if hr < 0 || raw_info.is_null() {
        return Err(DISP_E_TYPEMISMATCH);
    }
    let info = RecordInfoGuard(raw_info);
    // SAFETY: `psa` is a live descriptor.
    let elem_size = unsafe { SafeArrayGetElemsize(psa) } as usize;
    if elem_size != expected.size() {
        return Err(DISP_E_TYPEMISMATCH);
    }
    // SAFETY: `info` is a live `IRecordInfo*`.
    let info_size = unsafe { record_info_get_size(info.as_ptr()) }? as usize;
    if info_size != expected.size() {
        return Err(DISP_E_TYPEMISMATCH);
    }
    // SAFETY: `info` is a live `IRecordInfo*`.
    let info_guid = unsafe { record_info_get_guid(info.as_ptr()) }?;
    if let Some(expected_guid) = expected.type_guid()
        && expected_guid != GUID::default()
        && info_guid != expected_guid
    {
        return Err(DISP_E_TYPEMISMATCH);
    }
    // SAFETY: `info` is a live `IRecordInfo*`; if it is our own implementation, verify full layout
    // and nominal compatibility.
    if let Some(info_layout) = unsafe { layout_from_record_info(info.as_ptr()) }
        && !info_layout.matches_nominal(expected)
    {
        return Err(DISP_E_TYPEMISMATCH);
    }
    Ok(info)
}

/// Project one raw internal `VbaRecord` field into a staged Automation record buffer.
unsafe fn project_field_to_automation(
    src: *const u8,
    dst: *mut u8,
    kind: &VbaRecordFieldKind,
) -> Result<(), HResult> {
    match kind {
        VbaRecordFieldKind::Variant => {
            // SAFETY: `src` holds a live runtime `Variant`; `dst` is an initialized `VT_EMPTY` cell.
            unsafe {
                let value = &*src.cast::<Variant>();
                write_runtime_variant_to_automation_cell(dst.cast(), value)
            }
        }
        VbaRecordFieldKind::String => {
            // SAFETY: `src` holds a live BSTR pointer or null.
            let raw = unsafe { *src.cast::<*mut u16>() };
            if raw.is_null() {
                // SAFETY: `dst` is a writable pointer slot.
                unsafe { dst.cast::<*mut u16>().write(ptr::null_mut()) };
                return Ok(());
            }
            // SAFETY: `raw` is a live BSTR; clone its byte payload via `SysAllocStringByteLen`.
            let copy = unsafe {
                let len = SysStringByteLen(raw);
                SysAllocStringByteLen(raw.cast(), len)
            };
            if copy.is_null() {
                return Err(E_OUTOFMEMORY);
            }
            // SAFETY: `dst` is a writable pointer slot.
            unsafe { dst.cast::<*mut u16>().write(copy) };
            Ok(())
        }
        VbaRecordFieldKind::Object => {
            // SAFETY: `src` holds an `IUnknown*` or null; retain one reference for `dst`.
            unsafe {
                let raw = *src.cast::<*mut c_void>();
                interface_add_ref(raw);
                dst.cast::<*mut c_void>().write(raw);
            }
            Ok(())
        }
        VbaRecordFieldKind::DynamicArray { element } => {
            // SAFETY: `src` holds an owned runtime `SAFEARRAY*` or null.
            let raw = unsafe { *src.cast::<*mut c_void>() };
            if raw.is_null() {
                // SAFETY: `dst` is a writable pointer slot.
                unsafe { dst.cast::<*mut c_void>().write(ptr::null_mut()) };
                return Ok(());
            }
            let copied = match element.as_ref() {
                VbaRecordFieldKind::Record(sub_layout) => {
                    // SAFETY: `raw` is a live internal UDT `SAFEARRAY` of `sub_layout`.
                    unsafe {
                        vba_raw_record_safearray_to_automation(raw, sub_layout, ptr::null_mut())
                    }?
                }
                VbaRecordFieldKind::Variant => {
                    // SAFETY: `raw` is a live runtime `SAFEARRAY(VT_VARIANT)`.
                    unsafe { project_variant_safearray_to_automation(raw) }?
                }
                _ => {
                    let mut out = ptr::null_mut();
                    // SAFETY: `raw` is a live intrinsic `SAFEARRAY`.
                    let hr = unsafe { SafeArrayCopy(raw, &mut out) };
                    if hr < 0 {
                        return Err(hr);
                    }
                    out
                }
            };
            // SAFETY: `dst` is a writable pointer slot.
            unsafe { dst.cast::<*mut c_void>().write(copied) };
            Ok(())
        }
        VbaRecordFieldKind::Record(nested) => {
            for field in nested.fields() {
                // SAFETY: nested field offsets stay within both record buffers.
                unsafe {
                    project_field_to_automation(
                        src.add(field.offset()),
                        dst.add(field.offset()),
                        field.kind(),
                    )?;
                }
            }
            Ok(())
        }
        VbaRecordFieldKind::FixedArray { .. } => {
            let Some((elem_kind, len, stride)) =
                kind.fixed_array_layout().map_err(|_| DISP_E_TYPEMISMATCH)?
            else {
                return Err(DISP_E_TYPEMISMATCH);
            };
            for index in 0..len {
                // SAFETY: `index < len` stays within the fixed-array extent.
                unsafe {
                    project_field_to_automation(
                        src.add(index * stride),
                        dst.add(index * stride),
                        elem_kind,
                    )?;
                }
            }
            Ok(())
        }
        _ => {
            let (field_size, _) = kind.storage_shape().map_err(|_| DISP_E_TYPEMISMATCH)?;
            // SAFETY: plain scalar or fixed-string bytes; distinct buffers.
            unsafe { ptr::copy_nonoverlapping(src, dst, field_size) };
            Ok(())
        }
    }
}

/// Recover one Automation record field into an initialized runtime `VbaRecord` buffer.
unsafe fn recover_field_from_automation(
    src: *const u8,
    dst: *mut u8,
    kind: &VbaRecordFieldKind,
    resolve: &dyn Fn(&GUID, u32) -> Option<Arc<VbaRecordLayout>>,
) -> Result<(), HResult> {
    match kind {
        VbaRecordFieldKind::Variant => {
            // SAFETY: `src` is a live Automation `VARIANT` cell.
            let runtime_val =
                unsafe { automation_variant_to_runtime_variant(src.cast(), None, resolve) }?;
            // SAFETY: `dst` is an initialized `Variant` slot inside `VbaRecord`.
            let previous = unsafe { dst.cast::<Variant>().replace(runtime_val) };
            drop(previous);
            Ok(())
        }
        VbaRecordFieldKind::DynamicArray { element } => {
            // SAFETY: `src` holds an Automation `SAFEARRAY*` or null.
            let raw = unsafe { *src.cast::<*mut c_void>() };
            if raw.is_null() {
                return Ok(());
            }
            let runtime_psa = match element.as_ref() {
                VbaRecordFieldKind::Record(sub_layout) => {
                    // SAFETY: `raw` is a live Automation `SAFEARRAY(VT_RECORD)`.
                    let recovered = unsafe {
                        automation_safearray_to_vba_record_array(raw, sub_layout, false, resolve)
                    }?;
                    recovered
                        .try_as_safearray()
                        .map_err(|_| DISP_E_TYPEMISMATCH)?
                        .map_or(ptr::null_mut(), SafeArray::into_raw_safearray_ptr)
                }
                VbaRecordFieldKind::Variant => {
                    // SAFETY: `raw` is a live Automation `SAFEARRAY(VT_VARIANT)`.
                    unsafe { recover_variant_safearray_from_automation(raw, resolve) }?
                        .into_raw_safearray_ptr()
                }
                _ => {
                    // SAFETY: `raw` is a live intrinsic `SAFEARRAY`.
                    let copy = unsafe { SafeArray::try_clone_from_raw_safearray(raw, None) }
                        .map_err(|_| DISP_E_TYPEMISMATCH)?;
                    copy.map_or(ptr::null_mut(), SafeArray::into_raw_safearray_ptr)
                }
            };
            // SAFETY: `dst` is an initialized dynamic-array slot owning its previous descriptor.
            unsafe {
                let prev = dst.cast::<*mut c_void>().replace(runtime_psa);
                drop(SafeArray::from_raw_safearray_owned(
                    prev,
                    element.array_record_layout(),
                ));
            }
            Ok(())
        }
        VbaRecordFieldKind::Record(nested) => {
            for field in nested.fields() {
                // SAFETY: nested field offsets stay within both record buffers.
                unsafe {
                    recover_field_from_automation(
                        src.add(field.offset()),
                        dst.add(field.offset()),
                        field.kind(),
                        resolve,
                    )?;
                }
            }
            Ok(())
        }
        VbaRecordFieldKind::FixedArray { .. } => {
            let Some((elem_kind, len, stride)) =
                kind.fixed_array_layout().map_err(|_| DISP_E_TYPEMISMATCH)?
            else {
                return Err(DISP_E_TYPEMISMATCH);
            };
            for index in 0..len {
                // SAFETY: `index < len` stays within the fixed-array extent.
                unsafe {
                    recover_field_from_automation(
                        src.add(index * stride),
                        dst.add(index * stride),
                        elem_kind,
                        resolve,
                    )?;
                }
            }
            Ok(())
        }
        _ => {
            // SAFETY: ordinary scalar, FixedString, String (BSTR), or Object field.
            unsafe { VbaRecord::copy_field_from_raw(src, dst, kind) }
                .map_err(|_| DISP_E_TYPEMISMATCH)
        }
    }
}

/// Convert a non-null raw internal runtime UDT `SAFEARRAY` (`fFeatures` without `FADF_RECORD`)
/// into a fresh Automation `SAFEARRAY(VT_RECORD)` with `FADF_RECORD` and a nominal `IRecordInfo`.
///
/// If `record_info` is null, builds a CalcVBA `IRecordInfo` from `layout`.
unsafe fn vba_raw_record_safearray_to_automation(
    raw: *mut c_void,
    layout: &Arc<VbaRecordLayout>,
    record_info: *mut c_void,
) -> Result<*mut c_void, HResult> {
    if raw.is_null() {
        return Ok(ptr::null_mut());
    }
    // SAFETY: `raw` is a live internal UDT `SAFEARRAY` of `layout`.
    let Some(src_arr) = (unsafe { SafeArray::borrow_raw(raw, Some(Arc::clone(layout))) }) else {
        return Err(DISP_E_TYPEMISMATCH);
    };
    let bounds = src_arr.raw_bounds();
    if bounds.is_empty() {
        return Err(DISP_E_TYPEMISMATCH);
    }
    let owned_info = if record_info.is_null() {
        Some(RecordInfoGuard(create_record_info_for_layout(
            layout,
            "CalcVBA record",
            GUID::default(),
        )?))
    } else {
        None
    };
    let info = owned_info
        .as_ref()
        .map_or(record_info, RecordInfoGuard::as_ptr);

    // SAFETY: `bounds` has `bounds.len()` entries (dim 1 first) and `info` is a live `IRecordInfo*`.
    let psa =
        unsafe { SafeArrayCreateEx(VT_RECORD_VALUE, bounds.len() as u32, bounds.as_ptr(), info) };
    if psa.is_null() {
        return Err(E_OUTOFMEMORY);
    }
    let guard = SafeArrayRollbackGuard(psa);
    let fieldwise = layout.contains_variant_or_record_dynamic_array();

    for_each_nd_index(&bounds, |flat, indices| {
        let Some(src_elem) = src_arr.element_address(flat) else {
            return Err(DISP_E_TYPEMISMATCH);
        };
        if !fieldwise {
            // SAFETY: `psa` is a live `FADF_RECORD` array and `src_elem` is a live `layout` record.
            let hr = unsafe { SafeArrayPutElement(psa, indices.as_ptr(), src_elem.cast()) };
            if hr < 0 {
                return Err(hr);
            }
        } else {
            // SAFETY: `info` is a live `IRecordInfo*` with byte size `layout.size()`.
            let mut staged = unsafe { StagedAutomationRecord::new(info, layout.size()) }?;
            for field in layout.fields() {
                // SAFETY: offsets lie within both `layout`-sized buffers.
                unsafe {
                    project_field_to_automation(
                        src_elem.add(field.offset()),
                        staged.as_mut_ptr().add(field.offset()),
                        field.kind(),
                    )?;
                }
            }
            // SAFETY: `SafeArrayPutElement` copies `staged` into the element via `RecordCopy`.
            let hr = unsafe { SafeArrayPutElement(psa, indices.as_ptr(), staged.as_ptr().cast()) };
            if hr < 0 {
                return Err(hr);
            }
        }
        Ok(())
    })?;

    Ok(guard.into_raw())
}

/// Convert a runtime `Variant` holding a UDT array (`VT_ARRAY | VT_RECORD`) into an Automation
/// `SAFEARRAY(VT_RECORD)` carrying `FADF_RECORD` and a nominal `IRecordInfo` (or null if
/// unallocated).
///
/// # Safety
/// `record_info` must be null or a live `IRecordInfo*` matching `value`'s layout.
pub unsafe fn vba_record_array_to_automation_safearray_with_info(
    value: &Variant,
    expected: Option<&Arc<VbaRecordLayout>>,
    record_info: *mut c_void,
) -> Result<*mut c_void, HResult> {
    if value.core().raw_vt() != VT_ARRAY_RECORD {
        return Err(DISP_E_TYPEMISMATCH);
    }
    let layout = value.record_array_layout().ok_or(DISP_E_TYPEMISMATCH)?;
    if let Some(exp) = expected
        && !layout.matches_nominal(exp)
    {
        return Err(DISP_E_TYPEMISMATCH);
    }
    let raw = value.raw_safearray_ptr();
    if raw.is_null() {
        return Ok(ptr::null_mut());
    }
    // Prefer `expected`'s identity if `layout` was constructed anonymously.
    let effective_layout = match (layout.identity(), expected) {
        (None, Some(exp)) if exp.identity().is_some() => exp,
        _ => &layout,
    };
    // SAFETY: `raw` is this Variant's live UDT `SAFEARRAY`.
    unsafe { vba_raw_record_safearray_to_automation(raw, effective_layout, record_info) }
}

/// Convert a runtime `Variant` holding a UDT array (`VT_ARRAY | VT_RECORD`) into an Automation
/// `SAFEARRAY(VT_RECORD)` carrying `FADF_RECORD` and CalcVBA's nominal `IRecordInfo` (or null if
/// unallocated).
pub fn vba_record_array_to_automation_safearray(
    value: &Variant,
    expected: Option<&Arc<VbaRecordLayout>>,
) -> Result<*mut c_void, HResult> {
    // SAFETY: passing null `record_info` builds our nominal `IRecordInfo`.
    unsafe { vba_record_array_to_automation_safearray_with_info(value, expected, ptr::null_mut()) }
}

/// Convert an Automation `SAFEARRAY(VT_RECORD)` (or null for an unallocated array) into an
/// internal runtime UDT array `Variant`, validating `FADF_RECORD`, `IRecordInfo::GetSize`,
/// `cbElements`, and nominal `GetGuid` BEFORE copying any element data.
///
/// # Safety
/// `psa` must be null or a live Automation `SAFEARRAY*`.
pub unsafe fn automation_safearray_to_vba_record_array(
    psa: *mut c_void,
    expected: &Arc<VbaRecordLayout>,
    preserve_fixed_size: bool,
    resolve: &dyn Fn(&GUID, u32) -> Option<Arc<VbaRecordLayout>>,
) -> Result<Variant, HResult> {
    if psa.is_null() {
        return Ok(Variant::unallocated_record_array(Arc::clone(expected)));
    }
    // SAFETY: validates `FADF_RECORD`, `IRecordInfo`, size, and nominal GUID before reading data.
    let _info = unsafe { validate_record_safearray_against_layout(psa, expected) }?;
    // SAFETY: `psa` is a live `SAFEARRAY`.
    let bounds = unsafe { read_automation_safearray_bounds(psa) }?;
    let total_len = bounds.iter().try_fold(1usize, |acc, b| {
        acc.checked_mul(b.count as usize).ok_or(DISP_E_TYPEMISMATCH)
    })?;
    let fieldwise = expected.contains_variant_or_record_dynamic_array();
    let mut records = Vec::with_capacity(total_len);
    for_each_nd_index(&bounds, |_flat, indices| {
        let mut elem_ptr = ptr::null_mut();
        // SAFETY: `psa` is live and `indices` are in-bounds for `psa`.
        let hr = unsafe { SafeArrayPtrOfIndex(psa, indices.as_ptr(), &mut elem_ptr) };
        if hr < 0 || elem_ptr.is_null() {
            return Err(if hr < 0 { hr } else { DISP_E_TYPEMISMATCH });
        }
        let record = if !fieldwise {
            // SAFETY: `elem_ptr` is a live initialized `expected`-shaped Automation record.
            unsafe { VbaRecord::clone_from_raw(elem_ptr.cast(), Arc::clone(expected)) }
                .map_err(|_| DISP_E_TYPEMISMATCH)?
        } else {
            let mut rec =
                VbaRecord::new_default(Arc::clone(expected)).map_err(|_| E_OUTOFMEMORY)?;
            let src_base = elem_ptr.cast::<u8>();
            let dst_base = rec.data_mut_ptr();
            for field in expected.fields() {
                // SAFETY: field offsets lie within both `expected`-sized record buffers.
                unsafe {
                    recover_field_from_automation(
                        src_base.add(field.offset()),
                        dst_base.add(field.offset()),
                        field.kind(),
                        resolve,
                    )?;
                }
            }
            rec
        };
        records.push(record);
        Ok(())
    })?;
    let mut sa = SafeArray::from_vba_records_nd(bounds, Arc::clone(expected), records)
        .map_err(|_| E_OUTOFMEMORY)?;
    // SAFETY: `psa` is live.
    if preserve_fixed_size && (unsafe { SafeArrayFeatures(psa) } & FADF_FIXED_STORAGE) != 0 {
        sa.set_fixed_size(true);
    }
    Ok(Variant::from_safearray(sa))
}

/// Convert an Automation `SAFEARRAY(VT_RECORD)` whose UDT layout is resolved from its
/// `IRecordInfo` (either CalcVBA's own `IRecordInfo`, or `resolve(&guid, size)`).
///
/// Unknown foreign `SAFEARRAY(VT_RECORD)` values (and null descriptors with no known type) are
/// rejected with `DISP_E_TYPEMISMATCH` (VBA error 13).
///
/// # Safety
/// `psa` must be null or a live Automation `SAFEARRAY*`.
pub unsafe fn automation_safearray_to_vba_record_array_any(
    psa: *mut c_void,
    preserve_fixed_size: bool,
    resolve: &dyn Fn(&GUID, u32) -> Option<Arc<VbaRecordLayout>>,
) -> Result<Variant, HResult> {
    if psa.is_null() {
        return Err(DISP_E_TYPEMISMATCH);
    }
    // SAFETY: `psa` is a live non-null descriptor.
    if !unsafe { SafeArrayIsRecordArray(psa) } {
        return Err(DISP_E_TYPEMISMATCH);
    }
    let mut raw_info = ptr::null_mut();
    // SAFETY: `psa` is a live `FADF_RECORD` descriptor.
    let hr = unsafe { SafeArrayGetRecordInfo(psa, &mut raw_info) };
    if hr < 0 || raw_info.is_null() {
        return Err(DISP_E_TYPEMISMATCH);
    }
    let info = RecordInfoGuard(raw_info);
    // SAFETY: `info` is a live `IRecordInfo*`.
    let layout = if let Some(own_layout) = unsafe { layout_from_record_info(info.as_ptr()) } {
        own_layout
    } else {
        // SAFETY: `info` is a live `IRecordInfo*`.
        let (info_guid, info_size) = unsafe {
            (
                record_info_get_guid(info.as_ptr())?,
                record_info_get_size(info.as_ptr())?,
            )
        };
        resolve(&info_guid, info_size).ok_or(DISP_E_TYPEMISMATCH)?
    };
    drop(info);
    // SAFETY: `psa` is a live `FADF_RECORD` descriptor matching `layout`.
    unsafe { automation_safearray_to_vba_record_array(psa, &layout, preserve_fixed_size, resolve) }
}

/// Write a runtime UDT array `Variant` back into a caller's Automation `SAFEARRAY**` slot
/// (`ByRef items() As MyType` or retval `*mut *mut SAFEARRAY`).
///
/// - When `*slot` is non-null with identical bounds and nominal record identity, copies data in
///   place via `SafeArrayCopyData` (preserving descriptor identity and supporting fixed-size
///   arrays).
/// - When `*slot` has `FADF_STATIC | FADF_FIXEDSIZE` and the shape changed (or array was erased),
///   returns `DISP_E_TYPEMISMATCH` without modifying `*slot`.
/// - Otherwise replaces `*slot` with a freshly allocated `FADF_RECORD` `SAFEARRAY` (or null when
///   unallocated) and destroys the prior descriptor.
///
/// # Safety
/// `slot` must point to a live, writable, initialized Automation `SAFEARRAY*` slot (null or a
/// live `SAFEARRAY*` owned by the caller).
pub unsafe fn write_vba_record_array_to_automation_slot(
    slot: *mut *mut c_void,
    value: &Variant,
    expected: &Arc<VbaRecordLayout>,
) -> Result<(), HResult> {
    if slot.is_null() {
        return Err(E_POINTER);
    }
    if value.core().raw_vt() != VT_ARRAY_RECORD {
        return Err(DISP_E_TYPEMISMATCH);
    }
    let val_layout = value.record_array_layout().ok_or(DISP_E_TYPEMISMATCH)?;
    if !val_layout.matches_nominal(expected) {
        return Err(DISP_E_TYPEMISMATCH);
    }
    // SAFETY: caller owns this initialized `SAFEARRAY*` slot.
    let previous = unsafe { slot.read() };
    let new_raw = value.raw_safearray_ptr();

    if new_raw.is_null() {
        if !previous.is_null() {
            // SAFETY: `previous` is live.
            if (unsafe { SafeArrayFeatures(previous) } & FADF_FIXED_STORAGE) != 0 {
                return Err(DISP_E_TYPEMISMATCH);
            }
            // SAFETY: dynamic descriptor is destroyed and cleared to null.
            let hr = unsafe { SafeArrayDestroy(previous) };
            if hr < 0 {
                return Err(hr);
            }
            // SAFETY: `slot` is writable.
            unsafe { slot.write(ptr::null_mut()) };
        }
        return Ok(());
    }

    // SAFETY: `new_raw` is this Variant's live UDT `SAFEARRAY`.
    let new_bounds = unsafe { SafeArray::borrow_raw(new_raw, Some(Arc::clone(expected))) }
        .map(|arr| arr.raw_bounds())
        .ok_or(DISP_E_TYPEMISMATCH)?;

    if !previous.is_null() {
        // Check whether `previous` can be updated in place (same bounds and compatible `FADF_RECORD`).
        // SAFETY: `previous` is live.
        let same_shape_info = unsafe {
            validate_record_safearray_against_layout(previous, expected)
                .ok()
                .zip(read_automation_safearray_bounds(previous).ok())
        };
        if let Some((prev_info, prev_bounds)) = same_shape_info
            && prev_bounds == new_bounds
        {
            // Stage with `prev_info` so `SafeArrayCopyData` sees matching `IRecordInfo`.
            // SAFETY: `prev_info` is live and validated against `expected`.
            let staged = unsafe {
                vba_record_array_to_automation_safearray_with_info(
                    value,
                    Some(expected),
                    prev_info.as_ptr(),
                )
            }?;
            let staged_guard = SafeArrayRollbackGuard(staged);
            // SAFETY: both descriptors are live `FADF_RECORD` arrays with matching shape & info.
            let hr = unsafe { SafeArrayCopyData(staged_guard.0, previous) };
            return if hr < 0 { Err(hr) } else { Ok(()) };
        }

        // Shape changed (or previous wasn't a matching record array): fixed storage cannot reallocate.
        // SAFETY: `previous` is live.
        if (unsafe { SafeArrayFeatures(previous) } & FADF_FIXED_STORAGE) != 0 {
            return Err(DISP_E_TYPEMISMATCH);
        }

        // Build replacement BEFORE destroying `previous` so failure leaves caller's slot intact.
        let replacement = vba_record_array_to_automation_safearray(value, Some(expected))?;
        let guard = SafeArrayRollbackGuard(replacement);
        // SAFETY: caller's slot owns `previous`.
        let hr = unsafe { SafeArrayDestroy(previous) };
        if hr < 0 {
            return Err(hr);
        }
        // SAFETY: `slot` is writable; transfer ownership of `replacement`.
        unsafe { slot.write(guard.into_raw()) };
        return Ok(());
    }

    let replacement = vba_record_array_to_automation_safearray(value, Some(expected))?;
    // SAFETY: `slot` is writable; transfer ownership of `replacement`.
    unsafe { slot.write(replacement) };
    Ok(())
}

/// True when a raw `VARIANT` cell holds an Automation `VT_ARRAY | VT_RECORD` (zero tail pointer).
unsafe fn is_automation_record_array_cell(cell: *const c_void) -> bool {
    if cell.is_null() {
        return false;
    }
    // SAFETY: caller supplies a live `VARIANT` cell.
    unsafe {
        let vt = cell.cast::<u16>().read_unaligned();
        if vt != VT_ARRAY_RECORD {
            return false;
        }
        let tail = cell
            .cast::<u8>()
            .add(PAYLOAD_OFFSET + std::mem::size_of::<*mut c_void>())
            .cast::<*mut c_void>()
            .read_unaligned();
        tail.is_null()
    }
}

/// True when a `SAFEARRAY(VT_VARIANT)` contains any `VT_ARRAY | VT_RECORD` element (recursively).
unsafe fn variant_safearray_contains_record_array(
    psa: *mut c_void,
    automation_only: bool,
    depth: usize,
) -> bool {
    if psa.is_null() || depth > 16 {
        return false;
    }
    // SAFETY: `psa` is a live `SAFEARRAY(VT_VARIANT)`.
    let Ok(bounds) = (unsafe { read_automation_safearray_bounds(psa) }) else {
        return false;
    };
    let mut found = false;
    let _ = for_each_nd_index(&bounds, |_flat, indices| {
        if found {
            return Ok(());
        }
        let mut elem = ptr::null_mut();
        // SAFETY: `psa` is live and `indices` are in-bounds.
        if unsafe { SafeArrayPtrOfIndex(psa, indices.as_ptr(), &mut elem) } < 0 || elem.is_null() {
            return Ok(());
        }
        // SAFETY: `elem` points to a live `VARIANT` element.
        let vt = unsafe { elem.cast::<u16>().read_unaligned() };
        if vt == VT_ARRAY_RECORD {
            // SAFETY: `elem` is a live `0x2024` `VARIANT` cell.
            if !automation_only || unsafe { is_automation_record_array_cell(elem) } {
                found = true;
            }
        } else if vt == VT_ARRAY_VARIANT {
            // SAFETY: `elem` is a live `0x200C` `VARIANT` cell whose payload holds a `SAFEARRAY*`.
            let nested_has_record = unsafe {
                let nested_psa = elem
                    .cast::<u8>()
                    .add(PAYLOAD_OFFSET)
                    .cast::<*mut c_void>()
                    .read_unaligned();
                variant_safearray_contains_record_array(nested_psa, automation_only, depth + 1)
            };
            if nested_has_record {
                found = true;
            }
        }
        Ok(())
    });
    found
}

/// Project a runtime `SAFEARRAY(VT_VARIANT)` containing UDT arrays into an Automation
/// `SAFEARRAY(VT_VARIANT)` whose elements are valid Automation `VARIANT`s.
unsafe fn project_variant_safearray_to_automation(
    psa: *mut c_void,
) -> Result<*mut c_void, HResult> {
    if psa.is_null() {
        return Ok(ptr::null_mut());
    }
    // SAFETY: `psa` is a live runtime `SAFEARRAY(VT_VARIANT)`.
    let bounds = unsafe { read_automation_safearray_bounds(psa) }?;
    // Allocate via `SafeArray` so descriptor layout matches `oleaut::safe_array_destroy`.
    let out_sa = SafeArray::from_shape_typed(bounds.clone(), VT_VARIANT_VALUE)
        .and_then(|_| {
            let total = bounds
                .iter()
                .try_fold(1usize, |a, b| a.checked_mul(b.count as usize))
                .ok_or_else(|| "overflow".to_string())?;
            let empties = vec![Variant::empty(); total];
            SafeArray::from_typed_variants_nd(bounds.clone(), VT_VARIANT_VALUE, empties)
        })
        .map_err(|_| E_OUTOFMEMORY)?;
    let out_raw = out_sa.into_raw_safearray_ptr();
    let guard = SafeArrayRollbackGuard(out_raw);
    for_each_nd_index(&bounds, |_flat, indices| {
        let mut src_elem = ptr::null_mut();
        let mut dst_elem = ptr::null_mut();
        // SAFETY: both arrays have identical `bounds`.
        let (hr1, hr2) = unsafe {
            (
                SafeArrayPtrOfIndex(psa, indices.as_ptr(), &mut src_elem),
                SafeArrayPtrOfIndex(out_raw, indices.as_ptr(), &mut dst_elem),
            )
        };
        if hr1 < 0 || hr2 < 0 || src_elem.is_null() || dst_elem.is_null() {
            return Err(DISP_E_TYPEMISMATCH);
        }
        // SAFETY: `src_elem` is a live runtime `Variant`, `dst_elem` is an initialized `VT_EMPTY`.
        unsafe { write_runtime_variant_to_automation_cell(dst_elem, &*src_elem.cast::<Variant>()) }
    })?;
    Ok(guard.into_raw())
}

/// Recover an Automation `SAFEARRAY(VT_VARIANT)` (which may contain Automation `0x2024` elements)
/// into an owned runtime `SafeArray`.
unsafe fn recover_variant_safearray_from_automation(
    psa: *mut c_void,
    resolve: &dyn Fn(&GUID, u32) -> Option<Arc<VbaRecordLayout>>,
) -> Result<SafeArray, HResult> {
    // SAFETY: `psa` is a live `SAFEARRAY(VT_VARIANT)`.
    let bounds = unsafe { read_automation_safearray_bounds(psa) }?;
    let total = bounds.iter().try_fold(1usize, |acc, b| {
        acc.checked_mul(b.count as usize).ok_or(DISP_E_TYPEMISMATCH)
    })?;
    let mut values = Vec::with_capacity(total);
    for_each_nd_index(&bounds, |_flat, indices| {
        let mut elem = ptr::null_mut();
        // SAFETY: `psa` is live and `indices` are in-bounds.
        let hr = unsafe { SafeArrayPtrOfIndex(psa, indices.as_ptr(), &mut elem) };
        if hr < 0 || elem.is_null() {
            return Err(if hr < 0 { hr } else { DISP_E_TYPEMISMATCH });
        }
        // SAFETY: `elem` is a live Automation `VARIANT`.
        let v = unsafe { automation_variant_to_runtime_variant(elem, None, resolve) }?;
        values.push(v);
        Ok(())
    })?;
    SafeArray::from_typed_variants_nd(bounds, VT_VARIANT_VALUE, values).map_err(|_| E_OUTOFMEMORY)
}

/// Clear a live `VARIANT` cell that may hold either an Automation `0x2024` (or `0x200C` holding
/// Automation `0x2024`) or an ordinary runtime `Variant`.
///
/// # Safety
/// `cell` must point to a live, initialized, writable `VARIANT` cell.
pub unsafe fn clear_automation_variant_cell(cell: *mut c_void) -> Result<(), HResult> {
    if cell.is_null() {
        return Err(E_POINTER);
    }
    // SAFETY: `cell` is a live `VARIANT`.
    let vt = unsafe { cell.cast::<u16>().read_unaligned() };
    // SAFETY: `cell` is a live `VARIANT`.
    if vt == VT_ARRAY_RECORD && unsafe { is_automation_record_array_cell(cell) } {
        // SAFETY: Automation `0x2024` clears through `oleaut::variant_clear`.
        let hr = unsafe { VariantClear(cell) };
        return if hr < 0 { Err(hr) } else { Ok(()) };
    }
    if vt == VT_ARRAY_VARIANT {
        // SAFETY: read `parray` from the `VARIANT` and inspect nested `0x2024` elements.
        let (psa, has_auto_records) = unsafe {
            let psa = cell
                .cast::<u8>()
                .add(PAYLOAD_OFFSET)
                .cast::<*mut c_void>()
                .read_unaligned();
            (psa, variant_safearray_contains_record_array(psa, true, 0))
        };
        if has_auto_records {
            // Clear each Automation `0x2024` element before letting `SafeArray` drop the outer
            // `SAFEARRAY(VT_VARIANT)` descriptor (keeps `live_counters` balanced off Windows and
            // avoids oleaut32/Rust mismatches).
            // SAFETY: `psa` is a live `SAFEARRAY(VT_VARIANT)`.
            if let Ok(bounds) = unsafe { read_automation_safearray_bounds(psa) } {
                let _ = for_each_nd_index(&bounds, |_flat, indices| {
                    let mut elem = ptr::null_mut();
                    // SAFETY: `psa` is live and `indices` are in bounds; each element is a live `VARIANT`.
                    unsafe {
                        if SafeArrayPtrOfIndex(psa, indices.as_ptr(), &mut elem) >= 0
                            && !elem.is_null()
                        {
                            let _ = clear_automation_variant_cell(elem);
                        }
                    }
                    Ok(())
                });
            }
            // SAFETY: all inner Automation `0x2024` elements are now `VT_EMPTY`; clear the outer
            // cell via `Variant::drop` or `VariantClear`.
            let hr = unsafe { VariantClear(cell) };
            return if hr < 0 { Err(hr) } else { Ok(()) };
        }
    }
    // SAFETY: standard runtime `Variant` cell (or `VT_BYREF` which owns nothing).
    unsafe {
        let old = cell.cast::<Variant>().replace(Variant::empty());
        drop(old);
    }
    Ok(())
}

/// Write a runtime `Variant` into an initialized Automation `VARIANT` cell (e.g., `pVarResult`, a
/// `ByRef v As Variant` target, or an outbound `DISPPARAMS` argument cell), projecting any UDT
/// array (`0x2024`) to an Automation `SAFEARRAY(VT_RECORD)` with `FADF_RECORD` and a nominal
/// `IRecordInfo` (and zero tail).
///
/// # Safety
/// `dst` must point to a live, initialized, writable `VARIANT` cell.
pub unsafe fn write_runtime_variant_to_automation_cell(
    dst: *mut c_void,
    value: &Variant,
) -> Result<(), HResult> {
    if dst.is_null() {
        return Err(E_POINTER);
    }
    let vt = value.core().raw_vt();
    if vt == VT_ARRAY_RECORD {
        let psa = vba_record_array_to_automation_safearray(value, None)?;
        let guard = SafeArrayRollbackGuard(psa);
        // SAFETY: `dst` is a live initialized cell.
        unsafe { clear_automation_variant_cell(dst) }?;
        // SAFETY: `dst` is now `VT_EMPTY`; write Automation `0x2024` with zero tail.
        unsafe {
            ptr::write_bytes(dst.cast::<u8>(), 0, VARIANT_SIZE);
            dst.cast::<u16>().write(VT_ARRAY_RECORD);
            dst.cast::<u8>()
                .add(PAYLOAD_OFFSET)
                .cast::<*mut c_void>()
                .write(guard.into_raw());
        }
        return Ok(());
    }
    if vt == VT_ARRAY_VARIANT {
        let src_psa = value.raw_safearray_ptr();
        // SAFETY: `src_psa` is live or null.
        if unsafe { variant_safearray_contains_record_array(src_psa, false, 0) } {
            // SAFETY: `src_psa` is a live runtime `SAFEARRAY(VT_VARIANT)`.
            let projected = unsafe { project_variant_safearray_to_automation(src_psa) }?;
            let guard = SafeArrayRollbackGuard(projected);
            // SAFETY: `dst` is a live initialized cell.
            unsafe { clear_automation_variant_cell(dst) }?;
            // SAFETY: `dst` is now `VT_EMPTY`; write `0x200C` with projected `SAFEARRAY*`.
            unsafe {
                ptr::write_bytes(dst.cast::<u8>(), 0, VARIANT_SIZE);
                dst.cast::<u16>().write(VT_ARRAY_VARIANT);
                dst.cast::<u8>()
                    .add(PAYLOAD_OFFSET)
                    .cast::<*mut c_void>()
                    .write(guard.into_raw());
            }
            return Ok(());
        }
    }
    let cloned = value.clone();
    // SAFETY: `dst` is a live initialized cell; clearing handles any prior Automation `0x2024`.
    unsafe { clear_automation_variant_cell(dst) }?;
    // SAFETY: `dst` is `VT_EMPTY`; overwrite with `cloned`.
    unsafe {
        dst.cast::<Variant>().write(cloned);
    }
    Ok(())
}

/// Read an Automation `VARIANT` cell (following `VT_BYREF | VT_VARIANT` chains) into a runtime
/// `Variant`, converting Automation `SAFEARRAY(VT_RECORD)` (`0x2024` and `0x6024`) and nested
/// `Variant()` arrays holding them into internal runtime UDT array `Variant`s.
///
/// When `expected_record_array` is `Some(layout)`:
/// - Rejects any non-`VT_ARRAY | VT_RECORD` argument or mismatched nominal UDT with
///   `DISP_E_TYPEMISMATCH` BEFORE copying any data.
/// - Accepts null `SAFEARRAY*` as `Variant::unallocated_record_array(layout)`.
///
/// # Safety
/// `cell` must point to a live, readable Automation `VARIANT` cell.
pub unsafe fn automation_variant_to_runtime_variant(
    cell: *const c_void,
    expected_record_array: Option<&Arc<VbaRecordLayout>>,
    resolve: &dyn Fn(&GUID, u32) -> Option<Arc<VbaRecordLayout>>,
) -> Result<Variant, HResult> {
    if cell.is_null() {
        return Err(E_POINTER);
    }
    // Follow `VT_BYREF | VT_VARIANT` chains up to 32 hops.
    let mut cur = cell;
    for _ in 0..32 {
        // SAFETY: `cur` is a live `VARIANT` cell.
        let vt = unsafe { cur.cast::<u16>().read_unaligned() };
        if vt == (VT_BYREF | VT_VARIANT_VALUE) {
            // SAFETY: payload holds a `VARIANT*`.
            let next = unsafe {
                cur.cast::<u8>()
                    .add(PAYLOAD_OFFSET)
                    .cast::<*const c_void>()
                    .read_unaligned()
            };
            if next.is_null() {
                return Err(E_POINTER);
            }
            cur = next;
            continue;
        }
        if vt == VT_ARRAY_RECORD {
            // SAFETY: `cur` is a live `0x2024` cell.
            if !unsafe { is_automation_record_array_cell(cur) } {
                // SAFETY: non-null tail identifies an internal runtime `Variant`.
                let cloned = unsafe { &*cur.cast::<Variant>() }
                    .try_clone()
                    .map_err(|_| E_OUTOFMEMORY)?;
                if let Some(expected) = expected_record_array {
                    let actual = cloned.record_array_layout().ok_or(DISP_E_TYPEMISMATCH)?;
                    if !actual.matches_nominal(expected) {
                        return Err(DISP_E_TYPEMISMATCH);
                    }
                }
                return Ok(cloned);
            }
            // SAFETY: payload holds `SAFEARRAY*`.
            let psa = unsafe {
                cur.cast::<u8>()
                    .add(PAYLOAD_OFFSET)
                    .cast::<*mut c_void>()
                    .read_unaligned()
            };
            // SAFETY: `psa` is a live Automation `SAFEARRAY(VT_RECORD)` or null.
            return unsafe {
                match expected_record_array {
                    Some(expected) => {
                        automation_safearray_to_vba_record_array(psa, expected, false, resolve)
                    }
                    None => automation_safearray_to_vba_record_array_any(psa, false, resolve),
                }
            };
        }
        if vt == (VT_BYREF | VT_ARRAY_RECORD) {
            // SAFETY: payload holds `SAFEARRAY**`.
            let pparray = unsafe {
                cur.cast::<u8>()
                    .add(PAYLOAD_OFFSET)
                    .cast::<*const *mut c_void>()
                    .read_unaligned()
            };
            if pparray.is_null() {
                return Err(E_POINTER);
            }
            // SAFETY: `pparray` points to caller's live `SAFEARRAY*` slot; `*pparray` is a live
            // Automation `SAFEARRAY(VT_RECORD)` or null.
            return unsafe {
                let psa = *pparray;
                match expected_record_array {
                    Some(expected) => {
                        automation_safearray_to_vba_record_array(psa, expected, true, resolve)
                    }
                    None => automation_safearray_to_vba_record_array_any(psa, true, resolve),
                }
            };
        }
        if expected_record_array.is_some() {
            return Err(DISP_E_TYPEMISMATCH);
        }
        if vt == VT_ARRAY_VARIANT || vt == (VT_BYREF | VT_ARRAY_VARIANT) {
            // SAFETY: read `SAFEARRAY*` (dereferencing once when `VT_BYREF`).
            let psa = unsafe {
                if vt & VT_BYREF != 0 {
                    let pparray = cur
                        .cast::<u8>()
                        .add(PAYLOAD_OFFSET)
                        .cast::<*const *mut c_void>()
                        .read_unaligned();
                    if pparray.is_null() {
                        return Err(E_POINTER);
                    }
                    *pparray
                } else {
                    cur.cast::<u8>()
                        .add(PAYLOAD_OFFSET)
                        .cast::<*mut c_void>()
                        .read_unaligned()
                }
            };
            // SAFETY: `psa` is a live `SAFEARRAY(VT_VARIANT)` when non-null.
            if !psa.is_null() && unsafe { variant_safearray_contains_record_array(psa, true, 0) } {
                // SAFETY: `psa` is a live Automation `SAFEARRAY(VT_VARIANT)`.
                let recovered = unsafe { recover_variant_safearray_from_automation(psa, resolve) }?;
                return Ok(Variant::from_safearray(recovered));
            }
        }
        let mut copied = Variant::empty();
        // SAFETY: `cur` is a live non-record-array Automation `VARIANT`; `copied` is `VT_EMPTY`.
        let hr = unsafe { VariantCopyInd((&raw mut copied).cast::<c_void>(), cur) };
        if hr < 0 {
            return Err(hr);
        }
        return Ok(copied);
    }
    Err(DISP_E_TYPEMISMATCH)
}

/// Owning Automation `VARIANT` cell for staging outbound COM arguments and return values that may
/// carry `SAFEARRAY(VT_RECORD)` (`FADF_RECORD` + nominal `IRecordInfo`, zero tail).
#[repr(transparent)]
pub struct AutomationVariant {
    core: VariantCore,
}

impl AutomationVariant {
    pub fn empty() -> Self {
        Self {
            core: VariantCore::from_bytes(crate::VarType::Empty, [0; 8]),
        }
    }

    pub fn from_runtime(value: &Variant) -> Result<Self, HResult> {
        let mut cell = Self::empty();
        // SAFETY: `cell` is an initialized `VT_EMPTY` Automation cell.
        unsafe { write_runtime_variant_to_automation_cell(cell.as_mut_ptr(), value) }?;
        Ok(cell)
    }

    pub fn raw_vt(&self) -> u16 {
        self.core.raw_vt()
    }

    pub fn as_ptr(&self) -> *const c_void {
        (&raw const self.core).cast()
    }

    pub fn as_mut_ptr(&mut self) -> *mut c_void {
        (&raw mut self.core).cast()
    }

    pub fn to_runtime(
        &self,
        expected_record_array: Option<&Arc<VbaRecordLayout>>,
        resolve: &dyn Fn(&GUID, u32) -> Option<Arc<VbaRecordLayout>>,
    ) -> Result<Variant, HResult> {
        // SAFETY: `self.core` is a live Automation `VARIANT` cell.
        unsafe {
            automation_variant_to_runtime_variant(self.as_ptr(), expected_record_array, resolve)
        }
    }
}

impl Drop for AutomationVariant {
    fn drop(&mut self) {
        // SAFETY: `self.core` is a live Automation `VARIANT` cell.
        let _ = unsafe { clear_automation_variant_cell(self.as_mut_ptr()) };
    }
}

/// Owning Automation `SAFEARRAY(VT_RECORD)` slot (`*mut c_void`) for staging outbound `ByRef`
/// UDT-array arguments across `IDispatch::Invoke` or vtable calls.
pub struct AutomationRecordArraySlot {
    psa: *mut c_void,
    fixed_marker: bool,
}

impl AutomationRecordArraySlot {
    /// Stage a runtime UDT array into an owning Automation `SAFEARRAY(VT_RECORD)` slot.
    ///
    /// # Safety
    /// `record_info` must be null or a valid, live COM `IRecordInfo*`.
    pub unsafe fn from_runtime(
        value: &Variant,
        expected: Option<&Arc<VbaRecordLayout>>,
        record_info: *mut c_void,
    ) -> Result<Self, HResult> {
        // SAFETY: `record_info` is null or live.
        let psa = unsafe {
            vba_record_array_to_automation_safearray_with_info(value, expected, record_info)
        }?;
        let fixed_marker = value.is_fixed_size_array() && !psa.is_null();
        if fixed_marker {
            // SAFETY: `psa` is our freshly allocated descriptor.
            unsafe { crate::oleaut::safe_array_set_fixed_size(psa, true) };
        }
        Ok(Self { psa, fixed_marker })
    }

    pub fn slot_ptr(&mut self) -> *mut *mut c_void {
        &raw mut self.psa
    }

    pub fn raw_psa(&self) -> *mut c_void {
        self.psa
    }

    pub fn to_runtime(
        &self,
        expected: &Arc<VbaRecordLayout>,
        resolve: &dyn Fn(&GUID, u32) -> Option<Arc<VbaRecordLayout>>,
    ) -> Result<Variant, HResult> {
        // SAFETY: `self.psa` is null or a live `SAFEARRAY(VT_RECORD)`.
        unsafe {
            automation_safearray_to_vba_record_array(self.psa, expected, self.fixed_marker, resolve)
        }
    }
}

impl Drop for AutomationRecordArraySlot {
    fn drop(&mut self) {
        if !self.psa.is_null() {
            if self.fixed_marker {
                // SAFETY: undo our fixed-size marker so `SafeArrayDestroy` can release it.
                unsafe { crate::oleaut::safe_array_set_fixed_size(self.psa, false) };
            }
            // SAFETY: owns `self.psa`.
            let _ = unsafe { SafeArrayDestroy(self.psa) };
            self.psa = ptr::null_mut();
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{VbaRecordFieldSpec, current_thread_live_handle_counts};

    fn pair_layout(name: &str, guid: GUID) -> Arc<VbaRecordLayout> {
        Arc::new(
            VbaRecordLayout::new(vec![
                VbaRecordFieldSpec::named("Id", VbaRecordFieldKind::Long),
                VbaRecordFieldSpec::named("Label", VbaRecordFieldKind::String),
            ])
            .unwrap()
            .with_identity(name, guid),
        )
    }

    fn make_pair(layout: &Arc<VbaRecordLayout>, id: i32, label: &str) -> VbaRecord {
        let mut rec = VbaRecord::new_default(Arc::clone(layout)).unwrap();
        rec.write_field_variant(0, &Variant::from_i32(id)).unwrap();
        rec.write_field_variant(1, &Variant::from_string(label))
            .unwrap();
        rec
    }

    #[test]
    fn record_info_exposes_nominal_name_guid_size_and_matching() {
        let g1 = GUID::deterministic("Project1", "Record:slotpair");
        let g2 = GUID::deterministic("Project1", "Record:otherpair");
        let l1 = pair_layout("SlotPair", g1);
        let l2 = pair_layout("OtherPair", g2);

        let before = current_thread_live_handle_counts();
        let info1 = create_record_info_for_layout(&l1, "Fallback", GUID::default()).unwrap();
        let info1_dup = create_record_info_for_layout(&l1, "Fallback", GUID::default()).unwrap();
        let info2 = create_record_info_for_layout(&l2, "Fallback", GUID::default()).unwrap();
        // SAFETY: live `IRecordInfo*` pointers created above.
        unsafe {
            assert_eq!(record_info_get_size(info1), Ok(l1.size() as u32));
            assert_eq!(record_info_get_guid(info1), Ok(g1));
            assert_eq!(record_info_get_name(info1).as_deref(), Ok("SlotPair"));
            assert_eq!(record_info_get_name(info2).as_deref(), Ok("OtherPair"));
            assert!(record_info_is_matching_type(info1, info1));
            assert!(record_info_is_matching_type(info1, info1_dup));
            assert!(!record_info_is_matching_type(info1, info2));
            assert_eq!(
                layout_from_record_info(info1).and_then(|l| l.type_guid()),
                Some(g1)
            );

            // Verify RecordCreate / RecordCopy / RecordClear / RecordDestroy with null BSTR too.
            let table = *info1.cast::<*const RecordVtable>();
            let rec_a = ((*table).record_create)(info1);
            assert!(!rec_a.is_null());
            let mut rec_b = ptr::null_mut();
            assert_eq!(
                ((*table).record_create_copy)(info1, rec_a, &mut rec_b),
                S_OK
            );
            assert!(!rec_b.is_null());
            // Null BSTR stays null on RecordCopy.
            let bstr_offset = l1.fields()[1].offset();
            assert!(
                rec_b
                    .cast::<u8>()
                    .add(bstr_offset)
                    .cast::<*mut u16>()
                    .read()
                    .is_null()
            );
            assert_eq!(((*table).record_destroy)(info1, rec_b), S_OK);
            assert_eq!(((*table).record_destroy)(info1, rec_a), S_OK);

            release_record_info(info1);
            release_record_info(info1_dup);
            release_record_info(info2);
        }
        assert_eq!(
            before.balance_to(current_thread_live_handle_counts()),
            Default::default()
        );
    }

    #[test]
    fn safearray_vt_record_round_trips_nd_rejects_wrong_guid_and_handles_fixed_resize() {
        let g1 = GUID::deterministic("Project1", "Record:slotpair");
        let g2 = GUID::deterministic("Project1", "Record:otherpair");
        let l1 = pair_layout("SlotPair", g1);
        let l2 = pair_layout("OtherPair", g2);
        let no_resolve = |_: &GUID, _: u32| None;

        let before = current_thread_live_handle_counts();
        {
            let bounds = vec![
                SafeArrayBound { count: 2, lower: 1 },
                SafeArrayBound { count: 2, lower: 3 },
            ];
            let elems = vec![
                make_pair(&l1, 10, "a"),
                make_pair(&l1, 20, ""),
                make_pair(&l1, 30, "c"),
                make_pair(&l1, 40, "d"),
            ];
            let runtime = Variant::from_safearray(
                SafeArray::from_vba_records_nd(bounds, Arc::clone(&l1), elems).unwrap(),
            );

            let psa = vba_record_array_to_automation_safearray(&runtime, Some(&l1)).unwrap();
            assert!(!psa.is_null());

            // Reject identical-layout OtherPair BEFORE copying data!
            // SAFETY: `psa` is a live `SAFEARRAY(VT_RECORD)`.
            let mismatch =
                unsafe { automation_safearray_to_vba_record_array(psa, &l2, false, &no_resolve) };
            assert_eq!(mismatch, Err(DISP_E_TYPEMISMATCH));

            // Accept nominal SlotPair and check bounds + elements.
            // SAFETY: `psa` is a live `SAFEARRAY(VT_RECORD)`.
            let back =
                unsafe { automation_safearray_to_vba_record_array(psa, &l1, false, &no_resolve) }
                    .unwrap();
            assert_eq!(back.safearray_bounds_len().unwrap().0.len(), 2);
            let e0 = back.safearray_element(0).unwrap().unwrap();
            assert_eq!(e0.read_record_field_variant(0).unwrap().as_i32(), Some(10));
            assert_eq!(
                e0.read_record_field_variant(1)
                    .unwrap()
                    .as_bstr()
                    .unwrap()
                    .as_str(),
                "a"
            );
            let e3 = back.safearray_element(3).unwrap().unwrap();
            assert_eq!(e3.read_record_field_variant(0).unwrap().as_i32(), Some(40));
            assert_eq!(
                e3.read_record_field_variant(1)
                    .unwrap()
                    .as_bstr()
                    .unwrap()
                    .as_str(),
                "d"
            );

            // In-place writeback into a fixed-size slot with same bounds succeeds; resize fails.
            let mut slot = psa;
            // SAFETY: mark `slot` fixed for this test.
            unsafe { crate::oleaut::safe_array_set_fixed_size(slot, true) };
            let updated_same_shape = Variant::from_safearray(
                SafeArray::from_vba_records_nd(
                    vec![
                        SafeArrayBound { count: 2, lower: 1 },
                        SafeArrayBound { count: 2, lower: 3 },
                    ],
                    Arc::clone(&l1),
                    vec![
                        make_pair(&l1, 99, "updated"),
                        make_pair(&l1, 21, "b"),
                        make_pair(&l1, 31, "c"),
                        make_pair(&l1, 41, "d"),
                    ],
                )
                .unwrap(),
            );
            // SAFETY: `slot` is live and writable.
            unsafe {
                write_vba_record_array_to_automation_slot(&mut slot, &updated_same_shape, &l1)
            }
            .unwrap();
            assert_eq!(
                slot, psa,
                "in-place copy-back must preserve descriptor pointer"
            );

            let resized = Variant::from_safearray(
                SafeArray::from_vba_records_nd(
                    vec![SafeArrayBound { count: 1, lower: 0 }],
                    Arc::clone(&l1),
                    vec![make_pair(&l1, 1, "one")],
                )
                .unwrap(),
            );
            // SAFETY: `slot` is fixed-size; resize must fail with DISP_E_TYPEMISMATCH, then clear
            // fixed marker, resize dynamically, and erase back to null.
            unsafe {
                assert_eq!(
                    write_vba_record_array_to_automation_slot(&mut slot, &resized, &l1),
                    Err(DISP_E_TYPEMISMATCH)
                );
                crate::oleaut::safe_array_set_fixed_size(slot, false);
                write_vba_record_array_to_automation_slot(&mut slot, &resized, &l1).unwrap();
                assert!(!slot.is_null());
                let unalloc = Variant::unallocated_record_array(Arc::clone(&l1));
                write_vba_record_array_to_automation_slot(&mut slot, &unalloc, &l1).unwrap();
                assert!(slot.is_null());
            }
        }
        assert_eq!(
            before.balance_to(current_thread_live_handle_counts()),
            Default::default()
        );
    }

    #[test]
    fn automation_variant_round_trips_nested_udt_array_in_variant_field() {
        let g_inner = GUID::deterministic("Project1", "Record:innerpair");
        let g_outer = GUID::deterministic("Project1", "Record:outerbox");
        let inner = pair_layout("InnerPair", g_inner);
        let outer = Arc::new(
            VbaRecordLayout::new(vec![
                VbaRecordFieldSpec::named("Tag", VbaRecordFieldKind::Long),
                VbaRecordFieldSpec::named("Payload", VbaRecordFieldKind::Variant),
                VbaRecordFieldSpec::named(
                    "Items",
                    VbaRecordFieldKind::DynamicArray {
                        element: Box::new(VbaRecordFieldKind::Record(Arc::clone(&inner))),
                    },
                ),
            ])
            .unwrap()
            .with_identity("OuterBox", g_outer),
        );
        let no_resolve = |_: &GUID, _: u32| None;

        let before = current_thread_live_handle_counts();
        {
            let inner_arr = Variant::from_safearray(
                SafeArray::from_vba_records_nd(
                    vec![SafeArrayBound { count: 1, lower: 0 }],
                    Arc::clone(&inner),
                    vec![make_pair(&inner, 7, "nested")],
                )
                .unwrap(),
            );
            let mut rec = VbaRecord::new_default(Arc::clone(&outer)).unwrap();
            rec.write_field_variant(0, &Variant::from_i32(55)).unwrap();
            rec.write_field_variant(1, &inner_arr).unwrap();
            rec.write_field_variant(2, &inner_arr).unwrap();

            let outer_arr = Variant::from_safearray(
                SafeArray::from_vba_records_nd(
                    vec![SafeArrayBound { count: 1, lower: 1 }],
                    Arc::clone(&outer),
                    vec![rec],
                )
                .unwrap(),
            );

            let mut staged = AutomationVariant::from_runtime(&outer_arr).unwrap();
            assert_eq!(staged.raw_vt(), VT_ARRAY_RECORD);
            // Also test `VariantCopyInd` on the Automation `VARIANT` cell!
            let mut copy_cell = AutomationVariant::empty();
            // SAFETY: both are live Automation `VARIANT` cells.
            let copy_hr = unsafe { VariantCopyInd(copy_cell.as_mut_ptr(), staged.as_ptr()) };
            assert_eq!(copy_hr, S_OK);
            let recovered = copy_cell.to_runtime(None, &no_resolve).unwrap();
            let elem = recovered.safearray_element(0).unwrap().unwrap();
            assert_eq!(
                elem.read_record_field_variant(0).unwrap().as_i32(),
                Some(55)
            );
            let payload_arr = elem.read_record_field_variant(1).unwrap();
            let nested_elem = payload_arr.safearray_element(0).unwrap().unwrap();
            assert_eq!(
                nested_elem.read_record_field_variant(0).unwrap().as_i32(),
                Some(7)
            );
            assert_eq!(
                nested_elem
                    .read_record_field_variant(1)
                    .unwrap()
                    .as_bstr()
                    .unwrap()
                    .as_str(),
                "nested"
            );
            let items_arr = elem.read_record_field_variant(2).unwrap();
            let item0 = items_arr.safearray_element(0).unwrap().unwrap();
            assert_eq!(
                item0.read_record_field_variant(0).unwrap().as_i32(),
                Some(7)
            );

            // Overwrite `staged` with a plain scalar and make sure prior Automation `0x2024` clears cleanly.
            // SAFETY: `staged` is a live Automation cell.
            unsafe {
                write_runtime_variant_to_automation_cell(
                    staged.as_mut_ptr(),
                    &Variant::from_i32(123),
                )
            }
            .unwrap();
            assert_eq!(
                staged.to_runtime(None, &no_resolve).unwrap().as_i32(),
                Some(123)
            );
        }
        assert_eq!(
            before.balance_to(current_thread_live_handle_counts()),
            Default::default()
        );
    }
}
