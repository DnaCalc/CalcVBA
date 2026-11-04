//! Windows Automation record owners and registration-free IRecordInfo for
//! SAFEARRAY(VT_RECORD). The descriptor contains native byte extents and owning
//! field offsets; canonical VariantCore bytes never cross this boundary.
use std::{
    ffi::c_void,
    ptr,
    sync::atomic::{AtomicU32, Ordering},
};
use windows_sys::{
    Win32::{
        Foundation::{SysAllocStringByteLen, SysFreeString, SysStringByteLen},
        System::{
            Com::{CoTaskMemAlloc, CoTaskMemFree},
            Variant::{VARIANT, VariantClear, VariantCopyInd},
        },
    },
    core::GUID,
};

const E_NOINTERFACE: i32 = 0x80004002u32 as i32;
const E_NOTIMPL: i32 = 0x80004001u32 as i32;
const E_OUTOFMEMORY: i32 = 0x8007000eu32 as i32;

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
/// cell is initialized native record storage matching the retained descriptor.
pub unsafe fn clear_record(cell: *mut u8, descriptor: *const u32) -> i32 {
    // SAFETY: each owner offset and width was checked by the compiler.
    unsafe {
        let (size, fields) = shape(descriptor);
        for field in fields.chunks_exact(2) {
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
                _ => return 0x80070057u32 as i32,
            }
        }
        ptr::write_bytes(cell, 0, size);
    }
    0
}
/// # Safety
/// Both records are initialized and have the exact descriptor shape. Aliasing
/// is permitted. All destination owners remain valid on a failed clone.
pub unsafe fn copy_record(dst: *mut u8, src: *const u8, descriptor: *const u32) -> i32 {
    // SAFETY: stage every owner before replacing destination storage.
    unsafe {
        let (size, fields) = shape(descriptor);
        if fields.is_empty() {
            ptr::copy(src, dst, size);
            return 0;
        }
        let mut words = vec![0u64; size.div_ceil(8)];
        let stage = words.as_mut_ptr().cast::<u8>();
        ptr::copy_nonoverlapping(src, stage, size);
        for field in fields.chunks_exact(2) {
            ptr::write_bytes(
                stage.add(field[0] as usize),
                0,
                if field[1] == 12 { 24 } else { 8 },
            );
        }
        for field in fields.chunks_exact(2) {
            let offset = field[0] as usize;
            let hr = if field[1] == 12 {
                VariantCopyInd(stage.add(offset).cast(), src.add(offset).cast())
            } else {
                let source = src.add(offset).cast::<*mut u16>().read();
                let len = SysStringByteLen(source);
                let value = SysAllocStringByteLen(source.cast(), len).cast_mut();
                if len != 0 && value.is_null() {
                    E_OUTOFMEMORY
                } else {
                    stage.add(offset).cast::<*mut u16>().write(value);
                    0
                }
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
    0
}
#[repr(C)]
struct RecordVtable {
    query_interface: unsafe extern "system" fn(
        this: *mut c_void,
        iid: *const windows_sys::core::GUID,
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
    get_guid:
        unsafe extern "system" fn(this: *mut c_void, pguid: *mut windows_sys::core::GUID) -> i32,
    get_name: unsafe extern "system" fn(
        this: *mut c_void,
        pbstrname: *mut windows_sys::core::BSTR,
    ) -> i32,
    get_size: unsafe extern "system" fn(this: *mut c_void, pcbsize: *mut u32) -> i32,
    get_type_info:
        unsafe extern "system" fn(this: *mut c_void, pptypeinfo: *mut *mut c_void) -> i32,
    get_field: unsafe extern "system" fn(
        this: *mut c_void,
        pvdata: *const c_void,
        szfieldname: *const u16,
        pvarfield: *mut VARIANT,
    ) -> i32,
    get_field_no_copy: unsafe extern "system" fn(
        this: *mut c_void,
        pvdata: *const c_void,
        szfieldname: *const u16,
        pvarfield: *mut VARIANT,
        ppvdatacarray: *mut *mut c_void,
    ) -> i32,
    put_field: unsafe extern "system" fn(
        this: *mut c_void,
        wflags: u32,
        pvdata: *mut c_void,
        szfieldname: *const u16,
        pvarfield: *const VARIANT,
    ) -> i32,
    put_field_no_copy: unsafe extern "system" fn(
        this: *mut c_void,
        wflags: u32,
        pvdata: *mut c_void,
        szfieldname: *const u16,
        pvarfield: *const VARIANT,
    ) -> i32,
    get_field_names: unsafe extern "system" fn(
        this: *mut c_void,
        pcnames: *mut u32,
        rgbstrnames: *mut windows_sys::core::BSTR,
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
}

fn protect(work: impl FnOnce() -> i32) -> i32 {
    std::panic::catch_unwind(std::panic::AssertUnwindSafe(work)).unwrap_or(E_OUTOFMEMORY)
}
unsafe extern "system" fn query(this: *mut c_void, iid: *const GUID, out: *mut *mut c_void) -> i32 {
    // SAFETY: COM supplies a live object and writable output cell.
    unsafe {
        out.write(ptr::null_mut());
        let iid = &*iid;
        if (iid.data1 == 0 || iid.data1 == 0x2f)
            && iid.data2 == 0
            && iid.data3 == 0
            && iid.data4 == [0xc0, 0, 0, 0, 0, 0, 0, 0x46]
        {
            addref(this);
            out.write(this);
            0
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
    // SAFETY: the caller allocated GetSize bytes for this record type.
    unsafe {
        ptr::write_bytes(
            out.cast::<u8>(),
            0,
            (*this.cast::<RecordInfo>()).descriptor[0] as usize,
        );
    }
    0
}
unsafe extern "system" fn clear(this: *mut c_void, record: *const c_void) -> i32 {
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
    // SAFETY: live record info and writable GUID output.
    unsafe {
        out.write((*this.cast::<RecordInfo>()).guid);
    }
    0
}
unsafe extern "system" fn name(_this: *mut c_void, out: *mut windows_sys::core::BSTR) -> i32 {
    let units: Vec<_> = "CalcVBA native record".encode_utf16().collect();
    // SAFETY: caller owns the returned independent BSTR.
    unsafe {
        let value = SysAllocStringByteLen(units.as_ptr().cast(), units.len() as u32 * 2);
        out.write(value);
        if value.is_null() { E_OUTOFMEMORY } else { 0 }
    }
}
unsafe extern "system" fn size(this: *mut c_void, out: *mut u32) -> i32 {
    // SAFETY: live info and writable ULONG output.
    unsafe {
        out.write((*this.cast::<RecordInfo>()).descriptor[0]);
    }
    0
}
unsafe extern "system" fn type_info(_this: *mut c_void, out: *mut *mut c_void) -> i32 {
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
    _out: *mut VARIANT,
) -> i32 {
    E_NOTIMPL
}
unsafe extern "system" fn get_field_no_copy(
    _this: *mut c_void,
    _data: *const c_void,
    _name: *const u16,
    _out: *mut VARIANT,
    _array: *mut *mut c_void,
) -> i32 {
    E_NOTIMPL
}
unsafe extern "system" fn put_field(
    _this: *mut c_void,
    _flags: u32,
    _data: *mut c_void,
    _name: *const u16,
    _value: *const VARIANT,
) -> i32 {
    E_NOTIMPL
}
unsafe extern "system" fn field_names(
    _this: *mut c_void,
    count: *mut u32,
    _names: *mut windows_sys::core::BSTR,
) -> i32 {
    // SAFETY: there are no reflective field names in the owning ABI descriptor.
    unsafe {
        count.write(0);
    }
    0
}
unsafe extern "system" fn matches(this: *mut c_void, other: *mut c_void) -> i32 {
    if other.is_null() {
        return 0;
    }
    // SAFETY: COM supplies another live IRecordInfo. Query its public GetGuid.
    unsafe {
        let table = *other.cast::<*const RecordVtable>();
        let mut other_guid: GUID = std::mem::zeroed();
        if ((*table).get_guid)(other, &mut other_guid) < 0 {
            return 0;
        }
        let own = &(*this.cast::<RecordInfo>()).guid;
        i32::from(
            own.data1 == other_guid.data1
                && own.data2 == other_guid.data2
                && own.data3 == other_guid.data3
                && own.data4 == other_guid.data4,
        )
    }
}
unsafe extern "system" fn create(this: *mut c_void) -> *mut c_void {
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
    // SAFETY: allocate/clone before publishing ownership to the caller.
    unsafe {
        out.write(ptr::null_mut());
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
        0
    }
}
unsafe extern "system" fn destroy(this: *mut c_void, record: *const c_void) -> i32 {
    // SAFETY: RecordDestroy consumes this one RecordCreate allocation.
    unsafe {
        let hr = clear(this, record);
        if hr >= 0 {
            CoTaskMemFree(record);
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
/// descriptor spans size/count/owner pairs followed by sixteen identity bytes.
/// Every owner offset/width is valid for that native record type.
/// The returned IRecordInfo owns its metadata; release it with release_record_info.
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
        }))
        .cast()
    }
}
/// # Safety
/// pointer is a retained live IRecordInfo; consumes one reference.
pub unsafe fn release_record_info(pointer: *mut c_void) {
    if !pointer.is_null() {
        // SAFETY: every IRecordInfo starts with the standard COM reference slots.
        unsafe {
            let table = *pointer.cast::<*const RecordVtable>();
            ((*table).release)(pointer);
        }
    }
}
