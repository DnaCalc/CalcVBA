//! A COM record (`VT_RECORD`) value, owned the way oleaut32 owns the `BRECORD` of a `VARIANT`.
//!
//! The record memory comes from the task allocator and the value holds one `IRecordInfo`
//! reference. Copying follows `VariantCopy` (measured in Excel: `GetSize`, allocate,
//! `RecordCopy`, `AddRef`), and dropping follows `VariantClear` (`RecordClear`, free the
//! memory, `Release`). A runtime `VT_RECORD` cell holds these two pointers as its `BRECORD`,
//! so a cell can be handed to oleaut32 unchanged and cleared by it.

use calcvba_automation::com::IRecordInfoVtbl;
use core::ffi::c_void;

/// An owned COM record: `pvRecord` from the task allocator and one `IRecordInfo` reference.
pub struct ComRecord {
    record_data: *mut c_void,
    record_info: *mut c_void,
}

/// A COM record borrowed from a live `VT_RECORD` cell.
#[derive(Clone, Copy)]
pub struct ComRecordRef<'a> {
    record_data: *mut c_void,
    record_info: *mut c_void,
    _cell: core::marker::PhantomData<&'a ComRecord>,
}

/// # Safety
/// `record_info` must be a live `IRecordInfo` pointer.
unsafe fn vtable<'a>(record_info: *mut c_void) -> &'a IRecordInfoVtbl {
    // SAFETY: every COM interface starts with its vtable pointer.
    unsafe { &**record_info.cast::<*const IRecordInfoVtbl>() }
}

/// # Safety
/// `record_info` must be a live `IRecordInfo` pointer.
unsafe fn record_size(record_info: *mut c_void) -> Result<usize, String> {
    let mut size = 0u32;
    // SAFETY: guaranteed by the caller.
    let hr = unsafe { (vtable(record_info).get_size)(record_info, &mut size) };
    if hr < 0 {
        return Err(format!(
            "IRecordInfo::GetSize failed with HRESULT {:#010X}",
            hr as u32
        ));
    }
    Ok(size as usize)
}

/// Zeroed record memory from the task allocator (`CoTaskMemAlloc`, which `VariantClear`
/// frees), or the global allocator where there is no COM.
fn allocate_record(size: usize) -> Result<*mut c_void, String> {
    #[cfg(target_os = "windows")]
    {
        // SAFETY: plain allocation; a zero-size request still returns a unique block.
        let memory = unsafe { windows_sys::Win32::System::Com::CoTaskMemAlloc(size.max(1)) };
        if memory.is_null() {
            return Err("CoTaskMemAlloc failed for a COM record".to_string());
        }
        // SAFETY: `memory` is a fresh block of at least `size` bytes.
        unsafe { core::ptr::write_bytes(memory.cast::<u8>(), 0, size) };
        Ok(memory)
    }
    #[cfg(not(target_os = "windows"))]
    {
        let layout = fallback_layout(size)?;
        // SAFETY: the layout has a non-zero size.
        let memory = unsafe { std::alloc::alloc_zeroed(layout) };
        if memory.is_null() {
            return Err("allocation failed for a COM record".to_string());
        }
        Ok(memory.cast())
    }
}

#[cfg(not(target_os = "windows"))]
fn fallback_layout(size: usize) -> Result<std::alloc::Layout, String> {
    std::alloc::Layout::from_size_align(size.max(1), 16)
        .map_err(|_| "COM record size overflows".to_string())
}

/// # Safety
/// `memory` must come from [`allocate_record`] with the record size of `record_info`.
unsafe fn free_record(memory: *mut c_void, record_info: *mut c_void) {
    #[cfg(target_os = "windows")]
    {
        let _ = record_info;
        // SAFETY: guaranteed by the caller.
        unsafe { windows_sys::Win32::System::Com::CoTaskMemFree(memory) };
    }
    #[cfg(not(target_os = "windows"))]
    {
        // SAFETY: guaranteed by the caller; the record size is fixed for its type.
        let size = unsafe { record_size(record_info) }.unwrap_or(0);
        if let Ok(layout) = fallback_layout(size) {
            // SAFETY: allocated by `allocate_record` with this layout.
            unsafe { std::alloc::dealloc(memory.cast(), layout) };
        }
    }
}

impl ComRecord {
    /// A new record of `record_info`'s type, initialised by `RecordInit`. Takes over the
    /// caller's reference to `record_info`, also on failure.
    ///
    /// # Safety
    /// `record_info` must be a live `IRecordInfo` pointer whose reference the caller owns.
    pub unsafe fn create(record_info: *mut c_void) -> Result<Self, String> {
        if record_info.is_null() {
            return Err("COM record carried null IRecordInfo".to_string());
        }
        // SAFETY: guaranteed by the caller; the reference is released on every error path.
        unsafe {
            let created = record_size(record_info)
                .and_then(allocate_record)
                .and_then(|memory| {
                    let hr = (vtable(record_info).record_init)(record_info, memory);
                    if hr < 0 {
                        free_record(memory, record_info);
                        return Err(format!(
                            "IRecordInfo::RecordInit failed with HRESULT {:#010X}",
                            hr as u32
                        ));
                    }
                    Ok(memory)
                });
            match created {
                Ok(record_data) => Ok(Self {
                    record_data,
                    record_info,
                }),
                Err(error) => {
                    (vtable(record_info).unknown.release)(record_info);
                    Err(error)
                }
            }
        }
    }

    /// A copy of a borrowed record, as `VariantCopy` makes one: own memory, `RecordCopy`, and
    /// a new reference to `record_info`.
    ///
    /// # Safety
    /// `record_info` must be a live `IRecordInfo` pointer and `record_data` a live record of
    /// its type.
    pub unsafe fn copy_from(
        record_info: *mut c_void,
        record_data: *const c_void,
    ) -> Result<Self, String> {
        if record_info.is_null() {
            return Err("COM record carried null IRecordInfo".to_string());
        }
        if record_data.is_null() {
            return Err("COM record carried null record data".to_string());
        }
        // SAFETY: guaranteed by the caller.
        unsafe {
            let memory = allocate_record(record_size(record_info)?)?;
            let table = vtable(record_info);
            let hr = (table.record_copy)(record_info, record_data.cast_mut(), memory);
            if hr < 0 {
                free_record(memory, record_info);
                return Err(format!(
                    "IRecordInfo::RecordCopy failed with HRESULT {:#010X}",
                    hr as u32
                ));
            }
            (table.unknown.add_ref)(record_info);
            Ok(Self {
                record_data: memory,
                record_info,
            })
        }
    }

    /// Takes over a record: `record_data` from the task allocator (as `VariantClear` expects)
    /// and one reference to `record_info`.
    ///
    /// # Safety
    /// As described; both pointers non-null and live, owned by the caller.
    pub unsafe fn from_raw_parts(record_info: *mut c_void, record_data: *mut c_void) -> Self {
        Self {
            record_data,
            record_info,
        }
    }

    pub fn record_info_ptr(&self) -> *mut c_void {
        self.record_info
    }

    pub fn record_data_ptr(&self) -> *mut c_void {
        self.record_data
    }

    pub fn as_ref(&self) -> ComRecordRef<'_> {
        ComRecordRef {
            record_data: self.record_data,
            record_info: self.record_info,
            _cell: core::marker::PhantomData,
        }
    }

    pub fn try_clone(&self) -> Result<Self, String> {
        self.as_ref().try_to_owned()
    }

    /// Gives up ownership: `(record_info, record_data)`, released by `VariantClear` semantics.
    pub fn into_raw_parts(self) -> (*mut c_void, *mut c_void) {
        let parts = (self.record_info, self.record_data);
        core::mem::forget(self);
        parts
    }
}

impl<'a> ComRecordRef<'a> {
    /// # Safety
    /// Both pointers must stay live for `'a`, `record_data` a record of `record_info`'s type.
    pub unsafe fn from_raw_parts(record_info: *mut c_void, record_data: *mut c_void) -> Self {
        Self {
            record_data,
            record_info,
            _cell: core::marker::PhantomData,
        }
    }

    pub fn record_info_ptr(&self) -> *mut c_void {
        self.record_info
    }

    pub fn record_data_ptr(&self) -> *mut c_void {
        self.record_data
    }

    /// A `VariantCopy`-style copy of the borrowed record.
    pub fn try_to_owned(&self) -> Result<ComRecord, String> {
        // SAFETY: a `ComRecordRef` borrows a live record and its `IRecordInfo`.
        unsafe { ComRecord::copy_from(self.record_info, self.record_data) }
    }
}

impl Drop for ComRecord {
    fn drop(&mut self) {
        if self.record_info.is_null() {
            return;
        }
        // SAFETY: an owned record: memory from `allocate_record` (or adopted under the same
        // contract) and one `IRecordInfo` reference, released once here as `VariantClear` does.
        unsafe {
            let table = vtable(self.record_info);
            if !self.record_data.is_null() {
                let _ = (table.record_clear)(self.record_info, self.record_data);
                free_record(self.record_data, self.record_info);
            }
            (table.unknown.release)(self.record_info);
        }
    }
}

impl Clone for ComRecord {
    fn clone(&self) -> Self {
        self.try_clone().expect("COM record copy failed")
    }
}

impl core::fmt::Debug for ComRecord {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        self.as_ref().fmt(f)
    }
}

impl core::fmt::Debug for ComRecordRef<'_> {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        f.debug_struct("ComRecord")
            .field("record_info", &self.record_info)
            .field("record_data", &self.record_data)
            .finish()
    }
}

impl PartialEq for ComRecord {
    fn eq(&self, other: &Self) -> bool {
        self.record_info == other.record_info && self.record_data == other.record_data
    }
}

impl Eq for ComRecord {}

/// A minimal `IRecordInfo` for tests: records of `size` plain bytes, with counts of live
/// references and live records.
#[doc(hidden)]
pub mod test_support {
    use calcvba_automation::com::{Guid, HResult, IRecordInfoVtbl, IUnknownVtbl};
    use core::ffi::c_void;
    use std::sync::atomic::{AtomicIsize, Ordering};

    const E_NOTIMPL: HResult = 0x8000_4001_u32 as i32;

    /// Every instance, so none is unreachable (Miri's leak check) while tests still use it.
    static INSTANCES: std::sync::Mutex<Vec<&'static TestRecordInfo>> =
        std::sync::Mutex::new(Vec::new());

    #[repr(C)]
    pub struct TestRecordInfo {
        vtbl: *const IRecordInfoVtbl,
        size: u32,
        references: AtomicIsize,
        live_records: AtomicIsize,
    }

    // SAFETY: the vtable pointer names an immutable static and the counts are atomics.
    unsafe impl Sync for TestRecordInfo {}
    // SAFETY: as for `Sync`; nothing in the instance is tied to a thread.
    unsafe impl Send for TestRecordInfo {}

    impl TestRecordInfo {
        /// An instance holding one reference for the caller. It lives for the whole process
        /// (kept in [`INSTANCES`]), so a test can read its counts after the last release.
        pub fn create(size: u32) -> *mut c_void {
            let info: &'static Self = Box::leak(Box::new(Self {
                vtbl: &VTBL,
                size,
                references: AtomicIsize::new(1),
                live_records: AtomicIsize::new(0),
            }));
            INSTANCES
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .push(info);
            core::ptr::from_ref(info).cast_mut().cast()
        }

        /// # Safety
        /// `this` must come from [`Self::create`].
        pub unsafe fn references(this: *mut c_void) -> isize {
            // SAFETY: guaranteed by the caller.
            unsafe { (*this.cast::<Self>()).references.load(Ordering::Acquire) }
        }

        /// Records initialised or copied and not yet cleared.
        ///
        /// # Safety
        /// `this` must come from [`Self::create`].
        pub unsafe fn live_records(this: *mut c_void) -> isize {
            // SAFETY: guaranteed by the caller.
            unsafe { (*this.cast::<Self>()).live_records.load(Ordering::Acquire) }
        }
    }

    unsafe fn me<'a>(this: *mut c_void) -> &'a TestRecordInfo {
        // SAFETY: only reached through `VTBL`, whose `this` is a `TestRecordInfo`.
        unsafe { &*this.cast::<TestRecordInfo>() }
    }

    unsafe extern "system" fn query_interface(
        _this: *mut c_void,
        _iid: *const Guid,
        out: *mut *mut c_void,
    ) -> HResult {
        // SAFETY: the out pointer is supplied by the caller.
        unsafe { *out = core::ptr::null_mut() };
        E_NOTIMPL
    }

    unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
        // SAFETY: see `me`.
        (unsafe { me(this) }
            .references
            .fetch_add(1, Ordering::AcqRel)
            + 1) as u32
    }

    unsafe extern "system" fn release(this: *mut c_void) -> u32 {
        // The instance lives for the process (`INSTANCES`) so a test can read its counts.
        // SAFETY: see `me`.
        (unsafe { me(this) }
            .references
            .fetch_sub(1, Ordering::AcqRel)
            - 1) as u32
    }

    unsafe extern "system" fn record_init(this: *mut c_void, record: *mut c_void) -> HResult {
        // SAFETY: see `me`; `record` holds `size` writable bytes.
        unsafe {
            core::ptr::write_bytes(record.cast::<u8>(), 0, me(this).size as usize);
            me(this).live_records.fetch_add(1, Ordering::AcqRel);
        }
        0
    }

    unsafe extern "system" fn record_clear(this: *mut c_void, _record: *mut c_void) -> HResult {
        // SAFETY: see `me`.
        unsafe { me(this) }
            .live_records
            .fetch_sub(1, Ordering::AcqRel);
        0
    }

    unsafe extern "system" fn record_copy(
        this: *mut c_void,
        existing: *mut c_void,
        new: *mut c_void,
    ) -> HResult {
        // SAFETY: see `me`; both records hold `size` bytes.
        unsafe {
            core::ptr::copy_nonoverlapping(
                existing.cast::<u8>(),
                new.cast::<u8>(),
                me(this).size as usize,
            );
            me(this).live_records.fetch_add(1, Ordering::AcqRel);
        }
        0
    }

    unsafe extern "system" fn get_guid(_this: *mut c_void, _guid: *mut Guid) -> HResult {
        E_NOTIMPL
    }

    unsafe extern "system" fn get_name(_this: *mut c_void, _name: *mut *mut u16) -> HResult {
        E_NOTIMPL
    }

    unsafe extern "system" fn get_size(this: *mut c_void, size: *mut u32) -> HResult {
        // SAFETY: see `me`; the out pointer is supplied by the caller.
        unsafe { *size = me(this).size };
        0
    }

    unsafe extern "system" fn get_type_info(_this: *mut c_void, _out: *mut *mut c_void) -> HResult {
        E_NOTIMPL
    }

    unsafe extern "system" fn get_field(
        _this: *mut c_void,
        _data: *mut c_void,
        _name: *const u16,
        _field: *mut c_void,
    ) -> HResult {
        E_NOTIMPL
    }

    unsafe extern "system" fn get_field_no_copy(
        _this: *mut c_void,
        _data: *mut c_void,
        _name: *const u16,
        _field: *mut c_void,
        _carray: *mut *mut c_void,
    ) -> HResult {
        E_NOTIMPL
    }

    unsafe extern "system" fn put_field(
        _this: *mut c_void,
        _flags: u32,
        _data: *mut c_void,
        _name: *const u16,
        _field: *mut c_void,
    ) -> HResult {
        E_NOTIMPL
    }

    unsafe extern "system" fn get_field_names(
        _this: *mut c_void,
        _count: *mut u32,
        _names: *mut *mut u16,
    ) -> HResult {
        E_NOTIMPL
    }

    unsafe extern "system" fn is_matching_type(this: *mut c_void, other: *mut c_void) -> i32 {
        i32::from(this == other)
    }

    unsafe extern "system" fn record_create(_this: *mut c_void) -> *mut c_void {
        core::ptr::null_mut()
    }

    unsafe extern "system" fn record_create_copy(
        _this: *mut c_void,
        _source: *mut c_void,
        _dest: *mut *mut c_void,
    ) -> HResult {
        E_NOTIMPL
    }

    unsafe extern "system" fn record_destroy(_this: *mut c_void, _record: *mut c_void) -> HResult {
        E_NOTIMPL
    }

    static VTBL: IRecordInfoVtbl = IRecordInfoVtbl {
        unknown: IUnknownVtbl {
            query_interface,
            add_ref,
            release,
        },
        record_init,
        record_clear,
        record_copy,
        get_guid,
        get_name,
        get_size,
        get_type_info,
        get_field,
        get_field_no_copy,
        put_field,
        put_field_no_copy: put_field,
        get_field_names,
        is_matching_type,
        record_create,
        record_create_copy,
        record_destroy,
    };
}

#[cfg(test)]
mod tests {
    use super::test_support::TestRecordInfo;
    use super::*;

    #[test]
    fn create_copy_and_drop_follow_the_variant_protocol() {
        let info = TestRecordInfo::create(8);
        // SAFETY: `info` is a live test IRecordInfo whose reference moves into the record.
        let record = unsafe { ComRecord::create(info) }.expect("create");
        // SAFETY: `record` owns eight writable bytes.
        unsafe { *record.record_data_ptr().cast::<i64>() = 42 };
        let copy = record.try_clone().expect("copy");
        assert_ne!(copy.record_data_ptr(), record.record_data_ptr());
        // SAFETY: the copy owns eight bytes copied by RecordCopy.
        assert_eq!(unsafe { *copy.record_data_ptr().cast::<i64>() }, 42);
        // SAFETY: `info` is leaked by the test helper and stays readable.
        unsafe {
            assert_eq!(TestRecordInfo::references(info), 2);
            assert_eq!(TestRecordInfo::live_records(info), 2);
        }
        drop(record);
        drop(copy);
        // SAFETY: as above.
        unsafe {
            assert_eq!(TestRecordInfo::references(info), 0);
            assert_eq!(TestRecordInfo::live_records(info), 0);
        }
    }
}
