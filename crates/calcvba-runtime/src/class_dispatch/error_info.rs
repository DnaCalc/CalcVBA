//! Portable `ISupportErrorInfo` face and `IErrorInfo` publisher on `CompatObjectBase`
//! (dc-yqq0.27).
//!
//! Both late-bound (`IDispatch::Invoke` / `ITypeInfo::Invoke`) and early-bound (vtable slot
//! thunks on the default interface and on `Implements` interfaces) calls clear the thread's
//! `IErrorInfo` on entry and publish a rich `RuntimeErrorInfo` on VBA fault before returning
//! `DISP_E_EXCEPTION` or the VBA HRESULT (`0x800A0000 | Err.Number`).

use core::ffi::c_void;
use core::sync::atomic::{AtomicU32, Ordering};

use calcvba_automation::bstr::{BStr, Utf16String};
use calcvba_automation::com::{
    E_NOINTERFACE, E_POINTER, Guid, HResult, IErrorInfoVtbl, IID_IDISPATCH, IID_IERRORINFO,
    IID_IUNKNOWN, ISupportErrorInfoVtbl, IUnknownVtbl, S_FALSE, S_OK,
};
use calcvba_automation::oleaut;

use crate::events::MemberFault;
use crate::object_ref::{
    CompatObjectBase, RuntimeInterfaceId, compat_add_ref, compat_query_interface, compat_release,
};

/// Embedded `ISupportErrorInfo` COM face on `CompatObjectBase`.
///
/// Shares the owning `CompatObjectBase`'s single COM identity (`QI(IUnknown)` returns `&primary`)
/// and single atomic reference count.
#[repr(C)]
pub(crate) struct SupportErrorInfoFace {
    pub(crate) vtbl: *const ISupportErrorInfoVtbl,
    pub(crate) owner: *mut CompatObjectBase,
}

impl SupportErrorInfoFace {
    pub(crate) const fn new_unlinked() -> Self {
        Self {
            vtbl: &SUPPORT_ERROR_INFO_VTBL,
            owner: core::ptr::null_mut(),
        }
    }
}

pub(crate) static SUPPORT_ERROR_INFO_VTBL: ISupportErrorInfoVtbl = ISupportErrorInfoVtbl {
    unknown: IUnknownVtbl {
        query_interface: support_error_info_query_interface,
        add_ref: support_error_info_add_ref,
        release: support_error_info_release,
    },
    interface_supports_error_info: support_error_info_interface_supports_error_info,
};

pub(crate) fn is_support_error_info_query_interface(
    qi: unsafe extern "system" fn(*mut c_void, *const Guid, *mut *mut c_void) -> HResult,
) -> bool {
    core::ptr::fn_addr_eq(qi, SUPPORT_ERROR_INFO_VTBL.unknown.query_interface)
}

unsafe extern "system" fn support_error_info_query_interface(
    this: *mut c_void,
    riid: *const Guid,
    ppv: *mut *mut c_void,
) -> HResult {
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `SupportErrorInfoFace` embedded in `CompatObjectBase`.
    let owner = unsafe { (*this.cast::<SupportErrorInfoFace>()).owner };
    // SAFETY: delegates to `owner`'s primary `QueryInterface`, preserving single COM identity.
    unsafe { compat_query_interface(owner.cast(), riid, ppv) }
}

unsafe extern "system" fn support_error_info_add_ref(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `SupportErrorInfoFace` embedded in `CompatObjectBase`.
    let owner = unsafe { (*this.cast::<SupportErrorInfoFace>()).owner };
    // SAFETY: shares `owner`'s single atomic refcount.
    unsafe { compat_add_ref(owner.cast()) }
}

unsafe extern "system" fn support_error_info_release(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `SupportErrorInfoFace` embedded in `CompatObjectBase`.
    let owner = unsafe { (*this.cast::<SupportErrorInfoFace>()).owner };
    // SAFETY: shares `owner`'s single atomic refcount.
    unsafe { compat_release(owner.cast()) }
}

unsafe extern "system" fn support_error_info_interface_supports_error_info(
    this: *mut c_void,
    riid: *const Guid,
) -> HResult {
    if this.is_null() || riid.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `SupportErrorInfoFace` embedded in `CompatObjectBase`, and `riid`
    // is a non-null readable `Guid`.
    let (iid, descriptor) = unsafe {
        let owner = (*this.cast::<SupportErrorInfoFace>()).owner;
        (riid.read(), (*owner).class_descriptor)
    };
    if iid == IID_IDISPATCH || descriptor.default_interface_iid == Some(iid) {
        return S_OK;
    }
    if descriptor
        .interfaces
        .iter()
        .any(|iface| iface.id == RuntimeInterfaceId::Unsupported && iface.identity.guid == iid)
    {
        return S_OK;
    }
    S_FALSE
}

/// Portable refcounted COM `IErrorInfo` object published via `SetErrorInfo` on VBA faults.
#[repr(C)]
struct RuntimeErrorInfo {
    vtbl: *const IErrorInfoVtbl,
    ref_count: AtomicU32,
    guid: Guid,
    source: Utf16String,
    description: Utf16String,
    help_file: Utf16String,
    help_context: u32,
}

static RUNTIME_ERROR_INFO_VTBL: IErrorInfoVtbl = IErrorInfoVtbl {
    unknown: IUnknownVtbl {
        query_interface: error_info_query_interface,
        add_ref: error_info_add_ref,
        release: error_info_release,
    },
    get_guid: error_info_get_guid,
    get_source: error_info_get_source,
    get_description: error_info_get_description,
    get_help_file: error_info_get_help_file,
    get_help_context: error_info_get_help_context,
};

unsafe extern "system" fn error_info_query_interface(
    this: *mut c_void,
    riid: *const Guid,
    ppv: *mut *mut c_void,
) -> HResult {
    if ppv.is_null() {
        return E_POINTER;
    }
    // SAFETY: `ppv` is non-null writable out-slot.
    unsafe {
        *ppv = core::ptr::null_mut();
    }
    if this.is_null() || riid.is_null() {
        return E_POINTER;
    }
    // SAFETY: `riid` is non-null readable `Guid`.
    let iid = unsafe { riid.read() };
    if iid == IID_IUNKNOWN || iid == IID_IERRORINFO {
        // SAFETY: `this` is a live `RuntimeErrorInfo`; take one reference for `*ppv`.
        unsafe {
            error_info_add_ref(this);
            *ppv = this;
        }
        S_OK
    } else {
        E_NOINTERFACE
    }
}

unsafe extern "system" fn error_info_add_ref(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `RuntimeErrorInfo`.
    unsafe {
        (*this.cast::<RuntimeErrorInfo>())
            .ref_count
            .fetch_add(1, Ordering::AcqRel)
            + 1
    }
}

unsafe extern "system" fn error_info_release(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `RuntimeErrorInfo`; `AcqRel` synchronizes prior writes before drop.
    let previous = unsafe {
        (*this.cast::<RuntimeErrorInfo>())
            .ref_count
            .fetch_sub(1, Ordering::AcqRel)
    };
    let remaining = previous.saturating_sub(1);
    if remaining == 0 {
        // SAFETY: sole final reference to `Box::into_raw` allocation.
        unsafe {
            drop(Box::from_raw(this.cast::<RuntimeErrorInfo>()));
        }
    }
    remaining
}

unsafe extern "system" fn error_info_get_guid(this: *mut c_void, pguid: *mut Guid) -> HResult {
    if this.is_null() || pguid.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `RuntimeErrorInfo` and `pguid` is a writable out-slot.
    unsafe {
        pguid.write((*this.cast::<RuntimeErrorInfo>()).guid);
    }
    S_OK
}

fn alloc_optional_bstr(text: &Utf16String) -> *mut u16 {
    if text.is_empty() {
        core::ptr::null_mut()
    } else {
        BStr::from_utf16_units(text.as_units())
            .map(BStr::into_raw_bstr)
            .unwrap_or(core::ptr::null_mut())
    }
}

unsafe extern "system" fn error_info_get_source(this: *mut c_void, out: *mut *mut u16) -> HResult {
    if this.is_null() || out.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `RuntimeErrorInfo` and `out` is a writable out-slot.
    unsafe {
        *out = alloc_optional_bstr(&(*this.cast::<RuntimeErrorInfo>()).source);
    }
    S_OK
}

unsafe extern "system" fn error_info_get_description(
    this: *mut c_void,
    out: *mut *mut u16,
) -> HResult {
    if this.is_null() || out.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `RuntimeErrorInfo` and `out` is a writable out-slot.
    unsafe {
        *out = alloc_optional_bstr(&(*this.cast::<RuntimeErrorInfo>()).description);
    }
    S_OK
}

unsafe extern "system" fn error_info_get_help_file(
    this: *mut c_void,
    out: *mut *mut u16,
) -> HResult {
    if this.is_null() || out.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `RuntimeErrorInfo` and `out` is a writable out-slot.
    unsafe {
        *out = alloc_optional_bstr(&(*this.cast::<RuntimeErrorInfo>()).help_file);
    }
    S_OK
}

unsafe extern "system" fn error_info_get_help_context(this: *mut c_void, out: *mut u32) -> HResult {
    if this.is_null() || out.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `RuntimeErrorInfo` and `out` is a writable out-slot.
    unsafe {
        *out = (*this.cast::<RuntimeErrorInfo>()).help_context;
    }
    S_OK
}

/// Clear the current thread's `IErrorInfo` slot at entry to an Automation / vtable call.
pub(crate) fn clear_thread_error_info() {
    // SAFETY: passing null clears the current thread's COM error object.
    unsafe {
        let _ = oleaut::set_error_info(0, core::ptr::null_mut());
    }
}

/// Publish a VBA [`MemberFault`] as the current thread's COM `IErrorInfo` object.
pub(crate) fn publish_member_fault_error_info(interface_iid: Guid, fault: &MemberFault) {
    let raw = Box::into_raw(Box::new(RuntimeErrorInfo {
        vtbl: &RUNTIME_ERROR_INFO_VTBL,
        ref_count: AtomicU32::new(1),
        guid: interface_iid,
        source: fault.source.clone(),
        description: fault.description.clone(),
        help_file: fault.help_file.clone(),
        help_context: fault.help_context as u32,
    }))
    .cast::<c_void>();
    // SAFETY: `raw` is a live `IErrorInfo*` with refcount 1; `set_error_info` takes its own
    // reference and we release ours immediately.
    unsafe {
        let _ = oleaut::set_error_info(0, raw);
        error_info_release(raw);
    }
}
