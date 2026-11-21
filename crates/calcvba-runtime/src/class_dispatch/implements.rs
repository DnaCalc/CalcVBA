//! Extra COM vtables on `CompatObjectBase` for `Implements` interfaces (dc-yqq0.27).
//!
//! Every implemented interface on a VBA class (`Implements IShape`, `Implements Excel.IRtdServer`,
//! `Implements AddInDesignerObjects.IDTExtensibility2`, ...) gets its own COM face owned by
//! `CompatObjectBase`:
//! - `QueryInterface` delegates to `owner` (preserving one COM identity `&primary` for `IUnknown`).
//! - `AddRef` and `Release` delegate to `owner` (one shared atomic refcount).
//! - `IDispatch` (slots 3..6) and early-bound member slots (`7..`) dispatch against the
//!   implemented interface's [`RuntimeInterfaceDescriptor`], routing calls to the class's
//!   `Interface_Member` procedures with `Me = owner`.

use core::ffi::c_void;
use std::collections::HashMap;
use std::sync::{Mutex, OnceLock};

use calcvba_automation::com::{
    DISP_E_BADINDEX, DISP_E_UNKNOWNINTERFACE, E_POINTER, Guid, HResult, IID_NULL, IUnknownVtbl,
    S_OK,
};

use crate::object_ref::{
    CompatObjectBase, RuntimeClassDescriptor, RuntimeInterfaceDescriptor, RuntimeInterfaceId,
    compat_add_ref, compat_query_interface, compat_release,
};

use super::{
    SyncVtableSlice, compat_slot_not_implemented, invoke_on_interface,
    resolve_ids_of_names_for_interface, slot_thunks, type_info,
};

/// One implemented-interface COM face owned by `CompatObjectBase`.
///
/// Allocated in a fixed-size heap slice (`*mut [ImplementedInterfaceFace]`) for the lifetime of
/// `CompatObjectBase`, so every face pointer handed out by `QueryInterface` stays stable until
/// teardown.
#[repr(C)]
pub(crate) struct ImplementedInterfaceFace {
    pub(crate) vtbl: *const IUnknownVtbl,
    pub(crate) owner: *mut CompatObjectBase,
    pub(crate) interface: &'static RuntimeInterfaceDescriptor,
}

static IMPL_IUNKNOWN_VTBL: IUnknownVtbl = IUnknownVtbl {
    query_interface: impl_query_interface,
    add_ref: impl_add_ref,
    release: impl_release,
};

static IMPL_VTBLS: OnceLock<Mutex<HashMap<usize, SyncVtableSlice>>> = OnceLock::new();

fn impl_vtable_cache() -> &'static Mutex<HashMap<usize, SyncVtableSlice>> {
    IMPL_VTBLS.get_or_init(|| Mutex::new(HashMap::new()))
}

pub(crate) fn impl_vtable_slot_count(interface: &'static RuntimeInterfaceDescriptor) -> usize {
    impl_vtable_slice(interface).len()
}

fn impl_vtable_slice(interface: &'static RuntimeInterfaceDescriptor) -> &'static [*const c_void] {
    let key = interface as *const RuntimeInterfaceDescriptor as usize;
    let mut cache = impl_vtable_cache()
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    if let Some(existing) = cache.get(&key).copied() {
        return existing.0;
    }
    let member_end_slot = interface
        .members
        .iter()
        .filter_map(|member| member.vtable_slot.map(|slot| usize::from(slot) + 1))
        .max()
        .unwrap_or(7)
        .max(7 + interface.members.len());

    let mut slots: Vec<*const c_void> = Vec::with_capacity(member_end_slot);
    slots.push(IMPL_IUNKNOWN_VTBL.query_interface as *const c_void);
    slots.push(IMPL_IUNKNOWN_VTBL.add_ref as *const c_void);
    slots.push(IMPL_IUNKNOWN_VTBL.release as *const c_void);
    slots.push(
        impl_get_type_info_count as unsafe extern "system" fn(*mut c_void, *mut u32) -> HResult
            as *const c_void,
    );
    slots.push(
        impl_get_type_info
            as unsafe extern "system" fn(*mut c_void, u32, u32, *mut *mut c_void) -> HResult
            as *const c_void,
    );
    slots.push(
        impl_get_ids_of_names
            as unsafe extern "system" fn(
                *mut c_void,
                *const Guid,
                *const *const u16,
                u32,
                u32,
                *mut i32,
            ) -> HResult as *const c_void,
    );
    slots.push(
        impl_invoke
            as unsafe extern "system" fn(
                *mut c_void,
                i32,
                *const Guid,
                u32,
                u16,
                *mut c_void,
                *mut c_void,
                *mut c_void,
                *mut u32,
            ) -> HResult as *const c_void,
    );
    let placeholder = compat_slot_not_implemented
        as unsafe extern "system" fn(*mut c_void) -> HResult as *const c_void;
    while slots.len() < member_end_slot {
        slots.push(slot_thunks::slot_entry(slots.len()).unwrap_or(placeholder));
    }

    let leaked: &'static [*const c_void] = Box::leak(slots.into_boxed_slice());
    cache.insert(key, SyncVtableSlice(leaked));
    leaked
}

/// Build the per-instance array of implemented-interface COM faces for `descriptor` (unlinked;
/// `from_compat_object` stamps `owner` before handing out the object).
pub(crate) fn build_implements_faces(
    descriptor: &'static RuntimeClassDescriptor,
) -> Box<[ImplementedInterfaceFace]> {
    let faces: Vec<ImplementedInterfaceFace> = descriptor
        .interfaces
        .iter()
        .filter(|iface| {
            iface.id == RuntimeInterfaceId::Unsupported && iface.identity.guid != IID_NULL
        })
        .map(|interface| ImplementedInterfaceFace {
            vtbl: impl_vtable_slice(interface).as_ptr().cast::<IUnknownVtbl>(),
            owner: core::ptr::null_mut(),
            interface,
        })
        .collect();
    faces.into_boxed_slice()
}

pub(crate) fn is_impl_query_interface(
    qi: unsafe extern "system" fn(*mut c_void, *const Guid, *mut *mut c_void) -> HResult,
) -> bool {
    core::ptr::fn_addr_eq(qi, IMPL_IUNKNOWN_VTBL.query_interface)
}

unsafe extern "system" fn impl_query_interface(
    this: *mut c_void,
    riid: *const Guid,
    ppv: *mut *mut c_void,
) -> HResult {
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `ImplementedInterfaceFace` owned by `CompatObjectBase`.
    let owner = unsafe { (*this.cast::<ImplementedInterfaceFace>()).owner };
    // SAFETY: delegates to `owner`'s primary `QueryInterface`, preserving single COM identity.
    unsafe { compat_query_interface(owner.cast(), riid, ppv) }
}

unsafe extern "system" fn impl_add_ref(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `ImplementedInterfaceFace`.
    let owner = unsafe { (*this.cast::<ImplementedInterfaceFace>()).owner };
    // SAFETY: shares `owner`'s single atomic refcount.
    unsafe { compat_add_ref(owner.cast()) }
}

unsafe extern "system" fn impl_release(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `ImplementedInterfaceFace`.
    let owner = unsafe { (*this.cast::<ImplementedInterfaceFace>()).owner };
    // SAFETY: shares `owner`'s single atomic refcount.
    unsafe { compat_release(owner.cast()) }
}

unsafe extern "system" fn impl_get_type_info_count(
    this: *mut c_void,
    pctinfo: *mut u32,
) -> HResult {
    if this.is_null() || pctinfo.is_null() {
        return E_POINTER;
    }
    // SAFETY: `pctinfo` is a non-null writable out-slot.
    unsafe {
        *pctinfo = 1;
    }
    S_OK
}

unsafe extern "system" fn impl_get_type_info(
    this: *mut c_void,
    itinfo: u32,
    _lcid: u32,
    pptinfo: *mut *mut c_void,
) -> HResult {
    if this.is_null() || pptinfo.is_null() {
        return E_POINTER;
    }
    // SAFETY: `pptinfo` is a non-null writable out-slot.
    unsafe {
        *pptinfo = core::ptr::null_mut();
    }
    if itinfo != 0 {
        return DISP_E_BADINDEX;
    }
    // SAFETY: `this` is a live `ImplementedInterfaceFace`.
    let interface = unsafe { (*this.cast::<ImplementedInterfaceFace>()).interface };
    let ti = type_info::create_interface_type_info(
        interface.name,
        interface.identity.guid,
        interface,
        impl_vtable_slot_count(interface),
    );
    // SAFETY: `pptinfo` is non-null writable out-slot.
    unsafe {
        *pptinfo = ti;
    }
    S_OK
}

unsafe extern "system" fn impl_get_ids_of_names(
    this: *mut c_void,
    riid: *const Guid,
    names: *const *const u16,
    count: u32,
    _lcid: u32,
    dispids: *mut i32,
) -> HResult {
    if this.is_null() || names.is_null() || dispids.is_null() {
        return E_POINTER;
    }
    if !riid.is_null() {
        // SAFETY: `riid` is non-null readable `Guid`.
        let iid = unsafe { riid.read() };
        if iid != IID_NULL {
            return DISP_E_UNKNOWNINTERFACE;
        }
    }
    // SAFETY: `this` is a live `ImplementedInterfaceFace`.
    let interface = unsafe { (*this.cast::<ImplementedInterfaceFace>()).interface };
    // SAFETY: arguments forwarded from `IDispatch::GetIDsOfNames`.
    unsafe { resolve_ids_of_names_for_interface(interface, names, count, dispids) }
}

#[allow(clippy::too_many_arguments)]
unsafe extern "system" fn impl_invoke(
    this: *mut c_void,
    dispid: i32,
    riid: *const Guid,
    _lcid: u32,
    flags: u16,
    params: *mut c_void,
    result: *mut c_void,
    excep_info: *mut c_void,
    arg_err: *mut u32,
) -> HResult {
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `ImplementedInterfaceFace`.
    let (owner, interface) = unsafe {
        let face = &*this.cast::<ImplementedInterfaceFace>();
        (face.owner, face.interface)
    };
    // SAFETY: `owner` is a live `CompatObjectBase`; dispatch on `interface` with `Me = owner`.
    unsafe {
        invoke_on_interface(
            owner,
            interface,
            interface.identity.guid,
            dispid,
            riid,
            flags,
            params,
            result,
            excep_info,
            arg_err,
        )
    }
}
