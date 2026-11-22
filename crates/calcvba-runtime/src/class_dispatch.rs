//! `ClassDualVtbl` and portable table-driven `IDispatch` on `CompatObjectBase` (dc-yqq0.23).
//!
//! Every VBA project class instance carries a dual vtable built once per
//! [`RuntimeClassDescriptor`]: `IUnknown` (slots 0..=2), `IDispatch` (slots 3..=6), and one
//! entry per member in `vtable_slot` order (slots `7..`). `QueryInterface` answers `IUnknown`,
//! `IDispatch`, and the class's default interface IID by returning `&primary`, giving every
//! project object one COM identity with no wrapper allocation.

use core::ffi::c_void;
use core::sync::atomic::{AtomicU64, Ordering};
use std::collections::HashMap;
use std::sync::{Mutex, OnceLock};

use calcvba_automation::bstr::BStr;
use calcvba_automation::com::{
    DISP_E_BADINDEX, DISP_E_BADPARAMCOUNT, DISP_E_EXCEPTION, DISP_E_MEMBERNOTFOUND,
    DISP_E_PARAMNOTFOUND, DISP_E_PARAMNOTOPTIONAL, DISP_E_TYPEMISMATCH, DISP_E_UNKNOWNINTERFACE,
    DISP_E_UNKNOWNNAME, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT,
    DISPATCH_PROPERTYPUTREF, DISPID_PROPERTYPUT, DISPID_UNKNOWN, DispParams, E_FAIL, E_INVALIDARG,
    E_NOTIMPL, E_POINTER, ExcepInfo, Guid, HResult, IDispatchVtbl, IID_IDISPATCH, IID_NULL,
    IUnknownVtbl, RPC_E_WRONG_THREAD, S_OK,
};
use calcvba_automation::oleaut;
use calcvba_automation::safe_array::SafeArray;

use crate::events::{self, MemberFault};
use crate::object_ref::{
    COMPAT_OBJECT_VTBL, CompatObjectBase, ObjectRef, ObjectRefExt, RuntimeClassDescriptor,
    RuntimeInterfaceDescriptor, RuntimeMemberBacking, RuntimeMemberDescriptor,
    RuntimeMemberInvokeKind, RuntimeValueType, compat_owner_from_this,
};
use crate::{VarType, Variant, VariantCore};

mod connection_point;
mod error_info;
mod implements;
mod slot_thunks;
mod type_info;

pub use slot_thunks::invoke_early_bound_slot_raw;

pub(crate) use connection_point::{
    ConnectionPointContainerFace, ConnectionPointFace, ConnectionPointSink,
    fire_connection_point_event_for_owner, has_connection_point_sinks, is_cpc_query_interface,
    release_all_event_sinks,
};
pub(crate) use error_info::{SupportErrorInfoFace, is_support_error_info_query_interface};
pub(crate) use implements::{
    ImplementedInterfaceFace, build_implements_faces, is_impl_query_interface,
};

static NEXT_THREAD_ID: AtomicU64 = AtomicU64::new(1);

thread_local! {
    static CURRENT_THREAD_ID: u64 = NEXT_THREAD_ID.fetch_add(1, Ordering::Relaxed);
}

/// Portable monotonic identifier for the current OS thread (used for `CompatObjectBase`
/// owner-thread checks and foreign-thread release routing without OS-specific FFI).
pub fn current_thread_id() -> u64 {
    CURRENT_THREAD_ID.with(|id| *id)
}

#[derive(Clone, Copy)]
struct SyncVtableSlice(&'static [*const c_void]);

// SAFETY: every entry in a leaked `ClassDualVtbl` is an immutable static function pointer that
// lives for the process lifetime.
unsafe impl Send for SyncVtableSlice {}
// SAFETY: same immutable static function-pointer slice.
unsafe impl Sync for SyncVtableSlice {}

/// Interned per-class dual vtables (`&'static [*const c_void]`), keyed by the interned
/// `RuntimeClassDescriptor` pointer so every class builds its vtable once and allocations stay
/// rooted for the process lifetime (0 leaks under Miri).
static CLASS_DUAL_VTBLS: OnceLock<Mutex<HashMap<usize, SyncVtableSlice>>> = OnceLock::new();

fn vtable_cache() -> &'static Mutex<HashMap<usize, SyncVtableSlice>> {
    CLASS_DUAL_VTBLS.get_or_init(|| Mutex::new(HashMap::new()))
}

/// Build (or return cached) `ClassDualVtbl` for `descriptor`, as an `IUnknownVtbl` pointer
/// suitable for `CompatObjectBase.unknown.vtbl`.
pub(crate) fn class_dual_vtbl_for_descriptor(
    descriptor: &'static RuntimeClassDescriptor,
) -> *const IUnknownVtbl {
    class_dual_vtable_slice(descriptor)
        .as_ptr()
        .cast::<IUnknownVtbl>()
}

/// Return the cached `IDispatchVtbl` header pointer for `descriptor`'s `ClassDualVtbl`.
pub fn class_dual_vtable_ptr(descriptor: &'static RuntimeClassDescriptor) -> *const IDispatchVtbl {
    class_dual_vtable_slice(descriptor)
        .as_ptr()
        .cast::<IDispatchVtbl>()
}

/// Total number of slots in `descriptor`'s `ClassDualVtbl` (7 `IDispatch` slots + member slots).
pub fn class_dual_vtable_slot_count(descriptor: &'static RuntimeClassDescriptor) -> usize {
    class_dual_vtable_slice(descriptor).len()
}

fn class_dual_vtable_slice(
    descriptor: &'static RuntimeClassDescriptor,
) -> &'static [*const c_void] {
    let key = descriptor as *const RuntimeClassDescriptor as usize;
    let mut cache = vtable_cache()
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    if let Some(existing) = cache.get(&key).copied() {
        return existing.0;
    }
    let member_end_slot = descriptor
        .dispatch_interface()
        .map(|dispatch| {
            dispatch
                .members
                .iter()
                .filter_map(|member| member.vtable_slot.map(|slot| usize::from(slot) + 1))
                .max()
                .unwrap_or(7)
                .max(7 + dispatch.members.len())
        })
        .unwrap_or(7)
        + usize::from(descriptor.lifecycle.hidden_vtable_slots);

    let mut slots: Vec<*const c_void> = Vec::with_capacity(member_end_slot);
    slots.push(COMPAT_OBJECT_VTBL.query_interface as *const c_void);
    slots.push(COMPAT_OBJECT_VTBL.add_ref as *const c_void);
    slots.push(COMPAT_OBJECT_VTBL.release as *const c_void);
    slots.push(
        compat_get_type_info_count as unsafe extern "system" fn(*mut c_void, *mut u32) -> HResult
            as *const c_void,
    );
    slots.push(
        compat_get_type_info
            as unsafe extern "system" fn(*mut c_void, u32, u32, *mut *mut c_void) -> HResult
            as *const c_void,
    );
    slots.push(
        compat_get_ids_of_names
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
        compat_invoke
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

unsafe extern "system" fn compat_slot_not_implemented(_this: *mut c_void) -> HResult {
    E_NOTIMPL
}

unsafe extern "system" fn compat_get_type_info_count(
    this: *mut c_void,
    pctinfo: *mut u32,
) -> HResult {
    if this.is_null() || pctinfo.is_null() {
        return E_POINTER;
    }
    // SAFETY: per `IDispatch::GetTypeInfoCount`, `pctinfo` is a caller-owned writable out-slot.
    unsafe {
        *pctinfo = 1;
    }
    S_OK
}

unsafe extern "system" fn compat_get_type_info(
    this: *mut c_void,
    itinfo: u32,
    _lcid: u32,
    pptinfo: *mut *mut c_void,
) -> HResult {
    if this.is_null() || pptinfo.is_null() {
        return E_POINTER;
    }
    // SAFETY: per `IDispatch::GetTypeInfo`, `pptinfo` is a caller-owned writable out-slot.
    unsafe {
        *pptinfo = core::ptr::null_mut();
    }
    if itinfo != 0 {
        return DISP_E_BADINDEX;
    }
    let owner = compat_owner_from_this(this);
    // SAFETY: `this` is a live `CompatObjectBase` kept alive by the caller's COM reference.
    let descriptor = unsafe { (*owner).class_descriptor };
    let ti = type_info::create_class_type_info(descriptor);
    // SAFETY: `pptinfo` was verified non-null above.
    unsafe {
        *pptinfo = ti;
    }
    S_OK
}

unsafe fn wide_ptr_to_string(ptr: *const u16) -> Option<String> {
    if ptr.is_null() {
        return None;
    }
    let mut len = 0usize;
    // SAFETY: COM passes NUL-terminated UTF-16 strings (`LPCOLESTR`).
    unsafe {
        while *ptr.add(len) != 0 {
            len += 1;
        }
        let slice = core::slice::from_raw_parts(ptr, len);
        Some(String::from_utf16_lossy(slice))
    }
}

pub(super) unsafe fn resolve_ids_of_names_for_interface(
    interface: &'static RuntimeInterfaceDescriptor,
    names: *const *const u16,
    count: u32,
    dispids: *mut i32,
) -> HResult {
    if names.is_null() || dispids.is_null() {
        return E_POINTER;
    }
    if count == 0 {
        return E_INVALIDARG;
    }

    let count_usize = count as usize;
    // SAFETY: `dispids` has capacity for `count` writable `i32` slots per `GetIDsOfNames`.
    unsafe {
        for index in 0..count_usize {
            *dispids.add(index) = DISPID_UNKNOWN;
        }
    }

    // SAFETY: `count >= 1` and `names` is non-null.
    let Some(member_name) = (unsafe { wide_ptr_to_string(*names) }) else {
        return DISP_E_UNKNOWNNAME;
    };
    let member_name = member_name.trim();
    let Some(primary_member) = interface
        .members
        .iter()
        .find(|member| member.externally_visible && member.name.eq_ignore_ascii_case(member_name))
    else {
        return DISP_E_UNKNOWNNAME;
    };

    let dispid = primary_member.dispatch_id;
    // SAFETY: `dispids` has at least one slot.
    unsafe {
        *dispids = dispid;
    }

    if count_usize == 1 {
        return S_OK;
    }

    let matching_members: Vec<&'static RuntimeMemberDescriptor> = interface
        .members
        .iter()
        .filter(|member| member.externally_visible && member.dispatch_id == dispid)
        .collect();

    let mut status = S_OK;
    for index in 1..count_usize {
        // SAFETY: `names` has `count` readable wide string pointers.
        let param_name = unsafe { wide_ptr_to_string(*names.add(index)) }.unwrap_or_default();
        let param_name = param_name.trim();
        let resolved = resolve_param_dispid(&matching_members, param_name);
        // SAFETY: `dispids` has `count` writable slots.
        unsafe {
            *dispids.add(index) = resolved;
        }
        if resolved == DISPID_UNKNOWN {
            status = DISP_E_UNKNOWNNAME;
        }
    }
    status
}

unsafe extern "system" fn compat_get_ids_of_names(
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
        // SAFETY: non-null `riid` points to a readable `Guid`.
        let iid = unsafe { riid.read() };
        if iid != IID_NULL {
            return DISP_E_UNKNOWNINTERFACE;
        }
    }
    let owner = compat_owner_from_this(this);
    // SAFETY: `this` is a live `CompatObjectBase` kept alive by the caller's COM reference.
    let descriptor = unsafe { (*owner).class_descriptor };
    let Some(dispatch) = descriptor.dispatch_interface() else {
        return DISP_E_UNKNOWNNAME;
    };
    // SAFETY: arguments forwarded from `IDispatch::GetIDsOfNames`.
    unsafe { resolve_ids_of_names_for_interface(dispatch, names, count, dispids) }
}

fn resolve_param_dispid(members: &[&'static RuntimeMemberDescriptor], param_name: &str) -> i32 {
    // First look for a non-value (positional) parameter across matching accessors.
    for member in members {
        let is_put = matches!(
            member.invoke_kind,
            RuntimeMemberInvokeKind::PropertyLet | RuntimeMemberInvokeKind::PropertySet
        );
        let positional_len = if is_put && !member.params.is_empty() {
            member.params.len() - 1
        } else {
            member.params.len()
        };
        if let Some(pos) = member.params[..positional_len]
            .iter()
            .position(|param| param.name.eq_ignore_ascii_case(param_name))
        {
            return pos as i32;
        }
    }
    // Next check the trailing value parameter of a Property Let / Set accessor.
    for member in members {
        let is_put = matches!(
            member.invoke_kind,
            RuntimeMemberInvokeKind::PropertyLet | RuntimeMemberInvokeKind::PropertySet
        );
        if is_put
            && let Some(last) = member.params.last()
            && last.name.eq_ignore_ascii_case(param_name)
        {
            return DISPID_PROPERTYPUT;
        }
    }
    DISPID_UNKNOWN
}

fn select_member_for_invoke(
    interface: &'static RuntimeInterfaceDescriptor,
    dispid: i32,
    flags: u16,
    c_args: usize,
) -> Option<(&'static RuntimeMemberDescriptor, RuntimeMemberInvokeKind)> {
    let want_put_ref = flags & DISPATCH_PROPERTYPUTREF != 0;
    let want_put = flags & DISPATCH_PROPERTYPUT != 0;
    let want_get = flags & DISPATCH_PROPERTYGET != 0;
    let want_method = flags & DISPATCH_METHOD != 0;

    if want_put_ref
        && let Some(member) = interface.resolve_member_by_dispid(
            dispid,
            RuntimeMemberInvokeKind::PropertySet,
            c_args == 0,
            true,
        )
    {
        return Some((member, RuntimeMemberInvokeKind::PropertySet));
    }
    if want_put
        && let Some(member) = interface.resolve_member_by_dispid(
            dispid,
            RuntimeMemberInvokeKind::PropertyLet,
            c_args == 0,
            true,
        )
    {
        return Some((member, RuntimeMemberInvokeKind::PropertyLet));
    }
    if want_get && want_method {
        if let Some(member) = interface.members.iter().find(|member| {
            member.dispatch_id == dispid
                && member.invoke_kind == RuntimeMemberInvokeKind::PropertyGet
                && member.externally_visible
        }) {
            return Some((member, RuntimeMemberInvokeKind::PropertyGet));
        }
        if let Some(member) = interface.members.iter().find(|member| {
            member.dispatch_id == dispid
                && member.invoke_kind == RuntimeMemberInvokeKind::Method
                && member.externally_visible
        }) {
            return Some((member, RuntimeMemberInvokeKind::Method));
        }
        return None;
    }
    if want_get {
        return interface
            .resolve_member_by_dispid(
                dispid,
                RuntimeMemberInvokeKind::PropertyGet,
                c_args == 0,
                true,
            )
            .map(|member| (member, RuntimeMemberInvokeKind::PropertyGet));
    }
    if want_method {
        return interface
            .resolve_member_by_dispid(dispid, RuntimeMemberInvokeKind::Method, c_args == 0, true)
            .map(|member| (member, RuntimeMemberInvokeKind::Method));
    }
    None
}

struct MappedDispatchArgs {
    /// For each fixed formal parameter, the index into `DISPPARAMS.rgvarg` (`None` if omitted).
    slots: Vec<Option<usize>>,
    /// Extra positional `rgvarg` indices captured into a trailing `ParamArray`, in source order.
    param_array_slots: Vec<usize>,
    has_param_array: bool,
}

unsafe fn map_dispatch_params(
    member: &'static RuntimeMemberDescriptor,
    params: *const DispParams,
) -> Result<MappedDispatchArgs, (HResult, u32)> {
    let total_formals = member.params.len();
    let has_param_array = member.params.last().is_some_and(|param| param.param_array);
    let fixed_formals = if has_param_array {
        total_formals - 1
    } else {
        total_formals
    };

    let mut slots = vec![None; fixed_formals];
    let mut param_array_slots = Vec::new();

    if params.is_null() {
        return validate_required_slots(member, &slots, 0).map(|()| MappedDispatchArgs {
            slots,
            param_array_slots,
            has_param_array,
        });
    }

    // SAFETY: caller guarantees `params` points to a readable `DispParams`.
    let params = unsafe { &*params };
    let count = params.c_args as usize;
    let named = params.c_named_args as usize;
    if named > count
        || (count > 0 && params.rgvarg.is_null())
        || (named > 0 && params.rgdispid_named_args.is_null())
    {
        return Err((E_POINTER, 0));
    }

    let positional = count - named;
    if !has_param_array && positional > fixed_formals {
        return Err((DISP_E_BADPARAMCOUNT, 0));
    }

    let fixed_positional = positional.min(fixed_formals);
    for (logical, slot) in slots.iter_mut().enumerate().take(fixed_positional) {
        *slot = Some(count - 1 - logical);
    }
    if has_param_array && positional > fixed_formals {
        for logical in fixed_formals..positional {
            param_array_slots.push(count - 1 - logical);
        }
    }

    for index in 0..named {
        // SAFETY: `rgdispid_named_args` has `named` readable entries.
        let arg_dispid = unsafe { *params.rgdispid_named_args.add(index) };
        let formal = if arg_dispid == DISPID_PROPERTYPUT && fixed_formals > 0 {
            fixed_formals - 1
        } else {
            usize::try_from(arg_dispid).map_err(|_| (DISP_E_PARAMNOTFOUND, index as u32))?
        };
        match slots.get_mut(formal) {
            Some(slot @ None) => *slot = Some(index),
            _ => return Err((DISP_E_PARAMNOTFOUND, index as u32)),
        }
    }

    validate_required_slots(member, &slots, named)?;
    Ok(MappedDispatchArgs {
        slots,
        param_array_slots,
        has_param_array,
    })
}

fn validate_required_slots(
    member: &'static RuntimeMemberDescriptor,
    slots: &[Option<usize>],
    named_count: usize,
) -> Result<(), (HResult, u32)> {
    if let Some(missing) = slots.iter().enumerate().position(|(formal, slot)| {
        slot.is_none()
            && !member
                .params
                .get(formal)
                .is_some_and(|param| param.optional)
    }) {
        return Err((
            if named_count == 0 {
                DISP_E_BADPARAMCOUNT
            } else {
                DISP_E_PARAMNOTOPTIONAL
            },
            missing as u32,
        ));
    }
    Ok(())
}

unsafe fn copy_dispatch_arg_for_param(
    descriptor: &'static RuntimeClassDescriptor,
    member: &'static RuntimeMemberDescriptor,
    formal: Option<usize>,
    cell: *const VariantCore,
) -> Result<Variant, HResult> {
    let expected_record_array = formal.and_then(|param_index| {
        let param = member.params.get(param_index)?;
        if param.value_type != RuntimeValueType::Array {
            return None;
        }
        let RuntimeMemberBacking::Proc(proc_id) = member.backing else {
            return None;
        };
        descriptor.record_member_layout(proc_id, Some(param_index), true)
    });
    let resolve = |guid: &Guid, size: u32| descriptor.resolve_project_record(guid, size);
    // SAFETY: `cell` points to a valid caller-owned `VARIANT` cell.
    unsafe {
        calcvba_automation::automation_record::automation_variant_to_runtime_variant(
            cell.cast::<c_void>(),
            expected_record_array,
            &resolve,
        )
    }
}

unsafe fn copy_dispatch_arg(cell: *const VariantCore) -> Result<Variant, HResult> {
    let mut copied = Variant::empty();
    // SAFETY: `cell` points to a valid caller-owned `VARIANT` cell; `copied` is an empty writable
    // cell. `variant_copy_ind` copies ByVal or dereferences `VT_BYREF` safely on every platform.
    let hr = unsafe {
        oleaut::variant_copy_ind((&raw mut copied).cast::<c_void>(), cell.cast::<c_void>())
    };
    if hr < 0 {
        return Err(hr);
    }
    Ok(copied)
}

/// Resolve a `VT_BYREF` argument cell to its ultimate `(target_vt, target_ptr)`, unwrapping
/// Automation forwarding `VT_BYREF | VT_VARIANT` chains (up to 32 hops) whose inner `VARIANT` is
/// itself a `VT_BYREF`.
unsafe fn resolve_byref_target(mut raw: VariantCore) -> Result<Option<(u16, *mut u8)>, HResult> {
    const VT_BYREF: u16 = 0x4000;
    const VT_VARIANT: u16 = 0x000C;
    if raw.raw_vt() & VT_BYREF == 0 {
        return Ok(None);
    }
    for _ in 0..32 {
        let vt = raw.raw_vt();
        if vt & VT_BYREF == 0 {
            return Ok(None);
        }
        let target_vt = vt & !VT_BYREF;
        let target_ptr = raw.pointer().cast::<u8>();
        if target_ptr.is_null() {
            return Err(E_POINTER);
        }
        if target_vt == VT_VARIANT {
            // SAFETY: caller's `VT_BYREF | VT_VARIANT` points to a live `VARIANT` cell.
            let inner = unsafe { *target_ptr.cast::<VariantCore>() };
            if inner.raw_vt() & VT_BYREF != 0 {
                raw = inner;
                continue;
            }
        }
        return Ok(Some((target_vt, target_ptr)));
    }
    Err(DISP_E_TYPEMISMATCH)
}

unsafe fn compute_byref_aliases(
    member: &'static RuntimeMemberDescriptor,
    params: *const DispParams,
    slots: &[Option<usize>],
    total_args: usize,
) -> Vec<usize> {
    let mut aliases: Vec<usize> = (0..total_args).collect();
    if params.is_null() {
        return aliases;
    }
    // SAFETY: `params` was validated non-null and bounded above.
    let rgvarg = unsafe { (*params).rgvarg };
    if rgvarg.is_null() {
        return aliases;
    }
    let mut targets: Vec<Option<(u16, usize)>> = Vec::with_capacity(slots.len());
    for (formal, slot) in slots.iter().enumerate() {
        let by_ref = member.params.get(formal).is_some_and(|param| param.by_ref);
        let target = if by_ref && let Some(rg_idx) = *slot {
            // SAFETY: `rg_idx < c_args`.
            let raw = unsafe { *rgvarg.add(rg_idx) };
            // SAFETY: caller owns the argument `VARIANT` cell and any `VT_BYREF` target.
            unsafe { resolve_byref_target(raw) }
                .ok()
                .flatten()
                .map(|(vt, ptr)| (vt, ptr as usize))
        } else {
            None
        };
        targets.push(target);
    }
    for index in 0..targets.len() {
        let Some(target) = targets[index] else {
            continue;
        };
        if let Some(first) = targets[..index]
            .iter()
            .position(|other| *other == Some(target))
        {
            aliases[index] = first;
        }
    }
    aliases
}

unsafe fn write_back_byref_args(
    descriptor: &'static RuntimeClassDescriptor,
    member: &'static RuntimeMemberDescriptor,
    params: *const DispParams,
    slots: &[Option<usize>],
    args: &[Variant],
) -> Result<(), HResult> {
    const VT_ARRAY_RECORD: u16 = 0x2000 | 0x0024;
    if params.is_null() {
        return Ok(());
    }
    // SAFETY: `params` was validated above.
    let rgvarg = unsafe { (*params).rgvarg };
    if rgvarg.is_null() {
        return Ok(());
    }
    for (formal, slot) in slots.iter().enumerate() {
        if !member.params.get(formal).is_some_and(|param| param.by_ref) {
            continue;
        }
        let Some(rg_idx) = *slot else {
            continue;
        };
        let Some(updated) = args.get(formal) else {
            continue;
        };
        // SAFETY: `rg_idx < c_args`; caller owns the `VARIANT` cell and its `VT_BYREF` target.
        let raw = unsafe { *rgvarg.add(rg_idx) };
        // SAFETY: caller's `VT_BYREF` cell points to live writable storage for this `Invoke` call.
        let Some((target_vt, target_ptr)) = (unsafe { resolve_byref_target(raw) })? else {
            continue;
        };
        if target_vt == VT_ARRAY_RECORD {
            let expected = match member.backing {
                RuntimeMemberBacking::Proc(proc_id) => descriptor
                    .record_member_layout(proc_id, Some(formal), true)
                    .cloned(),
                _ => None,
            }
            .or_else(|| updated.record_array_layout())
            .ok_or(DISP_E_TYPEMISMATCH)?;
            // SAFETY: `target_ptr` is caller's live `SAFEARRAY*` slot (`VT_BYREF | VT_ARRAY | VT_RECORD`).
            unsafe {
                calcvba_automation::automation_record::write_vba_record_array_to_automation_slot(
                    target_ptr.cast::<*mut c_void>(),
                    updated,
                    &expected,
                )?;
            }
            continue;
        }
        if target_vt == 0x0C {
            // SAFETY: `target_ptr` is caller's live `VARIANT` cell (`VT_BYREF | VT_VARIANT`).
            unsafe {
                calcvba_automation::automation_record::write_runtime_variant_to_automation_cell(
                    target_ptr.cast::<c_void>(),
                    updated,
                )?;
            }
            continue;
        }
        // SAFETY: caller's `VT_BYREF` target is live writable storage for this `Invoke` call.
        let cell = unsafe { Variant::by_ref(target_vt, target_ptr) };
        let prepared = coerce_for_byref_writeback(updated, target_vt)?;
        // SAFETY: caller's `VT_BYREF` target is live writable storage for this `Invoke` call.
        unsafe { cell.write_through(prepared) }.map_err(|_| DISP_E_TYPEMISMATCH)?;
    }
    Ok(())
}

fn coerce_for_byref_writeback(value: &Variant, target_vt: u16) -> Result<Variant, HResult> {
    // `VT_VARIANT` accepts any Variant as-is.
    if target_vt == 0x0C || value.core().raw_vt() == target_vt {
        return Ok(value.clone());
    }
    match target_vt {
        0x02 => {
            let n = as_i64_scalar(value).ok_or(DISP_E_TYPEMISMATCH)?;
            i16::try_from(n)
                .map(Variant::from_i16)
                .map_err(|_| DISP_E_TYPEMISMATCH)
        }
        0x03 => {
            let n = as_i64_scalar(value).ok_or(DISP_E_TYPEMISMATCH)?;
            i32::try_from(n)
                .map(Variant::from_i32)
                .map_err(|_| DISP_E_TYPEMISMATCH)
        }
        0x04 => as_f64_scalar(value)
            .map(|v| Variant::from_f32(v as f32))
            .ok_or(DISP_E_TYPEMISMATCH),
        0x05 => as_f64_scalar(value)
            .map(Variant::from_f64)
            .ok_or(DISP_E_TYPEMISMATCH),
        0x0B => match value.vtype() {
            VarType::Boolean => Ok(value.clone()),
            _ => as_i64_scalar(value)
                .map(|n| Variant::from_bool(n != 0))
                .ok_or(DISP_E_TYPEMISMATCH),
        },
        0x11 => {
            let n = as_i64_scalar(value).ok_or(DISP_E_TYPEMISMATCH)?;
            u8::try_from(n)
                .map(Variant::from_u8)
                .map_err(|_| DISP_E_TYPEMISMATCH)
        }
        0x14 => as_i64_scalar(value)
            .map(Variant::from_i64)
            .ok_or(DISP_E_TYPEMISMATCH),
        _ => Ok(value.clone()),
    }
}

fn as_i64_scalar(value: &Variant) -> Option<i64> {
    match value.vtype() {
        VarType::Empty => Some(0),
        VarType::Byte => value.as_u8().map(i64::from),
        VarType::Integer => value.as_i16().map(i64::from),
        VarType::Long | VarType::Int => value.as_i32().map(i64::from),
        VarType::LongLong => value.as_i64(),
        VarType::Boolean => value.as_bool().map(|b| if b { -1 } else { 0 }),
        _ => None,
    }
}

fn as_f64_scalar(value: &Variant) -> Option<f64> {
    match value.vtype() {
        VarType::Single => value.as_f32().map(f64::from),
        VarType::Double => value.as_f64(),
        _ => as_i64_scalar(value).map(|n| n as f64),
    }
}

fn vba_error_to_hresult(number: i32) -> HResult {
    if number < 0 {
        number
    } else if (1..=0xFFFF).contains(&number) {
        (0x800A_0000u32 | (number as u32)) as i32
    } else {
        E_FAIL
    }
}

unsafe fn write_excep_info(excep_info: *mut ExcepInfo, fault: &MemberFault) {
    if excep_info.is_null() {
        return;
    }
    let alloc_bstr = |text: &crate::Utf16String| -> *mut u16 {
        if text.is_empty() {
            core::ptr::null_mut()
        } else {
            BStr::from_utf16_units(text.as_units())
                .map(BStr::into_raw_bstr)
                .unwrap_or(core::ptr::null_mut())
        }
    };
    let info = ExcepInfo {
        w_code: 0,
        w_reserved: 0,
        bstr_source: alloc_bstr(&fault.source),
        bstr_description: alloc_bstr(&fault.description),
        bstr_help_file: alloc_bstr(&fault.help_file),
        dw_help_context: fault.help_context as u32,
        pv_reserved: core::ptr::null_mut(),
        pfn_deferred_fill_in: None,
        scode: vba_error_to_hresult(fault.number),
    };
    // SAFETY: caller supplied a writable `EXCEPINFO` out-struct.
    unsafe {
        excep_info.write(info);
    }
}

fn invoke_field_directly(
    object: &ObjectRef,
    member: &'static RuntimeMemberDescriptor,
    field_token: i32,
    invoke_kind: RuntimeMemberInvokeKind,
    args: &[Variant],
) -> Result<Variant, MemberFault> {
    match invoke_kind {
        RuntimeMemberInvokeKind::PropertyGet | RuntimeMemberInvokeKind::Method => {
            if !args.is_empty() {
                return Err(MemberFault::from_number(450));
            }
            let raw = object
                .project_field_get(field_token)
                .unwrap_or_else(Variant::empty);
            if raw.vtype() == VarType::Empty {
                return Ok(default_zero_for_runtime_type(
                    member.return_type.unwrap_or(RuntimeValueType::Variant),
                ));
            }
            Ok(raw)
        }
        RuntimeMemberInvokeKind::PropertyLet => {
            let [value] = args else {
                return Err(MemberFault::from_number(450));
            };
            object.project_field_set(field_token, value.clone());
            Ok(Variant::empty())
        }
        RuntimeMemberInvokeKind::PropertySet => {
            let [value] = args else {
                return Err(MemberFault::from_number(450));
            };
            if value.vtype() != VarType::Object {
                return Err(MemberFault::from_number(424));
            }
            object.project_field_set(field_token, value.clone());
            Ok(Variant::empty())
        }
    }
}

fn default_zero_for_runtime_type(ty: RuntimeValueType) -> Variant {
    match ty {
        RuntimeValueType::Byte => Variant::from_u8(0),
        RuntimeValueType::Integer => Variant::from_i16(0),
        RuntimeValueType::Long => Variant::from_i32(0),
        RuntimeValueType::LongLong => Variant::from_i64(0),
        RuntimeValueType::LongPtr => {
            #[cfg(target_pointer_width = "64")]
            {
                Variant::from_i64(0)
            }
            #[cfg(not(target_pointer_width = "64"))]
            {
                Variant::from_i32(0)
            }
        }
        RuntimeValueType::Single => Variant::from_f32(0.0),
        RuntimeValueType::Double => Variant::from_f64(0.0),
        RuntimeValueType::Currency => Variant::from_currency_scaled_i64(0),
        RuntimeValueType::Date => Variant::from_date_f64(0.0),
        RuntimeValueType::Boolean => Variant::from_bool(false),
        RuntimeValueType::String => Variant::from_string(""),
        RuntimeValueType::Object => Variant::nothing(),
        RuntimeValueType::Variant
        | RuntimeValueType::Decimal
        | RuntimeValueType::Record
        | RuntimeValueType::Array => Variant::empty(),
    }
}

fn invoke_object_member(
    owner: *mut CompatObjectBase,
    object: ObjectRef,
    member: &'static RuntimeMemberDescriptor,
    invoke_kind: RuntimeMemberInvokeKind,
    args: &mut [Variant],
    aliases: &[usize],
) -> Result<Variant, MemberFault> {
    if let Some(outcome) = events::invoke_member(object.clone(), member, invoke_kind, args, aliases)
    {
        return outcome;
    }
    // Check an optional per-object weak invoker (installed by a session holder between runs).
    // SAFETY: `owner` is a live `CompatObjectBase` on its owning thread (checked before call).
    let per_object = unsafe { (*owner).invoker.borrow().as_ref().and_then(|w| w.upgrade()) };
    if let Some(invoker) = per_object {
        let Ok(mut guard) = invoker.try_borrow_mut() else {
            return Err(MemberFault::from_number(438));
        };
        return guard.invoke_member(object, member, invoke_kind, args, aliases);
    }
    // Direct fallback for field-backed members when no runtime executor is active (e.g. AOT
    // objects or standalone descriptor tests).
    if let RuntimeMemberBacking::Field(field_token) = member.backing {
        return invoke_field_directly(&object, member, field_token, invoke_kind, args);
    }
    Err(MemberFault::from_number(438))
}

#[allow(clippy::too_many_arguments)]
unsafe extern "system" fn compat_invoke(
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
    let owner = compat_owner_from_this(this);
    // SAFETY: `this` is a live `CompatObjectBase` kept alive by the caller's COM reference.
    let descriptor = unsafe { (*owner).class_descriptor };
    let Some(dispatch) = descriptor.dispatch_interface() else {
        return DISP_E_MEMBERNOTFOUND;
    };
    let interface_iid = descriptor.default_interface_iid.unwrap_or(IID_IDISPATCH);
    // SAFETY: `owner` is a live `CompatObjectBase` and arguments match `IDispatch::Invoke`.
    unsafe {
        invoke_on_interface(
            owner,
            dispatch,
            interface_iid,
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

#[allow(clippy::too_many_arguments)]
pub(super) unsafe fn invoke_on_interface(
    owner: *mut CompatObjectBase,
    interface: &'static RuntimeInterfaceDescriptor,
    interface_iid: Guid,
    dispid: i32,
    riid: *const Guid,
    flags: u16,
    params: *mut c_void,
    result: *mut c_void,
    excep_info: *mut c_void,
    arg_err: *mut u32,
) -> HResult {
    if !riid.is_null() {
        // SAFETY: non-null `riid` points to a readable `Guid`.
        let iid = unsafe { riid.read() };
        if iid != IID_NULL {
            return DISP_E_UNKNOWNINTERFACE;
        }
    }
    if !result.is_null() {
        // SAFETY: per `IDispatch::Invoke`, `pVarResult` is an `[out]` VARIANT slot.
        unsafe {
            result.cast::<Variant>().write(Variant::empty());
        }
    }

    // SAFETY: `owner` is a live `CompatObjectBase` kept alive by the caller's COM reference.
    let (owner_thread, descriptor) = unsafe { ((*owner).owner_thread, (*owner).class_descriptor) };
    if current_thread_id() != owner_thread {
        return RPC_E_WRONG_THREAD;
    }
    error_info::clear_thread_error_info();

    let params_ptr = params.cast::<DispParams>();
    let c_args = if params_ptr.is_null() {
        0
    } else {
        // SAFETY: caller supplied a readable `DispParams`.
        unsafe { (*params_ptr).c_args as usize }
    };
    let Some((member, requested_kind)) = select_member_for_invoke(interface, dispid, flags, c_args)
    else {
        return DISP_E_MEMBERNOTFOUND;
    };

    // SAFETY: `params_ptr` is null or a caller-owned `DispParams`.
    let mapped = match unsafe { map_dispatch_params(member, params_ptr) } {
        Ok(mapped) => mapped,
        Err((hr, position)) => {
            if !arg_err.is_null() {
                // SAFETY: caller supplied a writable `puArgErr` out-slot.
                unsafe {
                    *arg_err = position;
                }
            }
            return hr;
        }
    };

    let mut args = Vec::with_capacity(mapped.slots.len() + usize::from(mapped.has_param_array));
    for (formal, slot) in mapped.slots.iter().enumerate() {
        match slot {
            Some(rg_idx) => {
                // SAFETY: `rg_idx < c_args` and `rgvarg` was verified non-null.
                let cell = unsafe { (*params_ptr).rgvarg.add(*rg_idx) };
                // SAFETY: caller owns the argument `VARIANT` cell.
                match unsafe { copy_dispatch_arg_for_param(descriptor, member, Some(formal), cell) }
                {
                    Ok(value) => args.push(value),
                    Err(hr) => {
                        if !arg_err.is_null() {
                            // SAFETY: caller supplied a writable `puArgErr` out-slot.
                            unsafe {
                                *arg_err = formal as u32;
                            }
                        }
                        return hr;
                    }
                }
            }
            None => args.push(Variant::from_error_code(DISP_E_PARAMNOTFOUND)),
        }
    }
    if mapped.has_param_array {
        let mut tail = Vec::with_capacity(mapped.param_array_slots.len());
        for rg_idx in &mapped.param_array_slots {
            // SAFETY: `rg_idx < c_args` and `rgvarg` was verified non-null.
            let cell = unsafe { (*params_ptr).rgvarg.add(*rg_idx) };
            // SAFETY: caller owns the argument `VARIANT` cell.
            match unsafe { copy_dispatch_arg_for_param(descriptor, member, None, cell) } {
                Ok(value) => tail.push(value),
                Err(hr) => return hr,
            }
        }
        args.push(Variant::from_safearray(SafeArray::from_variants(tail)));
    }

    // SAFETY: `params_ptr` was validated above.
    let aliases = unsafe { compute_byref_aliases(member, params_ptr, &mapped.slots, args.len()) };

    // SAFETY: `owner` is a live `CompatObjectBase`; always pass `&primary` as `Me`.
    let Some(object) = (unsafe { ObjectRef::from_raw_iunknown_addref(owner.cast()) }) else {
        return E_POINTER;
    };

    let outcome = invoke_object_member(owner, object, member, requested_kind, &mut args, &aliases);

    // SAFETY: `params_ptr` and `mapped.slots` are valid for the duration of `Invoke`.
    let writeback =
        unsafe { write_back_byref_args(descriptor, member, params_ptr, &mapped.slots, &args) };

    match outcome {
        Ok(ret_val) => {
            if let Err(hr) = writeback {
                return hr;
            }
            if !result.is_null() {
                // SAFETY: `result` was initialized to `Variant::empty()` above; project any UDT
                // array to an Automation `SAFEARRAY(VT_RECORD)` (`FADF_RECORD`, zero tail).
                if let Err(hr) = unsafe {
                    calcvba_automation::automation_record::write_runtime_variant_to_automation_cell(
                        result, &ret_val,
                    )
                } {
                    return hr;
                }
            }
            S_OK
        }
        Err(fault) => {
            let _ = writeback;
            error_info::publish_member_fault_error_info(interface_iid, &fault);
            // SAFETY: `excep_info` is null or a caller-owned writable `ExcepInfo`.
            unsafe {
                write_excep_info(excep_info.cast::<ExcepInfo>(), &fault);
            }
            DISP_E_EXCEPTION
        }
    }
}

/// Early-bound member invocation helper matching the Win64 slot-thunk contract (clears thread
/// `IErrorInfo` on entry, publishes rich `IErrorInfo` tagged with `interface_iid` on VBA fault,
/// and returns `S_OK` or `0x800A0000 | Err.Number`).
pub fn invoke_early_bound_slot(
    object: &ObjectRef,
    interface_iid: Option<Guid>,
    slot: u16,
    args: &mut [Variant],
) -> Result<Variant, HResult> {
    if !object.is_compat_object() {
        return Err(E_POINTER);
    }
    let owner = crate::object_ref::compat_owner_from_unknown(object.raw_iunknown());
    // SAFETY: `object` retains `owner` for the duration of this call.
    let (owner_thread, descriptor) = unsafe { ((*owner).owner_thread, (*owner).class_descriptor) };
    if current_thread_id() != owner_thread {
        return Err(RPC_E_WRONG_THREAD);
    }
    error_info::clear_thread_error_info();
    let (interface, resolved_iid) = match interface_iid {
        Some(iid) if descriptor.default_interface_iid != Some(iid) && iid != IID_IDISPATCH => {
            let Some(iface) = descriptor
                .interfaces
                .iter()
                .find(|iface| iface.identity.guid == iid)
            else {
                return Err(E_NOTIMPL);
            };
            (iface, iid)
        }
        other => {
            let Some(dispatch) = descriptor.dispatch_interface() else {
                return Err(E_NOTIMPL);
            };
            (
                dispatch,
                other.unwrap_or_else(|| descriptor.default_interface_iid.unwrap_or(IID_IDISPATCH)),
            )
        }
    };
    let Some(member) = interface.member_by_vtable_slot(slot) else {
        return Err(E_NOTIMPL);
    };
    let aliases: Vec<usize> = (0..args.len()).collect();
    // SAFETY: `owner` is live; retain `&primary` as `Me`.
    let Some(me) = (unsafe { ObjectRef::from_raw_iunknown_addref(owner.cast()) }) else {
        return Err(E_POINTER);
    };
    match invoke_object_member(owner, me, member, member.invoke_kind, args, &aliases) {
        Ok(value) => Ok(value),
        Err(fault) => {
            error_info::publish_member_fault_error_info(resolved_iid, &fault);
            Err(vba_error_to_hresult(fault.number))
        }
    }
}

#[cfg(test)]
#[allow(clippy::undocumented_unsafe_blocks)]
mod tests {
    use super::*;
    use crate::class_descriptor::{
        RuntimeClassFieldSpec, RuntimeClassSpec, RuntimeMemberSpec, RuntimeParamSpec,
        build_runtime_class_descriptor,
    };
    use crate::events::{MemberInvoker, MemberInvokerScope};
    use crate::{RUNTIME_CLASS_LIFECYCLE_NONE, RUNTIME_GUID_IDISPATCH, RUNTIME_GUID_IUNKNOWN};
    use std::ptr::NonNull;

    fn sample_widget_descriptor() -> &'static RuntimeClassDescriptor {
        let spec = RuntimeClassSpec {
            unit_name: "ClassDispatchUnit".into(),
            class_index: 0,
            name: "Widget".into(),
            runtime_type_name: None,
            default_interface_iid: None,
            source_interface_iid: None,
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
            record_members: Vec::new(),
            project_records: Vec::new(),
            fields: vec![RuntimeClassFieldSpec {
                name: "Value".into(),
                token: 10,
                value_type: RuntimeValueType::Long,
                array_element_type: None,
                externally_visible: true,
                dispatch_id: None,
                get_vtable_slot: None,
                put_vtable_slot: None,
                putref_vtable_slot: None,
            }],
            as_new_fields: Vec::new(),
            implements: Vec::new(),
            methods: vec![
                RuntimeMemberSpec {
                    name: "Twice".into(),
                    dispatch_id: None,
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::Method,
                    params: Vec::new(),
                    return_type: Some(RuntimeValueType::Long),
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: true,
                    func_flags: 0,
                    backing: RuntimeMemberBacking::Proc(1),
                },
                RuntimeMemberSpec {
                    name: "AddPair".into(),
                    dispatch_id: None,
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::Method,
                    params: vec![
                        RuntimeParamSpec {
                            name: "Left".into(),
                            value_type: RuntimeValueType::Long,
                            by_ref: false,
                            optional: false,
                            param_array: false,
                        },
                        RuntimeParamSpec {
                            name: "Right".into(),
                            value_type: RuntimeValueType::Long,
                            by_ref: true,
                            optional: true,
                            param_array: false,
                        },
                    ],
                    return_type: Some(RuntimeValueType::Long),
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: true,
                    func_flags: 0,
                    backing: RuntimeMemberBacking::Proc(2),
                },
                RuntimeMemberSpec {
                    name: "Fail".into(),
                    dispatch_id: None,
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::Method,
                    params: Vec::new(),
                    return_type: None,
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: true,
                    func_flags: 0,
                    backing: RuntimeMemberBacking::Proc(3),
                },
                RuntimeMemberSpec {
                    name: "Scale".into(),
                    dispatch_id: None,
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::Method,
                    params: ["Factor", "Weight", "Bias", "Offset"]
                        .into_iter()
                        .enumerate()
                        .map(|(i, name)| RuntimeParamSpec {
                            name: name.into(),
                            value_type: if i == 1 {
                                RuntimeValueType::Single
                            } else {
                                RuntimeValueType::Double
                            },
                            by_ref: false,
                            optional: false,
                            param_array: false,
                        })
                        .collect(),
                    return_type: Some(RuntimeValueType::Double),
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: true,
                    func_flags: 0,
                    backing: RuntimeMemberBacking::Proc(5),
                },
                RuntimeMemberSpec {
                    name: "Hidden".into(),
                    dispatch_id: None,
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::Method,
                    params: Vec::new(),
                    return_type: None,
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: false,
                    func_flags: 0,
                    backing: RuntimeMemberBacking::Proc(4),
                },
            ],
            implemented_interfaces: Vec::new(),
            events: Vec::new(),
        };
        build_runtime_class_descriptor(&spec)
    }

    struct MockInvoker;
    impl MemberInvoker for MockInvoker {
        fn invoke_member(
            &mut self,
            object: ObjectRef,
            member: &'static RuntimeMemberDescriptor,
            invoke_kind: RuntimeMemberInvokeKind,
            args: &mut [Variant],
            _aliases: &[usize],
        ) -> Result<Variant, MemberFault> {
            match member.backing {
                RuntimeMemberBacking::Field(token) => {
                    invoke_field_directly(&object, member, token, invoke_kind, args)
                }
                RuntimeMemberBacking::Proc(1) => {
                    let v = object
                        .project_field_get(10)
                        .and_then(|v| v.as_i32())
                        .unwrap_or(0);
                    Ok(Variant::from_i32(v * 2))
                }
                RuntimeMemberBacking::Proc(2) => {
                    let left = args[0].as_i32().unwrap_or(0);
                    let right = if args[1].as_error_code() == Some(DISP_E_PARAMNOTFOUND) {
                        10
                    } else {
                        args[1].as_i32().unwrap_or(0)
                    };
                    args[1] = Variant::from_i32(right + 1);
                    Ok(Variant::from_i32(left + right))
                }
                RuntimeMemberBacking::Proc(3) => Err(MemberFault {
                    number: 5000,
                    description: "boom".into(),
                    source: "Widget".into(),
                    help_file: crate::Utf16String::new(),
                    help_context: 7,
                }),
                RuntimeMemberBacking::Proc(5) => {
                    let f = |i: usize| args[i].as_f64().or(args[i].as_f32().map(f64::from));
                    let (a, b, c, d) = (f(0), f(1), f(2), f(3));
                    Ok(Variant::from_f64(
                        a.unwrap_or(0.0) * b.unwrap_or(0.0) + c.unwrap_or(0.0) - d.unwrap_or(0.0),
                    ))
                }
                _ => Err(MemberFault::from_number(438)),
            }
        }
    }

    fn utf16_nul(s: &str) -> Vec<u16> {
        s.encode_utf16().chain(core::iter::once(0)).collect()
    }

    #[test]
    fn dual_vtable_qi_identity_get_ids_invoke_byref_exception_and_wrong_thread() {
        let descriptor = sample_widget_descriptor();
        let class_iid = descriptor.default_interface_iid.expect("default IID");
        let object = ObjectRef::from_project_instance(1, 0, 0, false, descriptor);
        assert!(object.is_compat_object());
        assert!(object.is_project_instance());
        assert!(object.has_dual_dispatch_vtable());
        assert!(class_dual_vtable_slot_count(descriptor) >= 7 + 6);

        let raw = object.raw_iunknown();
        let vtbl = unsafe { &*((*raw).vtbl.cast::<IDispatchVtbl>()) };

        // 1. QI(IUnknown), QI(IDispatch), QI(class IID) all return &primary and share one refcount.
        for iid in [&RUNTIME_GUID_IUNKNOWN, &RUNTIME_GUID_IDISPATCH, &class_iid] {
            let mut out = core::ptr::null_mut();
            let hr = unsafe { (vtbl.unknown.query_interface)(raw.cast(), iid, &mut out) };
            assert_eq!(hr, S_OK);
            assert_eq!(out, raw.cast::<c_void>());
            assert_eq!(object.strong_count(), 2);
            unsafe { (vtbl.unknown.release)(out) };
            assert_eq!(object.strong_count(), 1);
        }
        let unknown_iid = Guid::deterministic("Other", "INone");
        let mut bogus = core::ptr::null_mut();
        let hr = unsafe { (vtbl.unknown.query_interface)(raw.cast(), &unknown_iid, &mut bogus) };
        assert_eq!(hr, calcvba_automation::com::E_NOINTERFACE);
        assert!(bogus.is_null());

        // Round-tripping raw pointer as IDispatch preserves exact `is_identity()` and `is_compat_object()`.
        let cloned = unsafe { ObjectRef::from_raw_iunknown_addref(raw) }.expect("retained");
        assert!(cloned.is_compat_object());
        assert!(cloned.is_project_instance());
        assert_eq!(cloned.is_identity(), object.is_identity());
        drop(cloned);

        // 2. GetTypeInfoCount / GetTypeInfo (dc-yqq0.29).
        let mut tinfo_count = 99u32;
        assert_eq!(
            unsafe { (vtbl.get_type_info_count)(raw.cast(), &mut tinfo_count) },
            S_OK
        );
        assert_eq!(tinfo_count, 1);

        let mut bad_ti = core::ptr::null_mut();
        assert_eq!(
            unsafe { (vtbl.get_type_info)(raw.cast(), 1, 0, &mut bad_ti) },
            DISP_E_BADINDEX
        );
        assert!(bad_ti.is_null());

        let mut ti_ptr = core::ptr::null_mut();
        assert_eq!(
            unsafe { (vtbl.get_type_info)(raw.cast(), 0, 0, &mut ti_ptr) },
            S_OK
        );
        assert!(!ti_ptr.is_null());
        unsafe {
            use calcvba_automation::com::{
                IID_ITYPEINFO, ITypeInfo, TKIND_DISPATCH, TKIND_INTERFACE, TYPE_E_ELEMENTNOTFOUND,
                VBA_CLASS_TYPEFLAGS, VT_HRESULT,
            };
            let ti = ti_ptr.cast::<ITypeInfo>();
            let ti_vtbl = &*(*ti).vtbl;

            // QI(ITypeInfo) and QI(IUnknown) succeed and balance refcount.
            let mut ti_qi = core::ptr::null_mut();
            assert_eq!(
                (ti_vtbl.unknown.query_interface)(ti_ptr, &IID_ITYPEINFO, &mut ti_qi),
                S_OK
            );
            assert_eq!(ti_qi, ti_ptr);
            assert_eq!((ti_vtbl.unknown.release)(ti_qi), 1);

            // TYPEATTR on dispatch half matches Excel: TKIND_DISPATCH, flags 0x1C2, impl=1, vars=0.
            let mut pattr = core::ptr::null_mut();
            assert_eq!((ti_vtbl.get_type_attr)(ti_ptr, &mut pattr), S_OK);
            let attr = *pattr;
            assert_eq!(attr.typekind, TKIND_DISPATCH);
            assert_eq!(attr.w_type_flags, VBA_CLASS_TYPEFLAGS);
            assert_eq!(attr.w_type_flags, 0x1C2);
            assert_eq!(attr.guid, class_iid);
            assert_eq!(attr.c_funcs, 6); // Value get/let + Twice + AddPair + Fail + Scale (Hidden excluded)
            assert_eq!(attr.c_vars, 0);
            assert_eq!(attr.c_impl_types, 1);
            // 7 IDispatch + 6 public + 1 hidden slot = 14 slots.
            assert_eq!(
                attr.cb_size_vft as usize,
                14 * core::mem::size_of::<*const c_void>()
            );
            (ti_vtbl.release_type_attr)(ti_ptr, pattr);

            // GetDocumentation(-1) yields class name "Widget".
            let mut type_name = core::ptr::null_mut();
            assert_eq!(
                (ti_vtbl.get_documentation)(
                    ti_ptr,
                    -1,
                    &mut type_name,
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                ),
                S_OK
            );
            assert_eq!(BStr::from_raw_bstr(type_name).as_str(), "Widget");

            // GetRefTypeOfImplType(0) and GetImplTypeFlags(0) return TYPE_E_ELEMENTNOTFOUND.
            let mut href = 0u32;
            assert_eq!(
                (ti_vtbl.get_ref_type_of_impl_type)(ti_ptr, 0, &mut href),
                TYPE_E_ELEMENTNOTFOUND
            );
            let mut impl_flags = 0i32;
            assert_eq!(
                (ti_vtbl.get_impl_type_flags)(ti_ptr, 0, &mut impl_flags),
                TYPE_E_ELEMENTNOTFOUND
            );

            // Dual interface half via GetRefTypeOfImplType(-1) -> GetRefTypeInfo.
            assert_eq!(
                (ti_vtbl.get_ref_type_of_impl_type)(ti_ptr, u32::MAX, &mut href),
                S_OK
            );
            let mut iface_ptr = core::ptr::null_mut();
            assert_eq!(
                (ti_vtbl.get_ref_type_info)(ti_ptr, href, &mut iface_ptr),
                S_OK
            );
            assert!(!iface_ptr.is_null());
            let iface_vtbl = &*(*iface_ptr.cast::<ITypeInfo>()).vtbl;
            let mut iface_pattr = core::ptr::null_mut();
            assert_eq!(
                (iface_vtbl.get_type_attr)(iface_ptr, &mut iface_pattr),
                S_OK
            );
            let iface_attr = *iface_pattr;
            assert_eq!(iface_attr.typekind, TKIND_INTERFACE);
            assert_eq!(iface_attr.w_type_flags, 0x1C2);
            assert_eq!(iface_attr.guid, class_iid);
            assert_eq!(iface_attr.c_funcs, 6);
            assert_eq!(iface_attr.cb_size_vft, attr.cb_size_vft);
            (iface_vtbl.release_type_attr)(iface_ptr, iface_pattr);

            // Check Value Get FUNCDESC in interface half: HRESULT return, 1 [out, retval] param.
            let mut fd_ptr = core::ptr::null_mut();
            assert_eq!((iface_vtbl.get_func_desc)(iface_ptr, 0, &mut fd_ptr), S_OK);
            let fd = *fd_ptr;
            assert_eq!(fd.memid, 0x4003_0000);
            assert_eq!(fd.c_params, 1);
            assert_eq!(fd.elemdesc_func.tdesc.vt, VT_HRESULT);
            (iface_vtbl.release_func_desc)(iface_ptr, fd_ptr);

            assert_eq!((iface_vtbl.unknown.release)(iface_ptr), 0);
            assert_eq!((ti_vtbl.unknown.release)(ti_ptr), 0);
        }

        // 3. GetIDsOfNames resolves public members & parameter names, rejects Hidden/unknown.
        let n_value = utf16_nul("value");
        let n_twice = utf16_nul("TWICE");
        let n_add = utf16_nul("AddPair");
        let n_left = utf16_nul("left");
        let n_right = utf16_nul("RIGHT");
        let n_missing = utf16_nul("NoSuchParam");
        let n_hidden = utf16_nul("Hidden");

        let mut value_id = DISPID_UNKNOWN;
        let p_value = [n_value.as_ptr()];
        assert_eq!(
            unsafe {
                (vtbl.get_ids_of_names)(
                    raw.cast(),
                    &IID_NULL,
                    p_value.as_ptr(),
                    1,
                    0,
                    &mut value_id,
                )
            },
            S_OK
        );
        assert!(value_id > 0);

        let mut twice_id = DISPID_UNKNOWN;
        let p_twice = [n_twice.as_ptr()];
        assert_eq!(
            unsafe {
                (vtbl.get_ids_of_names)(
                    raw.cast(),
                    &IID_NULL,
                    p_twice.as_ptr(),
                    1,
                    0,
                    &mut twice_id,
                )
            },
            S_OK
        );
        assert_ne!(twice_id, value_id);

        let mut add_ids = [0i32; 4];
        let p_add = [
            n_add.as_ptr(),
            n_right.as_ptr(),
            n_missing.as_ptr(),
            n_left.as_ptr(),
        ];
        assert_eq!(
            unsafe {
                (vtbl.get_ids_of_names)(
                    raw.cast(),
                    &IID_NULL,
                    p_add.as_ptr(),
                    4,
                    0,
                    add_ids.as_mut_ptr(),
                )
            },
            DISP_E_UNKNOWNNAME
        );
        assert!(add_ids[0] > 0);
        assert_eq!(add_ids[1], 1);
        assert_eq!(add_ids[2], DISPID_UNKNOWN);
        assert_eq!(add_ids[3], 0);

        let mut hidden_id = 0i32;
        let p_hidden = [n_hidden.as_ptr()];
        assert_eq!(
            unsafe {
                (vtbl.get_ids_of_names)(
                    raw.cast(),
                    &IID_NULL,
                    p_hidden.as_ptr(),
                    1,
                    0,
                    &mut hidden_id,
                )
            },
            DISP_E_UNKNOWNNAME
        );
        assert_eq!(hidden_id, DISPID_UNKNOWN);

        // 4. Table-driven Invoke: PropertyPut on field, Method call, Named + ByRef writeback, Exception.
        let mut invoker = MockInvoker;
        let _scope = unsafe {
            MemberInvokerScope::enter(NonNull::from(&mut invoker as &mut dyn MemberInvoker))
        };

        // Put Value = 21 via DISPID_PROPERTYPUT.
        let mut put_val = Variant::from_i32(21);
        let mut put_named = [DISPID_PROPERTYPUT];
        let mut put_params = DispParams {
            rgvarg: (&raw mut put_val).cast::<VariantCore>(),
            rgdispid_named_args: put_named.as_mut_ptr(),
            c_args: 1,
            c_named_args: 1,
        };
        assert_eq!(
            unsafe {
                (vtbl.invoke)(
                    raw.cast(),
                    value_id,
                    &IID_NULL,
                    0,
                    DISPATCH_PROPERTYPUT,
                    (&raw mut put_params).cast(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                )
            },
            S_OK
        );

        // Call Twice() -> 42.
        let mut empty_params = DispParams::default();
        let mut ret_twice = Variant::empty();
        assert_eq!(
            unsafe {
                (vtbl.invoke)(
                    raw.cast(),
                    twice_id,
                    &IID_NULL,
                    0,
                    DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                    (&raw mut empty_params).cast(),
                    (&raw mut ret_twice).cast(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                )
            },
            S_OK
        );
        assert_eq!(ret_twice.as_i32(), Some(42));

        // Named args + ByRef writeback: AddPair(Right:=byref_target, Left:=30) -> 35, target becomes 6.
        let mut target_i32: i32 = 5;
        let mut arg_right = unsafe { Variant::by_ref(0x0003, (&raw mut target_i32).cast()) };
        let mut arg_left = Variant::from_i32(30);
        let mut rgvarg = [arg_right.core(), arg_left.core()];
        let mut named_ids = [1i32, 0i32];
        let mut named_params = DispParams {
            rgvarg: rgvarg.as_mut_ptr(),
            rgdispid_named_args: named_ids.as_mut_ptr(),
            c_args: 2,
            c_named_args: 2,
        };
        let mut ret_sum = Variant::empty();
        assert_eq!(
            unsafe {
                (vtbl.invoke)(
                    raw.cast(),
                    add_ids[0],
                    &IID_NULL,
                    0,
                    DISPATCH_METHOD,
                    (&raw mut named_params).cast(),
                    (&raw mut ret_sum).cast(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                )
            },
            S_OK
        );
        assert_eq!(ret_sum.as_i32(), Some(35));
        assert_eq!(target_i32, 6);
        let _ = (&mut arg_right, &mut arg_left);

        // Exception path populates ExcepInfo BSTRs and returns DISP_E_EXCEPTION.
        let n_fail = utf16_nul("Fail");
        let p_fail = [n_fail.as_ptr()];
        let mut fail_id = DISPID_UNKNOWN;
        assert_eq!(
            unsafe {
                (vtbl.get_ids_of_names)(raw.cast(), &IID_NULL, p_fail.as_ptr(), 1, 0, &mut fail_id)
            },
            S_OK
        );
        let mut excep = ExcepInfo::default();
        let hr = unsafe {
            (vtbl.invoke)(
                raw.cast(),
                fail_id,
                &IID_NULL,
                0,
                DISPATCH_METHOD,
                (&raw mut empty_params).cast(),
                core::ptr::null_mut(),
                (&raw mut excep).cast(),
                core::ptr::null_mut(),
            )
        };
        assert_eq!(hr, DISP_E_EXCEPTION);
        assert_eq!(excep.scode, (0x800A_0000u32 | 5000u32) as i32);
        assert_eq!(excep.dw_help_context, 7);
        let desc = unsafe { BStr::from_raw_bstr(excep.bstr_description) };
        let src = unsafe { BStr::from_raw_bstr(excep.bstr_source) };
        assert_eq!(desc.as_str(), "boom");
        assert_eq!(src.as_str(), "Widget");

        // 5. Off-thread Invoke returns RPC_E_WRONG_THREAD without entering the invoker.
        struct SendRawPtr(*mut crate::RawRuntimeIUnknown);
        unsafe impl Send for SendRawPtr {}
        let send_ptr = SendRawPtr(raw);
        let wrong_hr = std::thread::spawn(move || {
            let ptr = send_ptr;
            let vtbl = unsafe { &*((*ptr.0).vtbl.cast::<IDispatchVtbl>()) };
            let mut params = DispParams::default();
            unsafe {
                (vtbl.invoke)(
                    ptr.0.cast(),
                    twice_id,
                    &IID_NULL,
                    0,
                    DISPATCH_METHOD,
                    (&raw mut params).cast(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                )
            }
        })
        .join()
        .expect("worker thread");
        assert_eq!(wrong_hr, RPC_E_WRONG_THREAD);
    }
    /// dc-yqq0.24: an early-bound caller reaches each member through its vtable slot.
    #[cfg(all(target_os = "windows", target_arch = "x86_64"))]
    #[test]
    fn slot_thunks_call_members_through_the_vtable() {
        let descriptor = sample_widget_descriptor();
        let object = ObjectRef::from_project_instance(1, 0, 0, false, descriptor);
        let raw = object.raw_iunknown().cast::<c_void>();
        let members = descriptor.dispatch_interface().expect("dual").members;
        let slot_of = |name: &str, kind: RuntimeMemberInvokeKind| {
            let member = members
                .iter()
                .find(|m| m.name == name && m.invoke_kind == kind)
                .expect("member");
            usize::from(member.vtable_slot.expect("slot"))
        };
        let entry = |slot: usize| unsafe { *(*raw.cast::<*const *const c_void>()).add(slot) };
        let placeholder = compat_slot_not_implemented
            as unsafe extern "system" fn(*mut c_void) -> HResult
            as *const c_void;

        let mut invoker = MockInvoker;
        let _scope = unsafe {
            MemberInvokerScope::enter(NonNull::from(&mut invoker as &mut dyn MemberInvoker))
        };

        type PutLong = unsafe extern "system" fn(*mut c_void, i32) -> HResult;
        type GetLong = unsafe extern "system" fn(*mut c_void, *mut i32) -> HResult;
        type AddPair = unsafe extern "system" fn(*mut c_void, i32, *mut i32, *mut i32) -> HResult;
        type NoArgs = unsafe extern "system" fn(*mut c_void) -> HResult;

        let put = entry(slot_of("Value", RuntimeMemberInvokeKind::PropertyLet));
        assert_ne!(put, placeholder);
        let put: PutLong = unsafe { core::mem::transmute(put) };
        assert_eq!(unsafe { put(raw, 21) }, S_OK);

        let get: GetLong = unsafe {
            core::mem::transmute(entry(slot_of(
                "Value",
                RuntimeMemberInvokeKind::PropertyGet,
            )))
        };
        let mut value = 0;
        assert_eq!(unsafe { get(raw, &raw mut value) }, S_OK);
        assert_eq!(value, 21);

        let twice: GetLong = unsafe {
            core::mem::transmute(entry(slot_of("Twice", RuntimeMemberInvokeKind::Method)))
        };
        assert_eq!(unsafe { twice(raw, &raw mut value) }, S_OK);
        assert_eq!(value, 42);

        // ByVal 5 and ByRef 7: returns 12 and writes 8 back.
        let add: AddPair = unsafe {
            core::mem::transmute(entry(slot_of("AddPair", RuntimeMemberInvokeKind::Method)))
        };
        let mut right = 7;
        assert_eq!(unsafe { add(raw, 5, &raw mut right, &raw mut value) }, S_OK);
        assert_eq!((value, right), (12, 8));

        // Floating-point arguments: positions 1..=3 arrive in xmm1..xmm3, position 4 on the stack.
        type Scale =
            unsafe extern "system" fn(*mut c_void, f64, f32, f64, f64, *mut f64) -> HResult;
        let scale: Scale = unsafe {
            core::mem::transmute(entry(slot_of("Scale", RuntimeMemberInvokeKind::Method)))
        };
        let mut scaled = 0.0;
        assert_eq!(
            unsafe { scale(raw, 1.5, 4.0, 0.25, 2.0, &raw mut scaled) },
            S_OK
        );
        assert_eq!(scaled, 1.5 * 4.0 + 0.25 - 2.0);

        // A VBA error comes back as its HRESULT (0x800A0000 | number) and publishes rich
        // `IErrorInfo` on the current thread (dc-yqq0.27).
        let fail: NoArgs = unsafe {
            core::mem::transmute(entry(slot_of("Fail", RuntimeMemberInvokeKind::Method)))
        };
        assert_eq!(unsafe { fail(raw) }, 0x800A_1388u32 as i32);
        let mut err_ptr: *mut c_void = core::ptr::null_mut();
        assert_eq!(unsafe { oleaut::get_error_info(0, &mut err_ptr) }, S_OK);
        assert!(!err_ptr.is_null());
        let err_vtbl = unsafe { &*(*err_ptr.cast::<calcvba_automation::com::IErrorInfo>()).vtbl };
        let mut err_guid = Guid::default();
        assert_eq!(unsafe { (err_vtbl.get_guid)(err_ptr, &mut err_guid) }, S_OK);
        assert_eq!(
            err_guid,
            descriptor.default_interface_iid.expect("default IID")
        );
        let mut bstr_desc: *mut u16 = core::ptr::null_mut();
        assert_eq!(
            unsafe { (err_vtbl.get_description)(err_ptr, &mut bstr_desc) },
            S_OK
        );
        let desc = unsafe { BStr::from_raw_bstr(bstr_desc) };
        assert_eq!(desc.as_str(), "boom");
        unsafe {
            (err_vtbl.unknown.release)(err_ptr);
        }

        // Subsequent succeeding early-bound slot call clears thread error info.
        assert_eq!(unsafe { twice(raw, &raw mut value) }, S_OK);
        let mut cleared: *mut c_void = core::ptr::null_mut();
        assert_eq!(
            unsafe { oleaut::get_error_info(0, &mut cleared) },
            calcvba_automation::com::S_FALSE
        );
        assert!(cleared.is_null());
    }

    #[test]
    fn implements_faces_connection_points_and_error_info_share_one_identity() {
        use crate::class_descriptor::{
            RuntimeEventSpec, RuntimeImplementedInterfaceSpec, default_class_interface_iid,
            default_class_source_interface_iid,
        };
        use calcvba_automation::com::{
            CONNECT_E_CANNOTCONNECT, CONNECT_E_NOCONNECTION, ConnectData, E_NOINTERFACE,
            IConnectionPoint, IConnectionPointContainer, IEnumConnectionPoints, IEnumConnections,
            IErrorInfo, IID_ICONNECTIONPOINT, IID_ICONNECTIONPOINTCONTAINER, IID_IDISPATCH,
            IID_IENUMCONNECTIONPOINTS, IID_IENUMCONNECTIONS, IID_ISUPPORTERRORINFO, IID_IUNKNOWN,
            ISupportErrorInfo, ITypeInfo, S_FALSE,
        };
        use core::sync::atomic::AtomicU32;

        let shape_iid = default_class_interface_iid("Dispatch27Unit", "IShape");
        let square_iid = default_class_interface_iid("Dispatch27Unit", "Square");
        let source_iid = default_class_source_interface_iid("Dispatch27Unit", "Square");

        let spec = RuntimeClassSpec {
            unit_name: "Dispatch27Unit".into(),
            class_index: 0,
            name: "Square".into(),
            runtime_type_name: None,
            default_interface_iid: None,
            source_interface_iid: None,
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
            record_members: Vec::new(),
            project_records: Vec::new(),
            fields: vec![RuntimeClassFieldSpec {
                name: "Side".into(),
                token: 10,
                value_type: RuntimeValueType::Long,
                array_element_type: None,
                externally_visible: true,
                dispatch_id: None,
                get_vtable_slot: None,
                put_vtable_slot: None,
                putref_vtable_slot: None,
            }],
            as_new_fields: Vec::new(),
            implements: vec!["IShape".into()],
            methods: vec![RuntimeMemberSpec {
                name: "PrimaryFail".into(),
                dispatch_id: None,
                vtable_slot: None,
                invoke_kind: RuntimeMemberInvokeKind::Method,
                params: Vec::new(),
                return_type: None,
                is_default_member: false,
                is_enumerator_member: false,
                externally_visible: true,
                func_flags: 0,
                backing: RuntimeMemberBacking::Proc(100),
            }],
            implemented_interfaces: vec![RuntimeImplementedInterfaceSpec {
                name: "IShape".into(),
                qualified_name: "Dispatch27Unit.IShape".into(),
                guid: Some(shape_iid),
                members: vec![
                    RuntimeMemberSpec {
                        name: "Area".into(),
                        dispatch_id: Some(0x6003_0000),
                        vtable_slot: Some(7),
                        invoke_kind: RuntimeMemberInvokeKind::Method,
                        params: Vec::new(),
                        return_type: Some(RuntimeValueType::Long),
                        is_default_member: false,
                        is_enumerator_member: false,
                        externally_visible: true,
                        func_flags: 0,
                        backing: RuntimeMemberBacking::Proc(200),
                    },
                    RuntimeMemberSpec {
                        name: "Broken".into(),
                        dispatch_id: Some(0x6003_0001),
                        vtable_slot: Some(8),
                        invoke_kind: RuntimeMemberInvokeKind::Method,
                        params: Vec::new(),
                        return_type: None,
                        is_default_member: false,
                        is_enumerator_member: false,
                        externally_visible: true,
                        func_flags: 0,
                        backing: RuntimeMemberBacking::Proc(201),
                    },
                ],
            }],
            events: vec![
                RuntimeEventSpec {
                    name: "Changed".into(),
                    event_id: 0,
                    dispid: 0,
                    callback_arity: 2,
                    externally_visible: true,
                },
                RuntimeEventSpec {
                    name: "Secret".into(),
                    event_id: 1,
                    dispid: 1,
                    callback_arity: 0,
                    externally_visible: false,
                },
            ],
        };
        let descriptor = build_runtime_class_descriptor(&spec);
        assert_eq!(descriptor.default_interface_iid, Some(square_iid));
        assert_eq!(descriptor.source_interface_iid, Some(source_iid));

        struct SquareInvoker;
        impl MemberInvoker for SquareInvoker {
            fn invoke_member(
                &mut self,
                object: ObjectRef,
                member: &'static RuntimeMemberDescriptor,
                invoke_kind: RuntimeMemberInvokeKind,
                args: &mut [Variant],
                _aliases: &[usize],
            ) -> Result<Variant, MemberFault> {
                match member.backing {
                    RuntimeMemberBacking::Field(token) => {
                        invoke_field_directly(&object, member, token, invoke_kind, args)
                    }
                    RuntimeMemberBacking::Proc(100) => Err(MemberFault {
                        number: 513,
                        description: "primary fail".into(),
                        source: "Square".into(),
                        help_file: "sq.chm".into(),
                        help_context: 11,
                    }),
                    RuntimeMemberBacking::Proc(200) => {
                        let s = object
                            .project_field_get(10)
                            .and_then(|v| v.as_i32())
                            .unwrap_or(0);
                        Ok(Variant::from_i32(s * s))
                    }
                    RuntimeMemberBacking::Proc(201) => Err(MemberFault {
                        number: 514,
                        description: "shape broken".into(),
                        source: "IShape".into(),
                        help_file: "shape.chm".into(),
                        help_context: 22,
                    }),
                    _ => Err(MemberFault::from_number(438)),
                }
            }
        }

        type EventLog = std::rc::Rc<std::cell::RefCell<Vec<(i32, Vec<i32>)>>>;

        // Minimal COM event sink recording `(dispid, logical_args)`.
        #[repr(C)]
        struct TestSink {
            vtbl: *const IDispatchVtbl,
            ref_count: AtomicU32,
            events: EventLog,
        }
        unsafe extern "system" fn sink_qi(
            this: *mut c_void,
            riid: *const Guid,
            ppv: *mut *mut c_void,
        ) -> HResult {
            unsafe {
                *ppv = core::ptr::null_mut();
                let iid = *riid;
                if iid == IID_IUNKNOWN || iid == IID_IDISPATCH {
                    sink_add_ref(this);
                    *ppv = this;
                    S_OK
                } else {
                    E_NOINTERFACE
                }
            }
        }
        unsafe extern "system" fn sink_add_ref(this: *mut c_void) -> u32 {
            unsafe {
                (*this.cast::<TestSink>())
                    .ref_count
                    .fetch_add(1, Ordering::AcqRel)
                    + 1
            }
        }
        unsafe extern "system" fn sink_release(this: *mut c_void) -> u32 {
            unsafe {
                let prev = (*this.cast::<TestSink>())
                    .ref_count
                    .fetch_sub(1, Ordering::AcqRel);
                let rem = prev - 1;
                if rem == 0 {
                    drop(Box::from_raw(this.cast::<TestSink>()));
                }
                rem
            }
        }
        unsafe extern "system" fn sink_tic(_: *mut c_void, c: *mut u32) -> HResult {
            unsafe {
                *c = 0;
            }
            S_OK
        }
        unsafe extern "system" fn sink_ti(
            _: *mut c_void,
            _: u32,
            _: u32,
            out: *mut *mut c_void,
        ) -> HResult {
            unsafe {
                *out = core::ptr::null_mut();
            }
            E_NOTIMPL
        }
        unsafe extern "system" fn sink_ids(
            _: *mut c_void,
            _: *const Guid,
            _: *const *const u16,
            _: u32,
            _: u32,
            _: *mut i32,
        ) -> HResult {
            E_NOTIMPL
        }
        unsafe extern "system" fn sink_invoke(
            this: *mut c_void,
            dispid: i32,
            _: *const Guid,
            _: u32,
            _: u16,
            params: *mut c_void,
            _: *mut c_void,
            _: *mut c_void,
            _: *mut u32,
        ) -> HResult {
            unsafe {
                let sink = &*this.cast::<TestSink>();
                let mut logical = Vec::new();
                if !params.is_null() {
                    let dp = &*params.cast::<DispParams>();
                    for i in (0..dp.c_args as usize).rev() {
                        let cell = &*dp.rgvarg.add(i).cast::<Variant>();
                        logical.push(cell.as_i32().unwrap_or(-1));
                    }
                }
                sink.events.borrow_mut().push((dispid, logical));
            }
            S_OK
        }
        static SINK_VTBL: IDispatchVtbl = IDispatchVtbl {
            unknown: IUnknownVtbl {
                query_interface: sink_qi,
                add_ref: sink_add_ref,
                release: sink_release,
            },
            get_type_info_count: sink_tic,
            get_type_info: sink_ti,
            get_ids_of_names: sink_ids,
            invoke: sink_invoke,
        };

        // Non-dispatch IUnknown-only sink (must be rejected by Advise with CONNECT_E_CANNOTCONNECT).
        #[repr(C)]
        struct BadSink {
            vtbl: *const IUnknownVtbl,
            ref_count: AtomicU32,
        }
        unsafe extern "system" fn bad_qi(
            this: *mut c_void,
            riid: *const Guid,
            ppv: *mut *mut c_void,
        ) -> HResult {
            unsafe {
                *ppv = core::ptr::null_mut();
                if *riid == IID_IUNKNOWN {
                    bad_add_ref(this);
                    *ppv = this;
                    S_OK
                } else {
                    E_NOINTERFACE
                }
            }
        }
        unsafe extern "system" fn bad_add_ref(this: *mut c_void) -> u32 {
            unsafe {
                (*this.cast::<BadSink>())
                    .ref_count
                    .fetch_add(1, Ordering::AcqRel)
                    + 1
            }
        }
        unsafe extern "system" fn bad_release(this: *mut c_void) -> u32 {
            unsafe {
                let prev = (*this.cast::<BadSink>())
                    .ref_count
                    .fetch_sub(1, Ordering::AcqRel);
                let rem = prev - 1;
                if rem == 0 {
                    drop(Box::from_raw(this.cast::<BadSink>()));
                }
                rem
            }
        }
        static BAD_VTBL: IUnknownVtbl = IUnknownVtbl {
            query_interface: bad_qi,
            add_ref: bad_add_ref,
            release: bad_release,
        };

        let before = crate::current_thread_live_handle_counts();
        {
            // A class without events does not expose IConnectionPointContainer.
            let plain_desc = sample_widget_descriptor();
            let plain = ObjectRef::from_project_instance(99, 0, 0, false, plain_desc);
            let plain_raw = plain.raw_iunknown();
            let plain_vtbl = unsafe { &*(*plain_raw).vtbl };
            let mut no_cpc = core::ptr::null_mut();
            assert_eq!(
                unsafe {
                    (plain_vtbl.query_interface)(
                        plain_raw.cast(),
                        &IID_ICONNECTIONPOINTCONTAINER,
                        &mut no_cpc,
                    )
                },
                E_NOINTERFACE
            );
            assert!(no_cpc.is_null());

            let object = ObjectRef::from_project_instance(1, 0, 0, false, descriptor);
            object.project_field_set(10, Variant::from_i32(6));
            let primary = object.raw_iunknown();
            let primary_vtbl = unsafe { &*((*primary).vtbl.cast::<IDispatchVtbl>()) };

            let mut invoker = SquareInvoker;
            let _scope = unsafe {
                MemberInvokerScope::enter(NonNull::from(&mut invoker as &mut dyn MemberInvoker))
            };

            // 1. Implements face: distinct pointer, single COM identity, shared refcount.
            let mut shape_ptr: *mut c_void = core::ptr::null_mut();
            assert_eq!(
                unsafe {
                    (primary_vtbl.unknown.query_interface)(
                        primary.cast(),
                        &shape_iid,
                        &mut shape_ptr,
                    )
                },
                S_OK
            );
            assert!(!shape_ptr.is_null());
            assert_ne!(shape_ptr, primary.cast());
            assert_eq!(object.strong_count(), 2);

            let shape_vtbl = unsafe {
                &*((*shape_ptr.cast::<crate::RawRuntimeIUnknown>())
                    .vtbl
                    .cast::<IDispatchVtbl>())
            };
            let mut back_unk: *mut c_void = core::ptr::null_mut();
            assert_eq!(
                unsafe {
                    (shape_vtbl.unknown.query_interface)(shape_ptr, &IID_IUNKNOWN, &mut back_unk)
                },
                S_OK
            );
            assert_eq!(back_unk, primary.cast());
            unsafe { (primary_vtbl.unknown.release)(back_unk) };

            let mut back_disp: *mut c_void = core::ptr::null_mut();
            assert_eq!(
                unsafe {
                    (shape_vtbl.unknown.query_interface)(shape_ptr, &IID_IDISPATCH, &mut back_disp)
                },
                S_OK
            );
            assert_eq!(back_disp, primary.cast());
            unsafe { (primary_vtbl.unknown.release)(back_disp) };

            let mut same_shape: *mut c_void = core::ptr::null_mut();
            assert_eq!(
                unsafe {
                    (shape_vtbl.unknown.query_interface)(shape_ptr, &shape_iid, &mut same_shape)
                },
                S_OK
            );
            assert_eq!(same_shape, shape_ptr);
            unsafe { (shape_vtbl.unknown.release)(same_shape) };

            // ObjectRef adopted from secondary face preserves project identity (`Is`).
            let via_face =
                unsafe { ObjectRef::from_raw_iunknown_addref(shape_ptr.cast()) }.expect("adopted");
            assert!(via_face.is_compat_object());
            assert!(via_face.is_project_instance());
            assert_eq!(via_face.is_identity(), object.is_identity());
            assert_eq!(via_face.strong_count(), 3);
            drop(via_face);
            assert_eq!(object.strong_count(), 2);

            // IShape IDispatch: GetTypeInfo + GetIDsOfNames + Invoke("Area") -> 36.
            let mut tic = 0u32;
            assert_eq!(
                unsafe { (shape_vtbl.get_type_info_count)(shape_ptr, &mut tic) },
                S_OK
            );
            assert_eq!(tic, 1);
            let mut shape_ti: *mut c_void = core::ptr::null_mut();
            assert_eq!(
                unsafe { (shape_vtbl.get_type_info)(shape_ptr, 0, 0, &mut shape_ti) },
                S_OK
            );
            let ti_vtbl = unsafe { &*(*shape_ti.cast::<ITypeInfo>()).vtbl };
            let mut attr = core::ptr::null_mut();
            assert_eq!(
                unsafe { (ti_vtbl.get_type_attr)(shape_ti, &mut attr) },
                S_OK
            );
            assert_eq!(unsafe { (*attr).guid }, shape_iid);
            assert_eq!(unsafe { (*attr).c_funcs }, 2);
            unsafe {
                (ti_vtbl.release_type_attr)(shape_ti, attr);
                (ti_vtbl.unknown.release)(shape_ti);
            }

            let area_w: Vec<u16> = "Area\0".encode_utf16().collect();
            let area_names = [area_w.as_ptr()];
            let mut area_id = DISPID_UNKNOWN;
            assert_eq!(
                unsafe {
                    (shape_vtbl.get_ids_of_names)(
                        shape_ptr,
                        &IID_NULL,
                        area_names.as_ptr(),
                        1,
                        0,
                        &mut area_id,
                    )
                },
                S_OK
            );
            assert_eq!(area_id, 0x6003_0000);
            let mut empty_params = DispParams::default();
            let mut area_res = Variant::empty();
            assert_eq!(
                unsafe {
                    (shape_vtbl.invoke)(
                        shape_ptr,
                        area_id,
                        &IID_NULL,
                        0,
                        DISPATCH_METHOD,
                        (&raw mut empty_params).cast(),
                        (&raw mut area_res).cast(),
                        core::ptr::null_mut(),
                        core::ptr::null_mut(),
                    )
                },
                S_OK
            );
            assert_eq!(area_res.as_i32(), Some(36));

            // 2. ISupportErrorInfo + IErrorInfo (late-bound AND early-bound on both faces).
            let mut sei_ptr: *mut c_void = core::ptr::null_mut();
            assert_eq!(
                unsafe {
                    (shape_vtbl.unknown.query_interface)(
                        shape_ptr,
                        &IID_ISUPPORTERRORINFO,
                        &mut sei_ptr,
                    )
                },
                S_OK
            );
            let sei_vtbl = unsafe { &*(*sei_ptr.cast::<ISupportErrorInfo>()).vtbl };
            let mut sei_unk = core::ptr::null_mut();
            assert_eq!(
                unsafe { (sei_vtbl.unknown.query_interface)(sei_ptr, &IID_IUNKNOWN, &mut sei_unk) },
                S_OK
            );
            assert_eq!(sei_unk, primary.cast());
            unsafe { (primary_vtbl.unknown.release)(sei_unk) };
            assert_eq!(
                unsafe { (sei_vtbl.interface_supports_error_info)(sei_ptr, &IID_IDISPATCH) },
                S_OK
            );
            assert_eq!(
                unsafe { (sei_vtbl.interface_supports_error_info)(sei_ptr, &square_iid) },
                S_OK
            );
            assert_eq!(
                unsafe { (sei_vtbl.interface_supports_error_info)(sei_ptr, &shape_iid) },
                S_OK
            );
            assert_eq!(
                unsafe { (sei_vtbl.interface_supports_error_info)(sei_ptr, &source_iid) },
                S_FALSE
            );
            unsafe { (sei_vtbl.unknown.release)(sei_ptr) };

            let read_error_info = || -> (Guid, String, String, String, u32) {
                let mut err_ptr: *mut c_void = core::ptr::null_mut();
                assert_eq!(unsafe { oleaut::get_error_info(0, &mut err_ptr) }, S_OK);
                assert!(!err_ptr.is_null());
                let vtbl = unsafe { &*(*err_ptr.cast::<IErrorInfo>()).vtbl };
                let mut g = Guid::default();
                let mut s = core::ptr::null_mut();
                let mut d = core::ptr::null_mut();
                let mut h = core::ptr::null_mut();
                let mut ctx = 0u32;
                assert_eq!(unsafe { (vtbl.get_guid)(err_ptr, &mut g) }, S_OK);
                assert_eq!(unsafe { (vtbl.get_source)(err_ptr, &mut s) }, S_OK);
                assert_eq!(unsafe { (vtbl.get_description)(err_ptr, &mut d) }, S_OK);
                assert_eq!(unsafe { (vtbl.get_help_file)(err_ptr, &mut h) }, S_OK);
                assert_eq!(unsafe { (vtbl.get_help_context)(err_ptr, &mut ctx) }, S_OK);
                let out = (
                    g,
                    unsafe { BStr::from_raw_bstr(s) }.as_str().to_string(),
                    unsafe { BStr::from_raw_bstr(d) }.as_str().to_string(),
                    unsafe { BStr::from_raw_bstr(h) }.as_str().to_string(),
                    ctx,
                );
                unsafe { (vtbl.unknown.release)(err_ptr) };
                out
            };

            // Late-bound Invoke fault on IShape publishes IErrorInfo with shape_iid.
            let mut excep = ExcepInfo::default();
            assert_eq!(
                unsafe {
                    (shape_vtbl.invoke)(
                        shape_ptr,
                        0x6003_0001,
                        &IID_NULL,
                        0,
                        DISPATCH_METHOD,
                        (&raw mut empty_params).cast(),
                        core::ptr::null_mut(),
                        (&raw mut excep).cast(),
                        core::ptr::null_mut(),
                    )
                },
                DISP_E_EXCEPTION
            );
            // SAFETY: Invoke returned three owned BSTRs to this caller.
            let [source, description, help_file] = unsafe {
                [
                    excep.bstr_source,
                    excep.bstr_description,
                    excep.bstr_help_file,
                ]
                .map(|raw| BStr::from_raw_bstr(raw))
            };
            assert_eq!(source.as_str(), "IShape");
            assert_eq!(description.as_str(), "shape broken");
            assert_eq!(help_file.as_str(), "shape.chm");
            assert_eq!(
                read_error_info(),
                (
                    shape_iid,
                    "IShape".to_string(),
                    "shape broken".to_string(),
                    "shape.chm".to_string(),
                    22
                )
            );

            // Early-bound call on primary interface (`PrimaryFail`) publishes IErrorInfo with square_iid.
            let primary_fail_slot = descriptor
                .dispatch_interface()
                .unwrap()
                .members
                .iter()
                .find(|m| m.name == "PrimaryFail")
                .unwrap()
                .vtable_slot
                .unwrap();
            assert_eq!(
                invoke_early_bound_slot(&object, None, primary_fail_slot, &mut []),
                Err((0x800A_0000u32 | 513u32) as i32)
            );
            assert_eq!(
                read_error_info(),
                (
                    square_iid,
                    "Square".to_string(),
                    "primary fail".to_string(),
                    "sq.chm".to_string(),
                    11
                )
            );

            // Early-bound call on implemented interface (`IShape.Broken`, slot 8) publishes IErrorInfo with shape_iid.
            assert_eq!(
                invoke_early_bound_slot(&object, Some(shape_iid), 8, &mut []),
                Err((0x800A_0000u32 | 514u32) as i32)
            );
            assert_eq!(
                read_error_info(),
                (
                    shape_iid,
                    "IShape".to_string(),
                    "shape broken".to_string(),
                    "shape.chm".to_string(),
                    22
                )
            );

            // Early-bound succeeding call (`IShape.Area`, slot 7) returns 36 and clears IErrorInfo.
            assert_eq!(
                invoke_early_bound_slot(&object, Some(shape_iid), 7, &mut []),
                Ok(Variant::from_i32(36))
            );
            let mut none_err = core::ptr::null_mut();
            assert_eq!(unsafe { oleaut::get_error_info(0, &mut none_err) }, S_FALSE);
            assert!(none_err.is_null());

            // Real Win64 early-bound vtable slot calls on `shape_ptr` (slot 7 Area & slot 8 Broken).
            #[cfg(all(target_os = "windows", target_arch = "x86_64"))]
            {
                type GetLong = unsafe extern "system" fn(*mut c_void, *mut i32) -> HResult;
                type NoArgs = unsafe extern "system" fn(*mut c_void) -> HResult;
                let area_fn: GetLong = unsafe {
                    core::mem::transmute(*(*shape_ptr.cast::<*const *const c_void>()).add(7))
                };
                let mut area_out = 0i32;
                assert_eq!(unsafe { area_fn(shape_ptr, &raw mut area_out) }, S_OK);
                assert_eq!(area_out, 36);

                let broken_fn: NoArgs = unsafe {
                    core::mem::transmute(*(*shape_ptr.cast::<*const *const c_void>()).add(8))
                };
                assert_eq!(
                    unsafe { broken_fn(shape_ptr) },
                    (0x800A_0000u32 | 514u32) as i32
                );
                assert_eq!(
                    read_error_info(),
                    (
                        shape_iid,
                        "IShape".to_string(),
                        "shape broken".to_string(),
                        "shape.chm".to_string(),
                        22
                    )
                );
            }

            // 3. Connection points: IConnectionPointContainer, IConnectionPoint, enumerators, RaiseEvent.
            let mut cpc_ptr: *mut c_void = core::ptr::null_mut();
            assert_eq!(
                unsafe {
                    (shape_vtbl.unknown.query_interface)(
                        shape_ptr,
                        &IID_ICONNECTIONPOINTCONTAINER,
                        &mut cpc_ptr,
                    )
                },
                S_OK
            );
            let cpc_vtbl = unsafe { &*(*cpc_ptr.cast::<IConnectionPointContainer>()).vtbl };
            let mut cpc_unk = core::ptr::null_mut();
            assert_eq!(
                unsafe { (cpc_vtbl.unknown.query_interface)(cpc_ptr, &IID_IUNKNOWN, &mut cpc_unk) },
                S_OK
            );
            assert_eq!(cpc_unk, primary.cast());
            unsafe { (primary_vtbl.unknown.release)(cpc_unk) };

            let mut missing_cp = core::ptr::null_mut();
            assert_eq!(
                unsafe { (cpc_vtbl.find_connection_point)(cpc_ptr, &shape_iid, &mut missing_cp) },
                CONNECT_E_NOCONNECTION
            );
            assert!(missing_cp.is_null());

            let mut cp_ptr: *mut c_void = core::ptr::null_mut();
            assert_eq!(
                unsafe { (cpc_vtbl.find_connection_point)(cpc_ptr, &source_iid, &mut cp_ptr) },
                S_OK
            );
            assert!(!cp_ptr.is_null());
            let cp_vtbl = unsafe { &*(*cp_ptr.cast::<IConnectionPoint>()).vtbl };

            let mut got_iid = Guid::default();
            assert_eq!(
                unsafe { (cp_vtbl.get_connection_interface)(cp_ptr, &mut got_iid) },
                S_OK
            );
            assert_eq!(got_iid, source_iid);

            let mut got_cpc = core::ptr::null_mut();
            assert_eq!(
                unsafe { (cp_vtbl.get_connection_point_container)(cp_ptr, &mut got_cpc) },
                S_OK
            );
            assert_eq!(got_cpc, cpc_ptr);
            unsafe { (cpc_vtbl.unknown.release)(got_cpc) };

            // EnumConnectionPoints: Next -> S_OK, Next -> S_FALSE, Reset, Clone, Skip.
            let mut ecp_ptr = core::ptr::null_mut();
            assert_eq!(
                unsafe { (cpc_vtbl.enum_connection_points)(cpc_ptr, &mut ecp_ptr) },
                S_OK
            );
            let ecp_vtbl = unsafe { &*(*ecp_ptr.cast::<IEnumConnectionPoints>()).vtbl };
            let mut qi_ecp = core::ptr::null_mut();
            assert_eq!(
                unsafe {
                    (ecp_vtbl.unknown.query_interface)(
                        ecp_ptr,
                        &IID_IENUMCONNECTIONPOINTS,
                        &mut qi_ecp,
                    )
                },
                S_OK
            );
            unsafe { (ecp_vtbl.unknown.release)(qi_ecp) };
            let mut yielded_cp = core::ptr::null_mut();
            let mut fetched = 0u32;
            assert_eq!(
                unsafe { (ecp_vtbl.next)(ecp_ptr, 1, &mut yielded_cp, &mut fetched) },
                S_OK
            );
            assert_eq!(fetched, 1);
            assert_eq!(yielded_cp, cp_ptr);
            unsafe { (cp_vtbl.unknown.release)(yielded_cp) };
            assert_eq!(
                unsafe { (ecp_vtbl.next)(ecp_ptr, 1, &mut yielded_cp, &mut fetched) },
                S_FALSE
            );
            assert_eq!(fetched, 0);
            assert_eq!(unsafe { (ecp_vtbl.reset)(ecp_ptr) }, S_OK);
            let mut cloned_ecp = core::ptr::null_mut();
            assert_eq!(unsafe { (ecp_vtbl.clone)(ecp_ptr, &mut cloned_ecp) }, S_OK);
            assert_eq!(unsafe { (ecp_vtbl.skip)(cloned_ecp, 1) }, S_OK);
            assert_eq!(unsafe { (ecp_vtbl.skip)(cloned_ecp, 1) }, S_FALSE);
            unsafe {
                (ecp_vtbl.unknown.release)(cloned_ecp);
                (ecp_vtbl.unknown.release)(ecp_ptr);
            }

            // Advise rejects non-dispatch sink.
            let bad_sink = Box::into_raw(Box::new(BadSink {
                vtbl: &BAD_VTBL,
                ref_count: AtomicU32::new(1),
            }))
            .cast::<c_void>();
            let mut bad_cookie = 99u32;
            assert_eq!(
                unsafe { (cp_vtbl.advise)(cp_ptr, bad_sink, &mut bad_cookie) },
                CONNECT_E_CANNOTCONNECT
            );
            assert_eq!(bad_cookie, 0);
            unsafe { bad_release(bad_sink) };

            // Advise accepts IDispatch sink; EnumConnections enumerates it; RaiseEvent delivers.
            let seen = std::rc::Rc::new(std::cell::RefCell::new(Vec::new()));
            let sink = Box::into_raw(Box::new(TestSink {
                vtbl: &SINK_VTBL,
                ref_count: AtomicU32::new(1),
                events: seen.clone(),
            }))
            .cast::<c_void>();
            let mut cookie = 0u32;
            assert_eq!(unsafe { (cp_vtbl.advise)(cp_ptr, sink, &mut cookie) }, S_OK);
            assert_ne!(cookie, 0);
            assert!(object.has_connection_point_sinks());

            // EnumConnections.
            let mut ec_ptr = core::ptr::null_mut();
            assert_eq!(
                unsafe { (cp_vtbl.enum_connections)(cp_ptr, &mut ec_ptr) },
                S_OK
            );
            let ec_vtbl = unsafe { &*(*ec_ptr.cast::<IEnumConnections>()).vtbl };
            let mut qi_ec = core::ptr::null_mut();
            assert_eq!(
                unsafe {
                    (ec_vtbl.unknown.query_interface)(ec_ptr, &IID_IENUMCONNECTIONS, &mut qi_ec)
                },
                S_OK
            );
            unsafe { (ec_vtbl.unknown.release)(qi_ec) };
            let mut cd = ConnectData::default();
            let mut ec_fetched = 0u32;
            assert_eq!(
                unsafe { (ec_vtbl.next)(ec_ptr, 1, &mut cd, &mut ec_fetched) },
                S_OK
            );
            assert_eq!(ec_fetched, 1);
            assert_eq!(cd.dw_cookie, cookie);
            assert_eq!(cd.punk, sink);
            unsafe { sink_release(cd.punk) };
            assert_eq!(
                unsafe { (ec_vtbl.next)(ec_ptr, 1, &mut cd, &mut ec_fetched) },
                S_FALSE
            );
            assert_eq!(unsafe { (ec_vtbl.reset)(ec_ptr) }, S_OK);
            let mut cloned_ec = core::ptr::null_mut();
            assert_eq!(unsafe { (ec_vtbl.clone)(ec_ptr, &mut cloned_ec) }, S_OK);
            assert_eq!(unsafe { (ec_vtbl.skip)(cloned_ec, 1) }, S_OK);
            assert_eq!(unsafe { (ec_vtbl.skip)(cloned_ec, 1) }, S_FALSE);
            unsafe {
                (ec_vtbl.unknown.release)(cloned_ec);
                (ec_vtbl.unknown.release)(ec_ptr);
            }

            // Fire public event 0 with two arguments (must arrive in logical order) and private event 1 (must be ignored).
            object
                .fire_connection_point_event(0, &[Variant::from_i32(10), Variant::from_i32(20)])
                .expect("fire public event");
            object
                .fire_connection_point_event(1, &[])
                .expect("private event ignored");
            assert_eq!(*seen.borrow(), vec![(0, vec![10, 20])]);

            // Unadvise stops further delivery and rejects double-unadvise.
            assert_eq!(unsafe { (cp_vtbl.unadvise)(cp_ptr, cookie) }, S_OK);
            assert_eq!(
                unsafe { (cp_vtbl.unadvise)(cp_ptr, cookie) },
                CONNECT_E_NOCONNECTION
            );
            assert!(!object.has_connection_point_sinks());
            object
                .fire_connection_point_event(0, &[Variant::from_i32(30), Variant::from_i32(40)])
                .expect("no sinks");
            assert_eq!(*seen.borrow(), vec![(0, vec![10, 20])]);

            // Also verify IConnectionPoint QI identity.
            let mut cp_self = core::ptr::null_mut();
            assert_eq!(
                unsafe {
                    (cp_vtbl.unknown.query_interface)(cp_ptr, &IID_ICONNECTIONPOINT, &mut cp_self)
                },
                S_OK
            );
            assert_eq!(cp_self, cp_ptr);
            unsafe { (cp_vtbl.unknown.release)(cp_self) };

            unsafe {
                sink_release(sink);
                (cp_vtbl.unknown.release)(cp_ptr);
                (cpc_vtbl.unknown.release)(cpc_ptr);
                (shape_vtbl.unknown.release)(shape_ptr);
            }
            assert_eq!(object.strong_count(), 1);
        }
        assert_eq!(
            before.balance_to(crate::current_thread_live_handle_counts()),
            Default::default()
        );
    }

    #[test]
    fn record_array_early_bound_slots_and_late_bound_idispatch_round_trip() {
        use crate::automation_record::{
            AutomationVariant, layout_from_record_info, record_info_get_guid, record_info_get_name,
            record_info_get_size, release_record_info, vba_record_array_to_automation_safearray,
        };
        use crate::class_descriptor::{
            RuntimeProjectRecordSpec, RuntimeRecordMemberSpec, default_record_type_guid,
        };
        use crate::safe_array::SafeArrayBound;
        use crate::{VbaRecord, VbaRecordFieldKind, VbaRecordFieldSpec, VbaRecordLayout};
        use std::sync::Arc;

        let pair_guid = default_record_type_guid("RecUnit", "SlotPair");
        let other_guid = default_record_type_guid("RecUnit", "OtherPair");
        let base_layout = VbaRecordLayout::new(vec![
            VbaRecordFieldSpec::named("Id", VbaRecordFieldKind::Long),
            VbaRecordFieldSpec::named("Label", VbaRecordFieldKind::String),
        ])
        .unwrap();
        let schema = base_layout.boundary_schema()[1..].to_vec();
        let pair_layout = Arc::new(base_layout.clone().with_identity("SlotPair", pair_guid));
        let other_layout = Arc::new(base_layout.with_identity("OtherPair", other_guid));

        let spec = RuntimeClassSpec {
            unit_name: "RecUnit".into(),
            class_index: 0,
            name: "PairBox".into(),
            runtime_type_name: None,
            default_interface_iid: None,
            source_interface_iid: None,
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
            record_members: vec![
                // Proc 0: Sub MutatePairs(ByRef items() As SlotPair) -> mutates & resizes
                RuntimeRecordMemberSpec {
                    procedure: 0,
                    parameter: Some(0),
                    is_array: true,
                    name: "SlotPair".into(),
                    guid: Some(pair_guid),
                    schema: schema.clone(),
                },
                // Proc 1: Function EchoPairs(ByRef items() As SlotPair) As SlotPair()
                RuntimeRecordMemberSpec {
                    procedure: 1,
                    parameter: Some(0),
                    is_array: true,
                    name: "SlotPair".into(),
                    guid: Some(pair_guid),
                    schema: schema.clone(),
                },
                RuntimeRecordMemberSpec {
                    procedure: 1,
                    parameter: None,
                    is_array: true,
                    name: "SlotPair".into(),
                    guid: Some(pair_guid),
                    schema: schema.clone(),
                },
            ],
            project_records: vec![RuntimeProjectRecordSpec {
                name: "SlotPair".into(),
                guid: Some(pair_guid),
                schema: schema.clone(),
            }],
            fields: Vec::new(),
            as_new_fields: Vec::new(),
            implements: Vec::new(),
            methods: vec![
                RuntimeMemberSpec {
                    name: "MutatePairs".into(),
                    dispatch_id: Some(10),
                    vtable_slot: Some(7),
                    invoke_kind: RuntimeMemberInvokeKind::Method,
                    params: vec![RuntimeParamSpec {
                        name: "items".into(),
                        value_type: RuntimeValueType::Array,
                        by_ref: true,
                        optional: false,
                        param_array: false,
                    }],
                    return_type: None,
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: true,
                    func_flags: 0,
                    backing: RuntimeMemberBacking::Proc(0),
                },
                RuntimeMemberSpec {
                    name: "EchoPairs".into(),
                    dispatch_id: Some(11),
                    vtable_slot: Some(8),
                    invoke_kind: RuntimeMemberInvokeKind::Method,
                    params: vec![RuntimeParamSpec {
                        name: "items".into(),
                        value_type: RuntimeValueType::Array,
                        by_ref: true,
                        optional: false,
                        param_array: false,
                    }],
                    return_type: Some(RuntimeValueType::Array),
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: true,
                    func_flags: 0,
                    backing: RuntimeMemberBacking::Proc(1),
                },
                RuntimeMemberSpec {
                    name: "RoundTripVariant".into(),
                    dispatch_id: Some(12),
                    vtable_slot: Some(9),
                    invoke_kind: RuntimeMemberInvokeKind::Method,
                    params: vec![RuntimeParamSpec {
                        name: "v".into(),
                        value_type: RuntimeValueType::Variant,
                        by_ref: true,
                        optional: false,
                        param_array: false,
                    }],
                    return_type: Some(RuntimeValueType::Variant),
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: true,
                    func_flags: 0,
                    backing: RuntimeMemberBacking::Proc(2),
                },
            ],
            implemented_interfaces: Vec::new(),
            events: Vec::new(),
        };
        let descriptor = build_runtime_class_descriptor(&spec);

        let make_pair = |layout: &Arc<VbaRecordLayout>, id: i32, text: &str| {
            let mut rec = VbaRecord::new_default(Arc::clone(layout)).unwrap();
            rec.write_field_variant(0, &Variant::from_i32(id)).unwrap();
            rec.write_field_variant(1, &Variant::from_string(text))
                .unwrap();
            rec
        };

        struct Handler {
            pair_layout: Arc<VbaRecordLayout>,
            invocations: std::cell::Cell<usize>,
        }
        impl MemberInvoker for Handler {
            fn invoke_member(
                &mut self,
                _object: ObjectRef,
                member: &'static RuntimeMemberDescriptor,
                _kind: RuntimeMemberInvokeKind,
                args: &mut [Variant],
                _aliases: &[usize],
            ) -> Result<Variant, MemberFault> {
                self.invocations.set(self.invocations.get() + 1);
                match member.dispatch_id {
                    10 => {
                        // Mutate element 0 in place (`Id += 100`, `Label &= "!"`).
                        let mut elem = args[0].safearray_element(0).unwrap().unwrap();
                        let id = elem.read_record_field_variant(0).unwrap().as_i32().unwrap();
                        let label = elem
                            .read_record_field_variant(1)
                            .unwrap()
                            .as_bstr()
                            .unwrap()
                            .as_str();
                        elem.write_record_field_variant(0, &Variant::from_i32(id + 100))
                            .unwrap();
                        elem.write_record_field_variant(
                            1,
                            &Variant::from_string(format!("{label}!")),
                        )
                        .unwrap();
                        args[0].set_safearray_element(0, &elem).unwrap();
                        Ok(Variant::empty())
                    }
                    11 => Ok(args[0].clone()),
                    12 => {
                        // Mutate Variant holding UDT array by appending a second element and returning it.
                        let elem0 = args[0].safearray_element(0).unwrap().unwrap();
                        let mut elem1 =
                            VbaRecord::new_default(Arc::clone(&self.pair_layout)).unwrap();
                        elem1
                            .write_field_variant(0, &Variant::from_i32(77))
                            .unwrap();
                        elem1
                            .write_field_variant(1, &Variant::from_string("added"))
                            .unwrap();
                        let expanded = Variant::from_safearray(
                            SafeArray::from_vba_records_nd(
                                vec![SafeArrayBound { count: 2, lower: 1 }],
                                Arc::clone(&self.pair_layout),
                                vec![elem0.as_vba_record().unwrap(), elem1],
                            )
                            .unwrap(),
                        );
                        args[0] = expanded.clone();
                        Ok(expanded)
                    }
                    _ => unreachable!(),
                }
            }
        }

        let before = crate::current_thread_live_handle_counts();
        {
            let mut handler = Handler {
                pair_layout: Arc::clone(&pair_layout),
                invocations: std::cell::Cell::new(0),
            };
            let _scope = unsafe { MemberInvokerScope::enter(NonNull::from(&mut handler)) };
            let object = ObjectRef::from_project_instance(1, 0, 0, false, descriptor);
            let raw_this = object.raw_iunknown().cast::<c_void>();
            let disp_vtbl = unsafe { &*(*raw_this.cast::<*const IDispatchVtbl>()) };

            let arr_variant = Variant::from_safearray(
                SafeArray::from_vba_records_nd(
                    vec![SafeArrayBound { count: 1, lower: 1 }],
                    Arc::clone(&pair_layout),
                    vec![make_pair(&pair_layout, 5, "hi")],
                )
                .unwrap(),
            );
            let wrong_variant = Variant::from_safearray(
                SafeArray::from_vba_records_nd(
                    vec![SafeArrayBound { count: 1, lower: 1 }],
                    Arc::clone(&other_layout),
                    vec![make_pair(&other_layout, 5, "wrong")],
                )
                .unwrap(),
            );

            // 1. Early-bound slot rejects wrong nominal UDT BEFORE invoking member.
            let mut wrong_psa =
                vba_record_array_to_automation_safearray(&wrong_variant, Some(&other_layout))
                    .unwrap();
            let args_wrong = [
                raw_this.expose_provenance() as u64,
                (&raw mut wrong_psa).expose_provenance() as u64,
            ];
            let xmm = [0u64; 3];
            assert_eq!(
                unsafe { invoke_early_bound_slot_raw(7, args_wrong.as_ptr(), xmm.as_ptr()) },
                DISP_E_TYPEMISMATCH
            );
            assert_eq!(
                handler.invocations.get(),
                0,
                "mismatch must reject before invoking"
            );
            assert_eq!(unsafe { oleaut::safe_array_destroy(wrong_psa) }, S_OK);

            // 2. Early-bound slot MutatePairs (ByRef SAFEARRAY(VT_RECORD)) + EchoPairs (retval).
            let mut psa =
                vba_record_array_to_automation_safearray(&arr_variant, Some(&pair_layout)).unwrap();
            let orig_psa = psa;
            let args_mut = [
                raw_this.expose_provenance() as u64,
                (&raw mut psa).expose_provenance() as u64,
            ];
            assert_eq!(
                unsafe { invoke_early_bound_slot_raw(7, args_mut.as_ptr(), xmm.as_ptr()) },
                S_OK
            );
            assert_eq!(psa, orig_psa, "same-shape copy-back preserves descriptor");

            let mut out_psa: *mut c_void = core::ptr::null_mut();
            let args_echo = [
                raw_this.expose_provenance() as u64,
                (&raw mut psa).expose_provenance() as u64,
                (&raw mut out_psa).expose_provenance() as u64,
            ];
            assert_eq!(
                unsafe { invoke_early_bound_slot_raw(8, args_echo.as_ptr(), xmm.as_ptr()) },
                S_OK
            );
            assert!(!out_psa.is_null());
            let mut info = core::ptr::null_mut();
            assert_eq!(
                unsafe { oleaut::safe_array_get_record_info(out_psa, &mut info) },
                S_OK
            );
            assert_eq!(unsafe { record_info_get_guid(info) }, Ok(pair_guid));
            assert_eq!(
                unsafe { record_info_get_name(info) }.as_deref(),
                Ok("SlotPair")
            );
            assert_eq!(
                unsafe { record_info_get_size(info) },
                Ok(pair_layout.size() as u32)
            );
            assert!(unsafe { layout_from_record_info(info) }.is_some());
            unsafe { release_record_info(info) };

            let recovered = unsafe {
                crate::automation_record::automation_safearray_to_vba_record_array(
                    out_psa,
                    &pair_layout,
                    false,
                    &|_, _| None,
                )
            }
            .unwrap();
            let elem = recovered.safearray_element(0).unwrap().unwrap();
            assert_eq!(
                elem.read_record_field_variant(0).unwrap().as_i32(),
                Some(105)
            );
            assert_eq!(
                elem.read_record_field_variant(1)
                    .unwrap()
                    .as_bstr()
                    .unwrap()
                    .as_str(),
                "hi!"
            );
            assert_eq!(unsafe { oleaut::safe_array_destroy(out_psa) }, S_OK);

            // 3. Late-bound IDispatch::Invoke with VT_BYREF | VT_ARRAY | VT_RECORD (`0x6024`).
            let mut byref_arg = VariantCore::from_pointer(0x4000 | 0x2000 | 0x0024, &raw mut psa);
            let mut dp = DispParams {
                rgvarg: &raw mut byref_arg,
                rgdispid_named_args: core::ptr::null_mut(),
                c_args: 1,
                c_named_args: 0,
            };
            let mut excep = ExcepInfo::default();
            let mut arg_err = 0u32;
            assert_eq!(
                unsafe {
                    (disp_vtbl.invoke)(
                        raw_this,
                        10,
                        &IID_NULL,
                        0,
                        DISPATCH_METHOD,
                        (&raw mut dp).cast(),
                        core::ptr::null_mut(),
                        (&raw mut excep).cast(),
                        &mut arg_err,
                    )
                },
                S_OK
            );
            let after_late = unsafe {
                crate::automation_record::automation_safearray_to_vba_record_array(
                    psa,
                    &pair_layout,
                    false,
                    &|_, _| None,
                )
            }
            .unwrap();
            assert_eq!(
                after_late
                    .safearray_element(0)
                    .unwrap()
                    .unwrap()
                    .read_record_field_variant(0)
                    .unwrap()
                    .as_i32(),
                Some(205)
            );
            assert_eq!(unsafe { oleaut::safe_array_destroy(psa) }, S_OK);

            // 4. Variant holding UDT array across BOTH early-bound slot (9) and late-bound Invoke (12).
            let mut var_cell = AutomationVariant::from_runtime(&arr_variant).unwrap();
            let mut var_ret = AutomationVariant::empty();
            let args_var = [
                raw_this.expose_provenance() as u64,
                var_cell.as_mut_ptr().expose_provenance() as u64,
                var_ret.as_mut_ptr().expose_provenance() as u64,
            ];
            assert_eq!(
                unsafe { invoke_early_bound_slot_raw(9, args_var.as_ptr(), xmm.as_ptr()) },
                S_OK
            );
            assert_eq!(var_cell.raw_vt(), 0x2024);
            assert_eq!(var_ret.raw_vt(), 0x2024);
            let back_var = var_ret.to_runtime(None, &|_, _| None).unwrap();
            assert_eq!(back_var.safearray_bounds_len().unwrap().1, 2);
            assert_eq!(
                back_var
                    .safearray_element(1)
                    .unwrap()
                    .unwrap()
                    .read_record_field_variant(0)
                    .unwrap()
                    .as_i32(),
                Some(77)
            );

            // Late-bound Invoke with VT_BYREF | VT_VARIANT (`0x400C`) holding `0x2024`.
            let mut var_arg_byref =
                VariantCore::from_pointer(0x4000 | 0x000C, var_cell.as_mut_ptr());
            let mut late_ret = AutomationVariant::empty();
            let mut dp2 = DispParams {
                rgvarg: &raw mut var_arg_byref,
                rgdispid_named_args: core::ptr::null_mut(),
                c_args: 1,
                c_named_args: 0,
            };
            assert_eq!(
                unsafe {
                    (disp_vtbl.invoke)(
                        raw_this,
                        12,
                        &IID_NULL,
                        0,
                        DISPATCH_METHOD,
                        (&raw mut dp2).cast(),
                        late_ret.as_mut_ptr(),
                        (&raw mut excep).cast(),
                        &mut arg_err,
                    )
                },
                S_OK
            );
            assert_eq!(late_ret.raw_vt(), 0x2024);
            let back_late_var = late_ret.to_runtime(None, &|_, _| None).unwrap();
            assert_eq!(back_late_var.safearray_bounds_len().unwrap().1, 2);
        }
        assert_eq!(
            before.balance_to(crate::current_thread_live_handle_counts()),
            Default::default()
        );
    }
}
