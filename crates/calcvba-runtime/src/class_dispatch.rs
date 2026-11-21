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
    E_NOTIMPL, E_POINTER, ExcepInfo, Guid, HResult, IDispatchVtbl, IID_NULL, IUnknownVtbl,
    RPC_E_WRONG_THREAD, S_OK,
};
use calcvba_automation::oleaut;
use calcvba_automation::safe_array::SafeArray;

use crate::events::{self, MemberFault};
use crate::object_ref::{
    COMPAT_OBJECT_VTBL, CompatObjectBase, ObjectRef, ObjectRefExt, RuntimeClassDescriptor,
    RuntimeMemberBacking, RuntimeMemberDescriptor, RuntimeMemberInvokeKind, RuntimeValueType,
    compat_owner_from_this,
};
use crate::{VarType, Variant, VariantCore};

mod slot_thunks;
mod type_info;

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
        .unwrap_or(7);

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

pub(super) unsafe fn resolve_ids_of_names_for_descriptor(
    descriptor: &'static RuntimeClassDescriptor,
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

    let Some(dispatch) = descriptor.dispatch_interface() else {
        return DISP_E_UNKNOWNNAME;
    };

    // SAFETY: `count >= 1` and `names` is non-null.
    let Some(member_name) = (unsafe { wide_ptr_to_string(*names) }) else {
        return DISP_E_UNKNOWNNAME;
    };
    let member_name = member_name.trim();
    let Some(primary_member) = dispatch
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

    let matching_members: Vec<&'static RuntimeMemberDescriptor> = dispatch
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
    // SAFETY: arguments forwarded from `IDispatch::GetIDsOfNames`.
    unsafe { resolve_ids_of_names_for_descriptor(descriptor, names, count, dispids) }
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
    descriptor: &'static RuntimeClassDescriptor,
    dispid: i32,
    flags: u16,
    c_args: usize,
) -> Option<(&'static RuntimeMemberDescriptor, RuntimeMemberInvokeKind)> {
    let dispatch = descriptor.dispatch_interface()?;
    let want_put_ref = flags & DISPATCH_PROPERTYPUTREF != 0;
    let want_put = flags & DISPATCH_PROPERTYPUT != 0;
    let want_get = flags & DISPATCH_PROPERTYGET != 0;
    let want_method = flags & DISPATCH_METHOD != 0;

    if want_put_ref
        && let Some(member) = dispatch.resolve_member_by_dispid(
            dispid,
            RuntimeMemberInvokeKind::PropertySet,
            c_args == 0,
            true,
        )
    {
        return Some((member, RuntimeMemberInvokeKind::PropertySet));
    }
    if want_put
        && let Some(member) = dispatch.resolve_member_by_dispid(
            dispid,
            RuntimeMemberInvokeKind::PropertyLet,
            c_args == 0,
            true,
        )
    {
        return Some((member, RuntimeMemberInvokeKind::PropertyLet));
    }
    if want_get && want_method {
        if let Some(member) =
            descriptor.member_by_dispid(dispid, RuntimeMemberInvokeKind::PropertyGet, true)
        {
            return Some((member, RuntimeMemberInvokeKind::PropertyGet));
        }
        if let Some(member) =
            descriptor.member_by_dispid(dispid, RuntimeMemberInvokeKind::Method, true)
        {
            return Some((member, RuntimeMemberInvokeKind::Method));
        }
        return None;
    }
    if want_get {
        return dispatch
            .resolve_member_by_dispid(
                dispid,
                RuntimeMemberInvokeKind::PropertyGet,
                c_args == 0,
                true,
            )
            .map(|member| (member, RuntimeMemberInvokeKind::PropertyGet));
    }
    if want_method {
        return dispatch
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
    member: &'static RuntimeMemberDescriptor,
    params: *const DispParams,
    slots: &[Option<usize>],
    args: &[Variant],
) -> Result<(), HResult> {
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

    let owner = compat_owner_from_this(this);
    // SAFETY: `this` is a live `CompatObjectBase` kept alive by the caller's COM reference.
    let owner_thread = unsafe { (*owner).owner_thread };
    if current_thread_id() != owner_thread {
        return RPC_E_WRONG_THREAD;
    }

    // SAFETY: same live `CompatObjectBase` on its owning thread.
    let descriptor = unsafe { (*owner).class_descriptor };
    let params_ptr = params.cast::<DispParams>();
    let c_args = if params_ptr.is_null() {
        0
    } else {
        // SAFETY: caller supplied a readable `DispParams`.
        unsafe { (*params_ptr).c_args as usize }
    };
    let Some((member, requested_kind)) =
        select_member_for_invoke(descriptor, dispid, flags, c_args)
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
                match unsafe { copy_dispatch_arg(cell) } {
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
            match unsafe { copy_dispatch_arg(cell) } {
                Ok(value) => tail.push(value),
                Err(hr) => return hr,
            }
        }
        args.push(Variant::from_safearray(SafeArray::from_variants(tail)));
    }

    // SAFETY: `params_ptr` was validated above.
    let aliases = unsafe { compute_byref_aliases(member, params_ptr, &mapped.slots, args.len()) };

    // SAFETY: `this` is a live `CompatObjectBase`; retain one reference for the invocation.
    let Some(object) = (unsafe { ObjectRef::from_raw_iunknown_addref(this.cast()) }) else {
        return E_POINTER;
    };

    let outcome = invoke_object_member(owner, object, member, requested_kind, &mut args, &aliases);

    // SAFETY: `params_ptr` and `mapped.slots` are valid for the duration of `Invoke`.
    let writeback = unsafe { write_back_byref_args(member, params_ptr, &mapped.slots, &args) };

    match outcome {
        Ok(ret_val) => {
            if let Err(hr) = writeback {
                return hr;
            }
            if !result.is_null() {
                // Transfer ownership of `ret_val` directly into caller's `pVarResult`.
                // SAFETY: `result` points to an initialized empty `VariantCore`; `Variant` is
                // `#[repr(transparent)]` over `VariantCore`, so writing moves ownership once.
                unsafe {
                    result.cast::<Variant>().write(ret_val);
                }
            }
            S_OK
        }
        Err(fault) => {
            let _ = writeback;
            // SAFETY: `excep_info` is null or a caller-owned writable `ExcepInfo`.
            unsafe {
                write_excep_info(excep_info.cast::<ExcepInfo>(), &fault);
            }
            DISP_E_EXCEPTION
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
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
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

        // A VBA error comes back as its HRESULT (0x800A0000 | number).
        let fail: NoArgs = unsafe {
            core::mem::transmute(entry(slot_of("Fail", RuntimeMemberInvokeKind::Method)))
        };
        assert_eq!(unsafe { fail(raw) }, 0x800A_1388u32 as i32);
    }
}
