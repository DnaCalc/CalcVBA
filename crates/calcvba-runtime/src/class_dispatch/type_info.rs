//! Portable COM `ITypeInfo` for VBA project classes built from `&'static RuntimeClassDescriptor`
//! (dc-yqq0.29).
//!
//! Matches Excel VBA's measured `IDispatch::GetTypeInfo(0)` layout
//! (`docs/design/vba-class-dispids-excel-2026-11-21.md`):
//! - Primary half: `TKIND_DISPATCH`, `wTypeFlags = 0x1C2` (`FCANCREATE | FDUAL | FNONEXTENSIBLE |
//!   FOLEAUTOMATION`), named after the class, carrying the class's `default_interface_iid`,
//!   `cFuncs` public members, `cVars = 0`, `cImplTypes = 1`, and full `cbSizeVft` (including
//!   hidden `Friend`/`Private` slots).
//! - Interface half (reached via `GetRefTypeOfImplType(-1)` -> `GetRefTypeInfo(href)`):
//!   `TKIND_INTERFACE`, same name, GUID, flags `0x1C2`, `cFuncs`, `cVars = 0`, `cImplTypes = 1`,
//!   `cbSizeVft`, with every function returning `VT_HRESULT` and appending `[out, retval]` for
//!   functions and property getters.
//! - `GetRefTypeOfImplType(0)` and `GetImplTypeFlags(0)` return `TYPE_E_ELEMENTNOTFOUND`, as Excel
//!   does.

use core::ffi::c_void;
use core::mem::{align_of, size_of};
use core::sync::atomic::{AtomicU32, Ordering};
use std::collections::HashMap;
use std::sync::{Mutex, OnceLock};

use calcvba_automation::bstr::BStr;
use calcvba_automation::com::{
    CC_STDCALL, DispParams, E_INVALIDARG, E_NOINTERFACE, E_NOTIMPL, E_POINTER, ElemDesc,
    ElemDescPayload, ExcepInfo, FUNC_DISPATCH, FUNC_PUREVIRTUAL, FuncDesc, Guid, HResult,
    IDispatchVtbl, IID_ITYPEINFO, IID_IUNKNOWN, IID_NULL, INVOKE_FUNC, INVOKE_PROPERTYGET,
    INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF, ITypeInfoVtbl, IUnknownVtbl, IdlDesc, MEMBERID_NIL,
    PARAMFLAG_FIN, PARAMFLAG_FOPT, PARAMFLAG_FOUT, PARAMFLAG_FRETVAL, ParamDesc, S_OK,
    TKIND_DISPATCH, TKIND_INTERFACE, TYPE_E_BADMODULEKIND, TYPE_E_ELEMENTNOTFOUND, TypeAttr,
    TypeDesc, TypeDescPayload, VBA_CLASS_TYPEFLAGS, VT_HRESULT, VT_PTR, VT_SAFEARRAY, VT_VOID,
};

use crate::object_ref::{
    RawRuntimeIUnknown, RuntimeClassDescriptor, RuntimeMemberDescriptor, RuntimeMemberInvokeKind,
    RuntimeValueType,
};

use super::{class_dual_vtable_slot_count, resolve_ids_of_names_for_descriptor};

const HREF_INTERFACE_HALF: u32 = 1;
const HREF_DISPATCH_HALF: u32 = 2;

struct ClassTypeInfoHalf {
    descriptor: &'static RuntimeClassDescriptor,
    type_attr: TypeAttr,
    funcs: &'static [FuncDesc],
    func_members: &'static [&'static RuntimeMemberDescriptor],
    /// Keeps every leaked `ElemDesc` slice and nested `TypeDesc` reachable from the static root
    /// (zero leaks under Miri).
    _owned_elem_slices: &'static [&'static [ElemDesc]],
    _owned_type_descs: &'static [&'static TypeDesc],
}

struct ClassTypeInfoPair {
    dispatch: ClassTypeInfoHalf,
    interface: ClassTypeInfoHalf,
}

#[derive(Clone, Copy)]
struct SyncTypeInfoPair(&'static ClassTypeInfoPair);

// SAFETY: `ClassTypeInfoPair` holds only immutable `'static` COM descriptor tables built once and
// never mutated after publication.
unsafe impl Send for SyncTypeInfoPair {}
// SAFETY: same immutable `'static` COM descriptor tables.
unsafe impl Sync for SyncTypeInfoPair {}

static CLASS_TYPE_INFOS: OnceLock<Mutex<HashMap<usize, SyncTypeInfoPair>>> = OnceLock::new();

fn type_info_cache() -> &'static Mutex<HashMap<usize, SyncTypeInfoPair>> {
    CLASS_TYPE_INFOS.get_or_init(|| Mutex::new(HashMap::new()))
}

fn type_info_pair_for_descriptor(
    descriptor: &'static RuntimeClassDescriptor,
) -> &'static ClassTypeInfoPair {
    let key = descriptor as *const RuntimeClassDescriptor as usize;
    let mut cache = type_info_cache()
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    if let Some(existing) = cache.get(&key).copied() {
        return existing.0;
    }
    let pair: &'static ClassTypeInfoPair = Box::leak(Box::new(ClassTypeInfoPair {
        dispatch: build_half(descriptor, false),
        interface: build_half(descriptor, true),
    }));
    cache.insert(key, SyncTypeInfoPair(pair));
    pair
}

fn runtime_value_type_to_vt(ty: RuntimeValueType) -> u16 {
    match ty {
        RuntimeValueType::Integer => 0x0002,
        RuntimeValueType::Long => 0x0003,
        RuntimeValueType::Single => 0x0004,
        RuntimeValueType::Double => 0x0005,
        RuntimeValueType::Currency => 0x0006,
        RuntimeValueType::Date => 0x0007,
        RuntimeValueType::String => 0x0008,
        RuntimeValueType::Object => 0x0009,
        RuntimeValueType::Boolean => 0x000B,
        RuntimeValueType::Variant | RuntimeValueType::Record => 0x000C,
        RuntimeValueType::Decimal => 0x000E,
        RuntimeValueType::Byte => 0x0011,
        RuntimeValueType::LongLong => 0x0014,
        RuntimeValueType::LongPtr => {
            #[cfg(target_pointer_width = "64")]
            {
                0x0014
            }
            #[cfg(not(target_pointer_width = "64"))]
            {
                0x0003
            }
        }
        RuntimeValueType::Array => VT_SAFEARRAY,
    }
}

fn alloc_static_tdesc(
    tdesc: TypeDesc,
    owned_type_descs: &mut Vec<&'static TypeDesc>,
) -> *mut TypeDesc {
    let leaked: &'static TypeDesc = Box::leak(Box::new(tdesc));
    owned_type_descs.push(leaked);
    core::ptr::from_ref(leaked).cast_mut()
}

fn plain_tdesc(vt: u16) -> TypeDesc {
    TypeDesc {
        payload: TypeDescPayload {
            lptdesc: core::ptr::null_mut(),
        },
        vt,
    }
}

fn ptr_tdesc(inner: TypeDesc, owned_type_descs: &mut Vec<&'static TypeDesc>) -> TypeDesc {
    let inner_ptr = alloc_static_tdesc(inner, owned_type_descs);
    TypeDesc {
        payload: TypeDescPayload { lptdesc: inner_ptr },
        vt: VT_PTR,
    }
}

fn safearray_tdesc(elem_vt: u16, owned_type_descs: &mut Vec<&'static TypeDesc>) -> TypeDesc {
    let elem_ptr = alloc_static_tdesc(plain_tdesc(elem_vt), owned_type_descs);
    TypeDesc {
        payload: TypeDescPayload { lptdesc: elem_ptr },
        vt: VT_SAFEARRAY,
    }
}

fn value_tdesc(ty: RuntimeValueType, owned_type_descs: &mut Vec<&'static TypeDesc>) -> TypeDesc {
    if ty == RuntimeValueType::Array {
        safearray_tdesc(0x000C, owned_type_descs)
    } else {
        plain_tdesc(runtime_value_type_to_vt(ty))
    }
}

fn param_elem(tdesc: TypeDesc, w_param_flags: u16) -> ElemDesc {
    ElemDesc {
        tdesc,
        payload: ElemDescPayload {
            paramdesc: ParamDesc {
                pparamdescex: core::ptr::null_mut(),
                w_param_flags,
            },
        },
    }
}

fn invkind_for_member(kind: RuntimeMemberInvokeKind) -> i32 {
    match kind {
        RuntimeMemberInvokeKind::Method => INVOKE_FUNC,
        RuntimeMemberInvokeKind::PropertyGet => INVOKE_PROPERTYGET,
        RuntimeMemberInvokeKind::PropertyLet => INVOKE_PROPERTYPUT,
        RuntimeMemberInvokeKind::PropertySet => INVOKE_PROPERTYPUTREF,
    }
}

fn build_half(
    descriptor: &'static RuntimeClassDescriptor,
    is_interface_half: bool,
) -> ClassTypeInfoHalf {
    let public_members: Vec<&'static RuntimeMemberDescriptor> = descriptor
        .dispatch_interface()
        .map(|dispatch| {
            dispatch
                .members
                .iter()
                .filter(|member| member.externally_visible)
                .collect()
        })
        .unwrap_or_default();

    let mut funcs = Vec::with_capacity(public_members.len());
    let mut owned_elem_slices: Vec<&'static [ElemDesc]> = Vec::new();
    let mut owned_type_descs: Vec<&'static TypeDesc> = Vec::new();

    for &member in &public_members {
        let returns_value = matches!(
            member.invoke_kind,
            RuntimeMemberInvokeKind::Method | RuntimeMemberInvokeKind::PropertyGet
        ) && member.return_type.is_some();
        let has_param_array = member.params.last().is_some_and(|p| p.param_array);

        let mut params = Vec::with_capacity(
            member.params.len() + usize::from(is_interface_half && returns_value),
        );
        for param in member.params {
            let (tdesc, mut flags) = if param.param_array {
                let sa = safearray_tdesc(0x000C, &mut owned_type_descs);
                (
                    ptr_tdesc(sa, &mut owned_type_descs),
                    PARAMFLAG_FIN | PARAMFLAG_FOUT,
                )
            } else if param.by_ref {
                let inner = value_tdesc(param.value_type, &mut owned_type_descs);
                (
                    ptr_tdesc(inner, &mut owned_type_descs),
                    PARAMFLAG_FIN | PARAMFLAG_FOUT,
                )
            } else {
                (
                    value_tdesc(param.value_type, &mut owned_type_descs),
                    PARAMFLAG_FIN,
                )
            };
            if param.optional {
                flags |= PARAMFLAG_FOPT;
            }
            params.push(param_elem(tdesc, flags));
        }

        if is_interface_half && returns_value {
            let ret_ty = member.return_type.unwrap_or(RuntimeValueType::Variant);
            let inner = value_tdesc(ret_ty, &mut owned_type_descs);
            let tdesc = ptr_tdesc(inner, &mut owned_type_descs);
            params.push(param_elem(tdesc, PARAMFLAG_FOUT | PARAMFLAG_FRETVAL));
        }

        let c_params = params.len() as i16;
        let lprgelemdesc_param = if params.is_empty() {
            core::ptr::null_mut()
        } else {
            let leaked: &'static [ElemDesc] = Box::leak(params.into_boxed_slice());
            owned_elem_slices.push(leaked);
            leaked.as_ptr().cast_mut()
        };

        let ret_tdesc = if is_interface_half {
            plain_tdesc(VT_HRESULT)
        } else if returns_value {
            value_tdesc(
                member.return_type.unwrap_or(RuntimeValueType::Variant),
                &mut owned_type_descs,
            )
        } else {
            plain_tdesc(VT_VOID)
        };

        let slot = usize::from(member.vtable_slot.unwrap_or(7));
        let o_vft = (slot * size_of::<*const c_void>()) as i16;

        funcs.push(FuncDesc {
            memid: member.dispatch_id,
            lprgscode: core::ptr::null_mut(),
            lprgelemdesc_param,
            funckind: if is_interface_half {
                FUNC_PUREVIRTUAL
            } else {
                FUNC_DISPATCH
            },
            invkind: invkind_for_member(member.invoke_kind),
            callconv: CC_STDCALL,
            c_params,
            // Excel VBA leaves cParamsOpt at 0 for Optional parameters and sets -1 for ParamArray.
            c_params_opt: if has_param_array { -1 } else { 0 },
            o_vft,
            c_scodes: 0,
            elemdesc_func: ElemDesc {
                tdesc: ret_tdesc,
                payload: ElemDescPayload {
                    idldesc: IdlDesc::default(),
                },
            },
            w_func_flags: member.func_flags,
        });
    }

    let total_vtable_slots = class_dual_vtable_slot_count(descriptor);
    let cb_size_vft = (total_vtable_slots * size_of::<*const c_void>()) as u16;
    let type_attr = TypeAttr {
        guid: descriptor.default_interface_iid.unwrap_or(IID_NULL),
        lcid: 0,
        dw_reserved: 0,
        memid_constructor: MEMBERID_NIL,
        memid_destructor: MEMBERID_NIL,
        lpstr_schema: core::ptr::null_mut(),
        cb_size_instance: size_of::<*const c_void>() as u32,
        typekind: if is_interface_half {
            TKIND_INTERFACE
        } else {
            TKIND_DISPATCH
        },
        c_funcs: funcs.len() as u16,
        c_vars: 0,
        c_impl_types: 1,
        cb_size_vft,
        cb_alignment: align_of::<*const c_void>() as u16,
        w_type_flags: VBA_CLASS_TYPEFLAGS,
        w_major_ver_num: 0,
        w_minor_ver_num: 0,
        tdesc_alias: TypeDesc::default(),
        idldesc_type: IdlDesc::default(),
    };

    ClassTypeInfoHalf {
        descriptor,
        type_attr,
        funcs: Box::leak(funcs.into_boxed_slice()),
        func_members: Box::leak(public_members.into_boxed_slice()),
        _owned_elem_slices: Box::leak(owned_elem_slices.into_boxed_slice()),
        _owned_type_descs: Box::leak(owned_type_descs.into_boxed_slice()),
    }
}

#[repr(C)]
struct ClassTypeInfoObject {
    vtbl: *const ITypeInfoVtbl,
    ref_count: AtomicU32,
    pair: &'static ClassTypeInfoPair,
    is_interface_half: bool,
}

impl ClassTypeInfoObject {
    fn half(&self) -> &'static ClassTypeInfoHalf {
        if self.is_interface_half {
            &self.pair.interface
        } else {
            &self.pair.dispatch
        }
    }
}

static CLASS_TYPE_INFO_VTBL: ITypeInfoVtbl = ITypeInfoVtbl {
    unknown: IUnknownVtbl {
        query_interface: type_info_query_interface,
        add_ref: type_info_add_ref,
        release: type_info_release,
    },
    get_type_attr: type_info_get_type_attr,
    get_type_comp: type_info_get_type_comp,
    get_func_desc: type_info_get_func_desc,
    get_var_desc: type_info_get_var_desc,
    get_names: type_info_get_names,
    get_ref_type_of_impl_type: type_info_get_ref_type_of_impl_type,
    get_impl_type_flags: type_info_get_impl_type_flags,
    get_ids_of_names: type_info_get_ids_of_names,
    invoke: type_info_invoke,
    get_documentation: type_info_get_documentation,
    get_dll_entry: type_info_get_dll_entry,
    get_ref_type_info: type_info_get_ref_type_info,
    address_of_member: type_info_address_of_member,
    create_instance: type_info_create_instance,
    get_mops: type_info_get_mops,
    get_containing_type_lib: type_info_get_containing_type_lib,
    release_type_attr: type_info_release_type_attr,
    release_func_desc: type_info_release_func_desc,
    release_var_desc: type_info_release_var_desc,
};

fn allocate_type_info(pair: &'static ClassTypeInfoPair, is_interface_half: bool) -> *mut c_void {
    Box::into_raw(Box::new(ClassTypeInfoObject {
        vtbl: &CLASS_TYPE_INFO_VTBL,
        ref_count: AtomicU32::new(1),
        pair,
        is_interface_half,
    }))
    .cast()
}

/// Create a retained `ITypeInfo*` (dispatch half, `TKIND_DISPATCH`, flags `0x1C2`) for `descriptor`.
pub(super) fn create_class_type_info(descriptor: &'static RuntimeClassDescriptor) -> *mut c_void {
    let pair = type_info_pair_for_descriptor(descriptor);
    allocate_type_info(pair, false)
}

unsafe extern "system" fn type_info_query_interface(
    this: *mut c_void,
    riid: *const Guid,
    ppv: *mut *mut c_void,
) -> HResult {
    if ppv.is_null() {
        return E_POINTER;
    }
    // SAFETY: per COM `QueryInterface`, `ppv` is a writable out-slot.
    unsafe {
        *ppv = core::ptr::null_mut();
    }
    if this.is_null() || riid.is_null() {
        return E_POINTER;
    }
    // SAFETY: `riid` is a readable `Guid`.
    let iid = unsafe { riid.read() };
    if iid == IID_IUNKNOWN || iid == IID_ITYPEINFO {
        // SAFETY: `ppv` is non-null and writable; `this` is a live `ClassTypeInfoObject`.
        unsafe {
            *ppv = this;
            type_info_add_ref(this);
        }
        S_OK
    } else {
        E_NOINTERFACE
    }
}

unsafe extern "system" fn type_info_add_ref(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    let obj = this.cast::<ClassTypeInfoObject>();
    // SAFETY: `this` is a live `ClassTypeInfoObject` retained by the caller.
    unsafe { (*obj).ref_count.fetch_add(1, Ordering::AcqRel) + 1 }
}

unsafe extern "system" fn type_info_release(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    let obj = this.cast::<ClassTypeInfoObject>();
    // SAFETY: `this` is a live `ClassTypeInfoObject` releasing one reference; on 0 this caller
    // holds sole ownership of the `Box`.
    let prev = unsafe { (*obj).ref_count.fetch_sub(1, Ordering::AcqRel) };
    let remaining = prev.saturating_sub(1);
    if remaining == 0 {
        // SAFETY: sole final ownership of the allocation created by `allocate_type_info`.
        unsafe {
            drop(Box::from_raw(obj));
        }
    }
    remaining
}

unsafe extern "system" fn type_info_get_type_attr(
    this: *mut c_void,
    pp_type_attr: *mut *mut TypeAttr,
) -> HResult {
    if this.is_null() || pp_type_attr.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `ClassTypeInfoObject`; `pp_type_attr` is a writable out-slot.
    unsafe {
        let half = (*this.cast::<ClassTypeInfoObject>()).half();
        *pp_type_attr = core::ptr::from_ref(&half.type_attr).cast_mut();
    }
    S_OK
}

unsafe extern "system" fn type_info_get_type_comp(
    _this: *mut c_void,
    pp_tcomp: *mut *mut c_void,
) -> HResult {
    if !pp_tcomp.is_null() {
        // SAFETY: writable out-slot.
        unsafe {
            *pp_tcomp = core::ptr::null_mut();
        }
    }
    E_NOTIMPL
}

unsafe extern "system" fn type_info_get_func_desc(
    this: *mut c_void,
    index: u32,
    pp_func_desc: *mut *mut FuncDesc,
) -> HResult {
    if this.is_null() || pp_func_desc.is_null() {
        return E_POINTER;
    }
    // SAFETY: `pp_func_desc` is a writable out-slot; `this` is a live `ClassTypeInfoObject`.
    unsafe {
        *pp_func_desc = core::ptr::null_mut();
        let half = (*this.cast::<ClassTypeInfoObject>()).half();
        let Some(fd) = half.funcs.get(index as usize) else {
            return TYPE_E_ELEMENTNOTFOUND;
        };
        *pp_func_desc = core::ptr::from_ref(fd).cast_mut();
    }
    S_OK
}

unsafe extern "system" fn type_info_get_var_desc(
    _this: *mut c_void,
    _index: u32,
    pp_var_desc: *mut *mut c_void,
) -> HResult {
    if !pp_var_desc.is_null() {
        // SAFETY: writable out-slot.
        unsafe {
            *pp_var_desc = core::ptr::null_mut();
        }
    }
    TYPE_E_ELEMENTNOTFOUND
}

unsafe extern "system" fn type_info_get_names(
    this: *mut c_void,
    memid: i32,
    rg_bstr_names: *mut *mut u16,
    c_max_names: u32,
    pc_names: *mut u32,
) -> HResult {
    if this.is_null() || pc_names.is_null() || (c_max_names > 0 && rg_bstr_names.is_null()) {
        return E_POINTER;
    }
    // SAFETY: `pc_names` is a writable out-slot.
    unsafe {
        *pc_names = 0;
    }
    // SAFETY: `this` is a live `ClassTypeInfoObject`.
    let half = unsafe { (*this.cast::<ClassTypeInfoObject>()).half() };
    // Prefer a Method or PropertyGet when naming a member's parameters so Property Let/Set's
    // unnamed trailing RHS value parameter is never reported as a named parameter.
    let primary = half
        .func_members
        .iter()
        .copied()
        .find(|m| {
            m.dispatch_id == memid
                && matches!(
                    m.invoke_kind,
                    RuntimeMemberInvokeKind::Method | RuntimeMemberInvokeKind::PropertyGet
                )
        })
        .or_else(|| {
            half.func_members
                .iter()
                .copied()
                .find(|m| m.dispatch_id == memid)
        });
    let Some(member) = primary else {
        return TYPE_E_ELEMENTNOTFOUND;
    };

    let param_name_count = match member.invoke_kind {
        RuntimeMemberInvokeKind::PropertyLet | RuntimeMemberInvokeKind::PropertySet => {
            member.params.len().saturating_sub(1)
        }
        RuntimeMemberInvokeKind::Method | RuntimeMemberInvokeKind::PropertyGet => {
            member.params.len()
        }
    };
    let max_out = c_max_names as usize;
    let mut written = 0usize;
    let names_iter = core::iter::once(member.name).chain(
        member.params[..param_name_count]
            .iter()
            .map(|param| param.name),
    );
    for name in names_iter.take(max_out) {
        let raw = BStr::from(name).into_raw_bstr();
        // SAFETY: `written < c_max_names` and `rg_bstr_names` was verified non-null above.
        unsafe {
            *rg_bstr_names.add(written) = raw;
        }
        written += 1;
    }
    // SAFETY: `pc_names` is a writable out-slot.
    unsafe {
        *pc_names = written as u32;
    }
    S_OK
}

unsafe extern "system" fn type_info_get_ref_type_of_impl_type(
    this: *mut c_void,
    index: u32,
    p_ref_type: *mut u32,
) -> HResult {
    if this.is_null() || p_ref_type.is_null() {
        return E_POINTER;
    }
    // Dual interface toggle: index == -1 (u32::MAX) yields the opposite half. Index 0 fails with
    // TYPE_E_ELEMENTNOTFOUND, matching Excel VBA's measured ITypeInfo.
    if index == u32::MAX {
        // SAFETY: `this` is a live `ClassTypeInfoObject`; `p_ref_type` is writable.
        unsafe {
            let is_interface = (*this.cast::<ClassTypeInfoObject>()).is_interface_half;
            *p_ref_type = if is_interface {
                HREF_DISPATCH_HALF
            } else {
                HREF_INTERFACE_HALF
            };
        }
        S_OK
    } else {
        TYPE_E_ELEMENTNOTFOUND
    }
}

unsafe extern "system" fn type_info_get_impl_type_flags(
    this: *mut c_void,
    _index: u32,
    p_impl_type_flags: *mut i32,
) -> HResult {
    if this.is_null() || p_impl_type_flags.is_null() {
        return E_POINTER;
    }
    // Matches Excel VBA: GetImplTypeFlags(0) fails with TYPE_E_ELEMENTNOTFOUND.
    TYPE_E_ELEMENTNOTFOUND
}

unsafe extern "system" fn type_info_get_ids_of_names(
    this: *mut c_void,
    rgsz_names: *const *const u16,
    c_names: u32,
    p_mem_id: *mut i32,
) -> HResult {
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `ClassTypeInfoObject`; arguments follow `ITypeInfo::GetIDsOfNames`.
    unsafe {
        let half = (*this.cast::<ClassTypeInfoObject>()).half();
        resolve_ids_of_names_for_descriptor(half.descriptor, rgsz_names, c_names, p_mem_id)
    }
}

#[allow(clippy::too_many_arguments)]
unsafe extern "system" fn type_info_invoke(
    this: *mut c_void,
    pv_instance: *mut c_void,
    memid: i32,
    w_flags: u16,
    p_disp_params: *mut DispParams,
    p_var_result: *mut c_void,
    p_excep_info: *mut ExcepInfo,
    pu_arg_err: *mut u32,
) -> HResult {
    if this.is_null() {
        return E_POINTER;
    }
    if pv_instance.is_null() {
        return E_INVALIDARG;
    }
    // SAFETY: `pv_instance` is a live COM `IDispatch` object supplied by the caller.
    unsafe {
        let disp = pv_instance.cast::<RawRuntimeIUnknown>();
        let vtbl = &*((*disp).vtbl.cast::<IDispatchVtbl>());
        (vtbl.invoke)(
            pv_instance,
            memid,
            &IID_NULL,
            0,
            w_flags,
            p_disp_params.cast(),
            p_var_result,
            p_excep_info.cast(),
            pu_arg_err,
        )
    }
}

unsafe extern "system" fn type_info_get_documentation(
    this: *mut c_void,
    memid: i32,
    p_bstr_name: *mut *mut u16,
    p_bstr_doc_string: *mut *mut u16,
    pdw_help_context: *mut u32,
    p_bstr_help_file: *mut *mut u16,
) -> HResult {
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: each optional out-slot, when non-null, is writable.
    unsafe {
        if !p_bstr_name.is_null() {
            *p_bstr_name = core::ptr::null_mut();
        }
        if !p_bstr_doc_string.is_null() {
            *p_bstr_doc_string = core::ptr::null_mut();
        }
        if !pdw_help_context.is_null() {
            *pdw_help_context = 0;
        }
        if !p_bstr_help_file.is_null() {
            *p_bstr_help_file = core::ptr::null_mut();
        }
    }
    // SAFETY: `this` is a live `ClassTypeInfoObject`.
    let half = unsafe { (*this.cast::<ClassTypeInfoObject>()).half() };
    let name = if memid == MEMBERID_NIL {
        Some(half.descriptor.name)
    } else {
        half.func_members
            .iter()
            .copied()
            .find(|m| m.dispatch_id == memid)
            .map(|m| m.name)
    };
    let Some(name) = name else {
        return TYPE_E_ELEMENTNOTFOUND;
    };
    if !p_bstr_name.is_null() {
        let raw = BStr::from(name).into_raw_bstr();
        // SAFETY: `p_bstr_name` is non-null and writable.
        unsafe {
            *p_bstr_name = raw;
        }
    }
    S_OK
}

unsafe extern "system" fn type_info_get_dll_entry(
    _this: *mut c_void,
    _memid: i32,
    _invkind: i32,
    p_bstr_dll_name: *mut *mut u16,
    p_bstr_name: *mut *mut u16,
    pw_ordinal: *mut u16,
) -> HResult {
    // SAFETY: clear any non-null caller out-slots.
    unsafe {
        if !p_bstr_dll_name.is_null() {
            *p_bstr_dll_name = core::ptr::null_mut();
        }
        if !p_bstr_name.is_null() {
            *p_bstr_name = core::ptr::null_mut();
        }
        if !pw_ordinal.is_null() {
            *pw_ordinal = 0;
        }
    }
    TYPE_E_BADMODULEKIND
}

unsafe extern "system" fn type_info_get_ref_type_info(
    this: *mut c_void,
    h_ref_type: u32,
    pp_tinfo: *mut *mut c_void,
) -> HResult {
    if this.is_null() || pp_tinfo.is_null() {
        return E_POINTER;
    }
    // SAFETY: `pp_tinfo` is writable; `this` is a live `ClassTypeInfoObject`.
    unsafe {
        *pp_tinfo = core::ptr::null_mut();
        let pair = (*this.cast::<ClassTypeInfoObject>()).pair;
        match h_ref_type {
            HREF_INTERFACE_HALF => {
                *pp_tinfo = allocate_type_info(pair, true);
                S_OK
            }
            HREF_DISPATCH_HALF => {
                *pp_tinfo = allocate_type_info(pair, false);
                S_OK
            }
            _ => TYPE_E_ELEMENTNOTFOUND,
        }
    }
}

unsafe extern "system" fn type_info_address_of_member(
    _this: *mut c_void,
    _memid: i32,
    _invkind: i32,
    ppv: *mut *mut c_void,
) -> HResult {
    if !ppv.is_null() {
        // SAFETY: writable out-slot.
        unsafe {
            *ppv = core::ptr::null_mut();
        }
    }
    E_NOTIMPL
}

unsafe extern "system" fn type_info_create_instance(
    _this: *mut c_void,
    _outer: *mut c_void,
    _riid: *const Guid,
    ppv_obj: *mut *mut c_void,
) -> HResult {
    if !ppv_obj.is_null() {
        // SAFETY: writable out-slot.
        unsafe {
            *ppv_obj = core::ptr::null_mut();
        }
    }
    E_NOTIMPL
}

unsafe extern "system" fn type_info_get_mops(
    _this: *mut c_void,
    _memid: i32,
    p_bstr_mops: *mut *mut u16,
) -> HResult {
    if !p_bstr_mops.is_null() {
        // SAFETY: writable out-slot.
        unsafe {
            *p_bstr_mops = core::ptr::null_mut();
        }
    }
    S_OK
}

unsafe extern "system" fn type_info_get_containing_type_lib(
    _this: *mut c_void,
    pp_tlib: *mut *mut c_void,
    p_index: *mut u32,
) -> HResult {
    // SAFETY: clear any non-null caller out-slots.
    unsafe {
        if !pp_tlib.is_null() {
            *pp_tlib = core::ptr::null_mut();
        }
        if !p_index.is_null() {
            *p_index = 0;
        }
    }
    E_NOTIMPL
}

unsafe extern "system" fn type_info_release_type_attr(
    _this: *mut c_void,
    _p_type_attr: *mut TypeAttr,
) {
}

unsafe extern "system" fn type_info_release_func_desc(
    _this: *mut c_void,
    _p_func_desc: *mut FuncDesc,
) {
}

unsafe extern "system" fn type_info_release_var_desc(_this: *mut c_void, _p_var_desc: *mut c_void) {
}
