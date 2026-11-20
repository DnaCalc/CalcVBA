//! Early-bound calls over the runtime cell (design doc, step 8b): the compiler bound the member to
//! a `FUNCDESC` of the declared interface, so the call is built from that description, as VBA
//! builds it, and goes through the interface's vtable.
//!
//! Each argument cell is shaped by its declared parameter:
//! - A ByVal argument of a typed parameter is the value converted to that type (VBA's argument
//!   coercion); a String already is the BSTR, borrowed. A ByVal `VARIANT` is the caller's cell,
//!   borrowed.
//! - A ByRef argument whose reference already has the parameter's type is passed as that pointer,
//!   so the callee writes the caller's storage. Otherwise it gets a temporary of the parameter's
//!   type, and a reference is written back from it after the call.
//! - An omitted optional is `VT_ERROR` / `DISP_E_PARAMNOTFOUND` for a `VARIANT`, or the typelib
//!   default (`defaultvalue`); a hidden `[lcid]` is the call's locale; a ParamArray tail is a
//!   zero-based `SAFEARRAY` of the remaining values.
//! - The `[out, retval]` pointer is the result cell's payload (the cell itself for a `VARIANT`),
//!   and the result moves into the runtime.
//!
//! `oleaut32!DispCallFunc` performs the call, so every `VARTYPE` takes the platform's calling
//! convention (x64 and x86). A member the vtable cannot carry (a pure dispinterface, a proxy whose
//! vtable is not the type library's, metadata without the optional rules) is called through
//! `IDispatch::Invoke` with its DISPID, as VBA calls dispinterface members.
use super::{Dispatch, LOCALE_USER_DEFAULT, ProjectDispatch, adopt, fault_from_hresult};
use crate::call::{
    ComArg as Arg, ComArgValue as ArgValue, ComCallKind as CallKind, ComFault as Fault,
    ComMember as Member,
};
use crate::typelib::{
    OptionalParamDefault, SourceTypeKind, TypeLibMemberInvokeKind, TypeLibMemberMetadata,
    TypeLibParamType, TypeLibWireType,
};
use crate::windows_client::{IID_IDISPATCH, OwnedComInterface, query_interface_pointer};
use calcvba_runtime::ObjectRefExt;
use calcvba_runtime::{ObjectRef, VarType, Variant, WIRE_SIZE};
use std::cell::RefCell;
use std::collections::HashMap;
use windows_sys::Win32::System::Variant::VARIANT;

const VT_I2: u16 = 0x0002;
const VT_I4: u16 = 0x0003;
const VT_R4: u16 = 0x0004;
const VT_R8: u16 = 0x0005;
const VT_CY: u16 = 0x0006;
const VT_DATE: u16 = 0x0007;
const VT_BSTR: u16 = 0x0008;
const VT_DISPATCH: u16 = 0x0009;
const VT_ERROR: u16 = 0x000A;
const VT_BOOL: u16 = 0x000B;
const VT_VARIANT: u16 = 0x000C;
const VT_UNKNOWN: u16 = 0x000D;
const VT_DECIMAL: u16 = 0x000E;
const VT_UI1: u16 = 0x0011;
const VT_I8: u16 = 0x0014;
const VT_RECORD: u16 = 0x0024;
const VT_ARRAY: u16 = 0x2000;
const VT_BYREF: u16 = 0x4000;
const CC_STDCALL: u32 = 4;
const DISP_E_PARAMNOTFOUND: i32 = 0x8002_0004u32 as i32;
const E_NOINTERFACE: i32 = 0x8000_4002u32 as i32;

/// `IProxyManager`: every marshaling proxy answers it, a direct object does not.
const IID_IPROXYMANAGER: windows_sys::core::GUID = windows_sys::core::GUID {
    data1: 0x0000_0008,
    data2: 0x0000,
    data3: 0x0000,
    data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};

#[link(name = "oleaut32")]
unsafe extern "system" {
    fn DispCallFunc(
        instance: *mut core::ffi::c_void,
        offset: usize,
        callconv: u32,
        vt_return: u16,
        count: u32,
        types: *const u16,
        args: *const *mut VARIANT,
        result: *mut VARIANT,
    ) -> i32;
}

/// VBA's argument coercion of a value to a parameter's `VARTYPE` (scalar and String types), with
/// its errors (13 Type mismatch, 6 Overflow, 94 Invalid use of Null). The language layer supplies
/// it, so this module keeps no conversion rules of its own.
pub type Coerce<'a> = &'a dyn Fn(&Variant, u16) -> Result<Variant, Fault>;

/// How an early-bound call reached the object.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EarlyTransport {
    /// The interface's vtable slot.
    Vtable,
    /// `IDispatch::Invoke` with the member's DISPID.
    Dispatch,
}

/// Calls the compiler-bound `member` of `object` with `args` in source order (named arguments
/// name the member's parameters).
///
/// # Safety
/// Every `Reference` argument is a `VT_BYREF` cell whose target is live, writable storage of its
/// type for the whole call, and no Rust reference to that storage is held across it.
pub unsafe fn invoke(
    object: &ObjectRef,
    member: &TypeLibMemberMetadata,
    args: &[Arg<'_>],
    project: ProjectDispatch<'_>,
    coerce: Coerce<'_>,
) -> Result<(Variant, EarlyTransport), Fault> {
    let positional = arrange(member, args)?;
    if let Some(interface) = vtable_interface(object, member)? {
        let context = Context { project, coerce };
        // SAFETY: forwarded from the caller; `interface` is the declared interface, retained.
        let value = unsafe { call_vtable(&interface, member, &positional, &context) }?;
        return Ok((value, EarlyTransport::Vtable));
    }
    // The dispinterface shape: the visible parameters, then a ParamArray's values, trailing
    // omissions dropped.
    let mut dispatch_args: Vec<Arg<'_>> = positional
        .iter()
        .zip(rules(member))
        .filter(|(arg, rule)| match rule {
            Some(OptionalParamDefault::Lcid) => false,
            Some(OptionalParamDefault::ParamArray) => arg.is_some(),
            _ => true,
        })
        .map(|(arg, _)| arg.unwrap_or_else(Arg::missing))
        .collect();
    dispatch_args.extend(
        positional
            .iter()
            .skip(member.parameter_types.len())
            .flatten()
            .copied(),
    );
    while dispatch_args
        .last()
        .is_some_and(|arg| matches!(arg.value, ArgValue::Missing))
        && !matches!(
            member.invoke_kind,
            TypeLibMemberInvokeKind::PropertyPut | TypeLibMemberInvokeKind::PropertyPutRef
        )
    {
        dispatch_args.pop();
    }
    let dispatch = Dispatch::of(object)?;
    // SAFETY: forwarded from the caller.
    let value = unsafe {
        dispatch.invoke(
            Member::Id(member.token),
            call_kind(member.invoke_kind),
            &dispatch_args,
            project,
        )
    }?;
    Ok((value, EarlyTransport::Dispatch))
}

fn call_kind(kind: TypeLibMemberInvokeKind) -> CallKind {
    match kind {
        TypeLibMemberInvokeKind::Method => CallKind::Method,
        TypeLibMemberInvokeKind::PropertyGet => CallKind::Get,
        TypeLibMemberInvokeKind::PropertyPut => CallKind::Let,
        TypeLibMemberInvokeKind::PropertyPutRef => CallKind::Set,
    }
}

/// The optional rule of each `FUNCDESC` parameter, `None` for metadata without the rules.
fn rules(member: &TypeLibMemberMetadata) -> impl Iterator<Item = Option<&OptionalParamDefault>> {
    let complete = member.parameter_optional_defaults.len() == member.parameter_types.len();
    (0..member.parameter_types.len()).map(move |i| {
        complete
            .then(|| member.parameter_optional_defaults.get(i))
            .flatten()
    })
}

/// The arguments by `FUNCDESC` parameter (the hidden `[lcid]` gets none), then any ParamArray
/// extras. Named arguments go to the parameter of that name, as VBA's compiler places them.
fn arrange<'a>(
    member: &TypeLibMemberMetadata,
    args: &[Arg<'a>],
) -> Result<Vec<Option<Arg<'a>>>, Fault> {
    let visible: Vec<usize> = rules(member)
        .enumerate()
        .filter(|(_, rule)| !matches!(rule, Some(OptionalParamDefault::Lcid)))
        .map(|(i, _)| i)
        .collect();
    let param_array = matches!(
        member.parameter_optional_defaults.last(),
        Some(OptionalParamDefault::ParamArray)
    ) && member.parameter_optional_defaults.len() == member.parameter_types.len();
    let mut arranged: Vec<Option<Arg<'a>>> = vec![None; member.parameter_types.len()];
    let mut extras = Vec::new();
    let mut next = 0;
    // A property assignment's value is the last parameter; the rest are its indexes.
    let put = matches!(
        member.invoke_kind,
        TypeLibMemberInvokeKind::PropertyPut | TypeLibMemberInvokeKind::PropertyPutRef
    );
    let mut visible = visible;
    let args = if put {
        let (Some((value, indexes)), Some(last)) = (args.split_last(), visible.pop()) else {
            return Err(Fault::number(450));
        };
        if value.name.is_some() {
            return Err(Fault::number(448));
        }
        arranged[last] = Some(*value);
        indexes
    } else {
        args
    };
    for arg in args {
        if let Some(name) = arg.name {
            let index = member
                .parameter_names
                .iter()
                .position(|param| param.eq_ignore_ascii_case(name))
                .filter(|index| visible.contains(index))
                .ok_or_else(|| Fault::number(448))?;
            if arranged[index].is_some() {
                return Err(Fault::number(448));
            }
            arranged[index] = Some(Arg { name: None, ..*arg });
            continue;
        }
        match visible.get(next) {
            Some(&index) if !(param_array && index + 1 == member.parameter_types.len()) => {
                arranged[index] = Some(*arg);
                next += 1;
            }
            _ if param_array => extras.push(Some(*arg)),
            // "Wrong number of arguments or invalid property assignment".
            _ => return Err(Fault::number(450)),
        }
    }
    arranged.extend(extras);
    Ok(arranged)
}

/// The declared interface when the member is called through its vtable slot, `None` when it is
/// called through `IDispatch::Invoke` instead:
/// - A pure dispinterface member has no slot; VBA calls it by DISPID too.
/// - A parameter or result shape this marshaller does not build yet goes through `Invoke`, which
///   every dual member also answers.
/// - A proxy whose vtable is not the type library's (anything but `PSOAInterface`) has no slots
///   to call.
///
/// A dual member whose descriptor cannot be trusted (a reserved or out-of-bounds slot, no bound,
/// no interface IID, not `stdcall`) is refused with error 430 rather than called another way, so
/// a wrong descriptor is never hidden.
fn vtable_interface(
    object: &ObjectRef,
    member: &TypeLibMemberMetadata,
) -> Result<Option<OwnedComInterface>, Fault> {
    if !member.is_dual
        || member.source_typekind != Some(SourceTypeKind::Interface)
        || member.vtable_slot.is_none()
    {
        return Ok(None);
    }
    let refuse = |reason: &str| {
        Err(Fault {
            description: Some(format!("COM-E-EARLY-UNSUPPORTED: {reason}")),
            ..Fault::number(430)
        })
    };
    let slot = member.vtable_slot.unwrap_or_default();
    match member.vtable_slot_bound {
        _ if slot < 7 => return refuse("ReservedComSlot"),
        None => return refuse("MissingSlotBound"),
        Some(bound) if slot >= bound => return refuse("SlotOutOfBounds"),
        Some(_) => {}
    }
    let Some(iid) = member.interface_iid.filter(|iid| !iid.is_null()) else {
        return refuse("MissingInterfaceIid");
    };
    if !member.callconv_is_stdcall {
        return refuse("NonStdcall");
    }
    let shapes_supported = (0..member.parameter_types.len())
        .all(|index| wire_is_supported(&wire_type(member, index)))
        && member
            .return_wire_type
            .as_ref()
            .is_none_or(|wire| return_is_supported(wire, member.return_type));
    if !shapes_supported {
        return Ok(None);
    }
    let receiver = object.raw_iunknown().cast::<core::ffi::c_void>();
    // SAFETY: the ObjectRef retains a live IUnknown for the call.
    if unsafe { is_proxy(receiver) } && !proxy_vtable_is_typelib(&iid) {
        return Ok(None);
    }
    // SAFETY: as above.
    match unsafe { query_interface_pointer(receiver, &iid.to_guid()) } {
        // SAFETY: one owned reference from QueryInterface.
        Ok(pointer) => Ok(Some(unsafe { OwnedComInterface::from_owned(pointer) })),
        // "Class does not support Automation or does not support expected interface".
        Err(_) => Err(Fault {
            hresult: Some(E_NOINTERFACE),
            ..Fault::number(430)
        }),
    }
}

/// True for a marshaling proxy.
///
/// # Safety
/// `object` is a live IUnknown-derived pointer.
unsafe fn is_proxy(object: *mut core::ffi::c_void) -> bool {
    // SAFETY: forwarded; a successful QueryInterface hands one reference, released at once.
    match unsafe { query_interface_pointer(object, &IID_IPROXYMANAGER) } {
        Ok(manager) => {
            // SAFETY: the one reference QueryInterface returned.
            drop(unsafe { OwnedComInterface::from_owned(manager) });
            true
        }
        Err(_) => false,
    }
}

thread_local! {
    /// Per-interface answer of [`proxy_vtable_is_typelib`] (a registry read).
    static TYPELIB_MARSHALED: RefCell<HashMap<String, bool>> = RefCell::new(HashMap::new());
}

/// True when the interface is marshaled by the type-library marshaler, whose proxy vtable has the
/// type library's slots.
fn proxy_vtable_is_typelib(iid: &crate::ComInterfaceIid) -> bool {
    let key = crate::windows_typelib_loader::guid_to_string(&iid.to_guid());
    TYPELIB_MARSHALED.with(|cache| {
        if let Some(&known) = cache.borrow().get(&key) {
            return known;
        }
        let known = crate::windows_typelib_loader::interface_is_psoa_marshaled(&key);
        cache.borrow_mut().insert(key, known);
        known
    })
}

/// A parameter's wire shape; metadata without wire shapes has plain Automation types.
fn wire_type(member: &TypeLibMemberMetadata, index: usize) -> TypeLibWireType {
    member
        .parameter_wire_types
        .get(index)
        .cloned()
        .unwrap_or(TypeLibWireType::Automation(member.parameter_types[index]))
}

fn wire_is_supported(wire: &TypeLibWireType) -> bool {
    match wire {
        TypeLibWireType::Automation(_) | TypeLibWireType::InterfacePointer { .. } => true,
        TypeLibWireType::SafeArray { element_vt, .. }
        | TypeLibWireType::ByRefSafeArray { element_vt, .. } => *element_vt != VT_RECORD,
        TypeLibWireType::SafeArrayVariant | TypeLibWireType::ByRefSafeArrayVariant => true,
        TypeLibWireType::ByRefRecord { .. } => true,
        TypeLibWireType::Record { .. } => false,
    }
}

fn return_is_supported(wire: &TypeLibWireType, return_type: Option<TypeLibParamType>) -> bool {
    match wire {
        TypeLibWireType::Automation(_) | TypeLibWireType::InterfacePointer { .. } => true,
        TypeLibWireType::SafeArray { element_vt, .. } => *element_vt != VT_RECORD,
        TypeLibWireType::SafeArrayVariant => true,
        _ => return_type.is_none(),
    }
}

/// The `VARTYPE` of a parameter and whether it is passed by reference.
fn param_vt(ty: TypeLibParamType, wire: &TypeLibWireType) -> (u16, bool) {
    match wire {
        TypeLibWireType::SafeArray { element_vt, .. } => return (VT_ARRAY | element_vt, false),
        TypeLibWireType::SafeArrayVariant => return (VT_ARRAY | VT_VARIANT, false),
        TypeLibWireType::ByRefSafeArray { element_vt, .. } => return (VT_ARRAY | element_vt, true),
        TypeLibWireType::ByRefSafeArrayVariant => return (VT_ARRAY | VT_VARIANT, true),
        TypeLibWireType::ByRefRecord { .. } => return (VT_RECORD, true),
        _ => {}
    }
    use TypeLibParamType as T;
    let long_ptr = if cfg!(target_pointer_width = "64") {
        VT_I8
    } else {
        VT_I4
    };
    match ty {
        T::Variant => (VT_VARIANT, false),
        T::Long => (VT_I4, false),
        T::Integer => (VT_I2, false),
        T::String => (VT_BSTR, false),
        T::Boolean => (VT_BOOL, false),
        T::Double => (VT_R8, false),
        T::Single => (VT_R4, false),
        T::Currency => (VT_CY, false),
        T::Date => (VT_DATE, false),
        T::Decimal => (VT_DECIMAL, false),
        T::Object => (VT_DISPATCH, false),
        T::Byte => (VT_UI1, false),
        T::LongLong => (VT_I8, false),
        T::LongPtr => (long_ptr, false),
        T::Record => (VT_RECORD, false),
        T::ByRefVariant => (VT_VARIANT, true),
        T::ByRefLong => (VT_I4, true),
        T::ByRefInteger => (VT_I2, true),
        T::ByRefString => (VT_BSTR, true),
        T::ByRefDouble => (VT_R8, true),
        T::ByRefSingle => (VT_R4, true),
        T::ByRefCurrency => (VT_CY, true),
        T::ByRefDate => (VT_DATE, true),
        T::ByRefDecimal => (VT_DECIMAL, true),
        T::ByRefObject => (VT_DISPATCH, true),
        T::ByRefByte => (VT_UI1, true),
        T::ByRefBoolean => (VT_BOOL, true),
        T::ByRefLongLong => (VT_I8, true),
        T::ByRefLongPtr => (long_ptr, true),
        T::ByRefRecord => (VT_RECORD, true),
    }
}

/// The raw `vt` of a cell.
fn raw_vt(cell: &Variant) -> u16 {
    let bytes = cell.to_wire_bytes();
    u16::from_le_bytes([bytes[0], bytes[1]])
}

/// The cell a temporary of `vt` holds for an argument value (VBA's argument coercion; error 13
/// when it does not convert, 94 for Null to a typed parameter).
fn coerce_argument(value: &Variant, vt: u16, coerce: Coerce<'_>) -> Result<Variant, Fault> {
    if vt == VT_VARIANT {
        return Ok(value.clone());
    }
    if vt == VT_RECORD {
        return if value.vtype() == VarType::Record {
            Ok(value.clone())
        } else {
            Err(Fault::number(13))
        };
    }
    if vt & VT_ARRAY != 0 {
        return if value.holds_array() {
            Ok(value.clone())
        } else {
            Err(Fault::number(13))
        };
    }
    if value.vtype() == VarType::Null {
        return Err(Fault::number(94));
    }
    if raw_vt(value) == vt {
        return Ok(value.clone());
    }
    coerce(value, vt)
}

/// An object argument as the parameter's interface: a project object crosses as its `IDispatch`,
/// a specific interface is `QueryInterface`d (VBA error 13 when the object lacks it), and Nothing
/// is a null pointer.
fn object_argument(
    value: &Variant,
    iid: Option<&crate::ComInterfaceIid>,
    project: ProjectDispatch<'_>,
) -> Result<Variant, Fault> {
    if value.vtype() != VarType::Object {
        return Err(Fault::number(13));
    }
    let Some(object) = value.as_object_ref() else {
        return Ok(Variant::nothing());
    };
    if object.is_nothing_placeholder() {
        return Ok(Variant::nothing());
    }
    let crossed = if object.is_compat_object() {
        project(&object)?
    } else {
        value.clone()
    };
    let iid = match iid.filter(|iid| !iid.is_null()) {
        Some(iid) => iid.to_guid(),
        // An `IDispatch*` parameter: a `VT_UNKNOWN` object crosses as its IDispatch.
        None if raw_vt(&crossed) == VT_DISPATCH => return Ok(crossed),
        None => IID_IDISPATCH,
    };
    let Some(crossed_object) = crossed.as_object_ref() else {
        return Ok(Variant::nothing());
    };
    // SAFETY: the ObjectRef retains a live IUnknown for the call.
    let pointer = unsafe { query_interface_pointer(crossed_object.raw_iunknown().cast(), &iid) }
        .map_err(|_| Fault::number(13))?;
    // SAFETY: one owned reference from QueryInterface, which the ObjectRef takes over.
    unsafe { ObjectRef::from_raw_iunknown_owned(pointer.cast()) }
        .map(Variant::from_object_ref)
        .ok_or_else(|| Fault::number(13))
}

/// The cell an omitted parameter gets: the typelib default, `VT_ERROR` / `DISP_E_PARAMNOTFOUND`
/// for an optional `VARIANT`, an empty ParamArray, the locale for `[lcid]`; VBA error 449 for a
/// required one.
fn omitted_argument(
    rule: Option<&OptionalParamDefault>,
    vt: u16,
    coerce: Coerce<'_>,
) -> Result<Variant, Fault> {
    let value = match rule {
        Some(OptionalParamDefault::HasDefault(default)) => default.to_variant(),
        Some(OptionalParamDefault::OptionalVariant) => Variant::from_scode(DISP_E_PARAMNOTFOUND),
        Some(OptionalParamDefault::ParamArray) => Variant::from_safearray(
            calcvba_runtime::safe_array::SafeArray::from_variants(Vec::new()),
        ),
        Some(OptionalParamDefault::Lcid) => Variant::from_i32(LOCALE_USER_DEFAULT as i32),
        Some(OptionalParamDefault::OptionalNoDefault) if vt == VT_VARIANT => {
            Variant::from_scode(DISP_E_PARAMNOTFOUND)
        }
        Some(OptionalParamDefault::OptionalNoDefault) => Variant::empty(),
        // "Argument not optional".
        Some(OptionalParamDefault::Required) | None => return Err(Fault::number(449)),
    };
    if vt == VT_VARIANT || vt & VT_ARRAY != 0 {
        return Ok(value);
    }
    coerce_argument(&value, vt, coerce)
}

/// What argument shaping needs from the caller.
struct Context<'a> {
    project: ProjectDispatch<'a>,
    coerce: Coerce<'a>,
}

/// One argument of the native frame.
enum Slot {
    /// A cell passed with this `VARTYPE` (a borrowed bit copy or a temporary's).
    Cell(u16, [u8; WIRE_SIZE]),
    /// A `VT_BYREF | vt` argument pointing at `temps[index]`'s storage.
    Temp(u16, usize),
}

/// One vtable call through `DispCallFunc`.
///
/// # Safety
/// As [`invoke`]; `interface` is the member's declared interface, retained for the call.
unsafe fn call_vtable(
    interface: &OwnedComInterface,
    member: &TypeLibMemberMetadata,
    positional: &[Option<Arg<'_>>],
    context: &Context<'_>,
) -> Result<Variant, Fault> {
    let project = context.project;
    let count = member.parameter_types.len();
    // Temporaries live in a vector whose storage never moves during the call.
    let mut temps: Vec<Variant> = Vec::with_capacity(count + 1);
    // (temps index, reference) pairs written back after the call.
    let mut writebacks: Vec<(usize, &Variant)> = Vec::new();
    // Cells borrowed for the call (project objects' IDispatch, coerced values).
    let mut kept: Vec<Variant> = Vec::new();
    let mut slots: Vec<Slot> = Vec::with_capacity(count + 1);
    let param_array = matches!(
        member.parameter_optional_defaults.last(),
        Some(OptionalParamDefault::ParamArray)
    );
    let rules: Vec<Option<&OptionalParamDefault>> = rules(member).collect();
    for index in 0..count {
        let (vt, by_ref) = param_vt(member.parameter_types[index], &wire_type(member, index));
        let iid = member.parameter_iids.get(index).and_then(Option::as_ref);
        let arg = positional[index];
        if param_array && index + 1 == count {
            // The ParamArray tail: this argument, if any, and the extras, as Variants.
            let mut values: Vec<Variant> = Vec::new();
            for extra in std::iter::once(arg).chain(positional[count..].iter().copied()) {
                match extra.map(|arg| arg.value) {
                    Some(ArgValue::Value(value)) => values.push(value.clone()),
                    Some(ArgValue::Reference(reference)) => values.push(reference.clone()),
                    Some(ArgValue::Missing) => {
                        values.push(Variant::from_scode(DISP_E_PARAMNOTFOUND))
                    }
                    None => {}
                }
            }
            temps.push(Variant::from_safearray(
                calcvba_runtime::safe_array::SafeArray::from_variants(values),
            ));
            slots.push(if by_ref {
                Slot::Temp(vt, temps.len() - 1)
            } else {
                Slot::Cell(vt, temps[temps.len() - 1].to_wire_bytes())
            });
            continue;
        }
        let value = match arg.map(|arg| arg.value) {
            None | Some(ArgValue::Missing) => None,
            Some(value) => Some(value),
        };
        let Some(value) = value else {
            let omitted = omitted_argument(rules[index], vt, context.coerce)?;
            temps.push(omitted);
            slots.push(if by_ref {
                Slot::Temp(vt, temps.len() - 1)
            } else {
                Slot::Cell(vt, temps[temps.len() - 1].to_wire_bytes())
            });
            continue;
        };
        if by_ref {
            if let ArgValue::Reference(reference) = value {
                let target = reference.byref_target().map(|(target, _)| target);
                let object_target = matches!(target, Some(VT_DISPATCH | VT_UNKNOWN));
                if target == Some(vt) || (vt == VT_DISPATCH && object_target && iid.is_none()) {
                    // The caller's storage itself.
                    slots.push(Slot::Cell(VT_BYREF | vt, reference.to_wire_bytes()));
                    continue;
                }
                if vt == VT_VARIANT {
                    // A typed variable to a `VARIANT*`: a cell holding the reference.
                    temps.push(reference.clone_reference());
                    slots.push(Slot::Temp(vt, temps.len() - 1));
                    continue;
                }
                // A temporary of the parameter's type, written back after the call.
                let current = reference.clone();
                temps.push(typed_temp(&current, vt, iid, context)?);
                writebacks.push((temps.len() - 1, reference));
                slots.push(Slot::Temp(vt, temps.len() - 1));
                continue;
            }
            let ArgValue::Value(value) = value else {
                unreachable!("missing arguments were handled above")
            };
            // An expression passed ByRef: a temporary nothing reads back.
            temps.push(typed_temp(value, vt, iid, context)?);
            slots.push(Slot::Temp(vt, temps.len() - 1));
            continue;
        }
        // ByVal.
        let read;
        let value: &Variant = match value {
            ArgValue::Value(value) => value,
            ArgValue::Reference(reference) => {
                read = reference.clone();
                &read
            }
            ArgValue::Missing => unreachable!("missing arguments were handled above"),
        };
        if vt == VT_VARIANT {
            let crossed = match value.as_object_ref() {
                Some(object) if object.is_compat_object() => {
                    kept.push(project(&object)?);
                    kept[kept.len() - 1].to_wire_bytes()
                }
                _ => value.to_wire_bytes(),
            };
            slots.push(Slot::Cell(VT_VARIANT, crossed));
            continue;
        }
        if raw_vt(value) == vt && vt != VT_DISPATCH {
            // Already the parameter's type: borrowed.
            slots.push(Slot::Cell(vt, value.to_wire_bytes()));
            continue;
        }
        kept.push(typed_temp(value, vt, iid, context)?);
        slots.push(Slot::Cell(vt, kept[kept.len() - 1].to_wire_bytes()));
    }
    // The `[out, retval]` pointer.
    let result_vt = match member.invoke_kind {
        TypeLibMemberInvokeKind::Method | TypeLibMemberInvokeKind::PropertyGet => {
            member.return_type.map(|ty| {
                let wire = member
                    .return_wire_type
                    .clone()
                    .unwrap_or(TypeLibWireType::Automation(ty));
                param_vt(ty, &wire).0
            })
        }
        _ => None,
    };
    let mut result = [0u8; WIRE_SIZE];

    // The frame: one VARIANT per argument, each carrying its value or pointer.
    let mut cells: Vec<[u8; WIRE_SIZE]> = Vec::with_capacity(slots.len() + 1);
    let mut types: Vec<u16> = Vec::with_capacity(slots.len() + 1);
    for slot in &slots {
        match *slot {
            Slot::Cell(vt, bytes) => {
                types.push(vt);
                cells.push(bytes);
            }
            Slot::Temp(VT_RECORD, index) => {
                // A record's pointer is its `pvRecord`, already the cell's payload.
                let mut bytes = temps[index].to_wire_bytes();
                bytes[..2].copy_from_slice(&(VT_BYREF | VT_RECORD).to_le_bytes());
                types.push(VT_BYREF | VT_RECORD);
                cells.push(bytes);
            }
            Slot::Temp(vt, index) => {
                let cell = temps[index].as_variant_cell_ptr().cast::<u8>();
                // A VARIANT or DECIMAL is the whole cell; other types are its payload.
                let target = if vt == VT_VARIANT || vt == VT_DECIMAL {
                    cell
                } else {
                    // SAFETY: the payload is 8 bytes into the 16-byte cell.
                    unsafe { cell.add(8) }
                };
                types.push(VT_BYREF | vt);
                cells.push(by_ref_cell(VT_BYREF | vt, target));
            }
        }
    }
    if let Some(vt) = result_vt {
        let target = if vt == VT_VARIANT || vt == VT_DECIMAL {
            result.as_mut_ptr()
        } else {
            // SAFETY: the payload is 8 bytes into the 16-byte result cell.
            unsafe { result.as_mut_ptr().add(8) }
        };
        types.push(VT_BYREF | vt);
        cells.push(by_ref_cell(VT_BYREF | vt, target));
    }
    let mut pointers: Vec<*mut VARIANT> = cells
        .iter_mut()
        .map(|cell| cell.as_mut_ptr().cast::<VARIANT>())
        .collect();
    let slot = usize::from(member.vtable_slot.unwrap_or_default());
    // SAFETY: zero is VT_EMPTY.
    let mut returned: VARIANT = unsafe { core::mem::zeroed() };
    crate::windows_vtable::clear_thread_error_info();
    // SAFETY: `interface` is the declared interface and `slot` lies inside its vtable (checked by
    // `vtable_interface`); each argument VARIANT carries a value of its `VARTYPE` or a pointer to
    // live storage of it, all of which outlive the call.
    let called = unsafe {
        DispCallFunc(
            interface.as_ptr(),
            slot * core::mem::size_of::<*const core::ffi::c_void>(),
            CC_STDCALL,
            VT_ERROR,
            types.len() as u32,
            types.as_ptr(),
            pointers.as_mut_ptr(),
            &mut returned,
        )
    };
    drop(kept);
    // The method's HRESULT, in the VT_ERROR result.
    // SAFETY: the VARIANT layout is the runtime cell's.
    let returned = unsafe { core::mem::transmute::<VARIANT, [u8; WIRE_SIZE]>(returned) };
    let hr = if called < 0 {
        called
    } else {
        i32::from_le_bytes([returned[8], returned[9], returned[10], returned[11]])
    };
    // ByRef temporaries go back to the references they stand for, whether or not the call failed
    // (the callee writes the caller's storage directly in the other cases).
    let mut writeback_fault = None;
    for (index, reference) in writebacks {
        if let Err(fault) = write_back(reference, &temps[index], context.coerce) {
            writeback_fault.get_or_insert(fault);
        }
    }
    drop(temps);
    let value = result_vt.map(|vt| {
        if vt != VT_VARIANT {
            result[..2].copy_from_slice(&vt.to_le_bytes());
        }
        // SAFETY: the callee wrote a value of `vt` (or nothing, leaving it zero) into the result
        // cell, which is ours.
        unsafe { adopt(core::mem::transmute::<[u8; WIRE_SIZE], VARIANT>(result)) }
    });
    if hr < 0 {
        if let Some(Ok(value)) = value {
            drop(value);
        }
        return Err(fault_from_hresult(hr, None));
    }
    if let Some(fault) = writeback_fault {
        return Err(fault);
    }
    match value {
        Some(value) if result_vt == Some(VT_DISPATCH) => interface_result(value?),
        Some(value) => value,
        None => Ok(Variant::empty()),
    }
}

/// A `VT_BYREF | vt` argument cell pointing at `target`.
fn by_ref_cell(vt: u16, target: *mut u8) -> [u8; WIRE_SIZE] {
    let mut cell = [0u8; WIRE_SIZE];
    cell[..2].copy_from_slice(&vt.to_le_bytes());
    cell[8..8 + core::mem::size_of::<usize>()]
        .copy_from_slice(&(target.expose_provenance()).to_le_bytes());
    cell
}

/// A temporary of the parameter's type holding `value`.
fn typed_temp(
    value: &Variant,
    vt: u16,
    iid: Option<&crate::ComInterfaceIid>,
    context: &Context<'_>,
) -> Result<Variant, Fault> {
    if vt == VT_DISPATCH || vt == VT_UNKNOWN {
        return object_argument(value, iid, context.project);
    }
    coerce_argument(value, vt, context.coerce)
}

/// Stores what the callee left in a ByRef temporary through the caller's reference, with VBA's
/// assignment coercion (error 13 when it does not convert).
fn write_back(reference: &Variant, temp: &Variant, coerce: Coerce<'_>) -> Result<(), Fault> {
    let Some((target, _)) = reference.byref_target() else {
        return Ok(());
    };
    let value = if target == VT_VARIANT || target & VT_ARRAY != 0 || raw_vt(temp) == target {
        temp.clone()
    } else {
        coerce(temp, target)?
    };
    // SAFETY: the reference was built for this call over live storage of its type.
    unsafe { reference.write_through(value) }.map_err(|_| Fault::number(13))
}

/// An object result as the cell holds it: a dual interface is its `IDispatch`, anything else is a
/// `VT_UNKNOWN`.
fn interface_result(value: Variant) -> Result<Variant, Fault> {
    let Some(object) = value.as_object_ref() else {
        return Ok(value);
    };
    if object.is_nothing_placeholder() || raw_vt(&value) != VT_DISPATCH {
        return Ok(value);
    }
    // SAFETY: the ObjectRef retains a live IUnknown for the call.
    match unsafe { query_interface_pointer(object.raw_iunknown().cast(), &IID_IDISPATCH) } {
        // SAFETY: one owned reference from QueryInterface, which the ObjectRef takes over.
        Ok(dispatch) => unsafe { ObjectRef::from_raw_iunknown_owned(dispatch.cast()) }
            .map(Variant::from_object_ref)
            .ok_or_else(|| Fault::number(91)),
        Err(_) => {
            // Not an IDispatch: the cell is a VT_UNKNOWN for the same pointer.
            let mut bytes = value.to_wire_bytes();
            drop(object);
            core::mem::forget(value);
            bytes[..2].copy_from_slice(&VT_UNKNOWN.to_le_bytes());
            // SAFETY: the forgotten cell's one reference moves into the adopted one.
            unsafe { adopt(core::mem::transmute::<[u8; WIRE_SIZE], VARIANT>(bytes)) }
        }
    }
}
