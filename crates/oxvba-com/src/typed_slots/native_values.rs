//! Native cells for planned typed COM calls. Variant cells exist only for
//! declared VARIANT/SAFEARRAY signatures, independently of procedure storage.
use super::{
    factory::{NativeReceiver, make_native_object},
    native::map_hresult,
    plan, *,
};
use crate::{
    windows_client::OwnedComInterface,
    windows_ffi_bridge::{FfiArg, FfiReturnType, call_via_libffi},
};
use calcvba_host_api::{HostObject, TypedFault, Value, object::planned_slots::MemberArguments};
use oxvba_runtime::{VarType, Variant};
use std::ffi::c_void;
use windows_sys::Win32::{
    Foundation::SysFreeString,
    System::Variant::{VARIANT, VT_ARRAY, VT_VARIANT, VariantClear},
};

fn fault(number: i32) -> TypedFault {
    TypedFault {
        number,
        description: oxvba_runtime::default_error_message(number).into(),
        source: None,
        help_file: None,
        help_context: None,
    }
}
// Windows x64 requires 16-byte-aligned caller storage for indirect by-value
// aggregate arguments. The native cell includes VARIANT's full record union.
#[repr(C, align(16))]
struct NativeVariant(VARIANT);
impl NativeVariant {
    fn empty() -> Self {
        // SAFETY: an all-zero native VARIANT is VT_EMPTY.
        Self(unsafe { std::mem::zeroed() })
    }
    fn input(value: &Variant) -> Result<Self, TypedFault> {
        let mut cell = Self::empty();
        if value.vtype() == VarType::Object && value.as_object_ref().is_none() {
            // SAFETY: initialized union payload is null; tag makes this Nothing.
            cell.0.Anonymous.Anonymous.vt = windows_sys::Win32::System::Variant::VT_DISPATCH;
            return Ok(cell);
        }
        let value = crate::ComValue::from_variant(value).map_err(|_| fault(13))?;
        let mut dispatches = Vec::new();
        let mut resolve = |object: oxvba_runtime::ObjectRef| {
            let portable = HostObject::from_object_ref(&object)
                .ok_or("native argument is not a prepared host object")?;
            let native = portable
                .payload()
                .and_then(|v| v.downcast_ref::<NativeReceiver>())
                .ok_or("host object has no native receiver")?;
            // A prepared nominal interface need not itself be IDispatch. Keep
            // the exact QI'd dispatch owner until the VARIANT has retained it.
            // SAFETY: the wrapper retains its same-apartment native interface.
            let pointer = unsafe {
                crate::query_interface_pointer(
                    native.pointer.as_ptr(),
                    &crate::windows_client::IID_IDISPATCH,
                )
            }?;
            // SAFETY: successful QI transferred one owned dispatch reference.
            dispatches.push(unsafe { OwnedComInterface::from_owned(pointer) });
            Ok(pointer)
        };
        let mut retain = |pointer: *mut c_void| {
            // SAFETY: resolve returns a retained same-apartment dual interface.
            unsafe {
                crate::add_ref_dispatch(pointer.cast());
            }
        };
        // SAFETY: fresh initialized output; resolver/retain preserve native object
        // ownership, including elements recursively marshalled inside arrays.
        unsafe {
            crate::set_variant_from_com_value(&mut cell.0, &value, &mut resolve, &mut retain)
        }
        .map_err(|_| fault(13))?;
        Ok(cell)
    }
    fn value(&self, owner: &NativeReceiver) -> Result<Variant, TypedFault> {
        let mut query = |pointer| {
            // SAFETY: valid native unknown borrowed from an owned native cell.
            unsafe {
                crate::query_interface_pointer(pointer, &crate::windows_client::IID_IDISPATCH)
            }
        };
        let mut retain = |pointer: *mut c_void| {
            // SAFETY: decoder borrows a native cell with an owned interface ref.
            unsafe {
                crate::add_ref_dispatch(pointer.cast());
            }
        };
        let mut bind = |pointer: *mut c_void, _: &str, _: &'static str| {
            if pointer.is_null() {
                return Ok(Variant::nothing());
            }
            // SAFETY: decoder transfers one native IDispatch reference.
            let incoming = unsafe { OwnedComInterface::from_owned(pointer) };
            // SAFETY: make_native_object borrows this same-apartment reference
            // and retains its own independently QI'd native receiver.
            let object = unsafe {
                make_native_object(
                    owner.bridge.clone(),
                    incoming.as_ptr(),
                    None,
                    owner.library.clone(),
                )
            }?;
            Ok(Variant::from_object_ref(object.object_ref()))
        };
        // SAFETY: initialized native cell remains owned throughout decoding.
        unsafe {
            crate::variant_to_variant_value(
                &self.0,
                &mut query,
                &mut retain,
                &mut bind,
                "",
                "typed native result",
            )
        }
        .map_err(|_| fault(13))
    }
}
impl Drop for NativeVariant {
    fn drop(&mut self) {
        // SAFETY: this guard owns exactly one initialized native VARIANT.
        unsafe {
            VariantClear(&mut self.0);
        }
    }
}
struct NativePointer {
    raw: *mut c_void,
}
impl Drop for NativePointer {
    fn drop(&mut self) {
        if !self.raw.is_null() {
            // SAFETY: this initialized cell owns one same-apartment interface.
            unsafe {
                crate::release_unknown(self.raw);
            }
        }
    }
}
struct NativeBstr {
    raw: *mut u16,
}
impl Drop for NativeBstr {
    fn drop(&mut self) {
        // SAFETY: the cell owns this BSTR; null is an allowed empty BSTR.
        unsafe {
            SysFreeString(self.raw);
        }
    }
}
enum Cell {
    Empty,
    Long(i32),
    Double(f64),
    Boolean(i16),
    String(NativeBstr),
    Variant(NativeVariant),
    Array(NativeVariant),
    Object(NativePointer),
}
impl Cell {
    fn output(ty: Option<P>, wire: Option<&W>) -> Result<Self, TypedFault> {
        let carrier = match ty {
            Some(ty) => Some(plan::carrier(ty).ok_or_else(|| fault(13))?),
            None => None,
        };
        Ok(match carrier {
            None => Self::Empty,
            Some(T::Long) => Self::Long(0),
            Some(T::Double) => Self::Double(0.0),
            Some(T::Boolean) => Self::Boolean(0),
            Some(T::String) => Self::String(NativeBstr {
                raw: std::ptr::null_mut(),
            }),
            Some(T::HostObject) => Self::Object(NativePointer {
                raw: std::ptr::null_mut(),
            }),
            Some(T::Variant) => {
                let mut value = NativeVariant::empty();
                if let Some(element) = array_element(wire) {
                    // SAFETY: set only the initialized VARIANT tag; null parray
                    // is a valid empty output until the native call writes it.
                    value.0.Anonymous.Anonymous.vt = VT_ARRAY | element;
                    Self::Array(value)
                } else {
                    Self::Variant(value)
                }
            }
            _ => return Err(fault(13)),
        })
    }
    fn input(
        value: &Value,
        ty: P,
        wire: &W,
        iid: Option<crate::ComInterfaceIid>,
    ) -> Result<Self, TypedFault> {
        Ok(match value {
            Value::Long(v) if plan::carrier(ty) == Some(T::Long) => Self::Long(*v),
            Value::Double(v) if plan::carrier(ty) == Some(T::Double) => Self::Double(*v),
            Value::Boolean(v) if plan::carrier(ty) == Some(T::Boolean) => {
                Self::Boolean(if *v { -1 } else { 0 })
            }
            Value::String(v) if plan::carrier(ty) == Some(T::String) => Self::String(NativeBstr {
                raw: v.clone_raw_bstr().map_err(|_| fault(7))?,
            }),
            Value::Variant(v) if plan::carrier(ty) == Some(T::Variant) => {
                let cell = NativeVariant::input(v)?;
                if let Some(element) = array_element(Some(wire)) {
                    // SAFETY: cell is initialized by the native Variant adapter.
                    let tag = unsafe { cell.0.Anonymous.Anonymous.vt };
                    if tag != (VT_ARRAY | element) {
                        return Err(fault(13));
                    }
                    Self::Array(cell)
                } else {
                    Self::Variant(cell)
                }
            }
            Value::Object(value) if plan::carrier(ty) == Some(T::HostObject) => {
                let pointer = if let Some(object) = value {
                    let native = object
                        .payload()
                        .and_then(|v| v.downcast_ref::<NativeReceiver>())
                        .ok_or_else(|| fault(13))?;
                    // SAFETY: retained exact native interface; QI adds one ref.
                    unsafe {
                        crate::query_interface_pointer(
                            native.pointer.as_ptr(),
                            &iid.map(|i| i.to_guid())
                                .unwrap_or(crate::windows_client::IID_IDISPATCH),
                        )
                    }
                    .map_err(|_| fault(13))?
                } else {
                    std::ptr::null_mut()
                };
                Self::Object(NativePointer { raw: pointer })
            }
            _ => return Err(fault(13)),
        })
    }
    fn pointer(&mut self) -> *mut c_void {
        match self {
            Self::Empty => std::ptr::null_mut(),
            Self::Long(v) => std::ptr::from_mut(v).cast(),
            Self::Double(v) => std::ptr::from_mut(v).cast(),
            Self::Boolean(v) => std::ptr::from_mut(v).cast(),
            Self::String(v) => std::ptr::addr_of_mut!(v.raw).cast(),
            Self::Object(v) => std::ptr::addr_of_mut!(v.raw).cast(),
            Self::Variant(v) => std::ptr::addr_of_mut!(v.0).cast(),
            Self::Array(v) => {
                // SAFETY: a declared array cell keeps the parray union active.
                unsafe { std::ptr::addr_of_mut!(v.0.Anonymous.Anonymous.Anonymous.parray).cast() }
            }
        }
    }
    fn argument(&mut self, by_ref: bool) -> FfiArg {
        if by_ref {
            return FfiArg::Pointer(self.pointer());
        }
        match self {
            Self::Long(v) => FfiArg::Long(*v),
            Self::Double(v) => FfiArg::Double(*v),
            Self::Boolean(v) => FfiArg::Boolean(*v),
            Self::String(v) => FfiArg::Pointer(v.raw.cast()),
            Self::Object(v) => FfiArg::Pointer(v.raw),
            // Native x64 VARIANT by-value transport is an indirect struct copy.
            Self::Variant(v) => FfiArg::Pointer(std::ptr::addr_of_mut!(v.0).cast()),
            Self::Array(v) => {
                // SAFETY: array cell's initialized active union member is parray.
                FfiArg::Pointer(unsafe { v.0.Anonymous.Anonymous.Anonymous.parray }.cast())
            }
            Self::Empty => unreachable!("admitted native input"),
        }
    }
    fn value(
        &mut self,
        owner: &NativeReceiver,
        interface: Option<usize>,
    ) -> Result<Value, TypedFault> {
        Ok(match self {
            Self::Empty => Value::Empty,
            Self::Long(v) => Value::Long(*v),
            Self::Double(v) => Value::Double(*v),
            Self::Boolean(v) => Value::Boolean(*v != 0),
            Self::String(v) => {
                let raw = std::mem::replace(&mut v.raw, std::ptr::null_mut());
                // SAFETY: transfer exactly this native BSTR owner into Value.
                Value::String(unsafe { calcvba_host_api::BStr::from_raw_bstr(raw) })
            }
            Self::Variant(v) | Self::Array(v) => Value::Variant(v.value(owner)?),
            Self::Object(v) => {
                if v.raw.is_null() {
                    Value::Object(None)
                } else {
                    // SAFETY: output owns one native interface until this returns;
                    // factory QIs and retains the declared interface independently.
                    Value::Object(Some(
                        unsafe {
                            make_native_object(
                                owner.bridge.clone(),
                                v.raw,
                                interface,
                                owner.library.clone(),
                            )
                        }
                        .map_err(|_| fault(13))?,
                    ))
                }
            }
        })
    }
}
fn array_element(wire: Option<&W>) -> Option<u16> {
    match wire {
        Some(W::SafeArray { element_vt, .. } | W::ByRefSafeArray { element_vt, .. }) => {
            Some(*element_vt)
        }
        Some(W::SafeArrayVariant | W::ByRefSafeArrayVariant) => Some(VT_VARIANT),
        _ => None,
    }
}
fn default_input(member: &TypeLibMemberMetadata, index: usize) -> Result<Value, TypedFault> {
    let value = plan::default_value(member, index).ok_or_else(|| fault(450))?;
    Ok(
        match plan::carrier(member.parameter_types[index]).ok_or_else(|| fault(13))? {
            T::Variant => Value::Variant(value),
            T::Long => Value::Long(
                value
                    .as_i32()
                    .or_else(|| value.as_i64().and_then(|v| i32::try_from(v).ok()))
                    .ok_or_else(|| fault(13))?,
            ),
            T::Double => Value::Double(
                value
                    .as_f64()
                    .or_else(|| value.as_i32().map(f64::from))
                    .or_else(|| value.as_i64().map(|v| v as f64))
                    .ok_or_else(|| fault(13))?,
            ),
            T::Boolean => Value::Boolean(value.as_bool().ok_or_else(|| fault(13))?),
            T::String => Value::String(value.as_bstr().ok_or_else(|| fault(13))?),
            T::HostObject
                if value.vtype() == VarType::Object && value.as_object_ref().is_none() =>
            {
                Value::Object(None)
            }
            _ => return Err(fault(13)),
        },
    )
}
/// # Safety
/// owner retains the QI'd exact interface and authoritative member metadata;
/// the complete native signature and slot bound were validated before entry.
pub(super) unsafe fn invoke(
    owner: &NativeReceiver,
    prepared: &super::library::NativeMemberPlan,
    arguments: &mut MemberArguments<'_>,
) -> Result<Value, TypedFault> {
    let member = &prepared.member;
    let (parameters, _) =
        plan::visible_signature(member, arguments.len()).ok_or_else(|| fault(450))?;
    if parameters
        .iter()
        .enumerate()
        .any(|(i, ty)| arguments.get(i).is_none_or(|v| v.ty() != ty.carrier()))
    {
        return Err(fault(13));
    }
    let indices = plan::visible_indices(member);
    let mut cells: [Cell; plan::MAX_NATIVE_PARAMETERS] = std::array::from_fn(|_| Cell::Empty);
    let mut source = [None; plan::MAX_NATIVE_PARAMETERS];
    let mut canonical = [0; plan::MAX_NATIVE_PARAMETERS];
    for index in 0..member.parameter_types.len() {
        canonical[index] = index;
        let visible = indices.iter().position(|i| *i == index);
        source[index] = visible.filter(|i| *i < arguments.len());
        if let Some(i) = source[index]
            && plan::is_by_ref(member, index)
        {
            let alias = arguments.alias(i).ok_or_else(|| fault(13))?;
            if alias != i {
                let previous = indices[alias];
                if !plan::is_by_ref(member, previous)
                    || member.parameter_types[previous] != member.parameter_types[index]
                    || member.parameter_wire_types[previous] != member.parameter_wire_types[index]
                    || member.parameter_iids.get(previous) != member.parameter_iids.get(index)
                {
                    return Err(fault(13));
                }
                canonical[index] = previous;
                continue;
            }
        }
        let default;
        let value = if let Some(i) = source[index] {
            arguments.get(i).ok_or_else(|| fault(13))?
        } else {
            default = default_input(member, index)?;
            &default
        };
        cells[index] = Cell::input(
            value,
            member.parameter_types[index],
            &member.parameter_wire_types[index],
            member.parameter_iids.get(index).copied().flatten(),
        )?;
    }
    let mut output = Cell::output(member.return_type, member.return_wire_type.as_ref())?;
    let mut ffi = Vec::with_capacity(member.parameter_types.len() + 2);
    ffi.push(FfiArg::Pointer(owner.pointer.as_ptr()));
    // Derive each cell authority once after initialization. Repeated ByRef
    // arguments reuse that exact raw pointer, rather than retagging the cell.
    let cells_pointer = std::ptr::addr_of_mut!(cells).cast::<Cell>();
    for index in 0..member.parameter_types.len() {
        let argument = if canonical[index] != index {
            ffi[canonical[index] + 1].clone()
        } else {
            // SAFETY: one disjoint initialized array element, in bounds, whose
            // address remains stable until the native call returns.
            unsafe { (&mut *cells_pointer.add(index)).argument(plan::is_by_ref(member, index)) }
        };
        ffi.push(argument);
    }
    if member.return_type.is_some() {
        ffi.push(FfiArg::Pointer(output.pointer()));
    }
    // SAFETY: exact retained IID, checked slot bound and complete native signature.
    let entry = unsafe {
        *(*(owner.pointer.as_ptr().cast::<*const *const c_void>()))
            .add(member.vtable_slot.unwrap() as usize)
    };
    crate::windows_vtable::clear_thread_error_info();
    // SAFETY: each native cell is initialized and immobile until call completion;
    // the CIF uses exactly the imported Windows x64 parameter/retval ABI.
    let hr = unsafe { call_via_libffi(entry as usize, &ffi, FfiReturnType::Long) } as i32;
    // Capture IErrorInfo before any cell decoding, QI or cleanup can replace it.
    let failure = (hr < 0).then(|| map_hresult(hr));
    owner.bridge.record_typed_vtable_call();
    let mut writeback_failure = None;
    for index in 0..member.parameter_types.len() {
        if canonical[index] != index || !plan::is_by_ref(member, index) {
            continue;
        }
        if let Some(i) = source[index] {
            let decoded = cells[index]
                .value(owner, prepared.parameter_interfaces[index])
                .and_then(|value| arguments.set(i, value).map_err(fault));
            if let Err(error) = decoded {
                writeback_failure.get_or_insert(error);
            }
        }
    }
    if let Some(error) = failure.or(writeback_failure) {
        return Err(error);
    }
    output.value(owner, prepared.result_interface)
}
