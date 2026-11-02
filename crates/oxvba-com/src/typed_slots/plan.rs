//! Admission facts for descriptor-planned native members.
use super::*;

pub const MAX_NATIVE_PARAMETERS: usize = 32;

/// A dispinterface has only the seven IDispatch entries. Its typed table index
/// is a compiler metadata index, never an invented native vtable slot.
pub fn is_dispatch(member: &TypeLibMemberMetadata) -> bool {
    member.source_typekind == Some(SourceTypeKind::Dispatch)
        && !member.is_dual
        && member.vtable_slot.is_none()
        && member.vtable_slot_bound == Some(7)
}

pub fn table_slot(
    member: &TypeLibMemberMetadata,
    interface: &[TypeLibMemberMetadata],
) -> Option<usize> {
    if is_dispatch(member) {
        interface
            .iter()
            .position(|candidate| candidate == member)
            .filter(|i| *i < 4096)
    } else {
        member
            .vtable_slot
            .zip(member.vtable_slot_bound)
            .filter(|(slot, bound)| slot < bound)
            .map(|(slot, _)| usize::from(slot))
    }
}

pub fn carrier(ty: P) -> Option<T> {
    Some(match ty {
        P::Long | P::ByRefLong => T::Long,
        P::Double | P::ByRefDouble => T::Double,
        P::Boolean | P::ByRefBoolean => T::Boolean,
        P::String | P::ByRefString => T::String,
        P::Variant | P::ByRefVariant => T::Variant,
        P::Object | P::ByRefObject => T::HostObject,
        _ => return None,
    })
}
fn wire_supported(ty: P, wire: &W, result: bool) -> bool {
    match wire {
        W::Automation(w) => *w == ty && carrier(ty).is_some(),
        W::InterfacePointer { .. } => ty == P::Object,
        W::SafeArray {
            element_vt,
            record_info: None,
        } => {
            ty == P::Variant
                && oxvba_runtime::safe_array::SafeArray::supports_intrinsic_element_vartype(
                    *element_vt,
                )
        }
        W::SafeArrayVariant => ty == P::Variant,
        W::ByRefSafeArray {
            element_vt,
            record_info: None,
        } => {
            !result
                && ty == P::Variant
                && oxvba_runtime::safe_array::SafeArray::supports_intrinsic_element_vartype(
                    *element_vt,
                )
        }
        W::ByRefSafeArrayVariant => !result && ty == P::Variant,
        _ => false,
    }
}
pub fn validate(member: &TypeLibMemberMetadata) -> bool {
    let count = member.parameter_types.len();
    member.callconv_is_stdcall
        && member.interface_iid.is_some_and(|iid| !iid.is_null())
        && (is_dispatch(member)
            || (member.is_dual
                && member.source_typekind == Some(SourceTypeKind::Interface)
                && member
                    .vtable_slot
                    .zip(member.vtable_slot_bound)
                    .is_some_and(|(slot, bound)| slot >= 7 && slot < bound && bound <= 4096)))
        && count <= MAX_NATIVE_PARAMETERS
        && member.parameter_wire_types.len() == count
        && member.parameter_optional.len() == count
        && member.parameter_optional_defaults.len() == count
        && member
            .parameter_types
            .iter()
            .zip(&member.parameter_wire_types)
            .enumerate()
            .all(|(i, (&ty, wire))| {
                wire_supported(ty, wire, false)
                    && (member.parameter_optional_defaults[i] != D::Lcid || ty == P::Long)
            })
        && match (member.return_type, &member.return_wire_type) {
            (None, None) => true,
            (Some(ty), Some(wire)) => !ty.is_by_ref() && wire_supported(ty, wire, true),
            _ => false,
        }
}
pub fn visible_indices(member: &TypeLibMemberMetadata) -> Vec<usize> {
    (0..member.parameter_types.len())
        .filter(|&i| member.parameter_optional_defaults.get(i) != Some(&D::Lcid))
        .collect()
}
pub fn is_by_ref(member: &TypeLibMemberMetadata, native_index: usize) -> bool {
    member.parameter_types[native_index].is_by_ref()
        || matches!(
            member.parameter_wire_types[native_index],
            W::ByRefSafeArray { .. } | W::ByRefSafeArrayVariant
        )
}
pub fn default_value(
    member: &TypeLibMemberMetadata,
    native_index: usize,
) -> Option<oxvba_runtime::Variant> {
    Some(
        match member.parameter_optional_defaults.get(native_index)? {
            D::HasDefault(value) => value.to_variant(),
            D::OptionalVariant => oxvba_runtime::Variant::from_error_code(0x8002_0004u32 as i32),
            D::Lcid => oxvba_runtime::Variant::from_i32(0),
            _ => return None,
        },
    )
}
/// Native parameters may include many omitted trailing optionals. The supplied
/// source arguments retain the ordinary typed entry bound; native defaults do not
/// inflate the procedure frame or introduce a Variant execution frame.
pub fn visible_signature(member: &TypeLibMemberMetadata, supplied: usize) -> Option<(Vec<T>, T)> {
    if !validate(member) || supplied > calcvba_host_api::MAX_TYPED_ARGS {
        return None;
    }
    let indices = visible_indices(member);
    if supplied > indices.len()
        || indices[supplied..]
            .iter()
            .any(|&i| default_value(member, i).is_none())
    {
        return None;
    }
    let parameters = indices[..supplied]
        .iter()
        .map(|&i| carrier(member.parameter_types[i]))
        .collect::<Option<Vec<_>>>()?;
    let result = member.return_type.map(carrier).unwrap_or(Some(T::Void))?;
    Some((parameters, result))
}
