use crate::Fault;
use calcvba_runtime::{
    ObjectRef, ObjectRefExt, RuntimeClassDescriptor, RuntimeMemberDescriptor,
    RuntimeMemberInvokeKind, VarType, Variant,
};

/// Reserved project object route for VBA's built-in Collection.
pub const VBA_COLLECTION_ROUTE_KEY: i32 = i32::MIN;

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ObjectMemberRoute {
    Library,
    Collection,
    Project,
    Foreign,
}

pub fn object_member_route(object: &ObjectRef) -> ObjectMemberRoute {
    if calcvba_lib::objects::is_object(object) {
        ObjectMemberRoute::Library
    } else if object.route_key() == VBA_COLLECTION_ROUTE_KEY {
        ObjectMemberRoute::Collection
    } else if object.is_project_instance() {
        ObjectMemberRoute::Project
    } else {
        ObjectMemberRoute::Foreign
    }
}

pub fn object_receiver(value: &Variant) -> Result<ObjectRef, Fault> {
    value.as_object_ref().ok_or_else(|| {
        Fault::standard(if value.vtype() == VarType::Object {
            91
        } else {
            424
        })
    })
}

/// Collection's default member is Item, which requires an index.
pub fn collection_member_name(name: &str, args_empty: bool) -> Result<&str, Fault> {
    match (name.is_empty(), args_empty) {
        (true, true) => Err(Fault::standard(450)),
        (true, false) => Ok("Item"),
        (false, _) => Ok(name),
    }
}

pub fn resolve_project_default_member(
    descriptor: &RuntimeClassDescriptor,
    kind: RuntimeMemberInvokeKind,
    args_empty: bool,
    public_only: bool,
) -> Result<&'static RuntimeMemberDescriptor, Fault> {
    let member = descriptor
        .resolve_member_by_dispid(0, kind, args_empty, public_only)
        .ok_or_else(|| Fault::standard(438))?;
    if args_empty && kind == RuntimeMemberInvokeKind::PropertyGet && !takes_no_arguments(member) {
        return Err(Fault::standard(449));
    }
    Ok(member)
}

pub(crate) fn takes_no_arguments(member: &RuntimeMemberDescriptor) -> bool {
    member.params.iter().all(|p| p.optional || p.param_array)
}

/// Arithmetic and conversions follow object defaults until a non-object value
/// is found. A single explicit Let value read uses one dispatch instead.
pub fn read_default_value<E>(
    mut value: Variant,
    mut invoke: impl FnMut(Variant) -> Result<Variant, E>,
    error: impl FnOnce(Fault) -> E,
) -> Result<Variant, E> {
    for _ in 0..64 {
        if value.vtype() != VarType::Object {
            return Ok(value);
        }
        value = invoke(value)?;
    }
    Err(error(Fault::standard(28)))
}

pub fn project_has_default_get(value: &Variant) -> bool {
    value.as_object_ref().is_some_and(|object| {
        object.is_project_instance()
            && crate::object_types::has_parameterless_default_get_descriptor(
                object.class_descriptor(),
            )
    })
}
