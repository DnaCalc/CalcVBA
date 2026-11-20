//! Portable built-in library objects; dispatch stays on the owning runtime thread.
use crate::{LibError, LibResult};
use calcvba_bundle::ProjectMemberKind;
use calcvba_runtime::ObjectRefExt;
use calcvba_runtime::{
    ObjectRef, Variant,
    object_ref::{COMPAT_OBJECT_CLASS_DESCRIPTOR, RuntimeClassDescriptor},
    variant::VarType,
};

const ERR: i32 = i32::MIN + 6;
const USER_FORMS: i32 = i32::MIN + 5;
static USER_FORMS_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
    name: "Object",
    ..COMPAT_OBJECT_CLASS_DESCRIPTOR
};

pub fn class_name(name: &str) -> Option<&'static str> {
    crate::regexp::class_name(name).or_else(|| {
        let bare = match name.split_once('.') {
            Some((unit, bare)) if unit.eq_ignore_ascii_case("VBA") => bare,
            Some(_) => return None,
            None => name,
        };
        bare.eq_ignore_ascii_case("ErrObject")
            .then_some("ErrObject")
    })
}

pub fn new_err() -> Variant {
    static DESCRIPTOR: RuntimeClassDescriptor = RuntimeClassDescriptor {
        name: "ErrObject",
        ..COMPAT_OBJECT_CLASS_DESCRIPTOR
    };
    Variant::from_object_ref(ObjectRef::from_project_instance(
        ERR,
        ERR,
        0,
        false,
        &DESCRIPTOR,
    ))
}

pub fn is_err(object: &ObjectRef) -> bool {
    object.is_project_instance() && object.route_key() == ERR
}

pub fn new_user_forms() -> Variant {
    let object =
        ObjectRef::from_project_instance(USER_FORMS, USER_FORMS, 0, false, &USER_FORMS_CLASS);
    object.with_native_collection(|_| ());
    Variant::from_object_ref(object)
}

pub fn is_object(object: &ObjectRef) -> bool {
    crate::regexp::is_object(object)
        || is_err(object)
        || (object.is_project_instance() && object.route_key() == USER_FORMS)
}

pub fn dispatch(
    object: &ObjectRef,
    name: &str,
    kind: ProjectMemberKind,
    args: &[Variant],
    names: &[Option<String>],
) -> LibResult<Variant> {
    if crate::regexp::is_object(object) {
        return crate::regexp::dispatch(object, name, kind, args, names);
    }
    if !object.is_project_instance() || object.route_key() != USER_FORMS {
        return Err(LibError::new(
            438,
            "Object doesn't support this property or method",
        ));
    }
    if matches!(
        kind,
        ProjectMemberKind::PropertyLet | ProjectMemberKind::PropertySet
    ) {
        return Err(LibError::new(
            438,
            "Object doesn't support this property or method",
        ));
    }
    match name.to_ascii_lowercase().as_str() {
        "count" if args.is_empty() => Ok(Variant::from_i16(0)),
        "_newenum" if args.is_empty() => Ok(Variant::from_object_ref(object.clone())),
        "" | "item" => Err(LibError::new(9, "Subscript out of range")),
        _ => Err(LibError::new(
            438,
            "Object doesn't support this property or method",
        )),
    }
}

// No forms host is installed. Non-form instances cannot be loaded or unloaded.
pub fn load_unload(args: &[Variant]) -> LibResult<Variant> {
    let value = crate::need(args, 0)?;
    if value.vtype() != VarType::Object {
        return Err(LibError::new(424, "Object required"));
    }
    if value.as_object_ref().is_none() {
        return Err(LibError::new(
            91,
            "Object variable or With block variable not set",
        ));
    }
    Err(LibError::new(361, "Can't load or unload this object"))
}
