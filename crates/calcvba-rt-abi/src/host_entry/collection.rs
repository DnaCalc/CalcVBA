//! Owning execution of the base library's canonical Collection state.
use calcvba_runtime::{ObjectRef, Variant, object_ref::*};
static DESCRIPTOR: RuntimeClassDescriptor = RuntimeClassDescriptor {
    name: "Collection",
    project_identity: None,
    predeclared: false,
    lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
    fields: &[],
    as_new_fields: &[],
    implements: &[],
    interfaces: &[RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR],
};
pub fn new_collection() -> ObjectRef {
    let object = ObjectRef::from_compat_identity_with_descriptor(i32::MIN, &DESCRIPTOR);
    object.with_native_collection(|_| ());
    object
}
pub fn invoke(
    object: &ObjectRef,
    name: &str,
    values: &[Variant],
    arguments: &[calcvba_host_api::object::DynamicArgument],
) -> Result<Variant, i32> {
    use calcvba_eval::collection::{
        CollectionError as E, CollectionMethod as M, dispatch_collection,
    };
    let library = calcvba_bundle::vba_library_bundle();
    let class = library.classes.first().ok_or(438)?;
    let method = class
        .methods
        .iter()
        .find(|m| m.name.eq_ignore_ascii_case(name))
        .ok_or(438)?;
    let native = library
        .procedures
        .get(method.proc)
        .and_then(|p| p.native)
        .ok_or(438)?;
    let kind = match native {
        calcvba_bundle::NativeBody::Method(calcvba_bundle::NativeMethodId::CollectionAdd) => M::Add,
        calcvba_bundle::NativeBody::Method(calcvba_bundle::NativeMethodId::CollectionItem) => {
            M::Item
        }
        calcvba_bundle::NativeBody::Method(calcvba_bundle::NativeMethodId::CollectionCount) => {
            M::Count
        }
        calcvba_bundle::NativeBody::Method(calcvba_bundle::NativeMethodId::CollectionRemove) => {
            M::Remove
        }
        _ => return Err(438),
    };
    // Collection is a base-library class; its argument names are language
    // semantics, independent of any host application or COM object library.
    let names: &[&str] = match kind {
        M::Add => &["Item", "Key", "Before", "After"],
        M::Item | M::Remove => &["Index"],
        M::Count => &[],
    };
    let missing = || Variant::from_error_code(0x80020004u32 as i32);
    let mut ordered = vec![missing(); names.len()];
    let mut occupied = vec![false; names.len()];
    for (position, (value, argument)) in values.iter().zip(arguments).enumerate() {
        let index = match argument.name.as_deref() {
            Some(name) => names
                .iter()
                .position(|n| n.eq_ignore_ascii_case(name))
                .ok_or(448)?,
            None => position,
        };
        if index >= names.len() {
            return Err(450);
        }
        if occupied[index] {
            return Err(448);
        }
        occupied[index] = true;
        ordered[index] = value.clone();
    }
    object
        .with_native_collection(|data| dispatch_collection(kind, data, &ordered))
        .ok_or(438)?
        .map_err(|e| match e {
            E::NotFound => 9,
            E::DuplicateKey => 457,
            E::BadArgument => 5,
            E::ArgNotOptional => 449,
        })
}
