//! Owning execution of the base library's canonical Collection state.
use calcvba_runtime::{ObjectRef, Variant, object_ref::*};
static DESCRIPTOR: RuntimeClassDescriptor = RuntimeClassDescriptor {
    name: "Collection",
    runtime_type_name: None,
    project_identity: None,
    default_interface_iid: Some(calcvba_runtime::com::IID_VBA_COLLECTION),
    source_interface_iid: None,
    predeclared: false,
    lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
    record_members: &[],
    project_records: &[],
    fields: &[],
    as_new_fields: &[],
    implements: &[],
    interfaces: &[RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR],
    events: &[],
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
    let names: Vec<Option<String>> = arguments.iter().map(|a| a.name.clone()).collect();
    let ordered = calcvba_eval::collection::order_collection_args(kind, values, &names)?;
    object
        .with_native_collection(|data| dispatch_collection(kind, data, &ordered))
        .ok_or(438)?
        .map_err(E::vba_number)
}
