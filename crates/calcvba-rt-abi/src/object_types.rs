//! Nominal object checks shared by the interpreters and compiled entries.
use crate::Fault;
use calcvba_runtime::ObjectRef;

pub fn matches_type(
    host: &dyn calcvba_hal::HostServices,
    object: &ObjectRef,
    name: &str,
    iid: Option<[u8; 16]>,
) -> Result<bool, Fault> {
    let bare = name.rsplit('.').next().unwrap_or(name);
    // Every object is an `Object` (`TypeOf kk Is Object` on a project instance or a
    // Collection is True).
    if bare.eq_ignore_ascii_case("Object") && !object.is_nothing_placeholder() {
        return Ok(true);
    }
    if object.is_project_instance() || object.native_collection_snapshot().is_some() {
        let descriptor = object.class_descriptor();
        if let Some((unit, _)) = name.rsplit_once('.')
            && let Some(identity) = descriptor.project_identity
        {
            return Ok(identity.unit_name.eq_ignore_ascii_case(unit)
                && descriptor.name.eq_ignore_ascii_case(bare)
                || descriptor.implements.iter().any(|implemented| {
                    implemented.eq_ignore_ascii_case(name)
                        || identity.unit_name.eq_ignore_ascii_case(unit)
                            && implemented.eq_ignore_ascii_case(bare)
                }));
        }
        return Ok(descriptor.name.eq_ignore_ascii_case(bare)
            || descriptor
                .implements
                .iter()
                .any(|name| name.eq_ignore_ascii_case(bare)));
    }
    if let Some(prepared) = calcvba_host_api::HostObject::from_object_ref(object)
        && let Some(iid) = iid
    {
        return Ok(prepared.resolve_interface(iid).is_ok());
    }
    #[cfg(all(windows, target_arch = "x86_64"))]
    if let Some(iid) = iid {
        let pointer = if object.is_compat_object() {
            host.com()
                .retain_native_dispatch_object(object.clone())
                .map_err(Fault::from_hal)?
        } else {
            object.clone().into_raw_iunknown().cast()
        };
        // SAFETY: both routes transfer one same-apartment IUnknown reference.
        let _owner = unsafe { ObjectRef::from_raw_iunknown_owned(pointer.cast()) };
        let guid = calcvba_com::ComInterfaceIid {
            data1: u32::from_le_bytes(iid[..4].try_into().unwrap()),
            data2: u16::from_le_bytes(iid[4..6].try_into().unwrap()),
            data3: u16::from_le_bytes(iid[6..8].try_into().unwrap()),
            data4: iid[8..].try_into().unwrap(),
        }
        .to_guid();
        // SAFETY: retained native owner and an imported interface identity.
        return Ok(
            match unsafe { calcvba_com::query_interface_pointer(pointer, &guid) } {
                Ok(pointer) => {
                    // SAFETY: successful QueryInterface transfers one owned reference.
                    drop(unsafe { ObjectRef::from_raw_iunknown_owned(pointer.cast()) });
                    true
                }
                Err(_) => false,
            },
        );
    }
    Ok(host
        .com()
        .object_type_name(object.clone())
        .map_err(Fault::from_hal)?
        .is_some_and(|actual| {
            actual.eq_ignore_ascii_case(name) || actual.eq_ignore_ascii_case(bare)
        }))
}

/// The route key of a built-in `Collection` instance.
const COLLECTION_ROUTE_KEY: i32 = i32::MIN;

/// `Set target = object` where the target is declared `As target_type_name`: true when the
/// object cannot be stored there (error 13). An `Object` or `Variant` target, and a type this
/// check cannot see (a COM interface), takes any object. A project class or interface target
/// (declared in any of `programs`) takes only an instance of that class or of one implementing
/// it; a `Collection` target takes only a Collection. `source_program` is the program of the
/// object's own bundle, which describes a project instance's class.
pub fn set_type_mismatch<'p>(
    object: &ObjectRef,
    target_type_name: &str,
    source_program: Option<&calcvba_oxir::OxProgram>,
    programs: impl IntoIterator<Item = &'p calcvba_oxir::OxProgram>,
) -> bool {
    let bare = target_type_name
        .rsplit('.')
        .next()
        .unwrap_or(target_type_name);
    if bare.is_empty()
        || bare.eq_ignore_ascii_case("Object")
        || bare.eq_ignore_ascii_case("Variant")
        || object.is_nothing_placeholder()
    {
        return false;
    }
    let declares = |program: &calcvba_oxir::OxProgram| {
        program.classes.iter().any(|class| {
            class.name.eq_ignore_ascii_case(bare)
                || class
                    .implements
                    .iter()
                    .any(|i| i.eq_ignore_ascii_case(bare))
        })
    };
    let target_is_project =
        source_program.is_some_and(declares) || programs.into_iter().any(declares);
    if let Some(required) = calcvba_lib::objects::class_name(target_type_name)
        && !target_is_project
    {
        return !calcvba_lib::objects::is_object(object)
            || !object
                .class_descriptor()
                .name
                .eq_ignore_ascii_case(required);
    }
    if calcvba_lib::objects::is_object(object) {
        return target_is_project || bare.eq_ignore_ascii_case("Collection");
    }
    let is_collection = object.is_compat_object() && object.route_key() == COLLECTION_ROUTE_KEY;
    if is_collection {
        return !bare.eq_ignore_ascii_case("Collection") && target_is_project;
    }
    if object.is_project_instance() {
        let class = source_program.and_then(|p| p.classes.get(object.route_key() as usize));
        return match class {
            Some(class) => {
                let compatible = class.name.eq_ignore_ascii_case(bare)
                    || class
                        .implements
                        .iter()
                        .any(|i| i.eq_ignore_ascii_case(bare));
                !compatible && (target_is_project || bare.eq_ignore_ascii_case("Collection"))
            }
            None => false,
        };
    }
    // A foreign object is never a project class or a Collection.
    target_is_project || bare.eq_ignore_ascii_case("Collection")
}

/// Resolve local creation of the engine-provided Dictionary class. The public runtime name
/// marks the injected implementation; a user's class with the same source name is not one.
pub fn builtin_dictionary_class(
    program: &calcvba_oxir::OxProgram,
    id: calcvba_bundle::NativeImplId,
    args: &[calcvba_runtime::Variant],
) -> Option<usize> {
    use calcvba_runtime::VarType;
    let present = |index: usize| {
        args.get(index)
            .filter(|value| value.as_error_code() != Some(0x8002_0004u32 as i32))
    };
    let empty_string =
        |value: &calcvba_runtime::Variant| value.as_bstr().is_some_and(|text| text.is_empty());
    let class = match id {
        calcvba_bundle::NativeImplId::CreateObject
            if present(1)
                .is_none_or(|server| server.vtype() == VarType::Empty || empty_string(server)) =>
        {
            present(0)?
        }
        calcvba_bundle::NativeImplId::GetObject if present(0).is_some_and(empty_string) => {
            present(1)?
        }
        _ => return None,
    };
    if !class
        .as_bstr()?
        .as_str()
        .eq_ignore_ascii_case("Scripting.Dictionary")
    {
        return None;
    }
    program
        .classes
        .iter()
        .position(|class| class.runtime_type_name.as_deref() == Some("Dictionary"))
}

/// A zero-argument value read may use only a default getter whose parameters are optional.
/// In particular Dictionary.Item(Key) is not a default value for VarType.
pub fn has_parameterless_default_get(program: &calcvba_oxir::OxProgram, class: usize) -> bool {
    program.classes.get(class).is_some_and(|class| {
        class.methods.iter().any(|method| {
            method.is_default_member
                && matches!(
                    method.kind,
                    calcvba_bundle::ProjectMemberKind::PropertyGet
                        | calcvba_bundle::ProjectMemberKind::Method
                )
                && program.funcs.get(method.proc.0).is_some_and(|function| {
                    function
                        .locals
                        .iter()
                        .take(function.param_count)
                        .skip(crate::hidden_me_receiver_param_count(function))
                        .all(|local| {
                            local
                                .param
                                .as_ref()
                                .is_some_and(|param| param.optional || param.variadic)
                        })
                })
        })
    })
}
