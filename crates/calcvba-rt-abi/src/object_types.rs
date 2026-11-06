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
