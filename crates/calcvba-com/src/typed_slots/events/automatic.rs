//! Bridge a same-thread owning runtime callback to the native event sink.
use super::*;
use calcvba_host_api::{
    Value,
    object::{EventCallback, EventDescriptor, EventParameter, EventSubscription},
};
use std::rc::Rc;
pub(in crate::typed_slots) fn subscribe(
    owner: Rc<NativeReceiver>,
    descriptor: &EventDescriptor,
    callback: EventCallback,
) -> Result<EventSubscription, i32> {
    let native_iid = crate::ComInterfaceIid {
        data1: u32::from_le_bytes(descriptor.interface[..4].try_into().map_err(|_| 13)?),
        data2: u16::from_le_bytes(descriptor.interface[4..6].try_into().map_err(|_| 13)?),
        data3: u16::from_le_bytes(descriptor.interface[6..8].try_into().map_err(|_| 13)?),
        data4: descriptor.interface[8..].try_into().map_err(|_| 13)?,
    };
    let metadata = TypeLibEventMetadata {
        parameter_iids: Vec::new(),
        name: descriptor.name.clone(),
        token: descriptor.dispatch_id,
        callback_arity: u8::try_from(descriptor.parameters.len()).map_err(|_| 450)?,
        parameter_types: descriptor
            .parameters
            .iter()
            .map(|p| match p {
                EventParameter::Object(_) => crate::TypeLibParamType::Object,
                EventParameter::Boolean => crate::TypeLibParamType::Boolean,
                EventParameter::BooleanByRef => crate::TypeLibParamType::ByRefBoolean,
            })
            .collect(),
        dispatch_path: crate::TypeLibEventDispatchPath::Dispatch,
        connection_point_iid: Some(format!(
            "{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}",
            native_iid.data1,
            native_iid.data2,
            native_iid.data3,
            native_iid.data4[0],
            native_iid.data4[1],
            native_iid.data4[2],
            native_iid.data4[3],
            native_iid.data4[4],
            native_iid.data4[5],
            native_iid.data4[6],
            native_iid.data4[7]
        )),
        dispatch_member_id: Some(descriptor.dispatch_id),
        coclass: None,
    };
    // SAFETY: same-apartment native source is retained by owner; an opaque view
    // exposes no invented member plans and independently retains COM identity.
    let source = unsafe {
        make_native_object(
            owner.bridge.clone(),
            owner.pointer.as_ptr(),
            None,
            owner.library.clone(),
        )
    }
    .map_err(|_| 13)?;
    let parameters = descriptor.parameters.clone();
    let receive = Rc::new(move |arguments: &[SynchronousEventArgument]| {
        if arguments.len() != parameters.len() {
            return Err(450);
        }
        let mut values = Vec::with_capacity(arguments.len());
        for (argument, parameter) in arguments.iter().zip(&parameters) {
            values.push(match (argument, parameter) {
                (SynchronousEventArgument::Boolean(value), EventParameter::Boolean) => {
                    Value::Boolean(*value)
                }
                (SynchronousEventArgument::BooleanByRef(value), EventParameter::BooleanByRef) => {
                    Value::Boolean(value.get())
                }
                (SynchronousEventArgument::Object(pointer), EventParameter::Object(iid)) => {
                    let object = pointer
                        .as_ref()
                        .map(|pointer| {
                            let index = if let Some(iid) = iid {
                                Some(
                                    owner
                                        .library
                                        .metadata
                                        .iter()
                                        .position(|m| {
                                            m.iid.is_some_and(|native| {
                                                let mut bytes = [0; 16];
                                                bytes[..4]
                                                    .copy_from_slice(&native.data1.to_le_bytes());
                                                bytes[4..6]
                                                    .copy_from_slice(&native.data2.to_le_bytes());
                                                bytes[6..8]
                                                    .copy_from_slice(&native.data3.to_le_bytes());
                                                bytes[8..].copy_from_slice(&native.data4);
                                                bytes == *iid
                                            })
                                        })
                                        .ok_or(13)?,
                                )
                            } else {
                                None
                            };
                            // SAFETY: incoming argument retains its reference until callback return;
                            // the exact nominal view is selected from the retained library.
                            unsafe {
                                make_native_object(
                                    owner.bridge.clone(),
                                    pointer.0.as_ptr(),
                                    index,
                                    owner.library.clone(),
                                )
                            }
                            .map_err(|_| 13)
                        })
                        .transpose()?;
                    Value::Object(object)
                }
                _ => return Err(13),
            });
        }
        let result = callback(&mut values);
        for (argument, value) in arguments.iter().zip(&values) {
            if let (SynchronousEventArgument::BooleanByRef(target), Value::Boolean(value)) =
                (argument, value)
            {
                target.set(*value);
            }
        }
        result
    });
    // SAFETY: only checked IDispatch argument shapes are admitted; source IID,
    // DISPID and signature come from the compiler-owned event descriptor.
    let native = unsafe { advise_typed_event(&source, metadata, receive) }.map_err(|_| 438)?;
    Ok(EventSubscription::new(native))
}
