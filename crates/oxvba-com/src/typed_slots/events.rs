//! Owned native source subscriptions; all member/event metadata is preselected.
use super::factory::{NativeReceiver, make_native_object};
use crate::windows_connection_point::{
    SynchronousEventArgument, SynchronousEventCallback, SynchronousEventSinkConfig,
    try_advise_synchronous_event_sink,
};
use crate::{TypeLibEventMetadata, WindowsConnectionPointTransport, unadvise_connection_point};
use calcvba_host_api::{HostObject, InterfaceId};

pub struct TypedEventSubscription {
    transport: Option<WindowsConnectionPointTransport>,
    _source: HostObject,
}
impl TypedEventSubscription {
    pub fn close(mut self) -> Result<(), String> {
        // SAFETY: this guard owns the unique advise transport on this apartment.
        unsafe { unadvise_connection_point(self.transport.take().expect("live subscription")) }
    }
}
impl Drop for TypedEventSubscription {
    fn drop(&mut self) {
        if let Some(transport) = self.transport.take() {
            // SAFETY: this guard owns the one final Unadvise/Release operation.
            let _ = unsafe { unadvise_connection_point(transport) };
        }
    }
}
/// # Safety
/// metadata is the unchanged trusted imported event descriptor for this exact
/// source interface. Subscribe/close and all incoming calls stay on this apartment.
pub unsafe fn advise_typed_event(
    source: &HostObject,
    metadata: TypeLibEventMetadata,
    callback: SynchronousEventCallback,
) -> Result<TypedEventSubscription, String> {
    let owner = source
        .payload()
        .and_then(|p| p.downcast_ref::<NativeReceiver>())
        .ok_or("not a retained native source")?;
    // SAFETY: caller supplies unchanged trusted event metadata, checked below.
    let config = unsafe { SynchronousEventSinkConfig::new(metadata, callback) }?;
    // SAFETY: the native owner retains the source through this entire subscription.
    let transport =
        unsafe { try_advise_synchronous_event_sink(owner.pointer.as_ptr().cast(), config) }?
            .ok_or("source has no selected connection point")?;
    Ok(TypedEventSubscription {
        transport: Some(transport),
        _source: source.clone(),
    })
}
/// Resolve an incoming object only against the source's retained nominal library.
/// This performs QI for the exact declared IID; it performs no type/name discovery.
pub fn prepare_typed_event_object(
    source: &HostObject,
    argument: &SynchronousEventArgument,
    iid: InterfaceId,
) -> Result<Option<HostObject>, String> {
    let owner = source
        .payload()
        .and_then(|p| p.downcast_ref::<NativeReceiver>())
        .ok_or("not a retained native source")?;
    let SynchronousEventArgument::Object(pointer) = argument else {
        return Err("event argument is not an object".into());
    };
    let Some(pointer) = pointer else {
        return Ok(None);
    };
    let index = owner
        .library
        .metadata
        .iter()
        .position(|m| {
            m.iid.is_some_and(|native| {
                let mut bytes = [0; 16];
                bytes[..4].copy_from_slice(&native.data1.to_le_bytes());
                bytes[4..6].copy_from_slice(&native.data2.to_le_bytes());
                bytes[6..8].copy_from_slice(&native.data3.to_le_bytes());
                bytes[8..].copy_from_slice(&native.data4);
                bytes == iid
            })
        })
        .ok_or("event target IID is absent from retained library")?;
    // SAFETY: same-apartment input owns its native reference; exact retained IID.
    unsafe {
        make_native_object(
            owner.bridge.clone(),
            pointer.0.as_ptr(),
            Some(index),
            owner.library.clone(),
        )
    }
    .map(Some)
}
