//! Linux host services for native VBA code and its portable value operations.
use super::execution::Frame;
pub(super) use super::values::{read, write};
use calcvba_hal::{HalProfileId, HostPolicy, HostServices, adapters::builder::HostBuilder};
use calcvba_rt_abi::{ExecState, native_export::Context};
use calcvba_runtime::ObjectRef;
use std::{ptr, sync::Arc};

pub(super) struct Runtime {
    pub(super) exec: Box<ExecState<'static>>,
    _host: Arc<dyn HostServices>,
    pub(super) frames: Vec<Frame>,
    pub(super) members: super::members::Calls,
    pub(super) classes: super::classes::Classes,
}

impl Runtime {
    pub(super) fn end_frame(&mut self, _depth: u32) {}
}

pub(super) unsafe fn runtime(context: *mut Context) -> *mut Runtime {
    // SAFETY: this context and the allocated Runtime stay on the owning thread.
    unsafe {
        if (*context).services.is_null() {
            let host = HostBuilder::new()
                .profile(HalProfileId::Linux)
                .policy(HostPolicy::native())
                .build();
            // The Arc allocation stays alive until after the borrowing exec is dropped.
            let exec = Box::new(ExecState::new(&*Arc::as_ptr(&host)));
            let runtime = Runtime {
                exec,
                _host: host,
                frames: Vec::new(),
                members: Default::default(),
                classes: Default::default(),
            };
            (*context).services = Box::into_raw(Box::new(runtime)).cast();
        }
        (*context).services.cast()
    }
}

pub(super) unsafe fn shutdown(context: *mut Context) {
    // SAFETY: generated shutdown checked thread ownership and active depth.
    unsafe {
        let state = std::mem::replace(&mut (*context).services, ptr::null_mut());
        if !state.is_null() {
            drop(Box::from_raw(state.cast::<Runtime>()));
        }
    }
}

pub(super) unsafe fn host(context: *mut Context) -> Arc<dyn HostServices> {
    // SAFETY: only the owning thread can access services.
    unsafe { (*runtime(context))._host.clone() }
}

pub(super) unsafe fn foreign_type_matches(
    context: *mut Context,
    object: ObjectRef,
    name: &str,
    iid: Option<[u8; 16]>,
) -> Result<bool, i32> {
    // SAFETY: only the owning thread can access services.
    let host = unsafe { host(context) };
    calcvba_rt_abi::object_types::matches_type(host.as_ref(), &object, name, iid)
        .map_err(|fault| fault.code)
}
