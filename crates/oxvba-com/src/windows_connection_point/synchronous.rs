//! Synchronous, same-apartment incoming event arguments and Boolean copyback.
use super::*;
use crate::{TypeLibEventMetadata, TypeLibParamType, windows_client::OwnedComInterface};
use std::{
    cell::Cell,
    panic::{AssertUnwindSafe, catch_unwind},
    rc::Rc,
};

/// Native object arguments keep their own reference for this callback. No raw
/// ByRef pointer escapes Invoke; the Boolean cell is copied back synchronously.
pub struct IncomingEventObject(pub(crate) OwnedComInterface);

pub enum SynchronousEventArgument {
    Object(Option<IncomingEventObject>),
    Boolean(bool),
    BooleanByRef(Cell<bool>),
}
pub type SynchronousEventCallback = Rc<dyn Fn(&[SynchronousEventArgument]) -> Result<(), i32>>;
#[derive(Clone)]
pub struct SynchronousEventSinkConfig {
    metadata: TypeLibEventMetadata,
    callback: SynchronousEventCallback,
}
impl SynchronousEventSinkConfig {
    /// Only imported Object inputs and one incoming Boolean ByRef are admitted.
    /// # Safety
    /// metadata is the unchanged trusted source-interface descriptor selected
    /// by the compiler/host for the source's exact declared event interface.
    pub unsafe fn new(
        metadata: TypeLibEventMetadata,
        callback: SynchronousEventCallback,
    ) -> Result<Self, String> {
        if metadata.dispatch_path != crate::TypeLibEventDispatchPath::Dispatch
            || metadata.connection_point_iid.is_none()
            || metadata.dispatch_member_id.is_none()
            || metadata.parameter_types.len() != usize::from(metadata.callback_arity)
            || metadata.parameter_types.iter().any(|p| {
                !matches!(
                    p,
                    TypeLibParamType::Object
                        | TypeLibParamType::Boolean
                        | TypeLibParamType::ByRefBoolean
                )
            })
            || metadata
                .parameter_types
                .iter()
                .filter(|p| **p == TypeLibParamType::ByRefBoolean)
                .count()
                > 1
        {
            return Err("unsupported synchronous event signature".into());
        }
        Ok(Self { metadata, callback })
    }
}
/// # Safety
/// dispatch is retained on this apartment; config contains its trusted event metadata.
pub unsafe fn try_advise_synchronous_event_sink(
    dispatch: *mut RawIDispatch,
    config: SynchronousEventSinkConfig,
) -> Result<Option<WindowsConnectionPointTransport>, String> {
    let iid = config
        .metadata
        .connection_point_iid
        .clone()
        .expect("validated IID");
    let legacy = DispatchEventSinkConfig {
        event_dispatch_member: config
            .metadata
            .dispatch_member_id
            .expect("validated DISPID"),
        expected_arity: config.metadata.parameter_types.len(),
        connection_point_iid: None,
        on_event: Arc::new(|_| false),
    };
    // SAFETY: forward exact caller metadata and same-apartment source contract.
    unsafe { advise_dispatch_event_sink(dispatch, &iid, legacy, Some(config)) }
}
/// # Safety
/// rgvarg/raw indices have been validated against this exact event arity.
pub(super) unsafe fn invoke(
    config: &SynchronousEventSinkConfig,
    indices: &[usize],
    rgvarg: *mut VARIANT,
    arg_error: *mut u32,
) -> i32 {
    use windows_sys::Win32::System::Variant::{VT_BOOL, VT_BYREF, VT_DISPATCH, VT_UNKNOWN};
    let mut values = Vec::with_capacity(indices.len());
    let mut writeback = None;
    for (index, (&raw_index, ty)) in indices
        .iter()
        .zip(&config.metadata.parameter_types)
        .enumerate()
    {
        // SAFETY: common Invoke admission checked every raw index and array extent.
        let variant = unsafe { &*rgvarg.add(raw_index) };
        // SAFETY: live initialized native VARIANT tag.
        let tag = unsafe { variant.Anonymous.Anonymous.vt };
        let value = match ty {
            TypeLibParamType::Boolean if tag == VT_BOOL => {
                // SAFETY: the checked VT_BOOL tag selects the initialized i16 union member.
                SynchronousEventArgument::Boolean(unsafe {
                    variant.Anonymous.Anonymous.Anonymous.boolVal != 0
                })
            }
            TypeLibParamType::Object if matches!(tag, VT_DISPATCH | VT_UNKNOWN) => {
                // SAFETY: exact admitted native object union member.
                let raw = unsafe { variant.Anonymous.Anonymous.Anonymous.punkVal };
                if raw.is_null() {
                    SynchronousEventArgument::Object(None)
                } else {
                    // SAFETY: retains one dispatch reference from the live input.
                    let Some(pointer) = (unsafe { retain_dispatch_arg(variant) }) else {
                        return COM_DISP_E_TYPEMISMATCH;
                    };
                    // SAFETY: retain_dispatch_arg transferred one live native reference.
                    SynchronousEventArgument::Object(Some(IncomingEventObject(unsafe {
                        OwnedComInterface::from_owned(pointer.cast())
                    })))
                }
            }
            TypeLibParamType::ByRefBoolean if tag == (VT_BYREF | VT_BOOL) => {
                // SAFETY: exact BYREF BOOL tag selects one borrowed i16 pointer.
                let pointer = unsafe { variant.Anonymous.Anonymous.Anonymous.pboolVal };
                if pointer.is_null() {
                    return COM_E_INVALIDARG;
                }
                writeback = Some((index, pointer));
                // SAFETY: incoming Invoke owns this live Boolean cell until return.
                SynchronousEventArgument::BooleanByRef(Cell::new(unsafe { *pointer != 0 }))
            }
            _ => {
                if !arg_error.is_null() {
                    // SAFETY: Invoke supplies a writable optional argument-error cell.
                    unsafe {
                        *arg_error = raw_index as u32;
                    }
                }
                return COM_DISP_E_TYPEMISMATCH;
            }
        };
        values.push(value);
    }
    // Contain handler panics at the COM boundary. Mutations already made still
    // copy back, as on an ordinary VBA handler error.
    let result = catch_unwind(AssertUnwindSafe(|| (config.callback)(&values)));
    if let Some((index, pointer)) = writeback {
        let SynchronousEventArgument::BooleanByRef(value) = &values[index] else {
            unreachable!()
        };
        // SAFETY: the pointer never escaped this synchronous incoming Invoke.
        unsafe {
            *pointer = if value.get() { -1 } else { 0 };
        }
    }
    match result {
        Ok(Ok(())) => COM_S_OK,
        Ok(Err(_)) => windows_sys::Win32::Foundation::E_FAIL,
        Err(payload) => {
            if let Err(secondary) = catch_unwind(AssertUnwindSafe(|| drop(payload))) {
                std::mem::forget(secondary);
            }
            windows_sys::Win32::Foundation::E_FAIL
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows_sys::Win32::System::Variant::{VT_BOOL, VT_BYREF, VT_DISPATCH};

    fn metadata(parameters: Vec<TypeLibParamType>) -> TypeLibEventMetadata {
        TypeLibEventMetadata {
            parameter_iids: Vec::new(),
            name: "Cancel".into(),
            token: 77,
            callback_arity: parameters.len() as u8,
            parameter_types: parameters,
            dispatch_path: crate::TypeLibEventDispatchPath::Dispatch,
            connection_point_iid: Some(crate::IID_OXVBA_TEST_DISPATCH_EVENTS_STR.into()),
            dispatch_member_id: Some(77),
            coclass: Some("Fixture".into()),
        }
    }
    fn sink(
        parameters: Vec<TypeLibParamType>,
        callback: SynchronousEventCallback,
    ) -> *mut RawIDispatch {
        // SAFETY: these exact fixture descriptors declare the inputs supplied below.
        let config =
            unsafe { SynchronousEventSinkConfig::new(metadata(parameters), callback) }.unwrap();
        create_dispatch_event_sink(
            DispatchEventSinkConfig {
                event_dispatch_member: 77,
                expected_arity: config.metadata.parameter_types.len(),
                connection_point_iid: None,
                on_event: Arc::new(|_| false),
            },
            Some(config),
        )
        .cast()
    }
    // SAFETY CONTRACT: caller retains the sink and initialized input VARIANT array.
    unsafe fn call(sink: *mut RawIDispatch, arguments: &mut [VARIANT]) -> i32 {
        let mut params = DISPPARAMS {
            rgvarg: arguments.as_mut_ptr(),
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: arguments.len() as u32,
            cNamedArgs: 0,
        };
        // SAFETY: exact fixture DISPID and bounded synchronous input arrays.
        unsafe {
            ((*(*sink).vtbl).invoke)(
                sink.cast(),
                77,
                &crate::IID_NULL,
                0,
                1,
                &mut params,
                std::ptr::null_mut(),
                std::ptr::null_mut(),
                std::ptr::null_mut(),
            )
        }
    }
    // SAFETY CONTRACT: caller transfers the one factory-created sink reference.
    unsafe fn release(sink: *mut RawIDispatch) {
        // SAFETY: balances the one private sink factory allocation.
        assert_eq!(unsafe { ((*(*sink).vtbl).unknown.release)(sink.cast()) }, 0);
    }
    #[test]
    fn incoming_boolean_copyback_precedes_return_and_objects_balance_on_error() {
        let calls = Rc::new(Cell::new(0));
        let observed = calls.clone();
        let sink = sink(
            vec![TypeLibParamType::Object, TypeLibParamType::ByRefBoolean],
            Rc::new(move |args| {
                let [
                    SynchronousEventArgument::Object(Some(_)),
                    SynchronousEventArgument::BooleanByRef(cancel),
                ] = args
                else {
                    panic!("wrong argument order")
                };
                observed.set(observed.get() + 1);
                cancel.set(true);
                Err(5)
            }),
        );
        let object = crate::create_oxvba_test_dispatch();
        let mut cancel = 0i16;
        // SAFETY: zeroed VARIANT is VT_EMPTY; exact union fields are initialized before Invoke.
        let mut arguments: [VARIANT; 2] = unsafe { std::mem::zeroed() };
        arguments[0].Anonymous.Anonymous.vt = VT_BYREF | VT_BOOL;
        arguments[0].Anonymous.Anonymous.Anonymous.pboolVal = &mut cancel;
        arguments[1].Anonymous.Anonymous.vt = VT_DISPATCH;
        arguments[1].Anonymous.Anonymous.Anonymous.pdispVal = object.cast();
        // SAFETY: both borrowed input cells and the two native objects stay live.
        let actual = unsafe { call(sink, &mut arguments) };
        assert_eq!(actual, windows_sys::Win32::Foundation::E_FAIL);
        assert_eq!(cancel, -1);
        assert_eq!(calls.get(), 1);
        // SAFETY: the temporary incoming object owner already dropped on error.
        let actual = unsafe { ((*(*object).vtbl).unknown.release)(object.cast()) };
        assert_eq!(actual, 0);
        // SAFETY: final ownership transfer to the private sink releaser.
        unsafe {
            release(sink);
        }
    }
    #[test]
    fn invalid_shape_and_wrong_thread_refuse_before_handler_effects() {
        let calls = Rc::new(Cell::new(0));
        let observed = calls.clone();
        let sink = sink(
            vec![TypeLibParamType::ByRefBoolean],
            Rc::new(move |_| {
                observed.set(observed.get() + 1);
                Ok(())
            }),
        );
        // SAFETY: an initialized ByVal Boolean intentionally fails the ByRef contract.
        let mut arguments: [VARIANT; 1] = unsafe { std::mem::zeroed() };
        arguments[0].Anonymous.Anonymous.vt = VT_BOOL;
        // SAFETY: this valid input cell tests admission, without a writable pointer.
        let actual = unsafe { call(sink, &mut arguments) };
        assert_eq!(actual, COM_DISP_E_TYPEMISMATCH);
        let address = sink as usize;
        let result = std::thread::spawn(move || {
            // SAFETY: parent retains the immutable sink until join. The thread check
            // must reject this call before inspecting any input or running a handler.
            unsafe { call(address as *mut RawIDispatch, &mut []) }
        })
        .join()
        .unwrap();
        assert_eq!(result, 0x8001_010eu32 as i32);
        assert_eq!(calls.get(), 0);
        // SAFETY: join finished and the parent still owns this exact reference.
        unsafe {
            release(sink);
        }
    }
    #[test]
    fn panic_stays_inside_invoke_and_preserves_prior_boolean_mutation() {
        let sink = sink(
            vec![TypeLibParamType::ByRefBoolean],
            Rc::new(|args| {
                let [SynchronousEventArgument::BooleanByRef(cancel)] = args else {
                    unreachable!()
                };
                cancel.set(true);
                panic!("owned test panic");
            }),
        );
        let mut cancel = 0i16;
        // SAFETY: exact incoming initialized ByRef Boolean cell.
        let mut arguments: [VARIANT; 1] = unsafe { std::mem::zeroed() };
        arguments[0].Anonymous.Anonymous.vt = VT_BYREF | VT_BOOL;
        arguments[0].Anonymous.Anonymous.Anonymous.pboolVal = &mut cancel;
        // SAFETY: all cells and the private native sink remain live through return.
        let actual = unsafe { call(sink, &mut arguments) };
        assert_eq!(actual, windows_sys::Win32::Foundation::E_FAIL);
        assert_eq!(cancel, -1);
        // SAFETY: the panic was contained and ownership stays with this test.
        unsafe {
            release(sink);
        }
    }
    #[test]
    fn handler_can_release_the_subscription_owner_before_boolean_copyback() {
        let address = Rc::new(Cell::new(0));
        let callback_address = address.clone();
        let sink = sink(
            vec![TypeLibParamType::ByRefBoolean],
            Rc::new(move |args| {
                let [SynchronousEventArgument::BooleanByRef(cancel)] = args else {
                    unreachable!()
                };
                let source = callback_address.get() as *mut RawIDispatch;
                // SAFETY: this callback transfers the factory owner's reference,
                // analogous to Unadvise removing the last external sink owner.
                let remaining = unsafe { ((*(*source).vtbl).unknown.release)(source.cast()) };
                assert_eq!(remaining, 1);
                cancel.set(true);
                Ok(())
            }),
        );
        address.set(sink as usize);
        let mut cancel = 0i16;
        // SAFETY: initialized native Boolean cell retained until Invoke returns.
        let mut arguments: [VARIANT; 1] = unsafe { std::mem::zeroed() };
        arguments[0].Anonymous.Anonymous.vt = VT_BYREF | VT_BOOL;
        arguments[0].Anonymous.Anonymous.Anonymous.pboolVal = &mut cancel;
        // SAFETY: the incoming Invoke acquires its own lease before running VBA.
        assert_eq!(unsafe { call(sink, &mut arguments) }, COM_S_OK);
        assert_eq!(cancel, -1);
        // The callback released the factory owner and Invoke released its lease.
    }
}
