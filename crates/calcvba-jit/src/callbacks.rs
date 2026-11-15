//! `AddressOf` in the JIT (design doc, step 9; VM3's `prepare_native_callback_args` and
//! `run_native_callback_proc`): a procedure passed to a `ByVal LongPtr` parameter of a synchronous
//! callback API becomes a thunk address registered for the call, and a call through the thunk
//! runs the procedure on this run, with its raw arguments read as the parameters' types.
use super::*;
use crate::helpers::after_jit_frame_pop;
use calcvba_runtime::{CallbackExecutor, CallbackRegistration, register_callback};
use std::ptr::NonNull;

/// Runs `AddressOf` procedures on a JIT run.
pub(crate) struct JitCallbackExecutor {
    state: *mut RawExecState,
    run: *mut JitRun,
}

impl CallbackExecutor for JitCallbackExecutor {
    fn invoke_callback(&mut self, proc_token: usize, args: &[isize]) -> isize {
        // SAFETY: the executor is registered only while its Declare call runs on this thread,
        // and the JIT's helpers hold no borrow of the run or state across that call.
        unsafe { run_callback(self.state, self.run, proc_token, args) }
    }
}

/// The callback registrations of one Declare call, with their executor; dropping it ends them.
pub(crate) struct JitCallbacks {
    _registrations: Vec<CallbackRegistration>,
    _executor: Box<JitCallbackExecutor>,
}

/// A synchronous callback API: the callback runs inside the call (VM3's admission).
pub(crate) fn is_synchronous_callback_descriptor(
    descriptor: &calcvba_bundle::ExternalCallDescriptor,
) -> bool {
    let library = descriptor.library.to_ascii_lowercase();
    let symbol = descriptor.alias.to_ascii_lowercase();
    matches!(library.as_str(), "user32" | "user32.dll")
        && matches!(
            symbol.as_str(),
            "callwindowprocw" | "callwindowproca" | "callwindowproc"
        )
}

/// Replaces each `AddressOf` value among `cells` with a thunk address registered for this call.
/// `Err` carries the fault to raise (7 when the thunks are exhausted, 5 for an unsupported use).
///
/// # Safety
/// `state` and `run` are the live run's roots for the duration of the Declare call the
/// registrations serve.
pub(crate) unsafe fn register_callbacks(
    state: *mut RawExecState,
    run: *mut JitRun,
    descriptor: &calcvba_bundle::ExternalCallDescriptor,
    cells: &mut [Option<Variant>],
) -> Result<JitCallbacks, Fault> {
    let mut executor = Box::new(JitCallbackExecutor { state, run });
    let pointer = NonNull::from(&mut *executor);
    let mut registrations = Vec::new();
    for (index, cell) in cells.iter_mut().enumerate() {
        let Some(value) = cell.as_mut() else {
            continue;
        };
        let Some(proc_token) = value.as_proc_ref() else {
            continue;
        };
        let long_ptr =
            descriptor.param_types.get(index) == Some(&calcvba_bundle::DeclareParamType::LongPtr);
        let by_ref = descriptor.param_by_ref.get(index).copied().unwrap_or(false);
        if !long_ptr || by_ref || !is_synchronous_callback_descriptor(descriptor) {
            return Err(Fault::new(
                5,
                "AddressOf proc passed to an unsupported Declare parameter",
            ));
        }
        // SAFETY: the registration lives until the Declare call returns (`JitCallbacks`), and
        // the boxed executor outlives it.
        let registration = unsafe { register_callback(run as usize, proc_token, pointer) }
            .map_err(|error| match error {
                calcvba_runtime::CallbackThunkError::Exhausted => Fault::new(7, error.to_string()),
                calcvba_runtime::CallbackThunkError::UnsupportedPlatform => {
                    Fault::new(5, "AddressOf native callback thunks on this platform")
                }
            })?;
        *value = Variant::from_i64(registration.address() as i64);
        registrations.push(registration);
    }
    Ok(JitCallbacks {
        _registrations: registrations,
        _executor: executor,
    })
}

/// A raw callback argument as the parameter's type.
fn callback_arg(raw: isize, ty: &OxTy) -> Variant {
    match ty {
        OxTy::Byte => Variant::from_u8(raw as u8),
        OxTy::Integer => Variant::from_i16(raw as i16),
        OxTy::Long => Variant::from_i32(raw as i32),
        OxTy::LongLong => Variant::from_i64(raw as i64),
        OxTy::Bool => Variant::from_bool(raw != 0),
        _ => Variant::from_i64(raw as i64),
    }
}

/// A callback's result as the pointer-sized value the native caller reads.
fn callback_result(value: &Variant) -> isize {
    if let Some(value) = value.as_i64() {
        value as isize
    } else if let Some(value) = value.as_i32() {
        value as isize
    } else if let Some(value) = value.as_i16() {
        value as isize
    } else if let Some(value) = value.as_u8() {
        value as isize
    } else if let Some(value) = value.as_bool() {
        if value { -1 } else { 0 }
    } else {
        0
    }
}

/// Runs procedure `proc` of the current program with the raw callback arguments. An error the
/// callback leaves unhandled is dropped (0 is returned), as VM3 does; `End` ends the whole run
/// once the Declare that called back returns.
///
/// # Safety
/// As [`register_callbacks`]; called only from the thunk during that Declare call.
unsafe fn run_callback(
    state: *mut RawExecState,
    run: *mut JitRun,
    proc: usize,
    raw_args: &[isize],
) -> isize {
    let (program_index, image) = {
        // SAFETY: a short read-only borrow, ended before the entry runs.
        let run_ref = unsafe { &*run };
        let Some((index, image)) = current_program_image(run_ref) else {
            return 0;
        };
        (index, image)
    };
    if image.program.is_null() || image.functions.is_null() || proc >= image.function_count {
        return 0;
    }
    // SAFETY: installed from the owning CompiledImage for this run.
    let program = unsafe { &*image.program };
    let Some(func) = program.funcs.get(proc) else {
        return 0;
    };
    let Ok(mut frame) = new_jit_frame(program, program_index, func) else {
        return 0;
    };
    for (index, raw) in raw_args.iter().copied().enumerate().take(func.param_count) {
        if let (Some(local), Some(slot)) = (func.locals.get(index), frame.locals.get_mut(index)) {
            *slot = callback_arg(raw, &local.ty);
        }
    }
    let mut saved_err = RtSavedErrState::default();
    // SAFETY: the live state and initialized output storage.
    if unsafe { rt_err_enter_activation(state, &mut saved_err) } != ST_OK {
        return 0;
    }
    // SAFETY: no other typed run borrow is live; it ends before entry.
    unsafe { &mut *run }.frames.push(frame);
    // SAFETY: bounds were checked above.
    let entry = unsafe { *image.functions.add(proc) };
    // SAFETY: the JIT entry ABI with the run's stable roots; no typed borrow spans the call.
    let status = unsafe { entry(run, state) };
    let result = {
        // SAFETY: entry returned and no typed run borrow is live.
        let run_ref = unsafe { &mut *run };
        match run_ref.frames.pop() {
            Some(frame) => {
                let value = func
                    .return_local
                    .and_then(|ret| frame.locals.get(ret.0))
                    .map_or(0, callback_result);
                // SAFETY: frame cleanup invokes no entries.
                let _ = unsafe { after_jit_frame_pop(run_ref, state, &frame) };
                value
            }
            None => 0,
        }
    };
    // SAFETY: as above; `saved_err` came from the successful enter.
    let _ = unsafe { rt_err_restore_activation(state, &saved_err) };
    if status == ST_HALT {
        // SAFETY: the live state, not borrowed across this call.
        if let Some(exec) = unsafe { jit_exec_state_mut(state) } {
            exec.events
                .pending_abort
                .get_or_insert(calcvba_rt_abi::RunAbort::End);
        }
        return 0;
    }
    if status != ST_OK {
        return 0;
    }
    result
}
