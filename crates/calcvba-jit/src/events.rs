//! COM events in the JIT (design doc, step 8c): the run is the thread's event executor while it
//! executes, so a COM event fired inside a host call runs its WithEvents handler synchronously,
//! as VBA does (VM3 registers itself around each host call; the JIT's helpers hold no borrow of
//! the run across one, so one registration covers the run). An error the handler leaves
//! unhandled ends the run once the host call returns (`ST_ABORT` with the fault recorded), and
//! so does `End` in a handler (`ST_HALT`).
use super::*;
use crate::helpers::{after_jit_frame_pop, invoke_project_member_with_me};
use calcvba_runtime::default_error_message;
use calcvba_runtime::events::{EventExecutor, EventFault, ExecutorScope};
use std::ptr::NonNull;

/// Runs WithEvents handlers on a JIT run.
pub(crate) struct JitEventExecutor {
    state: *mut RawExecState,
    run: *mut JitRun,
}

impl EventExecutor for JitEventExecutor {
    fn deliver_event(
        &mut self,
        key: u64,
        dispid: i32,
        args: &[Variant],
    ) -> Result<bool, EventFault> {
        let (state, run) = (self.state, self.run);
        let call = {
            // SAFETY: the executor is registered only while its run executes on this thread, and a
            // host call (the only place an event arrives) holds no borrow of the state.
            let Some(exec) = (unsafe { jit_exec_state_mut(state) }) else {
                return Ok(false);
            };
            exec.com_event_call(key, dispid, args)
        };
        let Some(call) = call else {
            return Ok(false);
        };
        // The arguments live in a temporary frame; a reference (a ByRef argument, the server's
        // storage) is passed as an alias of its slot, so the handler writes through it.
        let descriptors: Vec<JitCallArgDesc> = call
            .args
            .iter()
            .enumerate()
            .map(|(index, arg)| JitCallArgDesc {
                kind: if arg.byref_target().is_some() {
                    JIT_CALL_ARG_BYREF_ALIAS
                } else {
                    JIT_CALL_ARG_BYVAL_VARIANT
                },
                aux: JIT_VARIANT_OPERAND_PLACE,
                value: 0,
                area: AREA_TEMP as i32,
                index: index as i32,
            })
            .collect();
        let names = vec![
            JitCallArgNameDesc {
                ptr: 0,
                len: -1,
                _pad: 0,
            };
            descriptors.len()
        ];
        let temps = call.args;
        let frame = JitFrame {
            program_index: call.program,
            local_types: Vec::new(),
            temp_types: vec![OxTy::Variant; temps.len()],
            locals: Vec::new(),
            temps,
            pointer_strings: Vec::new(),
            ref_locks: Vec::new(),
            loop_locks: Vec::new(),
            aliases: Vec::new(),
            gosub_stack: Vec::new(),
            saved_err: RtSavedErrState::default(),
            current_line: 0,
        };
        // SAFETY: as above; this push borrow ends before the handler is entered.
        unsafe { &mut *run }.frames.push(frame);
        // SAFETY: the temporary frame owns the arguments for this synchronous call; no run borrow
        // spans the handler.
        let result = unsafe {
            invoke_project_member_with_me(
                run,
                state,
                call.program,
                call.handler,
                call.me,
                &descriptors,
                &names,
            )
        };
        // SAFETY: the handler returned and the temporary frame is topmost again.
        let run_ref = unsafe { &mut *run };
        if let Some(frame) = run_ref.frames.pop() {
            // SAFETY: nothing still borrows the completed frame.
            let _ = unsafe { after_jit_frame_pop(run_ref, state, &frame) };
        }
        match result {
            Ok(_) => Ok(true),
            Err(ST_HALT) => {
                // `End` in the handler ends the whole run once the host call returns; the source
                // sees the event as handled (Excel, probe events8end).
                // SAFETY: the handler returned; the state may be borrowed again.
                if let Some(exec) = unsafe { jit_exec_state_mut(state) } {
                    exec.events
                        .pending_abort
                        .get_or_insert(calcvba_rt_abi::RunAbort::End);
                }
                Ok(true)
            }
            Err(_) => {
                // SAFETY: the handler returned; the state may be borrowed again.
                let Some(exec) = (unsafe { jit_exec_state_mut(state) }) else {
                    return Err(EventFault {
                        number: 51,
                        description: default_error_message(51).into(),
                        source: calcvba_runtime::Utf16String::new(),
                    });
                };
                let err = &exec.err_engine.err;
                let mut fault = Fault::new(err.number, err.description.clone());
                fault.source = Some(err.source.clone());
                let reported = EventFault {
                    number: err.number,
                    description: err.description.clone(),
                    source: err.source.clone(),
                };
                exec.events
                    .pending_abort
                    .get_or_insert(calcvba_rt_abi::RunAbort::Fault(fault));
                Err(reported)
            }
        }
    }
}

/// The JIT run's registration as the thread's event executor; dropping it ends it.
pub(crate) struct JitEventScope {
    _scope: ExecutorScope,
    _executor: Box<JitEventExecutor>,
}

/// Registers the run `(state, run)` as the thread's event executor for as long as the returned
/// scope lives.
///
/// # Safety
/// `state` and `run` are the live run's roots and stay valid until the scope is dropped; the
/// JIT's helpers hold no borrow of either across a host call.
pub(crate) unsafe fn enter_event_scope(
    state: *mut RawExecState,
    run: *mut JitRun,
) -> JitEventScope {
    let mut executor = Box::new(JitEventExecutor { state, run });
    let pointer: NonNull<dyn EventExecutor> =
        NonNull::from(&mut *executor as &mut dyn EventExecutor);
    // SAFETY: the boxed executor outlives the scope (dropped after it), and its pointers obey the
    // contract above.
    let scope = unsafe { ExecutorScope::enter(pointer) };
    JitEventScope {
        _scope: scope,
        _executor: executor,
    }
}

/// The status that ends the run when a WithEvents handler ended it during the host call that
/// just returned: `ST_ABORT` for an error it left unhandled (past every `On Error`), `ST_HALT`
/// for `End`.
///
/// # Safety
/// `state` is the live run's execution state, not borrowed across this call.
pub(crate) unsafe fn event_abort_status(state: *mut RawExecState) -> Option<i32> {
    // SAFETY: forwarded from the caller.
    let exec = unsafe { jit_exec_state_mut(state) }?;
    exec.events
        .pending_abort
        .as_ref()
        .map(calcvba_rt_abi::RunAbort::status)
}
