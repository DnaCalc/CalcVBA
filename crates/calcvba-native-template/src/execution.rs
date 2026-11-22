//! Shared native library calls and VBA activation/error state.
use super::services::{read, runtime, shutdown, write};
use calcvba_rt_abi::native_export::Context;
use calcvba_runtime::{VarType, Variant};

pub(super) type Frame = (u32, calcvba_rt_abi::SavedErrState, i32, Vec<u32>, String);

/// # Safety
/// No VBA frame is active and context belongs to the calling owning thread.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_services_clear(context: *mut Context) -> i32 {
    super::values::boundary(|| {
        // SAFETY: generated shutdown released globals and checked active depth.
        unsafe {
            shutdown(context);
        }
        Ok(())
    })
}

pub(super) unsafe fn finish(context: *mut Context, status: i32) -> Result<(), i32> {
    if status == calcvba_rt_abi::ST_OK {
        return Ok(());
    }
    // SAFETY: the runtime helper completed, leaving its fault in this context.
    unsafe {
        let error = &(*runtime(context)).exec.err_engine.err;
        let number = if error.number == 0 { 5 } else { error.number };
        let text = super::errors::ErrorText {
            source: error.source.to_bstr(),
            description: error.description.to_bstr(),
            help_file: error.help_file.to_bstr(),
            help_context: error.help_context,
        };
        super::errors::clear(context);
        (*context).error_owner = Box::into_raw(Box::new(text)).cast();
        (*context).error_number = number;
        (*context).custom_error = 0x10f;
        Err(number)
    }
}

/// # Safety
/// context belongs to this calling thread. metadata contains argc, result kind,
/// then argc input kinds; arguments contains that many retained native cells.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_library(
    context: *mut Context,
    native: u32,
    string_alias: u32,
    metadata: *const u32,
    arguments: *const *const u8,
    out: *mut u8,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: compiler-generated immutable descriptor and matching cells.
        unsafe {
            let count = metadata.read() as usize;
            let kind = metadata.add(1).read();
            let mut values = Vec::with_capacity(count);
            for index in 0..count {
                values.push(read(
                    context,
                    arguments.add(index).read(),
                    metadata.add(index + 2).read(),
                )?);
            }
            let mut result = Variant::empty();
            let state = &raw mut *(*runtime(context)).exec;
            let status = calcvba_rt_abi::rt_lib_invoke_with_policy(
                calcvba_rt_abi::exec_state_as_raw(&mut *state),
                native,
                values.as_ptr(),
                values.len(),
                string_alias as i32,
                &mut result,
            );
            finish(context, status)?;
            write(context, out, kind, &result)
        }
    })
}

/// # Safety
/// The generated native frame owns this context on the calling thread.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_frame_enter(
    context: *mut Context,
    source: *const u8,
    length: usize,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: generated frame entry runs before any callback-capable body operation.
        unsafe {
            let runtime = &mut *runtime(context);
            runtime.frames.try_reserve(1).map_err(|_| 7)?;
            // A completed public invocation cannot supply the pending fault for
            // a later call. Keep Err properties for explicit Raise inheritance;
            // nested calls retain pending faults until their caller dispatches.
            if (*context).depth == 1 {
                runtime.exec.err_engine.pending_fault = None;
            }
            let saved = runtime.exec.err_engine.enter_activation();
            let source =
                String::from_utf8_lossy(std::slice::from_raw_parts(source, length)).into_owned();
            let previous = std::mem::replace(&mut runtime.exec.default_error_source, source);
            runtime
                .frames
                .push(((*context).depth, saved, 0, Vec::new(), previous));
        }
        Ok(())
    })
}

/// # Safety
/// The generated native frame is exiting on its owning thread.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_frame_exit(context: *mut Context) -> i32 {
    super::values::boundary(|| {
        // SAFETY: depth identifies this accepted frame, including failed entry allocation.
        unsafe {
            let runtime = &mut *runtime(context);
            if runtime
                .frames
                .last()
                .is_some_and(|frame| frame.0 == (*context).depth)
            {
                let (depth, saved, _, _, source) = runtime.frames.pop().ok_or(5)?;
                runtime.end_frame(depth);
                runtime.exec.err_engine.restore(saved);
                runtime.exec.default_error_source = source;
            }
        }
        // SAFETY: all local owners were released before this frame exit helper.
        let number = unsafe { super::classes::cvb_drain_terminations(context) };
        if number == 0 { Ok(()) } else { Err(number) }
    })
}

/// # Safety
/// Context is live on the owning thread; action and argument are emitted constants.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_error_control(
    context: *mut Context,
    action: u32,
    argument: i32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: these state-only operations cannot call VBA or a host callback.
        unsafe {
            let runtime = &mut *runtime(context);
            match action {
                0..=3 => {
                    let state = calcvba_rt_abi::exec_state_as_raw(&mut runtime.exec);
                    let status =
                        calcvba_rt_abi::rt_set_error_handler(state, action, argument as u32);
                    finish(context, status)?;
                }
                4 => runtime.exec.err_engine.clear_err(),
                5 => runtime.frames.last_mut().ok_or(5)?.2 = argument,
                6 => {
                    if runtime
                        .exec
                        .err_engine
                        .pending_fault
                        .as_ref()
                        .is_none_or(|fault| fault.code != argument)
                    {
                        let source = runtime.exec.default_error_source.clone();
                        runtime.exec.err_engine.raise(
                            calcvba_rt_abi::Fault::new(
                                argument,
                                calcvba_rt_abi::default_error_message(argument),
                            ),
                            source,
                        );
                    }
                    let _ = finish(context, calcvba_rt_abi::ST_FAULT);
                }
                7 => runtime.exec.err_engine.clear_err_on_exit(),
                _ => return Err(5),
            }
        }
        Ok(())
    })
}

/// # Safety
/// out is an initialized native cell of kind, retained for this synchronous call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_error_get(
    context: *mut Context,
    out: *mut u8,
    kind: u32,
    field: u32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: clone error data before conversion can enter host services.
        unsafe {
            let error = &(*runtime(context)).exec.err_engine;
            let value = match field {
                0 => Variant::from_i32(error.err.number),
                1 => Variant::from_string(&error.err.description),
                2 => Variant::from_string(&error.err.source),
                3 => Variant::from_string(&error.err.help_file),
                4 => Variant::from_i32(error.err.help_context),
                5 => Variant::from_i32(super::thread::last_error() as i32),
                6 => Variant::from_i32(error.erl_line),
                _ => return Err(5),
            };
            write(context, out, kind, &value)
        }
    })
}

/// # Safety
/// input is a live native cell of kind; field identifies a shared Err property.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_error_set(
    context: *mut Context,
    field: u32,
    input: *const u8,
    kind: u32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: input conversion completes before the state-only error assignment.
        unsafe {
            let value = read(context, input, kind)?;
            let state = calcvba_rt_abi::exec_state_as_raw(&mut (*runtime(context)).exec);
            let status = calcvba_rt_abi::rt_err_set_field(state, field, &value);
            finish(context, status)
        }
    })
}

/// # Safety
/// out is a writable block-number cell. The compiler emits every branch destination.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_fault_route(
    context: *mut Context,
    resume: u32,
    next: u32,
    out: *mut u32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: shared routing only selects a block; compiled code performs the transfer.
        unsafe {
            let runtime = &mut *runtime(context);
            let line = runtime.frames.last().ok_or(5)?.2;
            let state = calcvba_rt_abi::exec_state_as_raw(&mut runtime.exec);
            let mut dispatch = 0;
            let status =
                calcvba_rt_abi::rt_route_fault(state, resume, next, line, &mut dispatch, out);
            finish(context, status)?;
            if dispatch == calcvba_rt_abi::RT_FAULT_DISP_UNWIND {
                return finish(context, calcvba_rt_abi::ST_FAULT);
            }
            Ok(())
        }
    })
}

/// # Safety
/// out is a writable block-number cell; kind/label are emitted Resume operands.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_resume(
    context: *mut Context,
    kind: u32,
    label: u32,
    out: *mut u32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: shared Resume updates Err and selects one compiled destination.
        unsafe {
            let state = calcvba_rt_abi::exec_state_as_raw(&mut (*runtime(context)).exec);
            let status = calcvba_rt_abi::rt_resume(state, kind, label, out);
            finish(context, status)
        }
    })
}

pub(super) unsafe fn raise_explicit(
    context: *mut Context,
    number: i32,
    text: &super::errors::ErrorText,
    flags: u32,
) -> i32 {
    // SAFETY: all explicit fields have been materialized into owned text before state access.
    unsafe {
        let runtime = &mut *runtime(context);
        let fault = runtime.exec.err_engine.explicit_fault(
            number,
            calcvba_rt_abi::RaiseFields {
                source: (flags & 1 != 0)
                    .then(|| calcvba_runtime::Utf16String::from_bstr(&text.source)),
                description: (flags & 2 != 0)
                    .then(|| calcvba_runtime::Utf16String::from_bstr(&text.description)),
                help_file: (flags & 4 != 0)
                    .then(|| calcvba_runtime::Utf16String::from_bstr(&text.help_file)),
                help_context: (flags & 8 != 0).then_some(text.help_context),
            },
            flags & 0x10 != 0,
        );
        let source = runtime.exec.default_error_source.clone();
        runtime.exec.err_engine.raise(fault, source);
        finish(context, calcvba_rt_abi::ST_FAULT).unwrap_err()
    }
}

/// # Safety
/// All cells have the exact declared native kinds and remain live for the call.
#[unsafe(no_mangle)]
#[allow(clippy::too_many_arguments)]
pub unsafe extern "C" fn cvb_value_binary(
    context: *mut Context,
    out: *mut u8,
    a: *const u8,
    b: *const u8,
    to: u32,
    ak: u32,
    bk: u32,
    family: u32,
    operation: u32,
    mode: u32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: generated operands are borrowed only while materializing values;
        // all runtime calls receive private, nonaliasing input/output Variants.
        unsafe {
            let a = read(context, a, ak)?;
            let b = read(context, b, bk)?;
            let mut result = Variant::empty();
            let state = &raw mut *(*runtime(context)).exec;
            let state = calcvba_rt_abi::exec_state_as_raw(&mut *state);
            let status = match family {
                0 => calcvba_rt_abi::rt_arith_v(state, operation, mode, &a, &b, &mut result),
                1 => calcvba_rt_abi::rt_logical_v(state, operation, &a, &b, &mut result),
                2 => {
                    let coercion = match mode >> 8 {
                        0 => calcvba_bundle::ComparisonCoercion::Strict,
                        1 => calcvba_bundle::ComparisonCoercion::String,
                        2 => calcvba_bundle::ComparisonCoercion::Numeric,
                        3 => calcvba_bundle::ComparisonCoercion::Variant,
                        _ => return Err(5),
                    };
                    calcvba_rt_abi::compare_declared_values(
                        state,
                        operation,
                        mode & 0xff,
                        coercion,
                        &a,
                        &b,
                        &mut result,
                    )
                }
                3 => match calcvba_eval::arith::concat(&a, &b) {
                    Ok(value) => {
                        result = value;
                        calcvba_rt_abi::ST_OK
                    }
                    Err(error) => {
                        let runtime = &mut *runtime(context);
                        let source = runtime.exec.default_error_source.clone();
                        runtime
                            .exec
                            .err_engine
                            .raise(calcvba_rt_abi::Fault::from_arith(error), source);
                        calcvba_rt_abi::ST_FAULT
                    }
                },
                _ => return Err(5),
            };
            if family == 0
                && result.vtype() == VarType::Double
                && result.as_f64().is_some_and(|f| !f.is_finite())
            {
                write(context, out, to, &result)?;
            }
            finish(context, status)?;
            write(context, out, to, &result)
        }
    })
}

/// # Safety
/// Cells have the exact declared native kinds and remain live for the call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_value_unary(
    context: *mut Context,
    out: *mut u8,
    input: *const u8,
    to: u32,
    from: u32,
    family: u32,
    mode: u32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: inputs are materialized before calling shared runtime helpers;
        // output is private and cannot alias the owned input Variant.
        unsafe {
            let input = read(context, input, from)?;
            let mut result = Variant::empty();
            let state = &raw mut *(*runtime(context)).exec;
            let state = calcvba_rt_abi::exec_state_as_raw(&mut *state);
            let status = match family {
                0 => calcvba_rt_abi::rt_neg_v(state, mode, &input, &mut result),
                1 => calcvba_rt_abi::rt_not_v(state, &input, &mut result),
                2 => {
                    let mut truth = 0;
                    let status = calcvba_rt_abi::rt_truthy_v(state, &input, &mut truth);
                    result = Variant::from_bool(truth != 0);
                    status
                }
                _ => return Err(5),
            };
            finish(context, status)?;
            write(context, out, to, &result)
        }
    })
}

/// # Safety
/// Inputs have the declared native kinds, and out owns a writable Boolean cell.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_variant_changed(
    context: *mut Context,
    out: *mut i16,
    current: *const u8,
    current_kind: u32,
    original: *const u8,
    original_kind: u32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: each value is materialized before comparing or writing output.
        unsafe {
            let current = read(context, current, current_kind)?;
            let original = read(context, original, original_kind)?;
            out.write(if calcvba_rt_abi::variant_changed(&current, &original) {
                -1
            } else {
                0
            });
        }
        Ok(())
    })
}

pub(super) unsafe fn typed_fault(
    context: *mut Context,
    fault: calcvba_host_api::TypedFault,
) -> Result<(), i32> {
    // SAFETY: the active native frame owns this error state, after member owners
    // have been released. No borrow spans a native/VBA callback.
    unsafe {
        let state = &mut *(*runtime(context)).exec;
        let source = state.default_error_source.clone();
        state.err_engine.raise(
            calcvba_rt_abi::Fault {
                code: fault.number,
                message: fault.description.into(),
                source: fault.source.map(Into::into),
                help_file: fault.help_file.map(Into::into),
                help_context: fault.help_context,
            },
            source,
        );
        finish(context, calcvba_rt_abi::ST_FAULT)
    }
}
