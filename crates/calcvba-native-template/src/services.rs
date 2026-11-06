//! Native cells enter the same value/library helpers used by compiled VBA.
//! This module executes no source instructions, procedure bodies or CFGs.
use calcvba_hal::{HalProfileId, HostPolicy, HostServices, adapters::builder::HostBuilder};
use calcvba_rt_abi::{ExecState, native_export::Context};
use calcvba_runtime::{VarType, Variant, bstr::BStr};
use std::{ffi::c_void, ptr, sync::Arc};
use windows_sys::Win32::System::Variant::{VARIANT, VariantClear};

pub(super) struct Runtime {
    // Drop the borrowing execution state before its retained host owner.
    exec: Box<ExecState<'static>>,
    _host: Arc<dyn HostServices>,
    frames: Vec<(u32, calcvba_rt_abi::SavedErrState, i32, Vec<u32>)>,
    iterators: std::collections::HashMap<usize, NativeIterator>,
    pub(super) classes: super::classes::Classes,
    pub(super) members: super::members::Calls,
    pub(super) param_arrays: super::parameters::Arrays,
}

pub(super) unsafe fn runtime(context: *mut Context) -> *mut Runtime {
    // SAFETY: generated wrappers have checked the DLL's owning thread. The
    // context retains this allocation until explicit same-thread shutdown.
    unsafe {
        if (*context).services.is_null() {
            let host = HostBuilder::new()
                .profile(HalProfileId::Windows)
                .policy(HostPolicy::native())
                .build();
            // The Arc allocation is stable; Runtime retains it after exec and
            // exposes no safe reference that can outlive that owning Runtime.
            let borrowed = &*Arc::as_ptr(&host);
            let mut exec = Box::new(ExecState::new(borrowed));
            if !(*context).error_source.is_null() {
                exec.default_error_source = String::from_utf8_lossy(std::slice::from_raw_parts(
                    (*context).error_source,
                    (*context).error_source_len,
                ))
                .into_owned();
            }
            let state = Runtime {
                exec,
                _host: host,
                iterators: Default::default(),
                frames: Vec::new(),
                classes: Default::default(),
                members: Default::default(),
                param_arrays: Default::default(),
            };
            (*context).services = Box::into_raw(Box::new(state)).cast();
        }
        (*context).services.cast()
    }
}

pub(super) unsafe fn host(context: *mut Context) -> Arc<dyn HostServices> {
    // SAFETY: the context retains its same-thread services until shutdown.
    unsafe { (*runtime(context))._host.clone() }
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
                message: fault.description,
                source: fault.source,
                help_file: fault.help_file,
                help_context: fault.help_context,
            },
            source,
        );
        finish(context, calcvba_rt_abi::ST_FAULT)
    }
}

pub(super) unsafe fn foreign_type_matches(
    context: *mut Context,
    object: calcvba_runtime::ObjectRef,
    name: &str,
    iid: Option<[u8; 16]>,
) -> Result<bool, i32> {
    // SAFETY: this independent host clone outlives the native nominal check.
    let host = unsafe { host(context) };
    calcvba_rt_abi::object_types::matches_type(host.as_ref(), &object, name, iid)
        .map_err(|fault| fault.code)
}

pub(super) unsafe fn shutdown(context: *mut Context) {
    // SAFETY: shutdown runs on the owning thread after all VBA frames exit.
    unsafe {
        let state = std::mem::replace(&mut (*context).services, ptr::null_mut());
        if !state.is_null() {
            drop(Box::from_raw(state.cast::<Runtime>()));
        }
    }
}

/// # Safety
/// key is a unique native iterator-cell address in the active frame; input
/// retains an initialized native cell of kind for the synchronous call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_iter_init(
    context: *mut Context,
    key: usize,
    input: *const u8,
    kind: u32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: materialize the source before borrowing iterator state.
        unsafe {
            let kind_value = kind & !0x1000_0000;
            let native = super::values::borrowed(input, kind_value);
            if native.Anonymous.Anonymous.vt & 0x4000 != 0 {
                return Err(13);
            }
            let iterator = if kind & 0x1000_0000 != 0 && native.Anonymous.Anonymous.vt & 0x2000 != 0
            {
                let array = native.Anonymous.Anonymous.Anonymous.parray;
                let view = ArrayView::new(array)?;
                super::values::status(windows_sys::Win32::System::Ole::SafeArrayLock(array))?;
                NativeIterator::Array(view, 0)
            } else {
                let source = read(context, input, kind_value)?;
                let values = if let Some(array) = source.as_safearray() {
                    array.variant_elements().unwrap_or_default()
                } else if let Some(object) = source.as_object_ref() {
                    if let Some(values) = object.native_collection_snapshot() {
                        values
                    } else {
                        (*runtime(context))
                            ._host
                            .com()
                            .enumerate_object(object)
                            .map_err(|error| error.host_error_code.unwrap_or(13))?
                    }
                } else {
                    return Err(13);
                };
                NativeIterator::Values(values, 0)
            };
            (*runtime(context)).iterators.insert(key, iterator);
            Ok(())
        }
    })
}

/// # Safety
/// key belongs to this frame; item and has_value are initialized writable cells.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_iter_next(
    context: *mut Context,
    key: usize,
    item: *mut u8,
    kind: u32,
    has_value: *mut i16,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: clone an element before any conversion can enter host code.
        unsafe {
            // The view is borrowed only during this state-only array read.
            // Materialize the item before writing it or entering host services.
            let (ready, array) = {
                let iterator = (*runtime(context)).iterators.get_mut(&key).ok_or(5)?;
                match iterator {
                    NativeIterator::Values(values, index) => {
                        let value = values.get(*index).cloned();
                        *index += usize::from(value.is_some());
                        (value, None)
                    }
                    NativeIterator::Array(view, index) => {
                        if *index >= view.count {
                            (None, None)
                        } else {
                            let position = *index;
                            *index += 1;
                            (None, Some((view.clone(), position)))
                        }
                    }
                }
            };
            let next = if let Some((view, position)) = array {
                Some(view.element(context, position)?)
            } else {
                ready
            };
            if let Some(value) = next {
                write(context, item, kind, &value)?;
                has_value.write(-1);
            } else {
                (*runtime(context)).iterators.remove(&key);
                has_value.write(0);
            }
            Ok(())
        }
    })
}

/// # Safety
/// context and key belong to the exiting native frame on the owning thread.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_iter_clear(context: *mut Context, key: usize) -> i32 {
    super::values::boundary(|| {
        // SAFETY: release only a snapshot owned by this exact frame cell.
        unsafe {
            if !(*context).services.is_null() {
                (*runtime(context)).iterators.remove(&key);
            }
        }
        Ok(())
    })
}

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

pub(super) unsafe fn read(
    context: *mut Context,
    cell: *const u8,
    kind: u32,
) -> Result<Variant, i32> {
    // SAFETY: the emitted call retains an exact native cell for this operation.
    unsafe {
        let value = super::values::borrowed(cell, kind);
        let data = value.Anonymous.Anonymous.Anonymous;
        let vt = value.Anonymous.Anonymous.vt;
        if vt & 0x4000 != 0 {
            let mut indirect = super::values::OwnedVariant::empty();
            super::values::status(windows_sys::Win32::System::Variant::VariantCopyInd(
                &mut indirect.0,
                &value,
            ))?;
            return read(context, (&indirect.0 as *const VARIANT).cast(), 12);
        }
        if vt & 0x2000 != 0 {
            if data.parray.is_null() {
                return Ok(Variant::unallocated_array(vt & 0x0fff));
            }
            return read_array(context, data.parray);
        }
        if matches!(vt, 9 | 13) {
            if data.punkVal.is_null() {
                return Ok(Variant::nothing());
            }
            let object = calcvba_runtime::ObjectRef::from_raw_iunknown_addref(data.punkVal.cast());
            if let Some(object) = object.filter(|object| {
                object.is_project_instance() || object.native_collection_snapshot().is_some()
            }) {
                return Ok(Variant::from_object_ref(object));
            }
        }
        // The COM value model deliberately normalizes narrow integer wire
        // arguments. Internal VBA values must retain their actual subtype.
        match vt {
            2 => return Ok(Variant::from_i16(data.iVal)),
            17 => return Ok(Variant::from_u8(data.bVal)),
            20 => return Ok(Variant::from_i64(data.llVal)),
            4 => return Ok(Variant::from_f32(data.fltVal)),
            _ => {}
        }
        let host = (*runtime(context))._host.clone();
        let mut query = |pointer| {
            calcvba_com::query_interface_pointer(
                pointer,
                &calcvba_com::windows_client::IID_IDISPATCH,
            )
        };
        let mut retain = |pointer: *mut c_void| {
            calcvba_com::add_ref_dispatch(pointer.cast());
        };
        let mut bind = |pointer: *mut c_void, hint: &str, _: &'static str| {
            if pointer.is_null() {
                return Ok(Variant::nothing());
            }
            host.com()
                .bind_native_dispatch_value(hint, pointer)
                .map_err(|error| error.to_string())
        };
        calcvba_com::windows_variant::variant_to_variant_value(
            &value,
            &mut query,
            &mut retain,
            &mut bind,
            "IDispatch",
            "native value",
        )
        .map_err(|_| 13)
    }
}

enum NativeIterator {
    Values(Vec<Variant>, usize),
    Array(ArrayView, usize),
}
impl Drop for NativeIterator {
    fn drop(&mut self) {
        if let Self::Array(view, _) = self {
            // SAFETY: Array entries own one lock on this exact native descriptor.
            unsafe {
                windows_sys::Win32::System::Ole::SafeArrayUnlock(view.pointer);
            }
        }
    }
}
#[derive(Clone)]
struct ArrayView {
    pointer: *mut windows_sys::Win32::System::Com::SAFEARRAY,
    bounds: Vec<calcvba_runtime::safe_array::SafeArrayBound>,
    count: usize,
    element: u16,
    ansi: bool,
}
impl ArrayView {
    unsafe fn new(array: *mut windows_sys::Win32::System::Com::SAFEARRAY) -> Result<Self, i32> {
        use windows_sys::Win32::System::Ole::{
            SafeArrayGetDim, SafeArrayGetLBound, SafeArrayGetUBound, SafeArrayGetVartype,
        };
        // SAFETY: the caller retains or locks the descriptor throughout this view.
        unsafe {
            if array.is_null() {
                return Err(13);
            }
            let rank = SafeArrayGetDim(array);
            if rank == 0 || rank > 60 {
                return Err(9);
            }
            let mut element = 0;
            super::values::status(SafeArrayGetVartype(array, &mut element))?;
            let ansi = element == 0 && (*array).fFeatures & 0x100 != 0;
            if ansi {
                element = 8;
            }
            if element == 0 || element == 36 || element == 14 {
                return Err(13);
            }
            let mut bounds = Vec::new();
            let mut count = 1usize;
            for dimension in 1..=rank {
                let mut lower = 0;
                let mut upper = -1;
                super::values::status(SafeArrayGetLBound(array, dimension, &mut lower))?;
                super::values::status(SafeArrayGetUBound(array, dimension, &mut upper))?;
                let length =
                    u32::try_from(i64::from(upper) - i64::from(lower) + 1).map_err(|_| 9)?;
                count = count.checked_mul(length as usize).ok_or(7)?;
                bounds.push(calcvba_runtime::safe_array::SafeArrayBound {
                    lower,
                    count: length,
                });
            }
            Ok(Self {
                pointer: array,
                bounds,
                count,
                element,
                ansi,
            })
        }
    }
    unsafe fn element(&self, context: *mut Context, linear: usize) -> Result<Variant, i32> {
        use windows_sys::Win32::System::Ole::SafeArrayGetElement;
        // SAFETY: linear is within count and this view's descriptor stays live.
        unsafe {
            let mut remainder = linear;
            let mut coordinates = Vec::with_capacity(self.bounds.len());
            for bound in &self.bounds {
                coordinates.push(bound.lower + (remainder % bound.count as usize) as i32);
                remainder /= bound.count as usize;
            }
            let mut cell = [0u64; 3];
            let pointer = cell.as_mut_ptr().cast::<u8>();
            super::values::status(SafeArrayGetElement(
                self.pointer,
                coordinates.as_ptr(),
                pointer.cast(),
            ))?;
            if self.ansi {
                let original = pointer.cast::<*mut u16>().read();
                let wide = super::values::ansi_to_bstr(
                    original.cast(),
                    windows_sys::Win32::Foundation::SysStringByteLen(original) as i32,
                );
                windows_sys::Win32::Foundation::SysFreeString(original);
                pointer.cast::<*mut u16>().write(wide?);
            }
            let value = read(context, pointer, u32::from(self.element));
            super::values::clear(pointer, u32::from(self.element))?;
            value
        }
    }
}
unsafe fn read_array(
    context: *mut Context,
    array: *mut windows_sys::Win32::System::Com::SAFEARRAY,
) -> Result<Variant, i32> {
    // SAFETY: the incoming native Variant retains this array through all copies.
    unsafe {
        if array.is_null() {
            return Ok(Variant::unallocated_array(12));
        }
        let view = ArrayView::new(array)?;
        let mut values = Vec::with_capacity(view.count);
        for index in 0..view.count {
            values.push(view.element(context, index)?);
        }
        let array = calcvba_runtime::safe_array::SafeArray::from_typed_variants_nd(
            view.bounds,
            view.element,
            values,
        )
        .map_err(|_| 13)?;
        Ok(Variant::from_safearray(array))
    }
}

pub(super) unsafe fn write(
    context: *mut Context,
    out: *mut u8,
    kind: u32,
    value: &Variant,
) -> Result<(), i32> {
    if out.is_null() || kind == 0 {
        return Ok(());
    }
    // SAFETY: emitted code supplies writable initialized exact result storage.
    unsafe {
        let mut native: VARIANT = std::mem::zeroed();
        if let Some(object) = value.as_object_ref().filter(|object| {
            object.is_project_instance() || object.native_collection_snapshot().is_some()
        }) {
            let pointer = object.into_raw_iunknown();
            let result = super::values::copy(out, std::ptr::from_ref(&pointer).cast(), kind, 9);
            drop(calcvba_runtime::ObjectRef::from_raw_iunknown_owned(pointer));
            return result;
        }
        match value.vtype() {
            VarType::Object if value.as_object_ref().is_none() => {
                native.Anonymous.Anonymous.vt = 9;
                native.Anonymous.Anonymous.Anonymous.pdispVal = std::ptr::null_mut();
            }
            VarType::Integer => {
                native.Anonymous.Anonymous.vt = 2;
                native.Anonymous.Anonymous.Anonymous.iVal = value.as_i16().ok_or(13)?;
            }
            VarType::Byte => {
                native.Anonymous.Anonymous.vt = 17;
                native.Anonymous.Anonymous.Anonymous.bVal = value.as_u8().ok_or(13)?;
            }
            VarType::LongLong => {
                native.Anonymous.Anonymous.vt = 20;
                native.Anonymous.Anonymous.Anonymous.llVal = value.as_i64().ok_or(13)?;
            }
            VarType::Single => {
                native.Anonymous.Anonymous.vt = 4;
                native.Anonymous.Anonymous.Anonymous.fltVal = value.as_f32().ok_or(13)?;
            }
            VarType::Error => {
                native.Anonymous.Anonymous.vt = 10;
                native.Anonymous.Anonymous.Anonymous.scode = value.as_error_code().ok_or(13)?;
            }
            _ => {
                let host = (*runtime(context))._host.clone();
                // Resolver transfers a reference; the marshaller's retain hook
                // is empty because its resulting VARIANT owns that reference.
                let mut resolve = |object| {
                    if let Some(prepared) = calcvba_host_api::HostObject::from_object_ref(&object) {
                        return calcvba_com::typed_slots::retain_prepared_native_dispatch(
                            &prepared,
                        );
                    }
                    host.com()
                        .retain_native_dispatch_object(object)
                        .map_err(|error| error.to_string())
                };
                let mut retain = |_: *mut c_void| {};
                calcvba_com::windows_variant::set_variant_from_com_value(
                    &mut native,
                    &calcvba_com::ComValue::from_variant(value).map_err(|_| 13)?,
                    &mut resolve,
                    &mut retain,
                )
                .map_err(|_| 13)?;
            }
        }
        let result = super::values::copy(out, (&native as *const VARIANT).cast(), kind, 12);
        VariantClear(&mut native);
        result
    }
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
            source: BStr::from(error.source.as_str()),
            description: BStr::from(error.description.as_str()),
            help_file: BStr::from(error.help_file.as_str()),
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
/// Name is compiler-retained UTF-8, and out is an initialized Object cell on
/// the image's owning thread. Activation occurs only while the cell is Nothing.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_com_ensure(
    context: *mut Context,
    name: *const u8,
    len: usize,
    out: *mut u8,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: all runtime borrows end before COM activation can reenter VBA.
        unsafe {
            if out.cast::<usize>().read() != 0 {
                return Ok(());
            }
            let name = std::str::from_utf8_unchecked(std::slice::from_raw_parts(name, len));
            let input = Variant::from_string(name);
            let mut result = Variant::empty();
            let native = calcvba_bundle::NativeImplId::ALL
                .iter()
                .position(|id| *id == calcvba_bundle::NativeImplId::CreateObject)
                .ok_or(5)? as u32;
            let state = &raw mut *(*runtime(context)).exec;
            let status = calcvba_rt_abi::rt_lib_invoke_with_policy(
                calcvba_rt_abi::exec_state_as_raw(&mut *state),
                native,
                &input,
                1,
                0,
                &mut result,
            );
            finish(context, status)?;
            write(context, out, 9, &result)
        }
    })
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
/// The generated native frame owns this context on the calling thread.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_frame_enter(context: *mut Context) -> i32 {
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
            runtime
                .frames
                .push(((*context).depth, saved, 0, Vec::new()));
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
                let (depth, saved, _, _) = runtime.frames.pop().ok_or(5)?;
                runtime
                    .param_arrays
                    .retain(|_, aliases| aliases.depth < depth);
                runtime.exec.err_engine.restore(saved);
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
                1 => Variant::from_string(error.err.description.as_str()),
                2 => Variant::from_string(error.err.source.as_str()),
                3 => Variant::from_string(error.err.help_file.as_str()),
                4 => Variant::from_i32(error.err.help_context),
                5 => Variant::from_i32(windows_sys::Win32::Foundation::GetLastError() as i32),
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
                source: (flags & 1 != 0).then(|| text.source.as_str()),
                description: (flags & 2 != 0).then(|| text.description.as_str()),
                help_file: (flags & 4 != 0).then(|| text.help_file.as_str()),
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
/// cell is initialized native storage of kind, retained by the calling VBA frame.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_array_erase(_context: *mut Context, cell: *mut u8, kind: u32) -> i32 {
    super::values::boundary(|| {
        // SAFETY: the native cell owns its descriptor; fixed arrays retain that
        // allocation while each Automation element is reset to its typed default.
        unsafe {
            let native = super::values::borrowed(cell, kind);
            if native.Anonymous.Anonymous.vt & 0x4000 != 0 {
                return Err(13);
            }
            if native.Anonymous.Anonymous.vt & 0x2000 == 0 {
                return super::values::clear(cell, kind);
            }
            let array = native.Anonymous.Anonymous.Anonymous.parray;
            if !array.is_null() && (*array).cLocks != 0 {
                return Err(10);
            }
            if array.is_null() || (*array).fFeatures & 0x10 == 0 {
                return super::values::clear(cell, kind);
            }
            let view = ArrayView::new(array)?;
            for linear in 0..view.count {
                let mut remainder = linear;
                let mut coordinates = Vec::with_capacity(view.bounds.len());
                for bound in &view.bounds {
                    coordinates.push(bound.lower + (remainder % bound.count as usize) as i32);
                    remainder /= bound.count as usize;
                }
                let zero = [0u64; 3];
                let value = if matches!(view.element, 8 | 9 | 13) {
                    ptr::null()
                } else {
                    zero.as_ptr().cast()
                };
                super::values::status(windows_sys::Win32::System::Ole::SafeArrayPutElement(
                    array,
                    coordinates.as_ptr(),
                    value,
                ))?;
            }
            Ok(())
        }
    })
}

/// # Safety
/// The native activation owns context. For pop, out is a writable block index;
/// push records only a compiler-validated return block in the current frame.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_gosub(
    context: *mut Context,
    pop: u32,
    value: u32,
    out: *mut u32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: no callback occurs while the active frame is borrowed.
        unsafe {
            let stack = &mut (*runtime(context)).frames.last_mut().ok_or(5)?.3;
            if pop == 0 {
                stack.try_reserve(1).map_err(|_| 7)?;
                stack.push(value);
            } else {
                out.write(stack.pop().ok_or(3)?);
            }
            Ok(())
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn variant_arithmetic_fault_retains_the_images_project_name() {
        // SAFETY: all cells belong to this test and remain live through the
        // synchronous helper call; only its retained error/runtime owners are freed.
        unsafe {
            let source = b"NativeProject";
            let mut context: Context = std::mem::zeroed();
            context.error_source = source.as_ptr();
            context.error_source_len = source.len();
            let left = 1i32;
            let right = 0i32;
            let mut output: VARIANT = std::mem::zeroed();
            assert_eq!(
                cvb_value_binary(
                    &raw mut context,
                    (&raw mut output).cast(),
                    (&raw const left).cast(),
                    (&raw const right).cast(),
                    12,
                    3,
                    3,
                    0,
                    calcvba_rt_abi::RT_ARITH_DIV,
                    calcvba_rt_abi::RT_NUMERIC_WIDENING,
                ),
                11
            );
            let error = &(*runtime(&raw mut context)).exec.err_engine.err;
            assert_eq!(error.source, "NativeProject");
            assert_eq!(error.description, "Division by zero");
            super::super::errors::clear(&raw mut context);
            assert_eq!(cvb_services_clear(&raw mut context), 0);
            assert_eq!(VariantClear(&raw mut output), 0);
        }
    }

    #[test]
    fn indirect_variant_array_is_copied_before_library_reads_and_typed_assignment() {
        use windows_sys::Win32::System::Ole::{SafeArrayCreateVector, SafeArrayDestroy};
        // SAFETY: every descriptor, pointer cell and VARIANT remains owned by this
        // test until synchronous helpers complete; only the recorded allocations are freed.
        unsafe {
            let mut array = SafeArrayCreateVector(3, 3, 3);
            assert!(!array.is_null());
            let mut input: VARIANT = std::mem::zeroed();
            input.Anonymous.Anonymous.vt = 0x6003;
            input.Anonymous.Anonymous.Anonymous.pparray = &raw mut array;
            let input_pointer = (&input as *const VARIANT).cast::<u8>();
            let mut context: Context = std::mem::zeroed();
            let implementation = calcvba_bundle::NativeImplId::ALL
                .iter()
                .position(|id| *id == calcvba_bundle::NativeImplId::VarType)
                .unwrap() as u32;
            let mut upper = 0i32;
            assert_eq!(
                cvb_library(
                    &raw mut context,
                    implementation,
                    0,
                    [1u32, 3, 12].as_ptr(),
                    [input_pointer].as_ptr(),
                    (&raw mut upper).cast()
                ),
                0
            );
            assert_eq!(upper, 0x2003);
            let mut copied: *mut windows_sys::Win32::System::Com::SAFEARRAY = ptr::null_mut();
            assert!(
                super::super::values::copy(
                    (&raw mut copied).cast::<u8>(),
                    input_pointer,
                    0x2003,
                    12
                )
                .is_ok()
            );
            assert_ne!(copied, array);
            assert_eq!(SafeArrayDestroy(copied), 0);
            assert_eq!(
                cvb_array_erase(&raw mut context, input_pointer.cast_mut(), 12),
                13
            );
            assert_eq!(
                cvb_iter_init(&raw mut context, 1, input_pointer, 0x1000_000c),
                13
            );
            assert_eq!(cvb_services_clear(&raw mut context), 0);
            assert_eq!(SafeArrayDestroy(array), 0);
        }
    }
}
