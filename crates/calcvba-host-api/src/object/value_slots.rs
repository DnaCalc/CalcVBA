//! Typed member slots added without changing the original two numeric ABIs.
use super::*;
use crate::{Value, ValueSlot, ValueType, value_slot_signature};

pub(super) struct TypedSlot {
    parameters: Vec<ValueType>,
    result: ValueType,
    callback: Box<crate::value::TypedCallback>,
}
impl HostSlot {
    pub fn typed(
        parameters: Vec<ValueType>,
        result: ValueType,
        callback: impl Fn(&mut TypedContext, &[ValueSlot<'_>]) -> Result<Value, i32> + 'static,
    ) -> Result<Self, crate::Error> {
        value_slot_signature(&parameters, result).ok_or_else(|| {
            crate::Error::UnsupportedSignature(
                "typed host slot exceeds its carrier/argument bounds".into(),
            )
        })?;
        Ok(Self {
            callback: Box::new(Callback::Typed(TypedSlot {
                parameters,
                result,
                callback: Box::new(callback),
            })),
        })
    }
}
impl TypedSlot {
    pub(super) fn native(&self) -> (*const u8, u32) {
        let signature =
            value_slot_signature(&self.parameters, self.result).expect("validated typed slot");
        let entry = match signature {
            SLOT_DOUBLE_0 => typed_double0 as *const u8,
            SLOT_DOUBLE_LONG => typed_double_long as *const u8,
            _ => typed_values as *const u8,
        };
        (entry, signature)
    }
    unsafe fn invoke(
        &self,
        context: &mut TypedContext,
        args: *const ScalarBits,
    ) -> Result<Value, i32> {
        let mut owners: [Value; crate::MAX_TYPED_ARGS] = std::array::from_fn(|_| Value::Empty);
        for (i, ty) in self.parameters.iter().enumerate() {
            // SAFETY: caller supplies this slot's exact initialized argument extent.
            let bits = unsafe { *args.add(i) };
            // SAFETY: admitted caller owns each exact carrier through this entry.
            owners[i] = unsafe { crate::value::retain_bits(*ty, bits) }?;
            if let Some(iid) = ty.interface()
                && !owners[i].has_interface(iid)
            {
                return Err(13);
            }
        }
        let mut slots = [ValueSlot::Long(0); crate::MAX_TYPED_ARGS];
        for (i, owner) in owners[..self.parameters.len()].iter().enumerate() {
            slots[i] = owner.borrow_slot().ok_or(13)?;
        }
        drop(crate::take_typed_fault(context, 0));
        let value = (self.callback)(context, &slots[..self.parameters.len()])?;
        drop(crate::take_typed_fault(context, 0));
        if value.ty().carrier() != self.result.carrier() {
            return Err(13);
        }
        if let Value::Double(number) = value
            && !number.is_finite()
        {
            return Err(6);
        }
        if let Some(iid) = self.result.interface()
            && !value.has_interface(iid)
        {
            return Err(13);
        }
        Ok(value)
    }
}

// SAFETY CONTRACT: data is a retained Callback::Typed of the advertised native
// signature; context/args/output are live, aligned and nonaliasing for the call.
unsafe extern "C" fn typed_values(
    data: *const (),
    context: *mut TypedContext,
    args: *const ScalarBits,
    output: *mut ScalarBits,
) -> i32 {
    let result = catch_unwind(AssertUnwindSafe(|| {
        // SAFETY: the factory pairs this callback and signature with native().
        let Callback::Typed(slot) = (unsafe { &*data.cast::<Callback>() }) else {
            return Err(5);
        };
        // SAFETY: generated caller retains context and the complete argument extent.
        unsafe { slot.invoke(&mut *context, args) }
    }));
    match result {
        Ok(Ok(value)) => {
            // SAFETY: caller supplies one initialized writable carrier cell.
            unsafe {
                crate::value::write_result(value, output);
            }
            0
        }
        Ok(Err(status)) if status == -1 || status > 0 => status,
        Ok(Err(_)) => 5,
        Err(payload) => {
            discard_panic(payload);
            5
        }
    }
}
// SAFETY CONTRACT: same as typed_values, for the original zero-argument Double ABI.
unsafe extern "C" fn typed_double0(
    data: *const (),
    context: *mut TypedContext,
    output: *mut f64,
) -> i32 {
    let result = catch_unwind(AssertUnwindSafe(|| {
        // SAFETY: exact retained factory callback and unique live context.
        let Callback::Typed(slot) = (unsafe { &*data.cast::<Callback>() }) else {
            return Err(5);
        };
        // SAFETY: this signature declares no input cells.
        match unsafe { slot.invoke(&mut *context, std::ptr::null()) }? {
            Value::Double(value) => Ok(value),
            _ => Err(13),
        }
    }));
    finish(result, output)
}
// SAFETY CONTRACT: same as typed_values, for the original Long-to-Double ABI.
unsafe extern "C" fn typed_double_long(
    data: *const (),
    context: *mut TypedContext,
    arg: i32,
    output: *mut f64,
) -> i32 {
    let result = catch_unwind(AssertUnwindSafe(|| {
        // SAFETY: exact retained factory callback and unique live context.
        let Callback::Typed(slot) = (unsafe { &*data.cast::<Callback>() }) else {
            return Err(5);
        };
        let argument = ScalarBits::long(arg);
        // SAFETY: this signature declares exactly one Long input cell.
        match unsafe { slot.invoke(&mut *context, &argument) }? {
            Value::Double(value) => Ok(value),
            _ => Err(13),
        }
    }));
    finish(result, output)
}

/// # Safety
/// Header is null or a retained factory table; args matches the checked signature,
/// and every argument's exact carrier owner spans this synchronous call.
pub unsafe fn invoke_value_slot(
    header: *const ObjectHeader,
    iid: InterfaceId,
    slot: usize,
    parameters: &[ValueType],
    result: ValueType,
    args: &[ScalarBits],
    context: &mut TypedContext,
) -> Result<Value, i32> {
    // SAFETY: every retained object carrier supplies the canonical IUnknown
    // prefix. Source objects must never be read as the larger factory header.
    let owner =
        unsafe { crate::value::retain_bits(ValueType::HostObject, ScalarBits::object(header)) }?;
    let Value::Object(owner) = owner else {
        return Err(13);
    };
    if owner.is_none() {
        return Err(91);
    }
    if args.len() != parameters.len() {
        return Err(13);
    }
    let signature = value_slot_signature(parameters, result).ok_or(13)?;
    if signature == SLOT_DOUBLE_0 || signature == SLOT_DOUBLE_LONG {
        // SAFETY: this retained signature uses the unchanged numeric native entry.
        let bits = unsafe {
            invoke_double_slot(
                header,
                iid,
                slot,
                args.first().map(|arg| arg.as_long()),
                context,
            )
        }?;
        return Ok(Value::Double(bits.as_double()));
    }
    // SAFETY: inherited factory-header lifetime contract.
    let header = unsafe { header.as_ref() }.ok_or(91)?;
    if header.iid != iid || slot >= header.slot_count {
        return Err(438);
    }
    // SAFETY: factory retains the complete initialized native table.
    let native = unsafe { &*header.slots.add(slot) };
    if native.signature != signature {
        return Err(13);
    }
    // SAFETY: the checked signature names this exact common typed-slot ABI.
    let call: unsafe extern "C" fn(
        *const (),
        *mut TypedContext,
        *const ScalarBits,
        *mut ScalarBits,
    ) -> i32 = unsafe { std::mem::transmute(native.entry) };
    let mut output = std::mem::MaybeUninit::<crate::Variant>::uninit();
    // SAFETY: caller retains inputs and the largest aligned typed-result extent.
    let status = unsafe {
        call(
            native.data,
            context,
            args.as_ptr(),
            output.as_mut_ptr().cast(),
        )
    };
    if status != 0 {
        return Err(status);
    }
    // SAFETY: success transfers one output owner of the declared result carrier.
    unsafe { crate::value::take_result(result, output.as_mut_ptr().cast()) }
}
