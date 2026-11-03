//! One retained descriptor-planned member operation, shared by both backends.
use super::{
    ScalarType,
    activation::{RawActivation, read_slot, slot_identity, write_owned, write_slot},
    value::{from_variant, read_operand, to_variant},
};
use calcvba_host_api::{
    TypedContext, Value,
    object::planned_slots::{CallerSyncPhase, MemberArguments, invoke_planned_member},
};
use oxvba_oxir::{OxCallArg, OxInst};

/// # Safety
/// raw/inst follow value::evaluate's matching leased frame/code contract. context
/// is the uniquely running synchronous typed context, including host reentry.
pub(super) unsafe fn invoke(
    raw: *mut RawActivation,
    inst: &OxInst,
    context: &mut TypedContext,
) -> Result<(), i32> {
    // SAFETY: the activation retains this immutable admitted layout.
    let layout = unsafe { &*(*raw).layout };
    let OxInst::ComCallEarly {
        method,
        recv,
        args,
        dst,
        ..
    } = inst
    else {
        return Err(5);
    };
    let plan = layout.host_calls.get(method).ok_or(5)?;
    let descriptor = plan.native.as_ref().ok_or(5)?;
    // SAFETY: checked receiver carrier and retained scratch lifetime.
    let header = unsafe { read_operand(raw, layout, recv, ScalarType::HostObject) }?.as_object();
    if header.is_null() {
        return Err(91);
    }
    // SAFETY: the retained carrier has a canonical IUnknown prefix; only a
    // factory-vtable proof permits reading the larger host slot table.
    let Value::Object(Some(receiver)) = (unsafe {
        calcvba_host_api::value::retain_bits(
            ScalarType::HostObject,
            super::ScalarBits::object(header),
        )
    })?
    else {
        return Err(13);
    };
    let mut values: [Value; calcvba_host_api::MAX_TYPED_ARGS] =
        std::array::from_fn(|_| Value::Empty);
    let mut aliases = std::array::from_fn::<_, { calcvba_host_api::MAX_TYPED_ARGS }, _>(|i| i);
    let mut targets = [None; calcvba_host_api::MAX_TYPED_ARGS];
    let mut identities = [None; calcvba_host_api::MAX_TYPED_ARGS];
    let indices = oxvba_com::typed_slots::plan::visible_indices(descriptor);
    for (i, (arg, ty)) in args.iter().zip(&plan.parameters).enumerate() {
        if let OxCallArg::Operand(operand) = arg
            && ty.carrier() == ScalarType::Variant
        {
            let actual = layout.operand_type(operand).map_err(|_| 5)?;
            if actual != ScalarType::Variant {
                // Box the actual carrier only at the declared VARIANT boundary.
                // A Range/object stays an object: passing a Variant parameter
                // must not evaluate its default member before the COM call.
                // SAFETY: the immutable layout proves the actual operand type;
                // the frame retains it through this owning input copy.
                let bits = unsafe { read_operand(raw, layout, operand, actual) }?;
                // SAFETY: exact actual carrier and its retained source owner.
                let value = unsafe { calcvba_host_api::value::retain_bits(actual, bits) }?;
                values[i] = Value::Variant(to_variant(value));
                continue;
            }
        }
        let bits = match arg {
            OxCallArg::Operand(operand) => {
                // SAFETY: admitted actual operand and exact carrier.
                unsafe { read_operand(raw, layout, operand, ty.carrier()) }?
            }
            OxCallArg::ByRef(place) => {
                let index = layout.slot(*place).map_err(|_| 5)?;
                if oxvba_com::typed_slots::plan::is_by_ref(descriptor, indices[i]) {
                    targets[i] = Some(index);
                    // SAFETY: admitted slot resolves only live ancestors/globals.
                    let identity = unsafe { slot_identity(raw, index) }?;
                    if let Some(previous) = identities[..i]
                        .iter()
                        .position(|old| *old == Some(identity))
                    {
                        aliases[i] = aliases[previous];
                    }
                    identities[i] = Some(identity);
                }
                // SAFETY: checked place and exact declared carrier.
                unsafe { read_slot(raw, index, ty.carrier()) }?
            }
            OxCallArg::Const(value) => {
                let owned = from_variant(ty.carrier(), oxvba_runtime::Variant::from_i32(*value))?;
                values[i] = owned;
                continue;
            }
            OxCallArg::Omitted => {
                let default = oxvba_com::typed_slots::plan::default_value(descriptor, indices[i])
                    .ok_or(450)?;
                values[i] = from_variant(ty.carrier(), default)?;
                continue;
            }
            _ => return Err(13),
        };
        // SAFETY: the retained frame owns each exact input through this clone.
        values[i] = unsafe { calcvba_host_api::value::retain_bits(ty.carrier(), bits) }?;
    }
    let mut synchronize = |phase, values: &mut [Value]| -> Result<(), i32> {
        for i in 0..args.len() {
            if aliases[i] != i {
                continue;
            }
            let ty = match descriptor.parameter_types[indices[i]] {
                oxvba_com::TypeLibParamType::ByRefLong => ScalarType::Long,
                oxvba_com::TypeLibParamType::ByRefDouble => ScalarType::Double,
                oxvba_com::TypeLibParamType::ByRefBoolean => ScalarType::Boolean,
                _ => continue,
            };
            let Some(target) = targets[i] else {
                continue;
            };
            match phase {
                CallerSyncPhase::WriteCaller => {
                    let slot = values[i].borrow_slot().ok_or(13)?;
                    if slot.ty() != ty {
                        return Err(13);
                    }
                    let bits = slot.bits();
                    // SAFETY: this exact primitive lvalue and its canonical owner
                    // stay leased; no frame reference crosses callback reentry.
                    unsafe { write_slot(raw, target, ty, &bits) }?;
                }
                CallerSyncPhase::ReadCaller => {
                    // SAFETY: callback finished; read the same exact primitive
                    // lvalue, including its ancestor/global alias authority.
                    let bits = unsafe { read_slot(raw, target, ty) }?;
                    // SAFETY: the retained source owns this exact primitive carrier.
                    values[i] = unsafe { calcvba_host_api::value::retain_bits(ty, bits) }?;
                }
            }
        }
        Ok(())
    };
    let mut arguments = MemberArguments::new(&mut values[..args.len()], &aliases[..args.len()])?
        .with_caller_sync(&mut synchronize);
    let result_type = plan.result_iid.map_or(plan.result, ScalarType::Interface);
    // SAFETY: factory receiver and retained plan plus typed argument owners span
    // this exact planned member callback; no frame/global reference crosses it.
    let result = unsafe {
        invoke_planned_member(
            receiver.argument().as_ptr(),
            plan.iid,
            plan.slot,
            result_type,
            &mut arguments,
            context,
        )
    };
    let mut copyback_error = None;
    for (i, target) in targets[..args.len()].iter().enumerate() {
        if aliases[i] != i {
            continue;
        }
        if let Some(index) = target {
            let value = arguments.get(i).ok_or(5)?;
            let bits = value.borrow_slot().ok_or(13)?.bits();
            // SAFETY: checked destination carrier and live owned callback cell.
            let written = unsafe { write_slot(raw, *index, plan.parameters[i].carrier(), &bits) };
            if let Err(error) = written {
                copyback_error.get_or_insert(error);
            }
        }
    }
    let result = result?;
    if let Some(error) = copyback_error {
        return Err(error);
    }
    if let Some(dst) = dst {
        let index = layout.slot(*dst).map_err(|_| 5)?;
        // SAFETY: one exact owned result transfers into the admitted destination.
        unsafe { write_owned(raw, index, layout.slots[index], result) }?;
    }
    Ok(())
}
