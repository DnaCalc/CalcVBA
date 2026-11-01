//! One retained descriptor-planned member operation, shared by both backends.
use super::{
    ScalarType,
    activation::{RawActivation, read_slot, slot_identity, write_owned, write_slot},
    value::{from_variant, read_operand},
};
use calcvba_host_api::{
    HostObject, TypedContext, Value,
    object::planned_slots::{MemberArguments, invoke_planned_member},
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
    // SAFETY: owning frame reads return retained factory headers only.
    let receiver = unsafe { HostObject::retain_header(header) };
    let mut values: [Value; calcvba_host_api::MAX_TYPED_ARGS] =
        std::array::from_fn(|_| Value::Empty);
    let mut aliases = std::array::from_fn::<_, { calcvba_host_api::MAX_TYPED_ARGS }, _>(|i| i);
    let mut targets = [None; calcvba_host_api::MAX_TYPED_ARGS];
    let mut identities = [None; calcvba_host_api::MAX_TYPED_ARGS];
    let indices = oxvba_com::typed_slots::plan::visible_indices(descriptor);
    for (i, (arg, ty)) in args.iter().zip(&plan.parameters).enumerate() {
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
    let mut arguments = MemberArguments::new(&mut values[..args.len()], &aliases[..args.len()])?;
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
