//! Shared value kernels for the additional owning typed carriers. This module
//! executes no CFG or procedure call and never calls VM3. Value coercions may
//! invoke compiler-retained default getters with the active typed context.
use super::activation::{RawActivation, pin_value, read_slot, write_slot};
use super::{ScalarBits, ScalarLayout, ScalarType};
use calcvba_host_api::{Value, value};
use oxvba_eval::arith;
use oxvba_oxir::{OxConst, OxInst, OxOperand, OxPlace};
use oxvba_runtime::{VarType, Variant};

pub fn to_variant(value: Value) -> Variant {
    match value {
        Value::Empty => Variant::empty(),
        Value::Long(v) => Variant::from_i32(v),
        Value::Double(v) => Variant::from_f64(v),
        Value::Boolean(v) => Variant::from_bool(v),
        Value::String(v) => Variant::from_string(v),
        Value::Variant(v) => v,
        Value::Object(Some(v)) => Variant::from_object_ref(v.object_ref()),
        Value::Object(None) => Variant::nothing(),
        Value::RuntimeObject(v) => Variant::from_object_ref(v),
        Value::Record(v) => Variant::from_vba_record(v),
        Value::Array(Some(v)) => Variant::from_safearray(v),
        Value::Array(None) => Variant::unallocated_array(12),
    }
}
pub fn from_variant(ty: ScalarType, value: Variant) -> Result<Value, i32> {
    use oxvba_bundle::NumericCoerceTarget as N;
    Ok(match ty {
        ScalarType::Variant => Value::Variant(value),
        ScalarType::Record => Value::Record(value.as_vba_record().ok_or(13)?),
        ScalarType::Array => {
            if value.vtype() != VarType::ArrayVariant {
                return Err(13);
            }
            Value::Array(value.as_safearray())
        }
        ScalarType::Long => Value::Long(
            arith::coerce_numeric(&value, N::Long)
                .map_err(|e| e.code)?
                .as_i32()
                .ok_or(13)?,
        ),
        ScalarType::Double => Value::Double(
            arith::coerce_numeric(&value, N::Double)
                .map_err(|e| e.code)?
                .as_f64()
                .ok_or(13)?,
        ),
        ScalarType::Boolean => Value::Boolean(
            arith::coerce_numeric(&value, N::Boolean)
                .map_err(|e| e.code)?
                .as_bool()
                .ok_or(13)?,
        ),
        ScalarType::String => Value::String(
            arith::coerce_string(&value)
                .map_err(|e| e.code)?
                .as_bstr()
                .ok_or(13)?,
        ),
        ScalarType::HostObject => {
            if value.vtype() != VarType::Object {
                return Err(13);
            }
            match value.as_object_ref() {
                None => Value::Object(None),
                Some(object) => match calcvba_host_api::HostObject::from_object_ref(&object) {
                    Some(host) => Value::Object(Some(host)),
                    None => Value::RuntimeObject(object),
                },
            }
        }
        _ => return Err(13),
    })
}

/// # Safety
/// raw/layout are the matching retained physical frame and immutable admission
/// facts. No reference into canonical storage survives this read.
pub(super) unsafe fn read_operand(
    raw: *mut RawActivation,
    layout: &ScalarLayout,
    operand: &OxOperand,
    to: ScalarType,
) -> Result<ScalarBits, i32> {
    let from = layout.operand_type(operand).map_err(|_| 5)?;
    let bits = match operand {
        OxOperand::Use(place) => {
            // SAFETY: checked slot and exact carrier.
            unsafe { read_slot(raw, layout.slot(*place).map_err(|_| 5)?, from) }?
        }
        OxOperand::Const(OxConst::Str(text)) => {
            // SAFETY: transfer one owned BSTR into this frame's scratch.
            unsafe { pin_value(raw, Value::String(text.as_str().into())) }?
        }
        OxOperand::Const(OxConst::Empty) => {
            // SAFETY: transfer one canonical Variant into stable scratch storage.
            unsafe { pin_value(raw, Value::Variant(Variant::empty())) }?
        }
        OxOperand::Const(OxConst::Null) => {
            // SAFETY: transfer one canonical Variant into stable scratch storage.
            unsafe { pin_value(raw, Value::Variant(Variant::null())) }?
        }
        _ => layout.read(&[], operand, from),
    };
    if from == to {
        return Ok(bits);
    }
    if from == ScalarType::Long && to == ScalarType::Double {
        return Ok(ScalarBits::double(f64::from(bits.as_long())));
    }
    // SAFETY: read above retains the exact operand through this conversion.
    let owned = unsafe { value::retain_bits(from, bits) }?;
    let result = from_variant(to, to_variant(owned))?;
    // SAFETY: scratch retains any newly created BSTR/Variant/object carrier.
    unsafe { pin_value(raw, result) }
}
pub(super) unsafe fn operand_value(
    raw: *mut RawActivation,
    layout: &ScalarLayout,
    operand: &OxOperand,
) -> Result<Variant, i32> {
    let ty = layout.operand_type(operand).map_err(|_| 5)?;
    // SAFETY: inherited retained frame/layout contract.
    let bits = unsafe { read_operand(raw, layout, operand, ty) }?;
    if ty == ScalarType::Array
        && bits.as_object().is_null()
        && let OxOperand::Use(place) = operand
    {
        let vt = *layout
            .array_slots
            .get(&layout.slot(*place).map_err(|_| 5)?)
            .ok_or(5)?;
        return Ok(Variant::unallocated_array(vt));
    }
    // SAFETY: this exact input remains owned until the completed instruction.
    Ok(to_variant(unsafe { value::retain_bits(ty, bits) }?))
}

pub fn needs_helper(inst: &OxInst, layout: &ScalarLayout) -> bool {
    if super::iteration::is_iteration(inst) || super::aggregate::is_aggregate(inst) {
        return true;
    }
    let extended = |op: &OxOperand| {
        matches!(
            layout.operand_type(op),
            Ok(ScalarType::String
                | ScalarType::Variant
                | ScalarType::HostObject
                | ScalarType::Record
                | ScalarType::Array)
        )
    };
    match inst {
        OxInst::ComCallLate { .. } => !super::stack_instruction(inst, layout),
        OxInst::ValidateAssignment { .. } | OxInst::Not { .. } => true,
        OxInst::CallNative {
            callee:
                oxvba_oxir::OxNativeCallee::Builtin(
                    oxvba_bundle::NativeImplId::IsMissing | oxvba_bundle::NativeImplId::IsArray,
                ),
            ..
        } => true,
        OxInst::ComCallEarly { method, .. } => layout
            .host_calls
            .get(method)
            .is_some_and(|p| p.native.is_some()),
        OxInst::Box { .. }
        | OxInst::CompareObjectIs { .. }
        | OxInst::Unbox { .. }
        | OxInst::Concat { .. }
        | OxInst::VariantChanged { .. } => true,
        OxInst::Coerce { dst, src, .. } => layout.operand_type(src).is_ok_and(|from| {
            layout.slot(*dst).is_ok_and(|i| {
                from != layout.slots[i]
                    && !(from == ScalarType::Long && layout.slots[i] == ScalarType::Double)
            })
        }),
        OxInst::Arith { dst, lhs, rhs, .. }
        | OxInst::Compare { dst, lhs, rhs, .. }
        | OxInst::Div { dst, lhs, rhs } => {
            extended(lhs)
                || extended(rhs)
                || layout
                    .slot(*dst)
                    .is_ok_and(|i| layout.slots[i] == ScalarType::Variant)
        }
        OxInst::Truthy { src, .. } => extended(src),
        OxInst::Neg { dst, src, .. } => {
            extended(src)
                || layout.slot(*dst).is_ok_and(|i| {
                    layout.slots[i] == ScalarType::Variant
                        || layout.operand_type(src).is_ok_and(|from| {
                            from != layout.slots[i]
                                && !(from == ScalarType::Long
                                    && layout.slots[i] == ScalarType::Double)
                        })
                })
        }
        _ => false,
    }
}

/// # Safety
/// raw is this instruction's leased, uniquely executing physical frame. inst
/// points into its retained original OxFunc, whose complete admission was checked.
/// context is the aligned, uniquely executing TypedContext for this activation;
/// no competing reference may be used while a default callback runs.
/// This helper handles selected value operations only, with no execution loop.
pub unsafe extern "C" fn evaluate(
    raw: *mut RawActivation,
    inst: *const OxInst,
    context: *mut super::TypedContext,
) -> i32 {
    crate::with_status(|| {
        // SAFETY: the generated entry retains both immutable descriptors.
        let layout = unsafe { &*(*raw).layout };
        // SAFETY: caller retains the original instruction allocation.
        let inst = unsafe { &*inst };
        if super::iteration::is_iteration(inst) {
            // SAFETY: admitted original instruction, retained frame and unique context.
            return unsafe { super::iteration::evaluate(raw, inst, &mut *context) }
                .map_or_else(|e| e, |()| 0);
        }
        if super::aggregate::is_aggregate(inst) {
            // SAFETY: original admitted instruction and leased physical frame.
            return unsafe { super::aggregate::evaluate(raw, inst) }.map_or_else(|e| e, |()| 0);
        }
        if let OxInst::ComCallEarly { .. } = inst {
            // SAFETY: same leased frame/original instruction and unique context.
            return unsafe { super::member::invoke(raw, inst, &mut *context) }
                .map_or_else(|e| e, |()| 0);
        }
        if let OxInst::ComCallLate { .. } = inst {
            // SAFETY: same admitted frame and original dynamic call descriptor.
            return unsafe { super::member::invoke_dynamic(raw, inst, &mut *context) }
                .map_or_else(|e| e, |()| 0);
        }
        // SAFETY: short operand reads retain their carriers in stable scratch.
        let read = |op: &OxOperand| unsafe { operand_value(raw, layout, op) };
        // SAFETY: the backend retains exclusive synchronous execution authority
        // over this context. No frame/global borrow spans a default callback.
        let read_value =
            |op: &OxOperand| resolve_default(read(op)?, layout, unsafe { &mut *context });
        if let OxInst::ValidateAssignment {
            src,
            target_type_name,
            ..
        } = inst
        {
            return (|| -> Result<(), i32> {
                let value = read(src)?;
                if value.vtype() != VarType::Object {
                    return Err(424);
                }
                if let Some(required) = layout.assignment_classes.get(target_type_name)
                    && let Some(object) = value.as_object_ref()
                    && !super::class_matches(&object, required)
                {
                    return Err(13);
                }
                if let Some(iid) = layout
                    .assignment_iids
                    .get(target_type_name)
                    .copied()
                    .flatten()
                    && let Some(reference) = value.as_object_ref()
                {
                    let object =
                        calcvba_host_api::HostObject::from_object_ref(&reference).ok_or(13)?;
                    object.resolve_interface(iid)?;
                }
                Ok(())
            })()
            .map_or_else(|e| e, |()| 0);
        }
        let result = (|| -> Result<(OxPlace, Value), i32> {
            let (dst, result) = match inst {
                OxInst::CallNative {
                    dst: Some(dst),
                    callee:
                        oxvba_oxir::OxNativeCallee::Builtin(
                            native @ (oxvba_bundle::NativeImplId::IsMissing
                            | oxvba_bundle::NativeImplId::IsArray),
                        ),
                    args,
                } => {
                    let operand = match &args[0] {
                        oxvba_oxir::OxCallArg::Operand(op) => op.clone(),
                        oxvba_oxir::OxCallArg::ByRef(p) => OxOperand::Use(*p),
                        _ => return Err(5),
                    };
                    let value = read(&operand)?;
                    (
                        *dst,
                        Variant::from_bool(if *native == oxvba_bundle::NativeImplId::IsArray {
                            value.vtype() == VarType::ArrayVariant
                        } else {
                            value.vtype() == VarType::Error
                                && value.as_error_code() == Some(0x8002_0004u32 as i32)
                        }),
                    )
                }
                OxInst::Box { dst, src, from: ty } => {
                    let value = read(src)?;
                    let value = if *ty == oxvba_oxir::OxTy::Integer {
                        arith::coerce_numeric(&value, oxvba_bundle::NumericCoerceTarget::Integer)
                            .map_err(|e| e.code)?
                    } else {
                        value
                    };
                    (*dst, value)
                }
                OxInst::CompareObjectIs { dst, lhs, rhs } => {
                    // Is never resolves an object's default member. Check both
                    // tags before reading identity; non-object Variants raise 424.
                    let left = read(lhs)?;
                    let right = read(rhs)?;
                    if left.vtype() != VarType::Object || right.vtype() != VarType::Object {
                        return Err(424);
                    }
                    let same = match (left.as_object_ref(), right.as_object_ref()) {
                        (None, None) => true,
                        (Some(left), Some(right)) => {
                            match (
                                calcvba_host_api::HostObject::from_object_ref(&left),
                                calcvba_host_api::HostObject::from_object_ref(&right),
                            ) {
                                (Some(left), Some(right)) => left.same_identity(&right),
                                // Runtime-owned and foreign runtime IUnknown
                                // carriers compare their retained pointer. Native
                                // COM interfaces enter through prepared tables,
                                // which retain canonical QI identity above.
                                (None, None) => left.raw_iunknown() == right.raw_iunknown(),
                                _ => false,
                            }
                        }
                        _ => false,
                    };
                    (*dst, Variant::from_bool(same))
                }
                OxInst::Unbox {
                    dst, src, checked, ..
                } => {
                    let result = read(src)?;
                    let ty = layout.slots[layout.slot(*dst).map_err(|_| 5)?];
                    if *checked && !matches_type(&result, ty) {
                        return Err(13);
                    }
                    (*dst, result)
                }
                OxInst::Coerce { dst, src, target } => {
                    let src = if matches!(target, oxvba_oxir::OxCoerceTarget::ImplicitVariant) {
                        read(src)?
                    } else {
                        read_value(src)?
                    };
                    let result = match target {
                        oxvba_oxir::OxCoerceTarget::Numeric(target) => {
                            arith::coerce_numeric(&src, *target)
                        }
                        oxvba_oxir::OxCoerceTarget::Str => arith::coerce_string(&src),
                        oxvba_oxir::OxCoerceTarget::ImplicitVariant => Ok(src),
                        _ => return Err(5),
                    }
                    .map_err(|e| e.code)?;
                    (*dst, result)
                }
                OxInst::Arith {
                    dst,
                    lhs,
                    rhs,
                    op,
                    mode,
                } => {
                    let lhs = read_value(lhs)?;
                    let rhs = read_value(rhs)?;
                    let result = match op {
                        oxvba_oxir::ArithOp::Add => arith::add(&lhs, &rhs, *mode),
                        oxvba_oxir::ArithOp::Sub => arith::sub(&lhs, &rhs, *mode),
                        oxvba_oxir::ArithOp::Mul => arith::mul(&lhs, &rhs, *mode),
                        _ => return Err(5),
                    }
                    .map_err(|e| e.code)?;
                    (*dst, result)
                }
                OxInst::Concat { dst, lhs, rhs } => (
                    *dst,
                    arith::concat(&read_value(lhs)?, &read_value(rhs)?).map_err(|e| e.code)?,
                ),
                OxInst::Div { dst, lhs, rhs } => (
                    *dst,
                    arith::div(&read_value(lhs)?, &read_value(rhs)?).map_err(|e| e.code)?,
                ),
                OxInst::Not { dst, src } => {
                    (*dst, arith::not(&read_value(src)?).map_err(|e| e.code)?)
                }
                OxInst::Neg { dst, src, mode } => (
                    *dst,
                    arith::neg(&read_value(src)?, *mode).map_err(|e| e.code)?,
                ),
                OxInst::Truthy { dst, src } => (
                    *dst,
                    Variant::from_bool(arith::is_truthy(&read_value(src)?).map_err(|e| e.code)?),
                ),
                OxInst::Compare {
                    dst,
                    lhs,
                    rhs,
                    mode,
                    op,
                } => {
                    let op = match op {
                        oxvba_oxir::CmpOp::Eq => arith::CmpOp::Eq,
                        oxvba_oxir::CmpOp::Ne => arith::CmpOp::Ne,
                        oxvba_oxir::CmpOp::Lt => arith::CmpOp::Lt,
                        oxvba_oxir::CmpOp::Le => arith::CmpOp::Le,
                        oxvba_oxir::CmpOp::Gt => arith::CmpOp::Gt,
                        oxvba_oxir::CmpOp::Ge => arith::CmpOp::Ge,
                    };
                    (
                        *dst,
                        arith::compare(&read_value(lhs)?, &read_value(rhs)?, *mode, op)
                            .map_err(|e| e.code)?,
                    )
                }
                _ => return Err(5),
            };
            let ty = layout.slots[layout.slot(dst).map_err(|_| 5)?];
            Ok((dst, from_variant(ty, result)?))
        })();
        match result {
            Ok((dst, value)) => {
                let ty = layout.slots[layout.slot(dst).expect("admitted destination")];
                let Some(slot) = value.borrow_slot() else {
                    return 5;
                };
                let bits = slot.bits();
                // SAFETY: the shared kernel produced the exact checked carrier;
                // its owned result remains live through this retaining store.
                unsafe {
                    write_slot(
                        raw,
                        layout.slot(dst).expect("admitted destination"),
                        ty,
                        &bits,
                    )
                }
                .map_or_else(|e| e, |()| 0)
            }
            Err(error) => error,
        }
    })
}
/// Execute admitted unary VBA library functions on their exact input carrier.
pub fn invoke_unary_builtin(
    id: oxvba_bundle::NativeImplId,
    value: Variant,
    layout: &ScalarLayout,
    context: &mut super::TypedContext,
) -> Result<Variant, i32> {
    use oxvba_bundle::NativeImplId::{CStr, IsArray, IsMissing, VarType as TypeCode};
    let value = if matches!(id, CStr | TypeCode) && value.as_object_ref().is_some() {
        resolve_default(value, layout, context)?
    } else {
        value
    };
    match id {
        IsArray => Ok(Variant::from_bool(value.vtype() == VarType::ArrayVariant)),
        IsMissing => Ok(Variant::from_bool(
            value.vtype() == VarType::Error && value.as_error_code() == Some(0x8002_0004u32 as i32),
        )),
        TypeCode => oxvba_lib::variant_type_code(&[value]),
        CStr => oxvba_lib::convert_to_string(&[value]),
        _ => return Err(5),
    }
    .map_err(|e| e.code)
}
pub fn resolve_default(
    mut value: Variant,
    layout: &ScalarLayout,
    context: &mut super::TypedContext,
) -> Result<Variant, i32> {
    for _ in 0..64 {
        if value.vtype() != VarType::Object {
            return Ok(value);
        }
        let reference = value.as_object_ref().ok_or(91)?;
        let object = calcvba_host_api::HostObject::from_object_ref(&reference).ok_or(13)?;
        if object.header().iid == calcvba_host_api::object::IDISPATCH {
            let call = calcvba_host_api::object::DynamicMemberCall {
                name: None,
                kind: calcvba_host_api::object::DynamicMemberKind::Get,
                arguments: vec![],
            };
            value = object
                .invoke_dynamic(&call, &mut [], context)
                .map_err(|e| e.publish(context))?;
            continue;
        }
        let plan = layout
            .default_calls
            .get(&object.header().iid)
            .ok_or_else(|| {
                layout
                    .default_faults
                    .get(&object.header().iid)
                    .copied()
                    .unwrap_or(438)
            })?;
        // SAFETY: the retained factory object and compiler-selected getter plan
        // supply the exact signature. The owned value survives the callback.
        let result = if plan.native.is_some() {
            let mut values = [];
            let mut arguments =
                calcvba_host_api::object::planned_slots::MemberArguments::new(&mut values, &[])?;
            // SAFETY: retained factory object and exact compiler-selected plan;
            // the callback supplies only metadata-declared optional arguments.
            unsafe {
                calcvba_host_api::object::planned_slots::invoke_planned_member(
                    object.argument().as_ptr(),
                    plan.iid,
                    plan.slot,
                    plan.result,
                    &mut arguments,
                    context,
                )
            }
        } else {
            // SAFETY: retained factory object and exact checked getter signature.
            unsafe {
                calcvba_host_api::object::value_slots::invoke_value_slot(
                    object.argument().header(),
                    plan.iid,
                    plan.slot,
                    &plan.parameters,
                    plan.result,
                    &[],
                    context,
                )
            }
        }?;
        value = to_variant(result);
    }
    if value.vtype() == VarType::Object {
        Err(28)
    } else {
        Ok(value)
    }
}
fn matches_type(value: &Variant, ty: ScalarType) -> bool {
    match ty {
        ScalarType::Long => value.vtype() == VarType::Long,
        ScalarType::Double => value.vtype() == VarType::Double,
        ScalarType::Boolean => value.vtype() == VarType::Boolean,
        ScalarType::String => value.vtype() == VarType::String,
        ScalarType::HostObject => value.vtype() == VarType::Object,
        ScalarType::Variant => true,
        ScalarType::Record => value.as_vba_record().is_some(),
        ScalarType::Array => value.vtype() == VarType::ArrayVariant,
        _ => false,
    }
}
