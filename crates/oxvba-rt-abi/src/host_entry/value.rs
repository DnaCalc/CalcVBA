//! Shared value kernels for the additional owning typed carriers. This module
//! executes no CFG, procedure call or host dispatch and never calls VM3.
use super::activation::{RawActivation, pin_value, read_slot, write_slot};
use super::{ScalarBits, ScalarLayout, ScalarType};
use calcvba_host_api::{Value, value};
use oxvba_eval::arith;
use oxvba_oxir::{OxConst, OxInst, OxOperand, OxPlace};
use oxvba_runtime::{VarType, Variant};

pub(super) fn to_variant(value: Value) -> Variant {
    match value {
        Value::Empty => Variant::empty(),
        Value::Long(v) => Variant::from_i32(v),
        Value::Double(v) => Variant::from_f64(v),
        Value::Boolean(v) => Variant::from_bool(v),
        Value::String(v) => Variant::from_string(v),
        Value::Variant(v) => v,
        Value::Object(Some(v)) => Variant::from_object_ref(v.object_ref()),
        Value::Object(None) => Variant::nothing(),
    }
}
pub(super) fn from_variant(ty: ScalarType, value: Variant) -> Result<Value, i32> {
    use oxvba_bundle::NumericCoerceTarget as N;
    Ok(match ty {
        ScalarType::Variant => Value::Variant(value),
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
            Value::Object(
                value
                    .as_object_ref()
                    .map(|object| calcvba_host_api::HostObject::from_object_ref(&object).ok_or(13))
                    .transpose()?,
            )
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
unsafe fn operand_value(
    raw: *mut RawActivation,
    layout: &ScalarLayout,
    operand: &OxOperand,
) -> Result<Variant, i32> {
    let ty = layout.operand_type(operand).map_err(|_| 5)?;
    // SAFETY: inherited retained frame/layout contract.
    let bits = unsafe { read_operand(raw, layout, operand, ty) }?;
    // SAFETY: this exact input remains owned until the completed instruction.
    Ok(to_variant(unsafe { value::retain_bits(ty, bits) }?))
}

pub fn needs_helper(inst: &OxInst, layout: &ScalarLayout) -> bool {
    let extended = |op: &OxOperand| {
        matches!(
            layout.operand_type(op),
            Ok(ScalarType::String | ScalarType::Variant)
        )
    };
    match inst {
        OxInst::Box { .. }
        | OxInst::Unbox { .. }
        | OxInst::Concat { .. }
        | OxInst::VariantChanged { .. } => true,
        OxInst::Coerce { dst, src, .. } => layout.operand_type(src).is_ok_and(|from| {
            layout.slot(*dst).is_ok_and(|i| {
                from != layout.slots[i]
                    && !(from == ScalarType::Long && layout.slots[i] == ScalarType::Double)
            })
        }),
        OxInst::Arith { dst, lhs, rhs, .. } | OxInst::Compare { dst, lhs, rhs, .. } => {
            extended(lhs)
                || extended(rhs)
                || layout
                    .slot(*dst)
                    .is_ok_and(|i| layout.slots[i] == ScalarType::Variant)
        }
        OxInst::Truthy { src, .. } | OxInst::Neg { src, .. } => extended(src),
        _ => false,
    }
}

/// # Safety
/// raw is this instruction's leased, uniquely executing physical frame. inst
/// points into its retained original OxFunc, whose complete admission was checked.
/// This helper handles selected value operations only, with no execution loop.
pub unsafe extern "C" fn evaluate(raw: *mut RawActivation, inst: *const OxInst) -> i32 {
    crate::with_status(|| {
        // SAFETY: the generated entry retains both immutable descriptors.
        let layout = unsafe { &*(*raw).layout };
        // SAFETY: caller retains the original instruction allocation.
        let inst = unsafe { &*inst };
        // SAFETY: short operand reads retain their carriers in stable scratch.
        let read = |op: &OxOperand| unsafe { operand_value(raw, layout, op) };
        let result = (|| -> Result<(OxPlace, Value), i32> {
            let (dst, result) = match inst {
                OxInst::Box { dst, src, .. } => (*dst, read(src)?),
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
                    let src = read(src)?;
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
                    let lhs = read(lhs)?;
                    let rhs = read(rhs)?;
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
                    arith::concat(&read(lhs)?, &read(rhs)?).map_err(|e| e.code)?,
                ),
                OxInst::Neg { dst, src, mode } => {
                    (*dst, arith::neg(&read(src)?, *mode).map_err(|e| e.code)?)
                }
                OxInst::Truthy { dst, src } => (
                    *dst,
                    Variant::from_bool(arith::is_truthy(&read(src)?).map_err(|e| e.code)?),
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
                        arith::compare(&read(lhs)?, &read(rhs)?, *mode, op).map_err(|e| e.code)?,
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
fn matches_type(value: &Variant, ty: ScalarType) -> bool {
    match ty {
        ScalarType::Long => value.vtype() == VarType::Long,
        ScalarType::Double => value.vtype() == VarType::Double,
        ScalarType::Boolean => value.vtype() == VarType::Boolean,
        ScalarType::String => value.vtype() == VarType::String,
        ScalarType::HostObject => value.vtype() == VarType::Object,
        ScalarType::Variant => true,
        _ => false,
    }
}
