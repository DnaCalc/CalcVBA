//! One-operation record/array kernels over the original admitted instruction.
use super::{ScalarLayout, ScalarType as T, activation::*, value::*};
use calcvba_host_api::Value;
use oxvba_oxir::{OxInst, OxOperand, OxPlace};
use oxvba_runtime::{
    Variant,
    safe_array::{SafeArray, SafeArrayBound},
    vba_record::VbaRecord,
};

pub(super) fn is_aggregate(inst: &OxInst) -> bool {
    matches!(
        inst,
        OxInst::NewRecord { .. }
            | OxInst::RecordGet { .. }
            | OxInst::RecordSet { .. }
            | OxInst::RecordLSet { .. }
            | OxInst::RecordArrayGet { .. }
            | OxInst::RecordArraySet { .. }
            | OxInst::ArrayLiteral { .. }
            | OxInst::ArrayRedim { .. }
            | OxInst::ArrayAppend { .. }
            | OxInst::ArrayGet { .. }
            | OxInst::ArraySet { .. }
            | OxInst::Bound { .. }
    )
}
pub(super) fn admit(layout: &ScalarLayout, inst: &OxInst) -> Result<bool, String> {
    if !is_aggregate(inst) {
        return Ok(false);
    }
    let operand = |op| layout.operand_type(op).map(|_| ());
    let place = |p| layout.slot(p).map(|_| ());
    let array = |op| -> Result<(), String> {
        match layout.operand_type(op)? {
            T::Array | T::Variant => Ok(()),
            _ => Err("typed array receiver requires an array carrier".into()),
        }
    };
    let record = |op| -> Result<(), String> {
        match layout.operand_type(op)? {
            T::Record => Ok(()),
            _ => Err("typed record receiver requires a record carrier".into()),
        }
    };
    match inst {
        OxInst::ArrayRedim {
            dst,
            upper_bounds,
            lower_bounds,
            ..
        } => {
            if upper_bounds.is_empty()
                || upper_bounds.len() > 60
                || lower_bounds.len() > upper_bounds.len()
            {
                return Err("invalid fixed array bound descriptor".into());
            }
            array(&OxOperand::Use(*dst))?;
            for bound in upper_bounds.iter().chain(lower_bounds) {
                operand(bound)?;
            }
        }
        OxInst::NewRecord { dst, fields } => {
            let slot = layout.slot(*dst)?;
            let expected = layout
                .record_slots
                .get(&slot)
                .ok_or("missing record slot descriptor")?;
            if **expected != *oxvba_bundle::vba_record_layout_for_fields(fields)? {
                return Err("record initializer layout mismatch".into());
            }
        }
        OxInst::RecordGet {
            dst,
            record: source,
            ..
        } => {
            record(source)?;
            place(*dst)?;
        }
        OxInst::RecordSet {
            record: target,
            value,
            ..
        }
        | OxInst::RecordLSet {
            record: target,
            value,
        } => {
            record(&OxOperand::Use(*target))?;
            operand(value)?;
        }
        OxInst::RecordArrayGet {
            dst,
            record: source,
            indices,
            ..
        } => {
            record(source)?;
            place(*dst)?;
            for index in indices {
                operand(index)?;
            }
        }
        OxInst::RecordArraySet {
            record: target,
            indices,
            value,
            ..
        } => {
            record(&OxOperand::Use(*target))?;
            operand(value)?;
            for index in indices {
                operand(index)?;
            }
        }
        OxInst::ArrayLiteral {
            dst,
            values,
            aliases,
            ..
        } => {
            array(&OxOperand::Use(*dst))?;
            if !aliases.is_empty() && aliases.len() != values.len() {
                return Err("array alias extent mismatch".into());
            }
            for value in values {
                operand(value)?;
            }
            for alias in aliases.iter().flatten() {
                place(*alias)?;
            }
        }
        OxInst::ArrayAppend {
            dst,
            array: source,
            item,
        } => {
            array(source)?;
            array(&OxOperand::Use(*dst))?;
            operand(item)?;
        }
        OxInst::ArrayGet {
            dst,
            array: source,
            indices,
        } => {
            array(source)?;
            place(*dst)?;
            for index in indices {
                operand(index)?;
            }
        }
        OxInst::ArraySet {
            array: target,
            indices,
            value,
        } => {
            array(&OxOperand::Use(*target))?;
            operand(value)?;
            for index in indices {
                operand(index)?;
            }
        }
        OxInst::Bound {
            dst,
            array: source,
            dimension,
            ..
        } => {
            array(source)?;
            place(*dst)?;
            if let Some(dimension) = dimension {
                operand(dimension)?;
            }
        }
        _ => return Ok(false),
    }
    Ok(true)
}
pub fn array_flat_index(bounds: &[SafeArrayBound], indices: &[i64]) -> Result<usize, i32> {
    if bounds.len() != indices.len() {
        return Err(9);
    }
    let mut offset = 0usize;
    for (bound, &index) in bounds.iter().zip(indices) {
        let index = index.checked_sub(i64::from(bound.lower)).ok_or(9)?;
        if index < 0 || index >= i64::from(bound.count) {
            return Err(9);
        }
        // Runtime SAFEARRAY values use first-dimension-outermost storage;
        // the native Automation boundary converts its descriptor ordering.
        offset = offset
            .checked_mul(bound.count as usize)
            .and_then(|offset| offset.checked_add(index as usize))
            .ok_or(9)?;
    }
    Ok(offset)
}
// SAFETY CONTRACT: raw and inst are the matching retained, uniquely executing
// activation and original admitted instruction. No borrowed carrier escapes.
pub(super) unsafe fn evaluate(raw: *mut RawActivation, inst: &OxInst) -> Result<(), i32> {
    // SAFETY: retained immutable admission facts span this operation.
    let layout = unsafe { &*(*raw).layout };
    // SAFETY: operand_value snapshots leased carriers and never runs a callback.
    let read = |op: &OxOperand| unsafe { operand_value(raw, layout, op) };
    let indices = |ops: &[OxOperand]| {
        ops.iter()
            .map(|op| oxvba_eval::arith::int(&read(op)?).map_err(|e| e.code))
            .collect::<Result<Vec<_>, i32>>()
    };
    let store = |dst: OxPlace, value: Variant| {
        let index = layout.slot(dst).map_err(|_| 5)?;
        let ty = layout.slots[index];
        let value = from_variant(ty, value)?;
        // SAFETY: exact owned result transfers into its admitted destination.
        unsafe { write_owned(raw, index, ty, value) }
    };
    match inst {
        OxInst::ArrayRedim {
            dst,
            upper_bounds,
            lower_bounds,
            element,
            preserve,
            fixed,
            explicit_element,
        } => {
            let previous = read(&OxOperand::Use(*dst))?;
            let selected = oxvba_bundle::array_runtime::resolve_redim_element(
                &previous,
                element,
                !*explicit_element
                    && layout.operand_type(&OxOperand::Use(*dst)).map_err(|_| 5)? == T::Variant,
                *preserve,
            )?;
            let element = &selected;
            let old = previous.as_safearray();
            if !fixed && old.as_ref().is_some_and(SafeArray::is_fixed_size) {
                return Err(10);
            }
            let mut bounds = Vec::with_capacity(upper_bounds.len());
            let mut count = 1usize;
            for (i, upper) in upper_bounds.iter().enumerate() {
                let lower = lower_bounds
                    .get(i)
                    .map(|op| {
                        i32::try_from(oxvba_eval::arith::int(&read(op)?).map_err(|e| e.code)?)
                            .map_err(|_| 6)
                    })
                    .transpose()?
                    .unwrap_or(0);
                let upper =
                    i32::try_from(oxvba_eval::arith::int(&read(upper)?).map_err(|e| e.code)?)
                        .map_err(|_| 6)?;
                if upper < lower {
                    return Err(9);
                }
                let length =
                    u32::try_from(i64::from(upper) - i64::from(lower) + 1).map_err(|_| 7)?;
                count = count.checked_mul(length as usize).ok_or(7)?;
                bounds.push(SafeArrayBound {
                    lower,
                    count: length,
                });
            }
            let mut values = Vec::new();
            values.try_reserve_exact(count).map_err(|_| 7)?;
            for _ in 0..count {
                values.push(oxvba_bundle::default_array_element(element).map_err(|_| 7)?);
            }
            if *preserve && let Some(old) = old {
                let old_bounds = old.bounds().ok_or(9)?;
                let rank = bounds.len();
                if old_bounds.len() != rank
                    || old_bounds[..rank - 1] != bounds[..rank - 1]
                    || old_bounds[rank - 1].lower != bounds[rank - 1].lower
                {
                    return Err(9);
                }
                let old_width = old_bounds[rank - 1].count as usize;
                let new_width = bounds[rank - 1].count as usize;
                for (index, value) in old.variant_elements().ok_or(13)?.into_iter().enumerate() {
                    let column = index % old_width;
                    if column < new_width {
                        values[index / old_width * new_width + column] = value;
                    }
                }
            }
            let array =
                oxvba_bundle::redim_safearray_from_elements(bounds, element, values, *fixed)
                    .map_err(|_| 13)?;
            store(*dst, Variant::from_safearray(array))
        }
        OxInst::NewRecord { dst, .. } => {
            let index = layout.slot(*dst).map_err(|_| 5)?;
            let record = VbaRecord::new_default(layout.record_slots.get(&index).ok_or(5)?.clone())
                .map_err(|_| 7)?;
            // SAFETY: the checked destination owns one exact record carrier.
            unsafe { write_owned(raw, index, T::Record, Value::Record(record)) }
        }
        OxInst::RecordGet { dst, record, index } => store(
            *dst,
            read(record)?
                .read_record_field_variant(*index)
                .map_err(|_| 13)?,
        ),
        OxInst::RecordSet {
            record,
            index,
            value,
        } => {
            let value = read(value)?;
            let mut target = read(&OxOperand::Use(*record))?;
            target
                .write_record_field_variant(*index, &value)
                .map_err(|_| 13)?;
            store(*record, target)
        }
        OxInst::RecordLSet { record, value } => {
            let value = read(value)?;
            let mut target = read(&OxOperand::Use(*record))?;
            target.lset_record_from(&value).map_err(|_| 13)?;
            store(*record, target)
        }
        OxInst::RecordArrayGet {
            dst,
            record,
            index,
            indices: subscripts,
        } => {
            let record = read(record)?;
            let indices = indices(subscripts)?;
            let (bounds, _) = record
                .record_array_field_bounds_len(*index)
                .map_err(|_| 13)?
                .ok_or(9)?;
            let item = record
                .record_array_field_element(*index, array_flat_index(&bounds, &indices)?)
                .map_err(|_| 9)?
                .ok_or(9)?;
            store(*dst, item)
        }
        OxInst::RecordArraySet {
            record,
            index,
            indices: subscripts,
            value,
        } => {
            let indices = indices(subscripts)?;
            let value = read(value)?;
            let mut target = read(&OxOperand::Use(*record))?;
            let (bounds, _) = target
                .record_array_field_bounds_len(*index)
                .map_err(|_| 13)?
                .ok_or(9)?;
            target
                .set_record_array_field_element(
                    *index,
                    array_flat_index(&bounds, &indices)?,
                    &value,
                )
                .map_err(|_| 13)?
                .ok_or(9)?;
            store(*record, target)
        }
        OxInst::ArrayLiteral {
            dst,
            values,
            lower_bound,
            aliases,
        } => {
            let values = values.iter().map(read).collect::<Result<Vec<_>, _>>()?;
            let count = u32::try_from(values.len()).map_err(|_| 7)?;
            let array = SafeArray::from_variants_nd(
                vec![SafeArrayBound {
                    lower: *lower_bound,
                    count,
                }],
                values,
            );
            store(*dst, Variant::from_safearray(array))?;
            let index = layout.slot(*dst).map_err(|_| 5)?;
            let aliases = aliases
                .iter()
                .map(|place| {
                    place
                        .map(|p| layout.slot(p).map(|index| Alias { frame: raw, index }))
                        .transpose()
                })
                .collect::<Result<Vec<_>, _>>()
                .map_err(|_| 5)?;
            // SAFETY: inert aliases refer only to this stable frame or ancestors.
            unsafe {
                (*raw).param_arrays.insert(index, aliases);
            }
            Ok(())
        }
        OxInst::ArrayAppend { dst, array, item } => {
            let array = read(array)?.as_safearray().ok_or(13)?;
            let mut values = array.variant_elements().ok_or(13)?;
            values.push(read(item)?);
            store(
                *dst,
                Variant::from_safearray(SafeArray::from_variants(values)),
            )
        }
        OxInst::Bound {
            dst,
            which,
            array,
            dimension,
        } => {
            let value = read(array)?;
            let (bounds, _) = value.safearray_bounds_len().ok_or(9)?;
            let dim = dimension
                .as_ref()
                .map(|op| oxvba_eval::arith::int(&read(op)?).map_err(|e| e.code))
                .transpose()?
                .unwrap_or(1);
            let dim = dim
                .checked_sub(1)
                .and_then(|dim| usize::try_from(dim).ok())
                .ok_or(9)?;
            let bound = bounds.get(dim).ok_or(9)?;
            let value = match which {
                oxvba_oxir::BoundWhich::Lower => bound.lower,
                oxvba_oxir::BoundWhich::Upper => {
                    i32::try_from(i64::from(bound.lower) + i64::from(bound.count) - 1)
                        .map_err(|_| 6)?
                }
            };
            store(*dst, Variant::from_i32(value))
        }
        OxInst::ArrayGet {
            dst,
            array,
            indices: subscripts,
        } => {
            let value = read(array)?;
            let (bounds, _) = value.safearray_bounds_len().ok_or(9)?;
            let flat = array_flat_index(&bounds, &indices(subscripts)?)?;
            let alias = if let OxOperand::Use(place) = array {
                // SAFETY: alias metadata is retained by this boxed stack.
                unsafe { array_aliases(raw, layout.slot(*place).map_err(|_| 5)?) }
                    .and_then(|v| v.get(flat).copied().flatten())
            } else {
                None
            };
            let item = if let Some(alias) = alias {
                // SAFETY: this authority names a live ancestor slot.
                unsafe { alias_value(alias) }?
            } else {
                value.safearray_element(flat).ok_or(9)?.map_err(|_| 9)?
            };
            store(*dst, item)
        }
        OxInst::ArraySet {
            array,
            indices: subscripts,
            value,
        } => {
            let indices = indices(subscripts)?;
            let value = read(value)?;
            let mut target = read(&OxOperand::Use(*array))?;
            let (bounds, _) = target.safearray_bounds_len().ok_or(9)?;
            let flat = array_flat_index(&bounds, &indices)?;
            // SAFETY: alias metadata is retained by this boxed stack.
            let aliases = unsafe { array_aliases(raw, layout.slot(*array).map_err(|_| 5)?) };
            if let Some(Some(alias)) = aliases.as_ref().and_then(|a| a.get(flat)) {
                // Validate and publish the exact caller type before changing the
                // array cell, so failed coercions cannot leave a partial write.
                // SAFETY: aliases retain their source slots for this whole call.
                unsafe { write_alias(*alias, value.clone()) }?;
                // SAFETY: read the coerced caller value, preserving Long rounding.
                let coerced = unsafe { alias_value(*alias) }?;
                for (index, other) in aliases
                    .as_ref()
                    .expect("present aliases")
                    .iter()
                    .enumerate()
                {
                    if let Some(other) = other {
                        // SAFETY: compare inert identities of retained caller cells.
                        let same = unsafe {
                            slot_identity(alias.frame, alias.index)?
                                == slot_identity(other.frame, other.index)?
                        };
                        if same {
                            target
                                .set_safearray_element(index, &coerced)
                                .map_err(|_| 13)?;
                        }
                    }
                }
            } else {
                target.set_safearray_element(flat, &value).map_err(|_| 13)?;
            }
            store(*array, target)
        }
        _ => Err(5),
    }
}
