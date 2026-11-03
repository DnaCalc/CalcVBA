//! Live For Each operations over host-owned enumerators, shared by VM3/JIT.
use super::{ScalarLayout, ScalarType as T, activation::*, value::*};
use calcvba_host_api::{HostObject, TypedContext, Value};
use calcvba_oxir::OxInst;

pub(super) fn is_iteration(inst: &OxInst) -> bool {
    matches!(
        inst,
        OxInst::ForEachInit { .. } | OxInst::ForEachNext { .. }
    )
}
pub(super) fn admit(layout: &ScalarLayout, inst: &OxInst) -> Result<bool, String> {
    let variant = |p| -> Result<(), String> {
        let index = layout.slot(p)?;
        if index >= layout.global_base || layout.slots[index] != T::Variant {
            return Err("typed iterator requires a frame-owned Variant slot".into());
        }
        Ok(())
    };
    match inst {
        OxInst::ForEachInit { iter, source } => {
            variant(*iter)?;
            if layout.operand_type(source)? != T::HostObject {
                return Err("typed host enumeration requires an object carrier".into());
            }
        }
        OxInst::ForEachNext {
            iter,
            item,
            has_value,
        } => {
            variant(*iter)?;
            variant(*item)?;
            if layout.slots[layout.slot(*has_value)?] != T::Boolean {
                return Err("typed enumeration condition requires Boolean".into());
            }
        }
        _ => return Ok(false),
    }
    Ok(true)
}
/// # Safety
/// raw/layout/instruction are the same retained, admitted frame. No frame
/// reference spans a host iterator callback or a native event reentry.
pub(super) unsafe fn evaluate(
    raw: *mut RawActivation,
    inst: &OxInst,
    context: &mut TypedContext,
) -> Result<(), i32> {
    // SAFETY: the activation retains its immutable admission facts.
    let layout = unsafe { &*(*raw).layout };
    match inst {
        OxInst::ForEachInit { iter, source } => {
            // SAFETY: read retains the actual source in frame-owned scratch.
            let value = unsafe { operand_value(raw, layout, source) }?;
            let reference = value.as_object_ref().ok_or(424)?;
            let object = HostObject::from_object_ref(&reference).ok_or(438)?;
            let index = layout.slot(*iter).map_err(|_| 5)?;
            // SAFETY: detach any previous iteration before starting a new one.
            drop(unsafe { (*raw).iterators.remove(&index) });
            let iterator = object.enumerator(context)?;
            // SAFETY: callbacks have returned; replace only this frame's owner.
            drop(unsafe { (*raw).iterators.insert(index, iterator) });
            Ok(())
        }
        OxInst::ForEachNext {
            iter,
            item,
            has_value,
        } => {
            let index = layout.slot(*iter).map_err(|_| 5)?;
            // Detach the owner so no map/frame borrow spans native Next/reentry.
            // SAFETY: the uniquely executing frame owns this iterator map.
            let mut iterator = unsafe { (*raw).iterators.remove(&index) }.ok_or(5)?;
            let value = iterator.advance(context)?;
            let present = value.is_some();
            if let Some(value) = value {
                // SAFETY: the source and complete frame remain retained.
                unsafe {
                    (*raw).iterators.insert(index, iterator);
                }
                let item = layout.slot(*item).map_err(|_| 5)?;
                // SAFETY: the admitted destination is one actual Variant cell.
                unsafe { write_owned(raw, item, T::Variant, Value::Variant(to_variant(value))) }?;
            }
            let dst = layout.slot(*has_value).map_err(|_| 5)?;
            // SAFETY: the admitted condition slot owns one Boolean value.
            unsafe { write_owned(raw, dst, T::Boolean, Value::Boolean(present)) }
        }
        _ => Err(5),
    }
}
