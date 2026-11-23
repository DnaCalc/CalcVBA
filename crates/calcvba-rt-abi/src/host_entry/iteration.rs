//! Live For Each operations over host-owned enumerators, shared by VM3/JIT.
use super::{ScalarLayout, ScalarType as T, activation::*, value::*};
use calcvba_host_api::{TypedContext, Value};
use calcvba_oxir::OxInst;

pub(super) fn is_iteration(inst: &OxInst) -> bool {
    matches!(
        inst,
        OxInst::ForEachInit { .. } | OxInst::ForEachNext { .. } | OxInst::ForEachEnd { .. }
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
            if !matches!(layout.operand_type(source)?, T::HostObject | T::Variant) {
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
        // Host enumeration holds no array lock, so the loop's end has nothing to release.
        OxInst::ForEachEnd { iter } => variant(*iter)?,
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
            let index = layout.slot(*iter).map_err(|_| 5)?;
            // SAFETY: detach any previous iteration before starting a new one.
            drop(unsafe { (*raw).iterators.remove(&index) });
            let iterator = if let Some(arr) = value.as_safearray() {
                let mut items = arr.enumeration_elements().unwrap_or_default().into_iter();
                calcvba_host_api::HostEnumerator::new(move |_| Ok(items.next().map(Value::Variant)))
            } else if let Some(reference) = value.as_object_ref() {
                if calcvba_runtime::ObjectRefExt::is_compat_object(&reference)
                    && calcvba_runtime::ObjectRefExt::route_key(&reference)
                        == crate::VBA_COLLECTION_ROUTE_KEY
                {
                    if let Some(cursor) =
                        calcvba_runtime::collection::CollectionCursor::new(&reference)
                    {
                        calcvba_host_api::HostEnumerator::new(move |_| {
                            Ok(cursor.next_item().map(Value::Variant))
                        })
                    } else {
                        let mut items =
                            calcvba_runtime::ObjectRefExt::native_collection_snapshot(&reference)
                                .unwrap_or_default()
                                .into_iter();
                        calcvba_host_api::HostEnumerator::new(move |_| {
                            Ok(items.next().map(Value::Variant))
                        })
                    }
                } else {
                    let object = super::member::host_view(&reference).ok_or(438)?;
                    object.enumerator(context)?
                }
            } else if value.vtype() == calcvba_runtime::VarType::Object {
                return Err(91);
            } else {
                return Err(13);
            };
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
        OxInst::ForEachEnd { .. } => Ok(()),
        _ => Err(5),
    }
}
