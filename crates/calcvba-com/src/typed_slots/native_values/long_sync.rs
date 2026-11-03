//! Exact Long/Double/Boolean native cells synchronized with their owning caller lvalues.
use super::{MemberArguments, TypedFault, Value, fault};

pub(super) enum Pointer {
    Long(*mut i32),
    Double(*mut f64),
    Boolean(*mut i16),
}

pub(super) struct Binding {
    pub source: usize,
    pub pointer: Pointer,
}
struct Call<'a, 'b> {
    arguments: *mut MemberArguments<'a>,
    bindings: &'b [Binding],
}

// SAFETY CONTRACT: data borrows the same-thread Call below; native pointers are
// the original canonical ByRef LONG/DOUBLE/VARIANT_BOOL authorities, live and immobile through the
// foreign call. No frame/cell reference survives either short synchronization.
unsafe fn publish(data: *mut ()) -> Result<(), i32> {
    // SAFETY: registration retains exactly this Call and exclusive arguments.
    let call = unsafe { &*data.cast::<Call<'_, '_>>() };
    // SAFETY: no competing argument access occurs during a native callback.
    let arguments = unsafe { &mut *call.arguments };
    for binding in call.bindings {
        // SAFETY: the pointer variant names the exact initialized native ABI cell.
        let value = unsafe {
            match binding.pointer {
                Pointer::Long(pointer) => Value::Long(pointer.read()),
                Pointer::Double(pointer) => Value::Double(pointer.read()),
                Pointer::Boolean(pointer) => Value::Boolean(pointer.read() != 0),
            }
        };
        arguments.set(binding.source, value)?;
    }
    arguments.write_caller_values()
}
unsafe fn refresh(data: *mut ()) -> Result<(), i32> {
    // SAFETY: same retained Call/argument lease as publish, after VBA returns.
    let call = unsafe { &*data.cast::<Call<'_, '_>>() };
    // SAFETY: callback body has completed; this is the only argument borrower.
    let arguments = unsafe { &mut *call.arguments };
    arguments.read_caller_values()?;
    for binding in call.bindings {
        // SAFETY: reuse the original native authority; no containing Cell is
        // reborrowed. The exact source carrier must still match that authority.
        unsafe {
            match (&binding.pointer, arguments.get(binding.source)) {
                (Pointer::Long(pointer), Some(Value::Long(value))) => pointer.write(*value),
                (Pointer::Double(pointer), Some(Value::Double(value))) => pointer.write(*value),
                (Pointer::Boolean(pointer), Some(Value::Boolean(value))) => {
                    pointer.write(if *value { -1 } else { 0 })
                }
                _ => return Err(13),
            }
        };
    }
    Ok(())
}

/// # Safety
/// Bindings retain initialized native cells, unique per repeated alias, and
/// arguments owns their exact source values for the whole foreign operation.
pub(super) unsafe fn with_sync<R>(
    arguments: &mut MemberArguments<'_>,
    bindings: &[Binding],
    body: impl FnOnce() -> R,
) -> Result<R, TypedFault> {
    if bindings.is_empty() {
        return Ok(body());
    }
    let mut call = Call {
        arguments: std::ptr::from_mut(arguments),
        bindings,
    };
    // SAFETY: this stack call and its source/native owners span exactly body;
    // hooks use short raw-cell/source operations and cannot enter VBA themselves.
    unsafe {
        calcvba_host_api::with_native_callback_sync(
            std::ptr::from_mut(&mut call).cast(),
            publish,
            refresh,
            body,
        )
    }
    .map_err(fault)
}
