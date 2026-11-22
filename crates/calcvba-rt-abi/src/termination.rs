use crate::{ExecState, Fault, ProcInvokeBridge, RawExecState};
use calcvba_oxir::FuncId;
use calcvba_runtime::Variant;
use std::ptr::NonNull;

type StateAccess<'h, T> = for<'a> fn(&'a mut T) -> &'a mut ExecState<'h>;

/// Drain queued Class_Terminate procedures and their releases on this session's
/// owning thread. The engine supplies its procedure entry and short state borrows.
pub fn drain_terminations<'h, T>(
    context: &mut T,
    state: StateAccess<'h, T>,
    mut invoke: impl FnMut(&mut T, usize, FuncId, Variant),
) {
    let exec = state(context);
    if exec.draining {
        return;
    }
    let _owner = exec.object_owner.enter();
    exec.draining = true;
    let guard = DrainGuard { context, state };

    while calcvba_runtime::has_pending_terminations() {
        // Deferred owner-thread drops may call host code before giving us the batch.
        let pending = calcvba_runtime::take_pending_terminations();
        for work in resolve_termination_batch(state(guard.context), pending) {
            if let (Some(proc), Some(object)) = (work.terminate, work.object) {
                let saved_err = (*state(guard.context).err_engine).clone();
                invoke(guard.context, work.bundle, proc, object);
                *state(guard.context).err_engine = saved_err;
            }
            calcvba_runtime::finish_pending_termination(work.instance_id);
            cleanup_terminated_owner(guard.context, state, work.instance_id);
        }
    }
}

/// One pending `Class_Terminate` callback resolved against the shared loaded-image tables.
struct TerminationWork {
    instance_id: i32,
    bundle: usize,
    terminate: Option<FuncId>,
    object: Option<Variant>,
}

/// Take the current pending-termination batch and resolve each item to its owning proc/object.
fn resolve_termination_batch(
    exec: &ExecState<'_>,
    pending: Vec<(i32, i32, i32)>,
) -> Vec<TerminationWork> {
    pending
        .into_iter()
        .map(|(instance_id, bundle_id, route_key)| {
            let bundle = bundle_id as usize;
            let terminate = exec
                .programs
                .get(bundle)
                .and_then(|lp| lp.program.classes.get(route_key as usize))
                .and_then(|class| class.terminate);
            let object = calcvba_runtime::retained_parked_termination_object(instance_id)
                .map(Variant::from_object_ref);
            TerminationWork {
                instance_id,
                bundle,
                terminate,
                object,
            }
        })
        .collect()
}

pub(super) fn cleanup_terminated_owner<'h, T>(
    context: &mut T,
    state: StateAccess<'h, T>,
    instance: i32,
) {
    let owned_key = |key: i64| (key >> 32) as i32 == instance;
    // Unadvise can deliver events that install another connection. Drop each
    // batch outside the state borrow and repeat until the owner has no sinks.
    loop {
        let connections = state(context).take_com_connections(owned_key);
        if connections.is_empty() {
            break;
        }
        drop(connections);
    }
    state(context)
        .events
        .withevents
        .retain(|key, _| !owned_key(*key));
}

struct DrainGuard<'a, 'h, T> {
    context: &'a mut T,
    state: StateAccess<'h, T>,
}

impl<T> Drop for DrainGuard<'_, '_, T> {
    fn drop(&mut self) {
        (self.state)(self.context).draining = false;
    }
}

struct BridgeState<'h>(NonNull<ExecState<'h>>);

impl<'h> BridgeState<'h> {
    fn exec(&mut self) -> &mut ExecState<'h> {
        // SAFETY: the raw drain owns a live same-thread state and permits synchronous
        // reentry only after each state borrow has ended.
        unsafe { self.0.as_mut() }
    }
}

// SAFETY CONTRACT: `state` is the live same-thread execution-state handle for the
// complete synchronous drain. No typed state borrow is retained across callbacks.
pub(super) unsafe fn maybe_drain_with_bridge(state: *mut RawExecState) -> Result<(), Fault> {
    let mut state = BridgeState(
        NonNull::new(state.cast::<ExecState<'_>>())
            .ok_or_else(|| Fault::new(5, "runtime ABI state pointer is null"))?,
    );
    let exec = state.exec();
    if exec.draining {
        return Ok(());
    }
    let Some(ProcInvokeBridge { ctx, invoke }) = exec.proc_invoker else {
        let _owner = exec.object_owner.enter();
        if calcvba_runtime::has_pending_terminations() {
            return Err(Fault::new(
                5,
                "rt_maybe_drain requires an installed ProcInvoker",
            ));
        }
        return Ok(());
    };
    drain_terminations(&mut state, BridgeState::exec, |_, bundle, proc, object| {
        // SAFETY: the bridge and owned receiver remain live for this call, and
        // the drain holds no typed state borrow while the engine can re-enter.
        let _ = unsafe { invoke(ctx, bundle, proc.0, &object, 1) };
    });
    Ok(())
}
