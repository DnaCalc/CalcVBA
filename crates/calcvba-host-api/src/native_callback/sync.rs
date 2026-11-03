//! Scoped synchronization of native ByRef cells around a synchronous callback.
use crate::Error;
use std::cell::Cell;

pub type NativeCallbackSync = unsafe fn(*mut ()) -> Result<(), i32>;

struct Frame {
    data: *mut (),
    before: NativeCallbackSync,
    after: NativeCallbackSync,
    error: Cell<i32>,
}
thread_local! {
    static CURRENT: Cell<*const Frame> = const { Cell::new(std::ptr::null()) };
}

pub(super) struct NativeCallScope(*const Frame);
impl Drop for NativeCallScope {
    fn drop(&mut self) {
        CURRENT.with(|slot| slot.set(self.0));
    }
}
pub(super) fn native_call_scope() -> NativeCallScope {
    NativeCallScope(CURRENT.with(|slot| slot.replace(std::ptr::null())))
}
struct Restore(*const Frame);
impl Drop for Restore {
    fn drop(&mut self) {
        CURRENT.with(|slot| slot.set(self.0));
    }
}

/// Synchronize live native aliases immediately before callback entry and after
/// callback return (also during unwinding). Nested native calls use their own
/// scope. Any synchronization fault remains a fault even if native code ignores
/// the callback's failure and reports success.
///
/// # Safety
/// Call inside with_native_call's foreign portion. data and both hooks remain
/// valid on this thread until call returns. Hooks must not panic, reenter, retain
/// references or expose competing Rust borrows of native/caller storage. They
/// must preserve every native pointer's provenance and the caller's exact types.
pub unsafe fn with_native_callback_sync<R>(
    data: *mut (),
    before: NativeCallbackSync,
    after: NativeCallbackSync,
    call: impl FnOnce() -> R,
) -> Result<R, i32> {
    if super::CONTEXT.with(Cell::get).is_null() || super::BORROWED.with(Cell::get) {
        return Err(5);
    }
    let frame = Frame {
        data,
        before,
        after,
        error: Cell::new(0),
    };
    let _restore = Restore(CURRENT.with(|slot| slot.replace(std::ptr::from_ref(&frame))));
    let result = call();
    match frame.error.get() {
        0 => Ok(result),
        number => Err(number),
    }
}

fn error(number: i32) -> Error {
    if number == -1 {
        Error::BudgetExceeded
    } else {
        Error::Runtime {
            number,
            description: oxvba_runtime::default_error_message(number).into(),
            vba_source: "NativeCallback".into(),
        }
    }
}
unsafe fn run(frame: *const Frame, before: bool) -> Result<(), Error> {
    // SAFETY: scoped registration retains this same-thread immutable frame.
    let frame = unsafe { &*frame };
    let hook = if before { frame.before } else { frame.after };
    // SAFETY: the unsafe registration contract lends the exact live hook data;
    // the callback borrow excludes overlapping synchronization/engine execution.
    if let Err(number) = unsafe { hook(frame.data) } {
        let number = if number == -1 || number > 0 {
            number
        } else {
            5
        };
        if frame.error.get() == 0 {
            frame.error.set(number);
        }
        return Err(error(number));
    }
    Ok(())
}

pub(super) struct CallbackSynchronization {
    frame: *const Frame,
}
impl CallbackSynchronization {
    fn complete(&mut self) -> Result<(), Error> {
        let frame = std::mem::replace(&mut self.frame, std::ptr::null());
        if frame.is_null() {
            return Ok(());
        }
        CURRENT.with(|slot| slot.set(frame));
        // SAFETY: callback completion still lies inside the retained native
        // registration; its exact caller/native storage leases remain live.
        unsafe { run(frame, false) }
    }
    pub(super) fn finish(mut self) -> Result<(), Error> {
        self.complete()
    }
}
impl Drop for CallbackSynchronization {
    fn drop(&mut self) {
        let _ = self.complete();
    }
}
pub(super) fn begin_callback() -> Result<CallbackSynchronization, Error> {
    let frame = CURRENT.with(Cell::get);
    if !frame.is_null() {
        // SAFETY: CURRENT is installed only by the retained same-thread scope;
        // the caller already holds the exclusive callback/context borrow.
        unsafe { run(frame, true) }?;
        // Do not expose an outer call's cells to nested native callbacks. Each
        // nested with_native_call starts with an empty synchronization scope.
        CURRENT.with(|slot| slot.set(std::ptr::null()));
    }
    Ok(CallbackSynchronization { frame })
}

#[cfg(test)]
mod tests;
