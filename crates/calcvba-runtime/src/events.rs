//! Delivery of COM events to the running VBA program (step 8 of the VARIANT-layout programme).
//!
//! A COM source fires an event by calling its sink's `Invoke`, which in VBA happens only while
//! VBA is inside a call to the host (the call that triggers the event, or its message loop). VM3
//! and the JIT register themselves as the thread's [`EventExecutor`] around each such call
//! ([`ExecutorScope`]); a sink hands the event to [`deliver`], which runs the handler on the
//! innermost executor, so an event fired inside a handler nests the same way. Outside any host
//! call there is no executor, and the event is not delivered.
use crate::{Utf16String, Variant};
use std::cell::RefCell;
use std::ptr::NonNull;

/// A handler failure, as VBA's `Err` sees it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EventFault {
    pub number: i32,
    pub description: Utf16String,
    pub source: Utf16String,
}

/// Runs WithEvents handlers for the subscriptions a VM made.
pub trait EventExecutor {
    /// Run the handler routed for `dispid` on subscription `key`, with `args` in source order: a
    /// `VT_BYREF` cell is the server's storage for a ByRef argument, any other cell a value the
    /// executor may copy. `Ok(false)` when the subscription has no handler for the event.
    fn deliver_event(
        &mut self,
        key: u64,
        dispid: i32,
        args: &[Variant],
    ) -> Result<bool, EventFault>;
}

thread_local! {
    static EXECUTORS: RefCell<Vec<NonNull<dyn EventExecutor>>> = const { RefCell::new(Vec::new()) };
}

/// The registration of an executor for the duration of one host call; dropping it ends it.
#[must_use = "the executor is registered only while the scope lives"]
pub struct ExecutorScope {
    depth: usize,
}

impl ExecutorScope {
    /// Registers `executor` as the thread's innermost event executor.
    ///
    /// # Safety
    /// `executor` stays valid, and is used by nothing else, until the scope is dropped: the
    /// caller makes the host call with no Rust reference into the executor held across it (the
    /// pointer is taken from `&mut` right before the call, as for `AddressOf` callbacks).
    pub unsafe fn enter(executor: NonNull<dyn EventExecutor + '_>) -> Self {
        // SAFETY: the lifetime is erased for the thread-local stack; the caller keeps the
        // executor valid until this scope (which pops it) is dropped.
        let executor: NonNull<dyn EventExecutor> = unsafe { core::mem::transmute(executor) };
        EXECUTORS.with(|stack| {
            let mut stack = stack.borrow_mut();
            stack.push(executor);
            Self { depth: stack.len() }
        })
    }
}

impl Drop for ExecutorScope {
    fn drop(&mut self) {
        EXECUTORS.with(|stack| {
            let mut stack = stack.borrow_mut();
            stack.truncate(self.depth.saturating_sub(1));
        });
    }
}

/// Hands an event to the innermost executor. `Ok(false)` when there is none or it has no handler
/// for the event.
pub fn deliver(key: u64, dispid: i32, args: &[Variant]) -> Result<bool, EventFault> {
    let Some(executor) = EXECUTORS.with(|stack| stack.borrow().last().copied()) else {
        return Ok(false);
    };
    // SAFETY: a registered executor is valid and unused by its owner while its host call runs
    // (`ExecutorScope::enter`); a nested event re-enters it through the same pointer, as a nested
    // call of the running VM. No borrow of the stack is held across the call.
    unsafe { (*executor.as_ptr()).deliver_event(key, dispid, args) }
}

/// Whether an executor is registered on this thread.
pub fn has_executor() -> bool {
    EXECUTORS.with(|stack| !stack.borrow().is_empty())
}
