//! Delivery of COM events and incoming COM member calls to the running VBA executor
//! (step 8c & step 10 / dc-yqq0.23).
//!
//! A COM source fires an event or calls back into a project object's `IDispatch::Invoke` while
//! VBA is inside a host call, or between runs on a COM session holder. VM3 and the JIT register
//! themselves as the thread's [`EventExecutor`] around active host calls ([`ExecutorScope`]),
//! and a session holder registers itself as a [`MemberInvoker`] ([`MemberInvokerScope`]) when
//! serving calls between runs. Incoming events go to the innermost [`EventExecutor`]; incoming
//! member invocations go to the innermost active executor or member invoker so callbacks during a
//! host call or `RaiseEvent` always run on the active run.
use crate::{
    ObjectRef, RuntimeMemberDescriptor, RuntimeMemberInvokeKind, Utf16String, Variant,
    default_error_message,
};
use std::cell::RefCell;
use std::ptr::NonNull;

/// A handler failure, as VBA's `Err` sees it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EventFault {
    pub number: i32,
    pub description: Utf16String,
    pub source: Utf16String,
}

/// A VBA member-invocation fault surfaced through `IDispatch::Invoke` (`EXCEPINFO`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MemberFault {
    pub number: i32,
    pub description: Utf16String,
    pub source: Utf16String,
    pub help_file: Utf16String,
    pub help_context: i32,
}

impl MemberFault {
    pub fn from_number(number: i32) -> Self {
        Self {
            number,
            description: default_error_message(number).into(),
            source: Utf16String::new(),
            help_file: Utf16String::new(),
            help_context: 0,
        }
    }
}

/// Runs a project class member selected by descriptor (`RuntimeMemberDescriptor`) on its owning
/// executor (VM3, JIT, session holder, or AOT).
pub trait MemberInvoker {
    /// Invoke `member` on `object` with `invoke_kind` and formal arguments `args` (with ByRef
    /// aliasing given by `aliases`: `aliases[i]` is the lowest index sharing `args[i]`'s storage).
    /// ByRef modifications are written back into `args`, even when the member raises a fault.
    fn invoke_member(
        &mut self,
        object: ObjectRef,
        member: &'static RuntimeMemberDescriptor,
        invoke_kind: RuntimeMemberInvokeKind,
        args: &mut [Variant],
        aliases: &[usize],
    ) -> Result<Variant, MemberFault>;
}

/// Runs WithEvents handlers and active-run incoming member callbacks for a running VM or JIT.
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

    /// Invoke a project class member on the active run (when foreign COM calls back into a
    /// project object's `IDispatch::Invoke` during a host call or `RaiseEvent`).
    fn invoke_member(
        &mut self,
        object: ObjectRef,
        member: &'static RuntimeMemberDescriptor,
        invoke_kind: RuntimeMemberInvokeKind,
        args: &mut [Variant],
        aliases: &[usize],
    ) -> Result<Variant, MemberFault> {
        let _ = (object, member, invoke_kind, args, aliases);
        Err(MemberFault::from_number(438))
    }

    /// Run the project procedure `proc` of linked unit `unit` (a Sub or Function with no hidden
    /// receiver) on the active run with already prepared `args`: the reentrant dynamic call a host
    /// member makes while a procedure of this run is inside it (`Application.Run`). The result is
    /// the Function's value, or Empty for a Sub.
    fn call_project_procedure(
        &mut self,
        unit: usize,
        proc: usize,
        args: Vec<Variant>,
    ) -> Result<Variant, MemberFault> {
        let _ = (unit, proc, args);
        Err(MemberFault::from_number(438))
    }
}

#[derive(Clone, Copy)]
enum ActiveExecutor {
    /// An event executor and the object owner (session) whose run it is.
    Event(NonNull<dyn EventExecutor>, u64),
    Member(NonNull<dyn MemberInvoker>),
}

thread_local! {
    static EXECUTORS: RefCell<Vec<ActiveExecutor>> = const { RefCell::new(Vec::new()) };
}

/// The registration of an executor for the duration of one host call; dropping it ends it.
#[must_use = "the executor is registered only while the scope lives"]
pub struct ExecutorScope {
    depth: usize,
}

impl ExecutorScope {
    /// Registers `executor` as the thread's innermost event and member executor.
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
            stack.push(ActiveExecutor::Event(
                executor,
                crate::object_owner::RuntimeObjectOwner::current().id(),
            ));
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

/// Registers a between-runs [`MemberInvoker`] (such as a COM session holder) for incoming
/// `IDispatch::Invoke` calls on this thread; dropping it ends the registration.
#[must_use = "the member invoker is registered only while the scope lives"]
pub struct MemberInvokerScope {
    depth: usize,
}

impl MemberInvokerScope {
    /// Registers `invoker` as the thread's innermost member invoker.
    ///
    /// # Safety
    /// `invoker` stays valid and unused by anything else until the scope is dropped.
    pub unsafe fn enter(invoker: NonNull<dyn MemberInvoker + '_>) -> Self {
        // SAFETY: the lifetime is erased for the thread-local stack; the caller keeps `invoker`
        // valid until this scope (which pops it) is dropped.
        let invoker: NonNull<dyn MemberInvoker> = unsafe { core::mem::transmute(invoker) };
        EXECUTORS.with(|stack| {
            let mut stack = stack.borrow_mut();
            stack.push(ActiveExecutor::Member(invoker));
            Self { depth: stack.len() }
        })
    }
}

impl Drop for MemberInvokerScope {
    fn drop(&mut self) {
        EXECUTORS.with(|stack| {
            let mut stack = stack.borrow_mut();
            stack.truncate(self.depth.saturating_sub(1));
        });
    }
}

/// Hands an event to the innermost event executor. `Ok(false)` when there is none or it has no
/// handler for the event.
pub fn deliver(key: u64, dispid: i32, args: &[Variant]) -> Result<bool, EventFault> {
    let executor = EXECUTORS.with(|stack| {
        stack.borrow().iter().rev().find_map(|entry| match *entry {
            ActiveExecutor::Event(executor, _) => Some(executor),
            ActiveExecutor::Member(_) => None,
        })
    });
    let Some(executor) = executor else {
        return Ok(false);
    };
    // SAFETY: a registered executor is valid and unused by its owner while its host call runs
    // (`ExecutorScope::enter`); a nested event re-enters it through the same pointer, as a nested
    // call of the running VM. No borrow of the stack is held across the call.
    unsafe { (*executor.as_ptr()).deliver_event(key, dispid, args) }
}

/// Invokes a project class member on the thread's innermost active executor or member invoker.
/// Returns `None` when no executor/invoker is registered on this thread.
pub fn invoke_member(
    object: ObjectRef,
    member: &'static RuntimeMemberDescriptor,
    invoke_kind: RuntimeMemberInvokeKind,
    args: &mut [Variant],
    aliases: &[usize],
) -> Option<Result<Variant, MemberFault>> {
    let active = EXECUTORS.with(|stack| stack.borrow().last().copied())?;
    Some(match active {
        ActiveExecutor::Event(executor, _) => {
            // SAFETY: registered via `ExecutorScope::enter`; valid and unborrowed across the call.
            unsafe {
                (*executor.as_ptr()).invoke_member(object, member, invoke_kind, args, aliases)
            }
        }
        ActiveExecutor::Member(invoker) => {
            // SAFETY: registered via `MemberInvokerScope::enter`; valid and unborrowed across call.
            unsafe { (*invoker.as_ptr()).invoke_member(object, member, invoke_kind, args, aliases) }
        }
    })
}

/// Runs a project procedure on the thread's innermost active event executor (the running VM or
/// JIT run that is inside a host call). `None` when no run is active on this thread.
pub fn call_project_procedure(
    unit: usize,
    proc: usize,
    args: Vec<Variant>,
) -> Option<Result<Variant, MemberFault>> {
    let executor = EXECUTORS.with(|stack| {
        stack.borrow().iter().rev().find_map(|entry| match *entry {
            ActiveExecutor::Event(executor, _) => Some(executor),
            ActiveExecutor::Member(_) => None,
        })
    })?;
    // SAFETY: as for `deliver`: the registered executor is valid and unused by its owner while its
    // host call runs, and no borrow of the stack is held across the call.
    Some(unsafe { (*executor.as_ptr()).call_project_procedure(unit, proc, args) })
}

/// [`call_project_procedure`] on the innermost run of the session whose object owner is `owner`:
/// the reentrant call of a procedure of that session, while a procedure of it is inside a host call
/// further up the stack (another session's run may be above it). `None` when no run of that
/// session is active on this thread.
pub fn call_project_procedure_for(
    owner: u64,
    unit: usize,
    proc: usize,
    args: Vec<Variant>,
) -> Option<Result<Variant, MemberFault>> {
    let executor = EXECUTORS.with(|stack| {
        stack.borrow().iter().rev().find_map(|entry| match *entry {
            ActiveExecutor::Event(executor, id) if id == owner => Some(executor),
            _ => None,
        })
    })?;
    // SAFETY: as for `call_project_procedure`.
    Some(unsafe { (*executor.as_ptr()).call_project_procedure(unit, proc, args) })
}

/// Whether an event executor is registered on this thread.
pub fn has_executor() -> bool {
    EXECUTORS.with(|stack| {
        stack
            .borrow()
            .iter()
            .any(|entry| matches!(entry, ActiveExecutor::Event(..)))
    })
}

/// Whether any member invoker (active run or session scope) is registered on this thread.
pub fn has_member_invoker() -> bool {
    EXECUTORS.with(|stack| !stack.borrow().is_empty())
}

#[cfg(test)]
mod tests {
    use super::*;

    struct Doubler;
    impl EventExecutor for Doubler {
        fn deliver_event(&mut self, _: u64, _: i32, _: &[Variant]) -> Result<bool, EventFault> {
            Ok(false)
        }
        fn call_project_procedure(
            &mut self,
            unit: usize,
            proc: usize,
            args: Vec<Variant>,
        ) -> Result<Variant, MemberFault> {
            if proc == 9 {
                return Err(MemberFault::from_number(11));
            }
            let n = args[0].as_i32().unwrap();
            Ok(Variant::from_i32(n * 2 + unit as i32 + proc as i32))
        }
    }

    #[test]
    fn a_project_procedure_call_reaches_the_innermost_event_executor() {
        assert!(call_project_procedure(0, 1, vec![Variant::from_i32(4)]).is_none());
        let mut executor = Doubler;
        // SAFETY: the executor outlives the scope and is not used while the scope is live.
        let scope =
            unsafe { ExecutorScope::enter(NonNull::from(&mut executor as &mut dyn EventExecutor)) };
        let value = call_project_procedure(2, 3, vec![Variant::from_i32(4)])
            .unwrap()
            .unwrap();
        assert_eq!(value.as_i32(), Some(13));
        let fault = call_project_procedure(2, 9, vec![Variant::from_i32(4)])
            .unwrap()
            .unwrap_err();
        assert_eq!(fault.number, 11);
        drop(scope);
        assert!(call_project_procedure(0, 1, vec![Variant::from_i32(4)]).is_none());
    }
}
