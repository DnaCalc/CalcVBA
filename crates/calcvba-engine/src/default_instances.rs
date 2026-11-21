//! Non-creating default-instance access for same-thread host lifecycle callbacks.
use super::{PhaseDiagnostic, ProjectRuntimeSession, SessionBackend};
use calcvba_rt_abi::DefaultInstanceContext;
use calcvba_runtime::ObjectRef;
use std::rc::Rc;

/// A checked view of a session's default-instance slots. It can be retained by a
/// synchronous host callback without borrowing or re-entering the running VM.
/// Looking up a missing default never allocates it or runs Initialize.
#[derive(Clone)]
pub struct DefaultInstanceAccess(Rc<Access>);
struct Access {
    classes: DefaultInstanceContext,
}
impl DefaultInstanceAccess {
    fn live(&self) -> Result<(), PhaseDiagnostic> {
        if !self.0.classes.is_alive() {
            return Err(PhaseDiagnostic::from_diagnostic(
                calcvba_diagnostics::Diagnostic::error(
                    "HOST-E-STALE-DEFAULTS",
                    calcvba_diagnostics::DiagnosticPhase::Host,
                    "default-instance access belongs to a disposed session",
                ),
            ));
        }
        Ok(())
    }
    /// Return one held reference to the existing default, if any. A qualified
    /// name selects its project; an unqualified name selects the entry project.
    pub fn peek(&self, class_name: &str) -> Result<Option<ObjectRef>, PhaseDiagnostic> {
        self.live()?;
        let (unit, class) = self
            .0
            .classes
            .find_predeclared_class(class_name)
            .ok_or_else(|| {
                super::vm3_runtime_diagnostic(calcvba_vm3::Vm3Error::Fault(
                    calcvba_rt_abi::Fault::new(429, "unknown predeclared class"),
                ))
            })?;
        // SAFETY: the weak session lease is live, access is confined by Rc to its
        // owning thread, and no map borrow crosses a callback or object release.
        Ok(unsafe { self.0.classes.predeclared(unit, class) }.and_then(|v| v.as_object_ref()))
    }
    /// Clear only the default slot that still denotes this exact object. Held
    /// references remain valid; Terminate waits for their final release.
    pub fn reset_if_same(
        &self,
        class_name: &str,
        expected: &ObjectRef,
    ) -> Result<bool, PhaseDiagnostic> {
        self.live()?;
        let (unit, class) = self
            .0
            .classes
            .find_predeclared_class(class_name)
            .ok_or_else(|| {
                super::vm3_runtime_diagnostic(calcvba_vm3::Vm3Error::Fault(
                    calcvba_rt_abi::Fault::new(429, "unknown predeclared class"),
                ))
            })?;
        // SAFETY: same checked, synchronous canonical-map lease as peek. The
        // expected reference pins the object until after removal has completed.
        Ok(unsafe {
            self.0
                .classes
                .clear_predeclared_if_same(unit, class, expected)
        })
    }
}
impl ProjectRuntimeSession {
    /// Obtain a non-creating view suitable for a host's unload callback. The view
    /// rejects access after this session is disposed and cannot move threads.
    pub fn default_instances(&mut self) -> DefaultInstanceAccess {
        let classes = match &mut self.backend {
            SessionBackend::Vm3(vm) => vm.default_instance_context(),
            SessionBackend::Jit(jit) => jit.session.default_instance_context(),
            SessionBackend::JitBusy => {
                panic!("default instances of a JIT session lent to an event")
            }
        };
        DefaultInstanceAccess(Rc::new(Access { classes }))
    }
    /// Release an owned host reference and run eligible finalizers now, on the
    /// owning VBA thread. Other VBA, host or control references keep it alive.
    pub fn release_class_instance(&mut self, object: ObjectRef) {
        let _owner = self.object_owner.enter();
        match &mut self.backend {
            SessionBackend::Vm3(vm) => vm.release_project_instance(object),
            SessionBackend::Jit(_) => {
                drop(object);
                self.drain_object_terminations();
            }
            SessionBackend::JitBusy => drop(object),
        }
    }
    /// Drain releases that occurred in host storage between VBA calls.
    pub fn drain_object_terminations(&mut self) {
        let _owner = self.object_owner.enter();
        match &mut self.backend {
            SessionBackend::Vm3(vm) => vm.drain_object_terminations(),
            // A JIT run drains queued Class_Terminate work as it ends; an empty one drains now.
            SessionBackend::Jit(jit) => {
                let _ = jit.image.drain_in_session(&mut jit.session);
            }
            SessionBackend::JitBusy => {}
        }
    }
}
