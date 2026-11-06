//! calcvba-engine: engine orchestration for the clean execution stack.

pub mod engine;
pub mod runner;

pub use engine::{
    DefaultInstanceAccess, DiagnosticPhase, Engine, ExecBackend, FinalErr, HostConfig,
    HostProfileProvider, JIT_NOT_IMPLEMENTED_MESSAGE, PhaseDiagnostic, ProjectRuntimeSession,
    SnapshotOutcome, Vm3Snapshot,
};
pub use runner::{
    ResolvedRunnerBootstrap, RunnerBootstrapFallbacks, RunnerBootstrapOptions, RuntimeProfileId,
    resolve_runner_bootstrap, resolve_runner_bootstrap_with_fallbacks,
};
