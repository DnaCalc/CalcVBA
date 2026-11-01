# formal/lean/

Lean 4 project scaffold for OxVBA formal specifications.

Planned modules (initial):
- `VarType.lean`
- `Coerce.lean`
- `Arithmetic.lean`
- `RefCount.lean`
- `ProcedureHandles.lean`: owner/generation rejection, independent of Rust code leases.
- `TypedReentry.lean`: typed context identity, depth limit and shared-budget admission;
  excludes Rust aliasing, native callback signatures and generated memory accesses.
- `TypedContinuation.lean`: deterministic slice composition, completed effect
  counts and retained code/state/view leases; excludes concrete CFG, Rust storage
  provenance and generated-code refinement.
- `ExecutionBudget.lean`: abstract finite/unlimited step model; zero is sticky,
  successful finite steps consume exactly one, and the budget never increases.
  This does not prove the Rust implementation or backend abort propagation.
