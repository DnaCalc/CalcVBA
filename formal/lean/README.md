# formal/lean/

Lean 4 project scaffold for OxVBA formal specifications.

Planned modules (initial):
- `VarType.lean`
- `Coerce.lean`
- `Arithmetic.lean`
- `RefCount.lean`
- `ExecutionBudget.lean`: abstract finite/unlimited step model; zero is sticky,
  successful finite steps consume exactly one, and the budget never increases.
  This does not prove the Rust implementation or backend abort propagation.
