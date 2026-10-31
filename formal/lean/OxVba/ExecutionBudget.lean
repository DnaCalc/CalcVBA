namespace OxVba.ExecutionBudget

-- Abstract model of ExecState::consume_execution_step. `none` is unlimited.
-- Rust uses u64; the model's finite input is embedded in Nat without truncation.
def step : Option Nat → Bool × Option Nat
  | none => (true, none)
  | some 0 => (false, some 0)
  | some (n + 1) => (true, some n)

theorem unlimited_is_unchanged : step none = (true, none) := rfl

theorem zero_is_sticky : step (some 0) = (false, some 0) := rfl

theorem success_consumes_exactly_one (n : Nat) :
    step (some (n + 1)) = (true, some n) := rfl

theorem finite_step_succeeds_iff_positive (n : Nat) :
    (step (some n)).1 = true ↔ 0 < n := by
  cases n <;> simp [step]

theorem finite_step_never_increases (n : Nat) :
    (step (some n)).2 = some (n - 1) := by
  cases n <;> simp [step]

end OxVba.ExecutionBudget
