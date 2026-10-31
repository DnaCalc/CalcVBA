namespace OxVba.TypedReentry

-- Abstract admission/budget contract only. This is not a proof of Rust aliases,
-- callback ABI, code leases or the generated call_indirect memory accesses.
structure Activation where
  depth : Nat
  context : Nat
  dynamic : Bool
  deriving DecidableEq

def canEnter (a : Activation) (context : Nat) : Bool :=
  !a.dynamic && (a.depth < 64) && (a.depth == 0 || a.context == context)

def enter (a : Activation) (context : Nat) : Option Activation :=
  if canEnter a context then some ⟨a.depth + 1, context, false⟩ else none

theorem dynamic_refuses (a : Activation) (c : Nat) (h : a.dynamic = true) :
    enter a c = none := by simp [enter, canEnter, h]

theorem depth_limit_refuses (a : Activation) (c : Nat) (h : 64 ≤ a.depth) :
    enter a c = none := by simp [enter, canEnter, Nat.not_lt.mpr h]

theorem wrong_context_refuses (a : Activation) (c : Nat)
    (active : a.depth ≠ 0) (wrong : a.context ≠ c) : enter a c = none := by
  simp [enter, canEnter, active, wrong]

theorem accepted_entry_increments_depth (a : Activation) (c : Nat)
    (h : canEnter a c = true) : enter a c = some ⟨a.depth + 1, c, false⟩ := by
  simp [enter, h]

theorem accepted_depth_stays_bounded (a : Activation) (c : Nat)
    (h : canEnter a c = true) : a.depth + 1 ≤ 64 := by
  have : a.depth < 64 := by
    simp [canEnter] at h
    exact h.1.2
  omega

-- Reentry borrows the same budget; there is no newly minted nested allowance.
theorem nested_work_consumes_outer_budget (remaining outer inner : Nat) :
    remaining - outer - inner = remaining - (outer + inner) := by omega

theorem nested_work_never_increases_budget (remaining outer inner : Nat) :
    remaining - outer - inner ≤ remaining := by omega

end OxVba.TypedReentry
