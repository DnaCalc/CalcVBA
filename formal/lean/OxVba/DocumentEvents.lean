namespace OxVba.DocumentEvents

-- Abstract subscription lifecycle only. The concrete native/TLS/Rust/JIT
-- refinement remains unproved. In particular, reset must retain nextEpoch.
structure State where
  nextEpoch : Nat
  attached : Option Nat
  deriving DecidableEq

def attach (s : State) : State :=
  match s.attached with
  | some _ => s
  | none => ⟨s.nextEpoch + 1, some (s.nextEpoch + 1)⟩

def detach (s : State) : State := { s with attached := none }
def accepts (s : State) (callback : Nat) : Bool := s.attached == some callback

theorem detach_rejects_all (s : State) (callback : Nat) :
    accepts (detach s) callback = false := by simp [accepts, detach]

theorem reattach_is_idempotent (s : State) : attach (attach s) = attach s := by
  cases h : s.attached <;> simp [attach, h]

theorem reset_rejects_previous_epoch (s : State) :
    accepts (attach (detach s)) s.nextEpoch = false := by
  simp [accepts, attach, detach]

theorem reset_keeps_epoch_monotonic (s : State) :
    s.nextEpoch < (attach (detach s)).nextEpoch := by
  simp [attach, detach]

end OxVba.DocumentEvents
