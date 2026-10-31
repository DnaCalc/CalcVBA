namespace OxVba.ProcedureHandles

-- Abstract contract only: this does not prove Rust ownership, native signatures
-- or Cranelift memory safety. Rust uses checked u64 generations and Rc identity.
structure Handle where
  owner : Nat
  generation : Nat

structure Session where
  owner : Nat
  generation : Nat
  live : Bool

def canCall (h : Handle) (s : Session) : Bool :=
  s.live && (h.owner == s.owner) && (h.generation == s.generation)

theorem stale_generation_cannot_call (h : Handle) (s : Session)
    (stale : h.generation ≠ s.generation) : canCall h s = false := by
  simp [canCall, stale]

theorem different_owner_cannot_call (h : Handle) (s : Session)
    (other : h.owner ≠ s.owner) : canCall h s = false := by
  simp [canCall, other]

theorem unloaded_cannot_call (h : Handle) (s : Session)
    (unloaded : s.live = false) : canCall h s = false := by
  simp [canCall, unloaded]

theorem reload_invalidates_old_handle (s : Session) :
    canCall ⟨s.owner, s.generation⟩ { s with generation := s.generation + 1 } = false := by
  simp [canCall]

end OxVba.ProcedureHandles
