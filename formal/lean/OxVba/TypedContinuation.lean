namespace OxVba.TypedContinuation

-- An abstract deterministic instruction transition. The state includes the
-- saved next instruction, locals and observable effect journal. This models
-- continuation composition, not Rust pointers or emitted native instructions.
def slice {State : Type} (step : State → State) : Nat → State → State
  | 0, state => state
  | n + 1, state => slice step n (step state)

theorem zero_slice {State : Type} (step : State → State) (s : State) :
    slice step 0 s = s := rfl

-- Splitting execution preserves every state component, including the journal:
-- the second slice receives the first slice's result, never the entry state.
theorem slices_compose {State : Type} (step : State → State) (a b : Nat) (s : State) :
    slice step (a + b) s = slice step b (slice step a s) := by
  induction a generalizing s with
  | zero => simp [slice]
  | succ a ih => simpa [Nat.succ_add, slice] using ih (step s)

structure Lease where
  code : Nat
  projectState : Nat
  view : Nat
  deriving DecidableEq

structure Frame where
  lease : Lease
  completed : Nat
  journal : List Nat
  deriving DecidableEq

def instruction (f : Frame) : Frame :=
  { f with completed := f.completed + 1, journal := f.journal ++ [f.completed] }

theorem step_keeps_lease (f : Frame) : (instruction f).lease = f.lease := rfl

theorem slice_keeps_lease (n : Nat) (f : Frame) :
    (slice instruction n f).lease = f.lease := by
  induction n generalizing f with
  | zero => rfl
  | succ n ih => simpa [slice, instruction] using ih (instruction f)

theorem slice_counts_completed_operations (n : Nat) (f : Frame) :
    (slice instruction n f).completed = f.completed + n := by
  induction n generalizing f with
  | zero => simp [slice]
  | succ n ih => simp [slice, ih, instruction, Nat.add_comm, Nat.add_left_comm]

theorem slice_appends_one_effect_per_operation (n : Nat) (f : Frame) :
    (slice instruction n f).journal.length = f.journal.length + n := by
  induction n generalizing f with
  | zero => simp [slice]
  | succ n ih => simp [slice, ih, instruction, Nat.add_comm, Nat.add_left_comm]

theorem slice_preserves_journal_prefix (n : Nat) (f : Frame) :
    ∃ suffix, (slice instruction n f).journal = f.journal ++ suffix := by
  induction n generalizing f with
  | zero => exact ⟨[], by simp [slice]⟩
  | succ n ih =>
    obtain ⟨suffix, h⟩ := ih (instruction f)
    exact ⟨[f.completed] ++ suffix, by simpa [slice, instruction, List.append_assoc] using h⟩

-- Publishing a replacement changes the current lease only. The old frame and
-- its saved state are still the exact arguments of its next slice.
def reload (replacement : Lease) (old : Frame) : Lease × Frame := (replacement, old)
theorem reload_preserves_suspended_activation (next : Lease) (old : Frame) :
    (reload next old).2 = old := rfl

-- A retained activation stack uses canonical locations for ByRef parameters.
-- Two aliases of the same cell observe each write immediately, including writes
-- made before a child faults. Copy-out at a later return is not this model.
abbrev Store := Nat → Nat
def write (s : Store) (location value : Nat) : Store :=
  fun other => if other = location then value else s other

theorem aliases_observe_write (s : Store) (first second value : Nat)
    (same : first = second) : write s first value second = value := by
  simp [write, same]

theorem distinct_cells_survive_write (s : Store) (first second value : Nat)
    (different : second ≠ first) : write s first value second = s second := by
  simp [write, different]

structure Activation where
  next : Nat
  handler : Nat
  activeError : Option Nat
  lease : Lease
  deriving DecidableEq

-- A child returns to the saved next operation. Pushing or discarding the child
-- never reconstructs a parent from the source entry or a replacement program.
theorem child_pop_keeps_parent (parent child : Activation) (ancestors : List Activation) :
    (child :: parent :: ancestors).tail = parent :: ancestors := rfl

theorem suspended_stack_reload_keeps_all_activations (replacement : Lease)
    (stack : List Activation) : (replacement, stack).2 = stack := rfl

-- Fault unwinding changes the stack, while completed ByRef writes and the
-- effect journal survive. The concrete Rust ErrEngine/CFG routing is excluded.
def unwind (stack : List Activation) (store : Store) (journal : List Nat) :=
  (stack.tail, store, journal)
theorem unwinding_keeps_writes_and_effects (stack : List Activation)
    (store : Store) (journal : List Nat) :
    (unwind stack store journal).2 = (store, journal) := rfl

end OxVba.TypedContinuation
