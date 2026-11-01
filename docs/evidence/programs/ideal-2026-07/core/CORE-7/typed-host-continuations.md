# Owning typed state, inspection and continuations — run 8

2026-11-01. Directed work: `bd-59co.2.10.10`, `bd-59co.2.9.13`,
`bd-59co.2.10.11`. Broader CORE-7/CORE-8 and canonical rows remain **in-progress**.
Authority: HOST-SESSION-001, JIT-CORE-001, JIT-PARITY-001, JIT-CACHE-001,
RUNTIME-VALUE-001; the current [hosting contract](../../../../../spec/HOSTING_AND_EARLY_BINDING.md).

The API is linked Rust API ABI 2, retaining native primary ABI 2. Owning typed
frames are an additional execution layout over the original checked OxFunc.
Both VM3 and generated native JIT entries consume that layout, with no JIT to
interpreter fallback. Required ByVal Long/Double/portable object Functions and
Property Gets are admitted, with numeric operations/control flow and existing
numeric host slots. ByRef, Subs, strings/arrays/records, handlers, internal call
stacks and suspension inside a native callback remain open under the broader
delivery owners. Canonical session globals are shared with dynamic calls; there
is no separate copy of module/Static state. No spreadsheet adapter members or
DNA Calc source changed.

The invocation owns arguments/locals, exact next-operation position, original
code/state lease, total budget and host-supplied id/parent/read-view metadata.
Finite slices preserve completed effects; zero performs no work. Terminal faults,
hard budget exhaustion and cancellation release owners and cannot be resumed.
Reset and unrelated calls refuse while the original state is suspended. Reload
publishes a replacement session while the old activation retains its state/code
and can finish with its old signature. Old handles refuse on the new generation.
Host calls complete synchronously before the next suspension point. This is
distinct from an actual new semantic invocation, which starts again.

Object assignments retain before releasing, including self-assignment. Local
slots retain owners, and global receiver reads retain a transient owner through
re-entrant replacement. Object results return an owning HostObject or Nothing;
the scalar slot API refuses raw object results. Portable-table recovery checks
the private IUnknown vtable identity before reading its header. Native COM uses
its existing path and never becomes a portable host table.
The language target follows Microsoft's [Set statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/set-statement)
and [Static statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/static-statement),
with the live Range observation below checking the bounded ownership sequence.

Object-typed local/global TypeName and IsObject now reach the existing descriptor route.
JIT TypeName also handles a project object in generic Object or global Variant
storage, while ordinary nonobjects retain Empty/Null/array behavior. The generic
library still owns native naming and other information functions. This change
does not claim general default-member coercion parity for every intrinsic.
The identity target follows [TypeName](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/typename-function)
and [IsObject](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/isobject-function);
Nothing remains an object variable even though it refers to no object.

Current focused evidence: all eleven `suspended_invocation` tests pass through
both runtimes. They cover module/Static/results and exact identity, dynamic clear
of typed state, reset/result lifetime, one-operation slices with one prefix host
effect, pinned id/parent/view, reload with a different String signature, stale
handles, inactive/wrong/deleted host views, callback replacement of the receiver,
wrong-IID assignment preservation, cancellation, faults/panics and total budgets.
Declared interface metadata is retained by the compiler before any member call,
including memberless interfaces and distinct qualified names with the same short
name. Function and Property Get object results retain their exact owner.
Unused interface-typed arguments reject a wrong nominal IID before body entry
on both the original native primary entry and the owning frame. Untyped Object
accepts the table. Slot::Nothing binds successfully and faults only on dereference.
The 47 project-object differential tests pass. Repository meta-check
`-Fast -NoArtifacts -NoRegistry` passes 3,717 tests in 194 suites; 180 are
ignored and four registry-writing fixtures are excluded without passing credit.
Formatting, strict workspace Clippy, governance and path checks pass. The
interpreter-only embed run passes 37 tests with one live case ignored; the
host API/embed wasm32 build passes.
The finite-input regression preserves universal exact carriers while typed host
arguments continue to reject nonfinite inputs.

Live Excel `calcvba8-final-parameter-source` passed explicit VBE compile command 578 with
the PID-scoped UIA guardian ready first. Real VBA module/Static/object-result
probe: `Range|Range|True|True|43|Range|True|13|7|7`. The final fields
show error 13 for an unused Range parameter given a Worksheet through Object,
and successful unused Nothing/Range arguments. This corrects the old portable
member-time error 438 to entry-time type mismatch 13 for incompatible arguments. The Object-typed local/global native
Range test passes both VM3 and JIT in that owned process (PID 32056, cleaned up). This is development
evidence, not Windows certification. The receipt retains the executable hash,
compile authority, guardian events and owned-process cleanup. Fresh review fixed
nominal metadata retention, exact-carrier regression, global pointer authority,
receiver ownership and unused-argument binding; no material issue remains in
the bounded delivery after the final checks.

Formal scope: all eight installed Lean 4.29.1 abstraction models pass.
Final default-mode Miri passes all eleven invocation tests, retaining the
exposed-provenance warnings. Separate strict-provenance object-only and
argument/Nothing cases each pass. TypedContinuation models deterministic slice composition, completed
effect counts and retained code/state/view leases; it does not prove concrete CFG
lowering, Rust aliasing, native signatures or generated memory accesses. Miri
does not execute JIT machine code. The canonical Variant's exposed-pointer byte
carrier limits strict-provenance claims for owning globals. Those refinements
remain in `docs/evidence/formal/EXTENDED_TODO.md`.

Scoped delivery leaves `bd-59co.2.10.10`, `bd-59co.2.9.13` and
`bd-59co.2.10.11` are closed on these results. General calls, ByRef and handler
continuations remain open under `bd-59co.2.10.12`; broader carriers, source-free
packages and host integration remain with `bd-59co.2.9.9`/`bd-59co.2.10.7`.
Canonical CORE-JIT-LOWERING and CORE-JIT-CACHE-OBJECT stay in-progress.

The [manifest](run8-manifest.json) binds retained raw receipts, commands,
failures without passing credit, native executable hashes and 400 exact source
files in [the archive](run8-validation.zip). Run
`python docs/evidence/programs/ideal-2026-07/core/CORE-7/verify-run8.py` from the repository
root to verify the bytes and current-source association. The archive also retains
the fresh review and the DNA Calc research context; it contains no prediction files.

The ignored build lock is retained separately as `build/Cargo.lock` in the
archive. Restore it to the repository root before reproducing Cargo commands;
source verification itself requires only the tracked files.
