# Typed call-stack continuations — run 9

2026-11-01. Delivery owner `bd-59co.2.10.12`; HOST-SESSION-001,
JIT-CORE-001, JIT-PARITY-001, JIT-CACHE-001, RUNTIME-VALUE-001.
CORE-JIT-LOWERING and CORE-JIT-CACHE-OBJECT remain **in-progress**.

The existing linked host API and native scalar entry ABI remain version 2.
The owning invocation now retains a bounded heap stack of original OxFunc
activations. Both runtimes consume the same checked layouts. VM3 executes the
typed instructions; JIT emits native arithmetic, storage and indirect callbacks.
Calls and Err/Resume transitions return to shared stack mechanics. There is no
backend fallback or source rewrite. All reachable callees are admitted before
the first body effect. Unused unsupported procedures do not become typed callees.

Required Long/Double/portable-object Function, Property Get and internal Sub calls
can suspend between any completed physical operations, including within child
calls and handlers. A parent's saved call site remains available for propagated
errors. Return stores the exact typed result once and advances past the call.
ByRef parameters refer to stable boxed ancestor frames or canonical globals;
repeated aliases observe writes immediately, including writes before an error.
ByVal expressions passed to a ByRef declaration remain private temporaries.
Host ByRef inputs remain private copies: this does not add a writable foreign
ByRef ABI. Typed host Sub entry still requires a future explicit void carrier.

The shared canonical ErrEngine retains handler policy, active Resume seeds,
line/Erl and error fields across slices. Child entry saves the caller policy;
return/unwind/cancel restores it. An unhandled child fault routes through the
parent's actual call-site fault pad. An active handler propagates its own error.
Explicit Resume retries its VBA statement; Resume Next skips that statement or
the call that propagated the error. Slice suspension itself never repeats effects.
The ordinary JIT Raise route now accepts runtime operands and explicit text/help
fields through retained operand descriptors; it no longer rejects such procedures
merely because the old bare-constant Raise route was narrower.

Nineteen focused invocation tests pass on VM3 and native JIT. New cases exercise
one-operation through unlimited slices, nested Subs, recursive Functions, repeated
numeric/object/global aliases, immediate error-edge mutation, handler callbacks,
explicit retry/skip, the 64-frame error-28 limit, pre-effect refusal and reload
while a descendant is suspended. Existing ownership/cancel/reentry checks remain.
The retained owned Excel receipt `calcvba9-continuation-initial` passes explicit
VBE compile command 578 with its PID-scoped UIA guardian armed first. Exact result:
`alias=10:0,1,2,|catch=27:0,7,8,2,|retry=43:7,8,7,|skip=14:7,2,|object=42:$B$1`.
Owned Excel PID 996 was cleaned up; PID 3144 was not stopped. This is a development
observation, not Windows certification. The source and supervisor wrapper are
`windows-x64/WIN-4/typed-continuation.bas` and
`scripts/run-typed-continuation-oracle.ps1`.

The target follows public [On Error](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/on-error-statement),
[Resume](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/resume-statement)
and [ByRef/ByVal](https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/passing-arguments-efficiently)
contracts, with the live result above checking the bounded scenarios.
Strings/arrays/records, Optional/ParamArray, cross-project calls and suspension
inside native/COM callbacks retain their broader owners. A host callback finishes
synchronously before the next suspension point; no asynchronous COM preemption is
claimed. DNA Calc owns its object library, read-view policy and glue on its machine.

Lean TypedContinuation now also models alias writes, distinct-cell preservation,
retained parent activations and write/effect preservation during unwind. Installed
Lean 4.29.1 accepts it. It does not prove concrete CFG routing, Rust pointer aliases,
COM signatures or generated code. Miri and final repository results are recorded
in the run-9 validation manifest; interrupted/failed attempts earn no passing credit.

Fresh review and a focused Miri rejection exposed a saved pointer into inline Err
storage invalidated by reentrant session access. Canonical Err is now heap-stable.
Both minimal alias and callback/handler cases pass strict-provenance Miri, including
a nested pure-entry fault followed by a successful stateful call. These are bounded
storage checks; the larger canonical Variant limitation recorded in run 8 remains.
The scoped call-stack delivery bead is closed; broader capability rows stay open.
