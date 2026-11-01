# Runtime default-member coercion — run 10

2026-11-01. Scoped delivery bd-59co.2.9.15; CORE-JIT-LOWERING remains in-progress.

The owning VM3/native JIT value operations retain a map from interface IID to
compiler-selected default getter. Provider-supplied host interfaces are retained
even when source mentions only Object/Variant. No spreadsheet names or adapter
members are involved. Default callbacks receive the active TypedContext and
invocation view, and produce owned results. The engine resolves object chains
without altering the original operand. Box/Variant copies and Set retain identity.
Supported operations: add/subtract/multiply/divide, unary minus, numeric/string
coercion, concatenation, value comparison and conditions within admitted bodies.

`cargo test -p oxvba-embed --test typed_defaults --test typed_value_slots` exercises
both runtimes at one-operation slices, callback order, nested defaults, ByRef use,
Nothing/Null/bad text/missing members, callback faults and retained original identity.
The neighboring value-slot ownership/error tests also pass.

`scripts/run-typed-defaults-oracle.ps1 -RunId calcvba10-defaults-initial` compiled
with VBE command 578, under the armed PID-scoped UIA supervisor, and observed:
`add=42|text=v=41|number=41|equal=True|condition=1|identity=True|nothing=91|null=94|bad=13|required-default=450`.
Excel PID 4360 was owned and cleaned up; PID 3144 was excluded. No security,
Trust Center, macro-security or registry changes occurred. These are development
observations, not certification. Raw supervisor evidence is retained in the run-10 archive.

Default chains are bounded at 64; exact deep recursion parity and general native
default signatures remain with bd-59co.2.9.9/WIN-4. Missing required arguments
report 450 without entering a member. Unsupported typed native signatures remain
refused; no dynamic backend, Invoke or name-lookup fallback is introduced.
Formal models abstract away callback dispatch and native code; they cannot prove
this new concrete path. Toolchain availability and Miri results are recorded in
the run receipt, without claiming native/JIT formal proof.

Fresh review and the full workspace run exposed an owning unary-minus admission
case with a Variant destination reaching scalar JIT emission. The helper selection
now includes that destination and other nontrivial numeric conversions. The existing
`tests::jit_matches_vm3_double_negation` regression passes after correction; the
rejected workspace transcript is retained. Eight Lean abstractions pass on the
installed 4.29.1 toolchain. Miri could not execute with the available offline
dependency/sysroot caches; see FTODO-TYPED-DEFAULTS-20261101 without passing credit.

Final `meta-check -Fast -NoArtifacts -NoRegistry`: 3739 passed in 197 suites,
181 ignored; four registry-writing fixtures excluded without credit. Strict
workspace/all-target Clippy, formatting and governance pass. Raw source/check/oracle
evidence is bound by [run10-defaults.json](run10-defaults.json) and
[run10-defaults.zip](run10-defaults.zip). Scoped bd-59co.2.9.15 is closed;
CORE-JIT-LOWERING and the broader owners remain in-progress.
