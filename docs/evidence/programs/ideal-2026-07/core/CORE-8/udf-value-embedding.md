# Value-call embedding and source-free execution

2026-10-31. Bounded delivery owners: `bd-59co.2.10.4` (UDF facade) and
`bd-59co.2.10.5` (OxImage CLI). The wider CORE-8 row remains in-progress/planned
as recorded in its canonical matrix; no persistent-session or cache completion
is claimed. The DnaCalc tree was read for context and was not edited.

Code checkpoint: `0113e7ee9d2675393cb2e88b551eb91171a66507`.

Run-6 note: the spreadsheet API/tests/browser example described below have moved
to `adapters/spreadsheet-compat`. The historical commands remain tied to this
checkpoint; current commands use `-p oxvba-spreadsheet-compat`. Current neutral
hosting and measurements are in [typed-host-entries.md](../CORE-7/typed-host-entries.md).

The contract is `docs/spec/UDF_EMBEDDING_V1.md`. Tests cover compiler-derived
public function/parameter metadata, the requested one-line Add2, scalar values,
worksheet errors versus runtime failures, mixed two-dimensional arrays,
ByRef private copies, fresh globals, repeated preparation/disposal, invalid
source/value/arity, backend declines, and noncatchable loop/recursion budgets.
Both VM3 and JIT run through the same value API. The interpreter-only feature
build runs in WebAssembly/V8 without imports, including Add2, a mixed matrix
and budget exhaustion. The wasm32 build also exposed and corrected an existing
8-GiB usize constant overflow by performing the limit calculation as u64.

The compiled-target step packages the existing typed OxImage and runs it with
`build --target OxImage` and `run-image --backend vm3|jit`. A real CLI process
test builds the Exe top-level-statement language extension, deletes its source
and project, then executes 42 through each backend and rejects corrupt images.
The host test additionally honors a nonfinal artifact entry and retains VBA
runtime error numbers. This is a backend-neutral executable artifact, not a
native DLL/EXE or an AOT output claim.

Commands:

```powershell
cargo test -p oxvba-embed
cargo test -p oxvba-embed --no-default-features
cargo test -p oxvba-cli --test packaged_image
cargo test -p oxvba-host --test execute_image
cargo build -p oxvba-embed --no-default-features --example browser_probe --target wasm32-unknown-unknown
node scripts/test-udf-wasm.mjs
```

Formal scope: all five Lean modules passed with the already installed Lean
4.29.1, including the new abstract ExecutionBudget model (sticky zero,
successful exact decrement, monotonicity). The repository's 4.19 pin was not
installed or changed. The proof concerns the abstract counter, not refinement
to Rust, native ABI behavior or backend abort propagation. Those are covered
here by exhaustive small-budget/u64 boundary transitions and executable VM/JIT
On Error and recursion tests. Native cargo-kani is absent (`cargo kani --version`:
no such command). Formal follow-up is recorded in the formal backlog.
Existing remote formal obligations are not treated as passing via empty test
filters.

Residual owners: `bd-59co.2.10.6` owns host objects, browser clocks, Optional/
ParamArray calls, persistent state, metadata and the real consumer exercise;
`.2/.3` retain general JIT sessions and cache work. DNA Calc `dc-snz.3` owns
calc-func/calc-fml registration. Native output remains the existing WIN-12 work.

Checkpoint verification: `scripts/meta-check.ps1 -Fast -NoArtifacts` passed
with RUST_TEST_THREADS=1: 3,634 tests, strict workspace Clippy, formatting and
governance. The default differential timeout includes its shared fixture-lock
queue; a previous concurrent run timed out on a fast compile diagnostic. That
harness issue is not a semantic golden change. Ten interpreter-only embedding
tests and fresh WebAssembly/V8 probes also passed. Fresh-eyes review corrected
missing procedure terminators, empty-array admission, nonfinite input handling,
artifact entry selection/dependency initialization, native ownership and stale
Windows decline assertions. Property terminators and negative Excel capture
remain under their explicit open owners.
