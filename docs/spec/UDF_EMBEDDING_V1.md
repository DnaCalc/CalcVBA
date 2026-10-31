# UDF value-call contract v1

2026-10-31. Bounded API delivered by `oxvba-embed`; DNA OneCalc value calls are delivered by DnaCalc `dc-snz.3` at `379d51f`. Wider integration remains in progress under OxVba `bd-59co.2.10.6` and DnaCalc `dc-snz.5`. API owner: `bd-59co.2.10.4`. Clauses: `HOST-SESSION-001`, `SYS-DUAL-001`, `SEC-BOUNDARY-001`. Consumers should pin a reviewed revision. The later CalcVba rename is deferred.

```rust
use oxvba_embed::{Backend, Module, Value};
let module = Module::load("Math", "Function Add2(a, b): Add2 = a + b: End Function\n")?;
for function in module.functions() {
    // name, qualified_name, parameter names/types/ByRef/Optional/ParamArray, return_type
}
let session = module.prepare(Backend::Interpreter)?; // or Backend::Jit
assert_eq!(session.call("Math.Add2", &[Value::Number(19.0), Value::Number(23.0)], 1000)?, Value::Number(42.0));
# Ok::<(), oxvba_embed::Error>(())
```

Load uses the existing compiler and OxIR verifier. The catalogue comes from compiler public exports, excluding Private functions and Subs. Calls accept bare or module-qualified names with ASCII case folding; broader VBA locale-sensitive name matching remains follow-up work under `bd-59co.2.10.6`. Calls are positional. Omitted Optional arguments use compiler-folded defaults, typed zero values, or the VBA Missing sentinel for Variant. `Value::Missing` explicitly omits a positional slot; it differs from `Empty`. A trailing ParamArray packs remaining arguments into a private zero-based Variant SAFEARRAY, including the empty 0..-1 case. Scalar types and Variant-array parameters are admitted; unsupported signatures return typed errors.

The caller owns Module; Session borrows it and retains JIT code until disposal. `CodeCache` shares verified immutable metadata/code among owned sessions on one thread; it never shares VBA state or host-object tables. **Calls in one session retain globals, Static storage and the runtime Err object.** `Module::into_session` uses a safe owning cell for hosts that need an owned session. Separate sessions isolate their mutable storage. `reset` clears state while retaining JIT code; unloading or replacing a module disposes its session. VM3 links lazily on the first bounded call; JIT retains common ExecState storage and initializes once, within the first call budget. Failed initialization is discarded and retried with a fresh budget on the next call, on both backends. `End` remains explicitly refused until whole-project termination/reset is supported. Classes/UDTs remain explicitly refused by this bounded facade. ByRef arguments receive private copies. Loading/calling a function never runs Main. The default Null HAL supplies no external services; `with_host` accepts explicit host-owned capabilities.

| Value | VBA boundary |
|---|---|
| `Missing` | Explicit omitted positional argument |
| `Range` | Live host-owned rectangle, with `load_spreadsheet` and a current SpreadsheetContext |
| `Empty` | Empty; formula blank coercion belongs to the caller |
| `Number(f64)` | Finite Double input; numeric/date/currency results become binary64 |
| `Text(String)` | BSTR preserving NUL and Unicode; invalid UTF-16/odd-byte results produce `ValueConversion` |
| `Boolean(bool)` | VBA Boolean |
| `Error(i32)` | CVErr code, such as 2042 for `#N/A`; remains a value |
| `Array { rows, columns, cells }` | Row-major scalar cells; two-dimensional inputs with lower bounds 1; rank-one results become one row and rank-two results retain shape |

Malformed, empty or nested arrays and Null/non-Range-object/record/procedure-reference results are rejected. `load_spreadsheet` supplies a bounded read-only Application/Range/WorksheetFunction host model. Range identities retain coordinates and read through the current call context; they never retain the borrowed reader. Members: Value2, Count, Row, Column, Address, Cells/Item and Rows/Columns Count. Coordinates and object/read quotas are checked. A returned Range materializes its live Value2. Indexed Rows/Columns and other members remain refused. Caller is the current host-supplied rectangle; Volatile reports a request to the host. Scheduling and callback admission belong to the host. This is not full Excel object-model parity. Host read/callback refusals remain `HostRefused`, even if VBA catches the resulting runtime error.

`Error::Runtime { number, description, vba_source }` is an unhandled VBA error: a formula host maps it to `#VALUE!` and retains its diagnostic. `Value::Error` preserves the worksheet error code. Compile failures, unknown names, wrong argument counts, value conversion, unsupported signatures/backends, absent JIT support and exhausted budgets have distinct error variants. JIT preparation compiles the whole module once or declines it before execution, with no interpreter fallback. Rich/dynamic Err.Raise metadata remains a JIT limitation; constant-number Err.Raise is tested on both backends.

`max_steps` counts OxIR basic-block entries, including initialization and nested calls. Zero prevents execution. Exhaustion returns `BudgetExceeded` to the host, bypassing On Error. Budgets reset per call. This is cooperative control; an executing host/native call is not interrupted. Full persistent project sessions and general JIT recursion/reentry guarantees remain existing core work.

For DNA OneCalc, register the admitted functions in the caller-owned **calc-func registry / calc-fml resolver**. OneCalc needs no calc-engine, calc-doc or calc-host dependency. The DnaCalc side owns name conflicts, workbook qualification, registration generations/invalidation, references and formula-error mapping. Reloading creates a new Module/session and invalidates old registrations. This repository has no calculation engine or DnaCalc dependency.

Interpreter-only consumers use `default-features = false`; desktop builds may retain the default `jit` feature.

```powershell
cargo test -p oxvba-embed
cargo test -p oxvba-embed --no-default-features
cargo run -p oxvba-embed --example add2
cargo build -p oxvba-embed --no-default-features --example browser_probe --target wasm32-unknown-unknown
node scripts/test-udf-wasm.mjs
```

The original V8 no-import example verifies Add2, a mixed matrix and a noncatchable budget. DNA OneCalc supplies the caller-owned CalcFunc registry, CalcFml invocation and separate `calc-vba` adapter. Native builds prefer JIT; unsupported admission may fall back before execution, never after a runtime failure or exhausted budget. Browser builds omit the JIT. The adapter admits SUM/MIN/MAX/AVERAGE/MATCH callbacks through the existing CalcFunc evaluator/reader. Other methods and unmeasured scalar-array shapes remain explicit refusals. UDF plans remain conservatively volatile; precise scheduling/effects parity remains open.

The bounded cache compares exact serialized verified OxIR, function signatures including folded defaults, backend, host policy/descriptor and spreadsheet mode, and a native target/configuration identity (ISA/CPU flags, codegen flags, helper/carrier ABI). Configuration is captured once and used for both key and compilation. LRU entry and description-byte limits bound retained cache owners, not executable arena sizes or code held by live sessions. Eviction/clear cannot invalidate live sessions. Compilation owns unpublished code through an unwind/error guard; published CompiledImage owns its allocation until its final session/cache owner drops. Cross-thread sharing, on-disk packaging, general JIT reentry, class disposal and descriptor lifetime remain separate open work. See [run 3 evidence](../evidence/programs/ideal-2026-07/core/CORE-8/udf-session-embedding.md).
