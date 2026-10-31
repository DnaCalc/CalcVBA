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

The caller owns Module; Session borrows it and retains JIT code until disposal. **Calls in one session retain globals, Static storage and the runtime Err object.** `Module::into_session` uses a safe owning cell for hosts that need an owned session. Separate sessions isolate their mutable storage. `reset` clears state while retaining JIT code; unloading or replacing a module disposes its session. VM3 links lazily on the first bounded call; JIT retains common ExecState storage and initializes once, within the first call budget. Failed initialization is discarded and retried with a fresh budget on the next call, on both backends. `End` remains explicitly refused until whole-project termination/reset is supported. Classes/UDTs remain explicitly refused by this bounded facade. ByRef arguments receive private copies. Loading/calling a function never runs Main. The default Null HAL supplies no external services; `with_host` accepts explicit host-owned capabilities.

| Value | VBA boundary |
|---|---|
| `Empty` | Empty; formula blank coercion belongs to the caller |
| `Number(f64)` | Finite Double input; numeric/date/currency results become binary64 |
| `Text(String)` | BSTR preserving NUL and Unicode; invalid UTF-16/odd-byte results produce `ValueConversion` |
| `Boolean(bool)` | VBA Boolean |
| `Error(i32)` | CVErr code, such as 2042 for `#N/A`; remains a value |
| `Array { rows, columns, cells }` | Row-major scalar cells; two-dimensional inputs with lower bounds 1; rank-one results become one row and rank-two results retain shape |

Malformed, empty or nested arrays and Null/object/record/procedure-reference results are rejected. **Range objects are not admitted or silently converted into arrays.** A future Range-object adapter belongs to the host boundary.

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

The same API runs Add2, a mixed matrix and a noncatchable budget in V8 WebAssembly without imports. DnaCalc `379d51f` now adds the caller-owned CalcFunc registry, CalcFml invocation and separate `calc-vba` adapter pinned to OxVba `74beca5d5bea0370103be3584f18825640eeb5c3`, with default features disabled. Native OneCalc controller/CLI and Chrome/Firefox execute `=Add2(A1,2)`; mounted UI tests, portable adapter tests and bundled browser smoke pass. Calls start with fresh state and exhaustion becomes a worksheet error with a diagnostic. The desktop GUI shell remains DnaCalc `dc-snz.1`.

Range/Caller/WorksheetFunction objects, browser time services, locale-sensitive names and richer effects/thread-safety metadata remain under `bd-59co.2.10.6` and DnaCalc `dc-snz.5`. Value-call delivery does not close those wider capabilities.
