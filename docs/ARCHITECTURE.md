# CalcVBA architecture

The product is an implementation of Office VBA for embedding, optimisation and
standalone output. The [charter](../CHARTER.md) defines scope. Current code and
executable tests describe what works; the implementation is not complete VBA.
Runtime execution and callbacks belong on one owning VBA thread. The compiler
may use parallel work internally.

## Compiler and execution

```text
source / .basproj / host project graph
  -> calcvba-project: modules and referenced-project closure
  -> calcvba-symbol: conditional compilation, declarations, providers, references
  -> calcvba-syntax: lossless VBA CST
  -> calcvba-bind: typed CoreProgram
  -> calcvba-oxir: typed CFG OxProgram / serialized OxImage (.oxi)
       |-- calcvba-vm3: interpreter
       `-- calcvba-jit: Cranelift native code
```

The symbol layer uses the shared parser. Providers describe source projects,
the VBA library, host libraries and imported COM type information. Binding
retains declared types, passing modes, property/default-member decisions and
native descriptors for both backends. Conditional compilation preserves active
source offsets. `calcvba-diagnostics` carries errors between layers.

`calcvba-bundle` holds Core IR and synthetic VBA-library metadata. Remaining
Bundle/Op types support that metadata; VM2 and .oxb are retired execution paths.
VM3 and the JIT execute the same typed OxIR. The JIT has no interpreter fallback.
OxImage stores a project closure; loading validates the image and its programs.
Validation is ordinary executable code, not a formal proof.

## Runtime values and language services

- `calcvba-runtime`: Variant, BSTR, SAFEARRAY, object identity, records and errors.
- `calcvba-eval`: shared value, arithmetic, coercion, collection and array semantics.
- `calcvba-lib`: VBA library functions over those values and host services.
- `calcvba-rt-abi`: shared VM/JIT helpers, physical typed slots and owning calls.
- `calcvba-vm3`: typed-CFG interpreter, frames, aliases, errors and class lifecycle.
- `calcvba-jit`: Cranelift lowering, compiled entries and runtime helper calls.
- `calcvba-aot`: Cranelift Module backend laying out direct native code for PE.
- `calcvba-pe`: template finishing, helper manifest and checked PE patching.
- `calcvba-native-template`: developer-built runtime DLL, with no compiler or VM.

BSTR/VARIANT/SAFEARRAY/IUnknown-shaped carriers preserve VBA/Automation layout
where implemented. Runtime ownership must survive calls, errors and teardown.
Typed arrays retain bounds and element types, including materialized UDT fields.
`On Error`, Err and Resume are runtime semantics shared by the execution paths;
see [the error model](OXIR_VM3_ERROR_MODEL.md).

The typed host entries use unboxed physical storage when admitted. An owning
frame supports objects, strings, Variants, tested aggregate/class families,
internal calls, ByRef aliases and error handling. Unsupported signatures and
bodies return explicit errors. The ordinary Variant entry remains available.
The source compiler is also the foundation for code transformation tools;
there is currently no language-service/LSP or debugger crate in the workspace.

## Embedding and project state

`calcvba-host-api` defines source projects, ordered references, signatures,
bound handles, values, host objects and provider interfaces without linking a
compiler or JIT. `calcvba-host::provider::Provider` composes the implementation.
`calcvba-engine` supplies source/project execution and VM3 package sessions.

Sessions initialise without calling Main, retain VBA globals/Static/Err/object
state and release it on disposal. The code cache shares code, never mutable
VBA state. Handles check session ownership and generation before invocation.
Reset replaces state; reload prepares replacement code before publishing it.

DNA Calc's `crates/calc-vba`, pinned at `953ece7f` when inspected on 2026-11-03,
uses retained invocation frames, read views, instruction budgets, sliced resume,
reload and generation checks. Old activations retain code/state until they
finish. These are used by the calculation scheduler and remain supported.
DNA Calc also implements the HAL host descriptor/policy interface and advertises
its supplied COM-shaped objects. Its checkout is an external consumer, not an
engine source directory. `adapters/spreadsheet-compat` retains an older adapter;
spreadsheet members and formula policy stay outside the engine.

## Host services and Windows interop

`calcvba-hal` delegates console/diagnostics, UI, events, files, environment/time,
native loading and COM operations to host adapters. Null and browser hosts can
refuse unavailable operations. DNA Calc and WASM use the host abstractions;
host policy and descriptors must not be confused with VBA language semantics.
The preset/configuration menu, maturity labels and record/replay adapters are
removed. Host/callback traits allow local state without Send/Sync bounds;
Arc ownership stays at the DNA Calc host API boundary. Private file, console,
directory, settings, portable-object and native-binding state uses local shared
ownership and checked borrows; LastDllError uses a shared local cell.

`calcvba-com` owns type libraries, native calls, dispatch, vtables, marshalling,
connection points and COM error translation. Imported declarations retain IIDs,
member kinds, DISPIDs, slots, argument/default and result information. Dual
members use native vtables without an implicit IDispatch retry. Pure
dispinterfaces use compiler-prebound Invoke; genuinely untyped Object/Variant
calls use dynamic dispatch. Shared native helpers preserve aliases and copyback.

Owning sessions wire document/WithEvents handlers from compiler metadata.
The host supplies document identities and lifecycle timing. Supported native
callbacks enter synchronously on the owning thread, including tested Boolean
ByRef cancellation. Idle native events require the host's STA message pump.
Handlers share canonical globals and Err state. General structural/custom event
signatures and the older package-session event path still need work; ordinary
out-of-process COM events remain in scope. MTA/free-threaded VBA execution does
not. VBA UserForms are intended to integrate with CalcForms/OxForms; there is
no forms-runtime crate here yet.

## Outputs and portability

`calcvba-build` emits OxImage and wrapper/COM-server artifacts.
`calcvba-comhost` is the reusable Windows in-process COM host, currently VM3 based.
`calcvba-cli` exposes source/project/package execution and builds. A wrapped DLL
runs the engine. NativeLibrary instead patches compiled code, named exports and
Windows unwind information into the shipped template without a user linker.
It currently admits ByVal Long/Double/Boolean Functions, direct local procedure
calls, numeric CLng/CDbl/CBool conversions and typed Abs/Sqr helpers.
Long integer division/modulo and Boolean/Long bitwise lowering is shared with VM3/JIT.
Native code uses baseline x64 instructions, 32-argument signatures and at most 128 active
procedure frames. One private result slot is reused for synchronous calls.
Wrappers call their procedure directly; there
is no interpreter or generic dispatch entry. Other widths, owning boundary
types, module state and native EXE output remain unfinished. No XLL target,
registration metadata or conversion code belongs to CalcVBA.
The existing bounded .vbp module-list importer is a compatibility adapter, not
support for the VB6 project system or forms.

Compiler/runtime code and the interpreter support non-Windows hosts. Browser
builds use `calcvba-host` without its default JIT feature; native COM and direct
filesystem access are supplied or refused by the host. Windows interop targets
64-bit Office and x64 native calls. Wider Windows architectures, broader native
signatures, nominal interface/record arrays, recursion limits and full JIT
COM serving remain incomplete.

## Tests

`./scripts/check.ps1` runs fmt, strict Clippy and workspace tests. Useful tests
include language/runtime regressions, VM3/JIT values and Err comparisons,
carrier ownership checks, controlled native COM fixtures, captured VBA probes,
and saved XLSM source tests. Live Excel comparisons are opt-in and keep dialog
handling and cleanup scoped to their owned processes. Fixtures live beside
tests, under `conformance/` or under `oracle/`; temporary runs belong in `temp/`.
