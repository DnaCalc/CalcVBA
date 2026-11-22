# CalcVBA architecture

The product is an implementation of Office VBA for embedding, optimisation and
standalone output. The [charter](../CHARTER.md) defines scope. Current code and
executable tests describe what works; the implementation is not complete VBA.
Runtime execution and callbacks belong on one owning VBA thread. The compiler
may use parallel work internally.

`calcvba-office` owns portable CFB/MS-OVBA and workbook-container I/O. Its
reader produces compiler manifests without VBOM or an Office installation;
UTF-8 `.basproj` folders retain VBE attributes and `.frm`/`.frx` resources.
Write-back validates the project codepage, replaces source streams and removes
or invalidates compiled caches. The original host workbook supplies its document
objects. Locked projects and encrypted files are refused; signed write-back is
not supported. Live checks use one recorded, owned Excel process across cases.

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

- `calcvba-automation`: the platform values themselves: the `VARIANT` cell (`Variant`), BSTR,
  SAFEARRAY, `BRECORD` records, interface references, Decimal and Currency, laid out as
  the platform's Automation ABI on every target.
- `calcvba-runtime`: VBA's runtime semantics over those values (project objects, collections,
  coercion, dates, errors, address-of), re-exporting `calcvba-automation`.
- `calcvba-eval`: shared value, arithmetic, coercion, collection and array semantics.
- `calcvba-lib`: VBA library functions over those values and host services.
- `calcvba-rt-abi`: shared VM/JIT helpers, physical typed slots and owning calls.
- `calcvba-vm3`: typed-CFG interpreter, frames, aliases, errors and class lifecycle.
- `calcvba-jit`: Cranelift lowering, compiled entries and runtime helper calls.
- `calcvba-aot`: Cranelift Module backend laying out direct native code for PE and ELF.
- `calcvba-pe`: template finishing, helper manifest and checked native image patching.
- `calcvba-native-template`: developer-built Windows and Linux runtime templates.
  Executable startup calls a patched native Main entry and shutdown.

BSTR/VARIANT/SAFEARRAY/IUnknown-shaped carriers preserve VBA/Automation layout
where implemented. Runtime ownership must survive calls, errors and teardown.
Typed arrays retain bounds and element types, including materialized UDT fields.
`On Error`, Err and Resume are runtime semantics shared by the execution paths;
see [the error model](OXIR_VM3_ERROR_MODEL.md).

Put every new VBA semantic decision in a helper shared by VM3 and the JIT.
Engine code supplies frames, procedure entry and control flow. Review changes
against both paths and cover observable changes with differential regressions.
VM3 supplies the reference behavior where Excel observations agree.

| Area | Shared implementation | Engine-specific work |
| --- | --- | --- |
| Numeric conversions and Variant coercion | `calcvba-eval` implements arithmetic and conversion kernels. `calcvba-rt-abi` applies declared-slot conversions and follows object defaults in value contexts. | Each engine evaluates operands and invokes default members. JIT lowering selects scalar instructions and helper calls. |
| Err and Resume | `ErrEngine` owns activation policy, explicit faults, Err field writes and resume transitions. `calcvba-runtime` maps standard error numbers to descriptions, including the shared array subscript fault. | VM3 and generated JIT code transfer control to their own handlers and restore activation state. |
| ByRef and ParamArray | OxIR identifies argument modes. Shared helpers bind named actuals, coerce private argument cells, construct typed ByRef Variant views and prepare ParamArray copy-out for every alias. | VM3 `Loc` and JIT `SlotAlias` locate and retain aliases in their own frame layouts. Typed and native entries use physical slot layouts. |
| Object lifetime | `object_owner` and `object_ref` own reference counts and queue final release on the owning thread. The shared drain runs queued finalizers through reentrant releases, restores Err and removes WithEvents subscriptions. | Each engine releases its frame values and supplies its Class_Initialize and Class_Terminate procedure entries. |
| Late-bound calls | Shared helpers select the receiver route and default member, and the runtime dispatch plan cache resolves named project members. Shared COM marshalling and Collection dispatch implement host argument and member rules. | Each engine evaluates operands, resolves frame aliases, calls project members and provides its own host callback entry. |

The JIT's typed entries and AOT both use `ScalarLayout` and `ScalarLower` to
lower the scalar instruction subset. The general JIT path uses `lowering.rs`
and its runtime slot helpers. AOT's `native.rs` and `native_values.rs` add
separate admission, frame setup, calls, error transfer and owning value
lowering for native exports. AOT's template also implements its own finalizer
drain and ParamArray write-through for native storage. The `calcvba-aot` crate
supplies the PE and ELF module backends used by that compiler.

The typed host entries use unboxed physical storage when admitted. An owning
frame supports objects, strings, Variants, tested aggregate/class families,
internal calls, ByRef aliases and error handling. Unsupported signatures and
bodies return explicit errors. The ordinary Variant entry remains available.
The source compiler is also the foundation for code transformation tools;
there is currently no language-service/LSP or debugger crate in the workspace.
The [VM3 debugger design](design/vm3-debugger.md) describes the source and
runtime hooks needed by DNA Calc and Formula Explorer.

`tools/calcvba-udf` reads bound signatures and generates ordinary VBA XLL shims
and native export declarations. It uses the existing native builder; Excel's
registration and value conversion remain VBA source outside the compiler.

### JIT refusals

The JIT compiles every procedure in the linked image. It reports the first
refused shape with its procedure when available. Typed host entries and native
exports have additional admission rules for their own physical calling
conventions.

| Shape | Reason for a JIT refusal |
| --- | --- |
| Frame, global, parameter and result carriers | A declared array with an array or procedure-reference element type has no JIT SAFEARRAY carrier. Variant elements can hold arrays. Record operations need a retained layout. The destination of a compiled procedure call must have the return type or be a Variant. |
| Arrays, UDTs and compound places | Array get/set and ReDim require a supported element carrier or a Variant and a shape consistent with the instruction. Typed element operations use an exact destination carrier or a Variant. The helper ABI bounds dimensions, index counts and field indices. |
| Numeric, logical, comparison and string instructions | The scalar paths admit specific operand and result lanes. Variant paths require operands that can be represented by the runtime helper descriptors. The lowering reports refused coercion, boxing, unboxing and operator shapes at the instruction. |
| Direct and linked procedure calls | Static calls need argument metadata, the formal argument count and exact ByRef aliases. A raw omitted operand uses the declared Optional default, Missing for an Optional Variant without a default, or the initialized value for an Optional numeric, Boolean, Date or String parameter. Other parameter types require a declared default for a raw omission. A ParamArray must be last and use a ByVal Variant or dynamic Variant-array pack. Linked calls need the referenced unit and matching procedure export. |
| Indirect procedure calls | The target must be a procedure-reference place. A statically known signature uses the supported direct-call carriers. The unknown-signature lowering has narrower Long and String call paths and limited result carriers. It needs an admitted signature candidate for other shapes. |
| Classes, object tests and dispatch | Project class imports need a linked class export. The VBA-library `NewExtern` path has a Collection case. Other classes need a supported constructor instruction. A single-project `TypeOf` on an untyped receiver can be refused if the class or interface identity is unavailable. Project and COM calls have to fit the argument descriptors, which currently refuse an OxIR `Const` call argument on that path. |
| VBA library imports | Legacy VBA `CallExtern` accepts the explicit library argument and result shapes in the lowering. `CallNative` built-ins use the shared runtime library. |
| Native host calls and target support | Native COM and Declare calls depend on the available host, platform and call descriptor. Linux x64 supports the tested native Declare and export signatures. Host failures can surface as VBA runtime errors after compilation. WebAssembly hosts use VM3. |
| Internal representation | Unsupported instructions and terminators, metadata that exceeds the helper ABI and malformed internal places are diagnosed during compilation. A compiler defect and a VBA runtime error have distinct outcomes. |

The source parity runner's `inspect` mode checks admission after shared binding
and elaboration. Its results describe the compiler boundary. Runtime comparisons
with VM3 require a runnable project entry and the host services the case uses.
See [procedure admission](../crates/calcvba-jit/src/admission.rs) and
[general JIT lowering](../crates/calcvba-jit/src/lowering.rs).

The package round-trip regression exercises `Omitted` for a typed optional
Long through ByVal and ByRef calls on both engines, with a declared default
of 7 and with the implicit zero. The binder materializes defaults for
ordinary statically resolved source calls. Late-bound project calls can
carry an explicit omission to the callee and use the same JIT frame setup.


## Embedding and project state

`calcvba-host-api` defines source projects, ordered references, signatures,
bound handles, values, host objects and provider interfaces without linking a
compiler or JIT. `calcvba-host::provider::Provider` composes the implementation.
`calcvba-engine` supplies source/project execution and VM3 package sessions.

At the host API boundary, `RuntimeProvider::select_backend(BackendPolicy::Automatic)`
selects JIT on Windows, Linux and macOS when the provider was built with JIT.
WebAssembly (browser and WASI) and interpreter-only builds select VM3.
Hosts select `Debugging` for a session with VBA debugging or `NoCodeGeneration`
when their deployment prohibits executable memory. Pass the selected `Backend`
to `load`, `load_graph` or `load_with_object_libraries`. Explicit `Interpreter`
and `Jit` requests remain available. The selection happens before compilation.
JIT refusals and failures are returned to the host, and reload keeps the
session's selected engine.

Guarded mode uses the same engine selection as ordinary execution. Host
capability checks apply to both engines. A Guarded deployment that also
prohibits native code generation selects `NoCodeGeneration`.

Sessions initialise without calling Main, retain VBA globals/Static/Err/object
state and release it on disposal. The code cache shares code, never mutable
VBA state. Handles check session ownership and generation before invocation.
Reset replaces state; reload prepares replacement code before publishing it.

The VM3-backed `ProjectRuntimeSession` also exposes a checked default-instance
view for same-thread host callbacks. Lookup creates nothing, reset compares the
expected identity, and explicit host release drains eligible finalizers while
preserving held object/control references. The view expires with its runtime state.

DNA Calc's `crates/calc-vba`, pinned at `953ece7f` when inspected on 2026-11-03,
uses retained invocation frames, read views, instruction budgets, sliced resume,
reload and generation checks. Old activations retain code/state until they
finish. These are used by the calculation scheduler and remain supported.
DNA Calc also implements the HAL host descriptor/policy interface and advertises
its supplied COM-shaped objects. Its checkout is an external consumer, not an
engine source directory. Spreadsheet members and formula policy stay outside
the engine.

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
`calcvba-comhost` is the reusable Windows in-process COM host. It loads the
project image into a VM3 or JIT session and returns the project instance's own
COM interface from its class factory.
`calcvba-cli` exposes source/project/package execution and builds. A wrapped DLL
runs the engine. NativeLibrary patches compiled code, named exports and unwind
information into the shipped Windows PE or GNU/Linux ELF template without a user linker.
It currently admits all VBA native numeric/Boolean boundary widths, Functions and Subs, direct local procedure
calls, numeric CLng/CDbl/CBool conversions and typed Abs/Sqr helpers.
Long integer division/modulo and Boolean/Long bitwise lowering is shared with VM3/JIT.
Native code uses baseline x64 instructions, 32-argument signatures and at most 128 active
procedure frames. One private result slot is reused for synchronous calls.
Wrappers call their procedure directly; there
is no interpreter or generic dispatch entry. BSTR/UTF-16/ANSI inputs, native UDT
layout, SAFEARRAYs, VARIANTs and COM references use direct value-operation helpers.
Project class construction, fields, properties and methods use owning objects
and compiled entries, including local `As New` recreation and same-thread
Initialize/Terminate calls. Collection and Implements calls use the same owning
carriers. Early COM calls retain imported interface/vtable descriptors and share
the JIT's native marshalling; late COM calls share its host services. Literal
project CallByName calls select compiled members.
TypeOf uses project identity or the imported COM IID. Unsupported class field
initializers are refused. `scripts/test-native-excel-members.ps1` compares an
Excel COMReference body on VM3, JIT, AOT and Office using an existing owned session.
Internal optional arguments and ParamArray packs are compiled directly. Native
ParamArray aliases coerce to declared caller storage before changing either value;
overflow leaves both unchanged and fixed String copy-out pads or truncates.
Fixed-length String defaults contain NULs; assignments pad or truncate.
GoSub uses a per-frame return stack and branches in compiled code.
Writable `.cvbdata` holds module state and the once-only initialization status.
Each DLL retains the identity of its first calling thread; foreign threads return
zero/E_WRONG_THREAD without touching VBA or caller storage. ErrorPolicy and
CalcVbaGetLastError describe failures without unwinding through the foreign ABI.
Native exits share one owning-value cleanup block. Wider numeric carriers are
admitted to direct native cells; retained host activations keep their existing
carrier set. Exe enters compiled Sub Main directly. ComExe supplies native class
factories and dual vtables; Automation invokes those same compiled entries through
type information. Constructors, methods, finalizers and message dispatch run on
the owning STA. The server loads metadata and a TLB, without an executable OxImage.
Public scalar/BSTR boundaries include ByRef arguments; optional parameters,
Variant/Object boundaries, exported events and custom interfaces are refused.
The private Counter example tests early-bound and Automation clients with
temporary HKCU registration and exact owned-process cleanup. Broader
language coverage remains unfinished. No XLL target,
registration metadata or conversion code belongs to CalcVBA.
The existing bounded .vbp module-list importer is a compatibility adapter, not
support for the VB6 project system or forms.

Compiler/runtime code and the interpreter support non-Windows hosts. Browser
builds use `calcvba-host` without its default JIT feature; native COM and direct
filesystem access are supplied or refused by the host. WASI builds also use
VM3. The embedder maps Windows drives to preopened guest paths and supplies
`XDG_CONFIG_HOME` inside a writable preopen for VBA settings. It can set
`CALCVBA_WASI_UTC_OFFSET_MINUTES` to a fixed UTC offset for local dates and times.
Windows interop targets
64-bit Office and x64 native calls. Wider Windows architectures, broader native
signatures, nominal interface/record arrays and recursion limits remain incomplete.

## Tests

`./scripts/check.ps1` runs fmt, strict Clippy and workspace tests. Useful tests
include language/runtime regressions, VM3/JIT values and Err comparisons,
carrier ownership checks, controlled native COM fixtures, captured VBA probes,
and saved XLSM source tests. Live Excel comparisons are opt-in and keep dialog
handling and cleanup scoped to their owned processes. Fixtures live beside
tests, under `conformance/` or under `oracle/`; temporary runs belong in `temp/`.
