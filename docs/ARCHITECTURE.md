# OxVba Architecture

Date: 2026-07-10
Status: current implementation realization and gap map
Destination authority: [`spec/OXVBA_SYSTEM_CONTRACT_V1.md`](spec/OXVBA_SYSTEM_CONTRACT_V1.md)
Evidence review: [`OXVBA_POST_JIT_STATUS_REVIEW_2026-07-10.md`](OXVBA_POST_JIT_STATUS_REVIEW_2026-07-10.md)

## 1. Role of this document

The system contract states what completed OxVba must become. This document states what the repository is now, how the active crates fit together and where implementation diverges from that destination.

Current code and executable tests are implementation truth. A green subset is not a completed capability claim. The canonical worksets and validation matrices own delivery/evidence status; historical plans and reports are not architecture authority.

## 2. Current system summary

The active production path is:

```text
.basproj/.vbp/source
  -> oxvba-project project/reference closure
  -> oxvba-symbol target-aware conditional preprocessing
  -> oxvba-syntax lossless CST of the supplied active-source view
  -> oxvba-symbol declarations/providers/resolution environment
  -> oxvba-bind typed binding and CoreProgram
  -> oxvba-oxir::elaborate typed CFG OxProgram
  -> OxImage (.oxi project-closure artifact)
       |-- oxvba-vm3 interpreter
       \-- oxvba-jit Cranelift compiler
```

This is the sole production compiler/execution stack. The legacy compiler, HIR/front-end fallback, source-rewrite compiler, VM2, CoreProgram-to-Bundle linearizer and `.oxb` product execution path have been removed or retired.

The architecture is broad and real but not complete. The compiler/binder is the most mature layer. VM3 and the JIT cover a large platform-neutral language/runtime/library surface. The JIT now executes bounded Windows x64 scalar early/late-COM and Declare subsets. Early dual calls preserve compiler metadata through native vtable transport on both runtimes; the owning typed-slot path additionally supports a bounded imported scalar/BSTR native family with policy, IID, apartment and proxy checks; general early COM, pointer/callback execution and Windows parity remain open. There is no active clean-stack language-service implementation.

## 3. Workspace ownership

### Compiler and project

- `oxvba-project` parses `.basproj`, the supported `.vbp` subset and convention projects, loads modules and constructs referenced-project closures.
- `oxvba-syntax` provides the lossless lexer/parser and green/red CST.
- `oxvba-symbol` owns conditional compilation, declaration scanning, scopes, symbols, signatures and providers for projects, referenced surfaces, the VBA library, host references, COM typelibs and Declare descriptors.
- `oxvba-bind` performs typed binding and emits `oxvba_bundle::coreir::CoreProgram`.
- `oxvba-diagnostics` provides the shared cross-layer diagnostic DTO; producing crates retain ownership of semantic error meaning.

### Executable semantics and runtime

- `oxvba-bundle` now primarily owns Core IR plus bounded legacy Bundle/Op types used by the synthetic VBA-library metadata path. It is not the current executable package producer.
- `oxvba-oxir` elaborates Core IR into typed CFG OxIR, defines OxProgram/OxImage, normalization/analysis passes and the current program verifier.
- `oxvba-runtime` owns Variant, BStr, SafeArray, ObjectRef, record and related value/lifecycle carriers.
- `oxvba-eval` owns shared value-semantic kernels used by VM3 and JIT for the extracted operation families.
- `oxvba-rt-abi` owns VM/JIT-neutral runtime cells and helper decisions used across the backend boundary.
- `oxvba-lib` implements VBA base-library bodies over runtime values and HAL services.
- `oxvba-vm3` is the sole typed-CFG interpreter and the JIT reference backend.
- `oxvba-jit` directly lowers linked OxProgram sets to Cranelift native code with a hard no-fallback boundary.
- `oxvba-differential` owns VM3 golden and VM3/JIT/oracle comparison harnesses.

### Host, platform and outputs

- `oxvba-hal` owns host capability/profile/policy and adapter contracts.
- `oxvba-com` owns Windows-first COM metadata, invocation, dynamic-object and wire-boundary services.
- `oxvba-host` orchestrates compiler/backend execution and VM3-backed package sessions.
- `oxvba-build` emits OxImage and wrapper/COM-server artifacts.
- `oxvba-comhost` is the reusable Windows in-process COM-server host, currently VM3-backed.
- `oxvba-cli` exposes source/project run and build commands.
- `calcvba-host-api` owns the small host contract: source modules and ordered references, signatures, opaque generation-checked handles, typed slots, host-object tables and the runtime-provider interface. It depends on exact runtime carriers and links no compiler, HAL, VM3, JIT or libffi. `oxvba-embed::provider::Provider` is the application's composition edge to those implementations. `oxvba-embed` retains compiler/session/cache ownership. The legacy spreadsheet model stays in `adapters/spreadsheet-compat`; no members were added there. DNA Calc migration remains external; see [`spec/HOSTING_AND_EARLY_BINDING.md`](spec/HOSTING_AND_EARLY_BINDING.md).

No active `oxvba-languageservice`, `oxvba-lsp`, `oxvba-debug` or forms-runtime crate exists in the current workspace.

## 4. Source and compiler realization

### Current shape

For each supplied module, `oxvba-symbol` applies target-aware conditional preprocessing and calls the shared parser. Length-preserving blanking keeps active token offsets relative to the supplied module text. Parsed modules are retained in one resolution environment and scanned into project/provider surfaces before binding.

The binder walks resolved CST once, infers types, inserts coercions and emits Core IR with explicit places, assignment intent, calls, properties/default members, arrays, records, objects, events, errors and import/export descriptors.

The provider architecture is the right target shape: the binder asks what a symbol means rather than branching on whether it came from source, the VBA library, a host or COM metadata.

### Material gaps against the contract

- identifier scanning can panic on valid non-ASCII UTF-8 input;
- malformed conditional expressions/directives can fail open;
- original file provenance is incomplete after some project normalization/generation;
- referenced-project public data fields are not represented across compiled surfaces;
- declared return types are erased on important referenced/library/native call routes;
- ByVal/ByRef, arrays/UDTs and Declare legality checks are incomplete;
- many symbol/bind diagnostics lack source locations;
- grammar/language matrices do not yet provide current-route evidence for the full surface;
- the compiler does not yet publish the immutable `AnalysisResultV1`/`AnalysisMode::{Strict, Editor}` fact contract, stable identities or half-open UTF-8 span/provenance maps required by language services.

The current detailed target is [`spec/OXVBA_COMPILER_AND_SEMANTIC_ANALYSIS_CONTRACT_V2.md`](spec/OXVBA_COMPILER_AND_SEMANTIC_ANALYSIS_CONTRACT_V2.md).

## 5. Core IR, OxIR and OxImage realization

Core IR is the resolved semantic tree emitted by the binder. `oxvba-oxir::elaborate` converts it into OxProgram: typed locals/places, basic blocks, instructions, terminators, fault edges, functions, globals, classes, records, external descriptors, COM interfaces, imports and exports.

OxImage serializes a project closure as pretty JSON `.oxi` with a schema magic/version, program list and entry index. VM3-backed host/build paths can load it and create package sessions.

### Material gaps against the contract

- `OxImage::validate` checks only basic header/count/entry conditions and does not seal fully verified programs;
- production VM3/JIT APIs can still receive raw OxProgram values;
- VM3 linking follows a last-program convention instead of treating image entry as fully authoritative;
- the verifier is incomplete for several ID, type, arity, descriptor, export, event and effect families;
- duplicate/ambiguous case-folded link identities are not comprehensively rejected;
- OxImage lacks content digest, required helper-catalog/carrier ABI identities and digests, target/capability requirements, full provenance and source/debug maps;
- a few OxIR operations have divergent VM3/JIT dispositions;
- product consumers do not yet share a sealed `VerifiedOxProgram`/`VerifiedOxImage` boundary.

The current target is [`spec/OXVBA_OXIR_AND_IMAGE_CONTRACT_V1.md`](spec/OXVBA_OXIR_AND_IMAGE_CONTRACT_V1.md).

## 6. Runtime, library and host realization

Variant is the canonical execution carrier. BStr uses BSTR-shaped UTF-16 storage; SafeArray represents typed/dynamic arrays; ObjectRef provides IUnknown-shaped identity; records use descriptor-backed storage. The representation direction is exact VBA/OLE carrier layout rather than boundary-only projections.

`oxvba-eval` shares a meaningful value core across VM3 and JIT, while runtime, library, array/object/call/lifecycle/error behavior is still distributed among runtime, rt-abi, VM3, JIT, lib and host. `oxvba-lib` exposes a broad VBA-library implementation; host-sensitive operations delegate to HAL.

### Material gaps against the contract

- the shared semantic-kernel extraction is incomplete;
- VM3 and rt-abi duplicate some class/interface descriptor projection;
- descriptor and host/image session paths contain process-lifetime leaks, and helper registration is not yet generated everywhere from one versioned `oxvba-rt-abi` catalog;
- some public rt-abi functions hide raw-pointer safety contracts behind safe Rust signatures;
- panic/fault and manual drain/reentrancy state need RAII hardening;
- the base library has no member-by-member typed/compiler/VM3/JIT/oracle completion matrix;
- stateful file I/O, locale and several host-sensitive families remain bounded subsets;
- host denial, unsupported implementation and VBA runtime failure are not consistently separated in all paths.

The exact layout doctrine remains [`spec/OXVBA_REPRESENTATION_LAYOUT_DOCTRINE_V1.md`](spec/OXVBA_REPRESENTATION_LAYOUT_DOCTRINE_V1.md); library/host completion is governed by system clauses `RUNTIME-*`, `LIB-VBA-001` and `HOST-*`.

## 7. VM3 realization

VM3 is the sole product interpreter and the JIT reference backend. It interprets typed OxIR with heap-owned frames, typed places, ByRef aliases, error/Resume routing, class lifecycle, project events and broad library/runtime support. Its admitted scalar entry reads the same OxFunc through unboxed stack storage without a Variant frame; universal scalar calls adapt to that entry.

Focused VM/runtime suites are broad, and the VM3 golden contains hundreds of value/error rows. This is useful regression evidence, not automatic VBA authority.

### Material gaps against the contract

- the current golden gate has a reproducible BSTR-balance failure on a policy error path;
- VM3 does not implement every verifier-accepted OxIR operation;
- loader verification and explicit image-entry handling are incomplete;
- class/interface descriptor ownership and repeated-session lifetime need hardening;
- some value/error edges still need live Excel/VBA clarification;
- the differential observable does not yet structurally compare every carrier/lifecycle family.

VM3 destination behavior is specified with OxIR/OxImage in [`spec/OXVBA_OXIR_AND_IMAGE_CONTRACT_V1.md`](spec/OXVBA_OXIR_AND_IMAGE_CONTRACT_V1.md).

## 8. JIT realization

The JIT is a real Cranelift backend. It directly lowers linked OxProgram blocks, compiles whole accepted program sets without VM fallback and supports broad control flow, calls, values, arrays, records, project classes, lifecycle and project events.

A bounded typed primary family now accepts zero to four required ByVal Long/Double
arguments, a runtime budget context and a typed result pointer, returning an explicit
status. Admitted pure numeric bodies use unboxed physical slots; successful warm
host calls allocate no heap storage. VM3 has a matching unboxed interpreter route.
The run-7 family additionally accepts borrowed host objects and compiler-bound
Double-returning slots with zero arguments or one required ByVal Long. JIT emits
an exact typed indirect call; VM3 calls the matching checked native signature.
Neither route constructs Variant arguments or uses libffi. These are explicit
portable host tables with a runtime IUnknown prefix, separate from native COM's
Automation vtables. Same-session typed callback reentry requires the active
context, shares its budget and has a 64-activation limit across sessions on the
thread. Dynamic and general COM-event same-session reentry remain refused.
The dynamic JIT entry adapts its Variant frame to that primary family where admitted;
image run-entry bodies retain their existing diagnostic local snapshots.

The neutral embedding API binds handles once, checks owner identity and generation
on each call, initializes without running Main, and retains independent project
state over shared cached code. Reload prepares a replacement before invalidating
old generations; a handle never publishes an executable address. The source-module
facade still refuses project classes/UDTs and does not replace general package
sessions. Unsupported typed bodies/signatures return an explicit refusal; they do
not silently use the dynamic frame path. See the bounded
[typed-host evidence](evidence/programs/ideal-2026-07/core/CORE-7/typed-host-entries.md).
Run-7 [callback evidence](evidence/programs/ideal-2026-07/core/CORE-7/typed-host-callbacks.md)
also covers generic arithmetic default-member resolution while preserving the
original object argument. Spreadsheet names and conversions remain host-owned.

The run-8 owning typed frame retains portable host objects in module/Static state
and typed object results. It uses canonical session globals and unboxed local
slots, with a saved next-operation position for finite VM3/native-JIT slices.
Completed effects do not replay on resume. Invocation metadata retains a
host-supplied read view; code and project-state leases let an old activation
finish after reload publishes a new signature/state. Reset and unrelated calls
on paused state refuse; bounded typed callback reentry shares the active context.
Object-typed/global inspection preserves project and native object identity.
See [the bounded continuation contract](spec/HOSTING_AND_EARLY_BINDING.md) and
[evidence](evidence/programs/ideal-2026-07/core/CORE-7/typed-host-continuations.md).
Run 9 adds a retained heap call stack, immediate scalar/object/global ByRef aliases
and handler/Resume continuations. JIT body operations remain native; call/error
transitions use common stack mechanics and the canonical ErrEngine. Entire typed
call graphs are checked before body effects. The bounded proof is
[typed-host-calls-run9](evidence/programs/ideal-2026-07/core/CORE-7/typed-host-calls-run9.md).
Run 10 adds retained runtime default-getter dispatch for host Objects and object
payloads in Variant during owning value operations, with the active typed context
and owned intermediate results on both backends. Boxing and object copies preserve
identity. See [the scoped proof](evidence/programs/ideal-2026-07/core/CORE-7/typed-defaults-run10.md).

### Material gaps against the contract

- the public compiler does not require sealed verified image/program input;
- no inspectable procedure-lowering plan currently separates semantic OxIR from physical/codegen decisions;
- the direct native primary family remains bounded to required ByVal Long/Double/borrowed host objects and numeric slots; the additional owning frame supports portable typed object globals/Static/results and instruction-boundary continuation with four arguments/128 physical slots. The run-9 owning family additionally retains bounded internal calls, live ByRef aliases and Err/Resume continuations through shared stack mechanics on both backends. The additive value API now carries exact BSTR/Boolean/Variant and nominal interface results in required portable member slots without a universal Variant frame; typed arrays/records, Optional/ParamArray, cross-project calls, other external calls and native recursion remain open;
- helper registration is not a versioned descriptor catalog;
- line/Erl, writable Err fields and full dynamic Err.Raise metadata are incomplete;
- deep source recursion relies on the native stack and is not safely proven;
- borrowed JIT sessions retain ExecState for repeated calls; general owned verified-package sessions, descriptor lifetimes and product cache remain open;
- general early COM signatures/proxies, pointer/callback operations and native calls outside bounded x64 subsets remain incomplete; scalar dual calls now use the same preserved-descriptor native route on both backends;
- the implementation is concentrated in a very large single source module;
- differential evidence is often status/tag based rather than fully structural.

The destination is [`spec/OXVBA_JIT_ARCHITECTURE_V1.md`](spec/OXVBA_JIT_ARCHITECTURE_V1.md). The older `JIT_V2_*` planning documents are historical design inputs, not current authority.

## 9. Windows COM and native realization

VM3, HAL, COM, host, build and comhost contain substantial Windows work: typelib loading, late/early bridge scaffolding, Declare execution, wrapper packaging, type-library emission and bounded COM serving.

The JIT lowers scalar late COM through the shared verified interop plan and supports dynamic CreateObject ProgIDs and positional GetObject forms. An explicitly owned Excel workbook probe exercises native IDispatch properties, scalar writes and mixed two-dimensional Range.Value2 arrays on both backends. The Windows VARIANT adapter maps standard error numbers to/from the Automation 0x800A0000 prefix while preserving other HRESULTs. Native development tests also exercise a test-owned x64 DLL, including Function/Sub, ByRef mutation, LastDllError and unload. General Windows rows remain in progress: ComCallEarly now carries compiler-owned IID/slot/signature metadata through a strict native-vtable HAL entry on both backends, while general signatures/proxies, pointer operations, callbacks and broader carrier/session families remain unfinished. Its typed route never retries through IDispatch. Native ByRef cells preserve repeated scalar aliases and private argument temporaries; IErrorInfo crosses HAL as separate VBA Err fields. See the [hosting stock-take](spec/HOSTING_AND_EARLY_BINDING.md) and [bounded evidence](evidence/programs/ideal-2026-07/windows-x64/WIN-4/early_scalar.md). Project `WithEvents`/`RaiseEvent` proves internal event semantics, not native connection points.

Native scalar calls carry canonical caller-storage IDs in `RuntimeCallArgument`.
VM3 resolves frame/global aliases; JIT derives call-local IDs from resolved slots.
The shared HAL reuses one native scalar cell for repeated aliases and copies its
final value back to each argument. Equal-valued distinct variables stay separate.
Currency uses scaled i64 ABI storage, Date uses f64, and Byte keeps its runtime tag.
Win64 LastError is captured before libffi cleanup, and entry-fault snapshots retain
it. General field/element aliases, noncanonical foreign Boolean payloads and
initializer-failure error-state capture remain open under the broader owners.

### Material gaps against the contract

- authoritative registry/file typelib selection is not yet one stable metadata service for all consumers;
- scalar early COM consumes one shared descriptor-driven COM marshalling plan; a total sealed backend-neutral plan covering every Windows transport remains open;
- scalar late COM executes through IDispatch and the bounded early dual subset uses native vtables; broader late/early signatures and pure-dispinterface static binding remain incomplete;
- synchronous COM-event ByRef writeback is absent;
- JIT-backed COM serving/vtable generation is absent;
- exact nominal interface arrays and broad VT_RECORD shapes are incomplete;
- bounded scalar JIT Declare works; pointer helpers and AddressOf callbacks remain unfinished;
- x64 artifact and 64-bit Office certification is incomplete;
- JIT wrapper sessions and genuine native DLL/EXE exports do not exist.

The destination is [`spec/OXVBA_WINDOWS_INTEROP_ARCHITECTURE_V1.md`](spec/OXVBA_WINDOWS_INTEROP_ARCHITECTURE_V1.md).

Run 10's bounded native owning member path adds Variant, aliased output cells,
nominal interface results and intrinsic SAFEARRAY transport using retained metadata.
Dictionary Add/Exists/Item and registration-free output/error/cleanup cases run on
both backends. Native Application interface results have owned Excel development
evidence. The run-12 owner ruling admits one compiler-prebound typed Invoke for
pure dispinterfaces. The owning path carries the exact retained IID/DISPID,
accessor, arguments/defaults and nominal results through both backends; native
table indices remain distinct from vtable slots. Range default/Value/Value2
scalars/arrays and WorksheetFunction calls have controlled transport and owned
Excel/VBA development proof, with no runtime name/type-info lookup or dual retry.
See [run-12 evidence](evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-prebound-dispatch-run12.md).
See [the scoped evidence](evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-native-values-run10.md).

## 10. Host sessions, builds and outputs

The owning typed family now executes `Is` on VM3/native JIT. Native interface and
Variant wrappers retain canonical IUnknown identity before publication; comparison
calls no default getter or COM transport. An owned Excel probe checks Nothing and
non-object errors. See [the scoped proof](evidence/programs/ideal-2026-07/core/CORE-7/typed-identity-run11.md).

The neutral linked provider now accepts explicit procedural source closures via
`ProjectGraph`/`load_graph`, retaining separate OxPrograms, canonical stores and
prepared import indices on VM3/native JIT. Owning Optional calls support declared
defaults, private omitted ByRef cells and Missing/IsMissing; source-level names
are resolved before execution. Graph-wide cache identity and suspended code/store
leases survive replacement. This bounded source API does not close verified
source-free sessions, classes/records, ParamArray or the broader CORE-7/CORE-8 work.
See [the project-call proof](evidence/programs/ideal-2026-07/core/CORE-7/typed-project-calls-run10.md).

`oxvba-host` compiles source/project closures for VM3 or one-shot JIT execution. CLI `build --target OxImage` emits a source-free project closure; `run-image` executes it on either backend and honors its entry index. Program-style top-level statements are exercised through the packaged JIT path. `ProjectRuntimeSession` remains VM3-backed. `oxvba-build` also emits wrapper/COM metadata; `oxvba-comhost` loads those packages through VM3.

Current wrapped output infrastructure is valuable, but output labels must remain exact:

- `.oxi` is the current typed serialized artifact;
- WrapperExe/WrapperLibrary are runtime-backed wrappers where implemented;
- WrappedComServer is a reusable runtime-backed COM host;
- native DLL/EXE program outputs do not yet exist.

The target adds a backend-neutral verified session, JIT cache, JIT-backed wrappers and distinct genuine native outputs without loader-lock work.

## 11. Language-service realization

There is no active clean-stack language-service or LSP implementation. The former crates were removed from the workspace and later deleted. The VS Code extension and several older documents still describe the deleted surface and are deprecated by this architecture sweep.

Reusable foundations exist: lossless CST, symbols/scopes/signatures/declaration spans, providers, project closure loading, diagnostics, Core IR facts and historical tests/designs. Missing are the compiler `AnalysisResultV1` fact stream and fixed UTF-8 span boundary, semantic snapshots, overlays, indices, invalidation, direct query API, LSP transport and editor integration.

The destination is [`spec/OXVBA_LANGUAGE_SERVICE_ARCHITECTURE_V1.md`](spec/OXVBA_LANGUAGE_SERVICE_ARCHITECTURE_V1.md).

## 12. Debugging, forms, portability and security

No active debugger or forms-runtime crate exists. Older debugger/direct-host documents are design history rather than current capability. The system contract retains semantic debugger, forms runtime/designer and security as explicit extended profiles so their absence cannot be confused with either completion or permanent exclusion.

VM3 and much of the compiler/runtime are designed for portable hosts. Browser/WASM and desktop-shell documents describe earlier integration programs, but those targets were not certified in the 2026-07-10 review. Portable COM-shaped objects are not native Windows COM evidence.

Security currently appears mainly as host policy, artifact checks and unsafe-boundary discipline. A broader runtime security profile remains future work and must build on the same verified image and host capability model.

## 13. Capability status and accepted delivery

| profile | current status | accepted workset |
|---|---|---|
| Core VBA toolchain | broad, in-progress | [`worksets/WORKSET_2026-07-10_POST_JIT_CORE_CONFORMANCE_AND_READINESS.md`](worksets/WORKSET_2026-07-10_POST_JIT_CORE_CONFORMANCE_AND_READINESS.md) |
| Windows VBA compatibility | VM3 substrate plus missing JIT/general parity, in-progress | [`worksets/WORKSET_2026-07-10_JIT_WINDOWS_COM_NATIVE_INTEROP_AND_BINARY_EXPORT.md`](worksets/WORKSET_2026-07-10_JIT_WINDOWS_COM_NATIVE_INTEROP_AND_BINARY_EXPORT.md) |
| IDE foundation | not implemented on clean stack | [`worksets/WORKSET_2026-07-10_LANGUAGE_SERVICES_CLEAN_STACK_BASELINE.md`](worksets/WORKSET_2026-07-10_LANGUAGE_SERVICES_CLEAN_STACK_BASELINE.md) |
| Standalone tooling | `.oxi` and bounded wrappers only, in-progress | core and Windows worksets |
| Extended profiles | not assessed or not implemented as a unified profile | system contract plus future accepted worksets |

The dated review and these three accepted worksets are the current umbrella-program entry under `bd-59co`. Older ladders, worksets and handoffs remain provenance unless PROGRAM-0 explicitly consumes their residuals.

## 14. Documentation authority

The current hierarchy is:

1. `CHARTER.md` â€” mission and scope;
2. `OPERATIONS.md` â€” execution and evidence doctrine;
3. `docs/spec/OXVBA_SYSTEM_CONTRACT_V1.md` â€” destination architecture and capability clauses;
4. this document â€” current realization and gaps;
5. current subsystem specifications;
6. accepted active worksets and canonical validation/evidence artifacts.

Superseded designs and guidance are classified in [`spec/DEPRECATION_LEDGER_2026-07-10.md`](spec/DEPRECATION_LEDGER_2026-07-10.md). A historical document remains useful for provenance but cannot override this hierarchy.


Run 11 owning source aggregates retain physical record/SAFEARRAY owners and live
ParamArray element aliases. Bounded source classes use canonical runtime object
fields/identity and original typed procedures, with session/reset-scoped deferred
termination queues. Static source member targets are retained during preparation;
unknown dynamic receivers stay refused. Both backends and owned Excel agree on
the [bounded cases](evidence/programs/ideal-2026-07/core/CORE-7/typed-aggregates-classes-run11.md).
Broader classes, native records/callbacks and source-free publication remain open.


Prepared native COM calls now lend the active typed context to synchronous host
callbacks. A same-thread scope permits one callback borrow and restores nested
state on return/unwinding. Bound handlers execute the selected original VM3/JIT
functions, share canonical globals/budget and inherit invocation metadata.
Real native function callbacks and scalar connection-point/IDispatch sinks have
[development evidence](evidence/programs/ideal-2026-07/core/CORE-7/typed-native-callbacks-run11.md).
General WithEvents, incoming ByRef cancellation and cross-apartment policy remain open.

Exact native ByRef Long cells now synchronize with retained caller lvalues around
synchronous callbacks, including error/unwind return and nested native calls.
Native slots and prebound Invoke share this boundary; repeated aliases reuse
their original native pointer. The retained global counterexample produces 13:13
on VM3/native JIT and owned Excel VBA. See [run 12 evidence](evidence/programs/ideal-2026-07/core/CORE-7/typed-native-long-callbacks-run12.md).
Other carriers and cleanup/event reentry remain broader open work.


Run 13 adds explicit retained native invocation locales and QI'd nominal views
for typed assignments from generic COM results. The installed typelib loader now
reads DWORD HREFTYPE values through the correct union member. Owned Excel/VBA and
both owning backends agree on the bounded typed Range geometry/formula/display
and application option/calculation/Evaluate helpers. NumberFormat currency
classification is explained by setter locale, with 1033 explicitly selected by
the Excel host proof. Existing default invocation locale policy is unchanged;
automatic locale selection, genuine Variant-result member chains and collection
iteration remain open. See [run-13 evidence](evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-excel-surface-run13.md).
