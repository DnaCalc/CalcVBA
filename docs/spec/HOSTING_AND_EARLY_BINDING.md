# Hosting and early binding

2026-10-31. Owner direction; current code and bounded evidence, not full Windows parity.
Authority: `CHARTER.md`, `OXVBA_SYSTEM_CONTRACT_V1.md` HOST-SESSION-001,
COMP-BIND-001, COM-CLIENT-001 and JIT-CORE-001. Names remain OxVba; master only.

## Engine boundary

CalcVBA is an independent VBA engine. Worksheet functions, calculation dependencies,
volatility, cell errors, Range, Application and WorksheetFunction belong to DNA Calc's
glue/object library. DNA Calc publishes that library as COM type information and supplies
its objects to the engine. The engine must not contain an Excel-shaped substitute.
Windows targets full VBA/COM compatibility; Linux excludes general COM references;
web hosts expose a declared subset with no direct filesystem I/O.

## Current stock-take

| Capability | Interpreter (VM3) | JIT |
|---|---|---|
| Typelib import | Shared symbol/provider loader resolves registered/file-backed libraries, coclasses, dual partner interfaces, IIDs, member kinds, slots, bounds, parameter/wire types and defaults. Selection/provenance completeness remains open. | Same compiler metadata reaches OxIR. No separate JIT importer. |
| Native early calls before this change | `ComCallEarly` retained metadata in OxIR but passed only token/name through HAL; runtime rediscovered metadata and PreferVtable could fall back to Invoke. | Same token/name reduction through the project-member helper; no preserved-descriptor native contract. |
| Native early calls in this change | Descriptor reaches `ComHal::invoke_early_bound`; native x64 dual calls use declared IID/slot/signature, QI and shared vtable marshalling. No implicit Invoke retry. | Compiled call embeds an immutable member descriptor retained by CompiledImage and uses that same HAL/marshaller contract. |
| Dual/vtable signatures | Existing marshaller covers Automation scalars, strings, Variant, several object/array/record and ByRef shapes. New end-to-end proof is bounded to in-process scalar/string/Variant methods/properties, private ByRef temporaries and errors. | Same native marshaller and bounded proof; broad signature/proxy parity is still open. |
| Implements | Project-class interface routing and descriptors exist; bounded VM3-backed COM serving emits dual vtables. Imported-interface completeness and general serving identity remain open. | Internal project interface/class routing exists. General JIT-backed COM serving/generated export vtables are missing. |
| Events | Internal WithEvents/RaiseEvent plus native connection-point metadata/subscription/queued delivery. General synchronous native ByRef event writeback and safe same-VM reentry remain open. | Internal events execute; native event delivery/reentry/serving are not parity-complete. |
| Host objects | Native retained IDispatch roots can be supplied through ComHal; the host owns its library and activation. Existing HostInjected providers also exist. | Same host boundary; typed host calls use the preserved descriptor. |

Code anchors: `oxvba-com/src/windows_typelib_loader.rs`, `typelib.rs`,
`windows_invoke.rs`, `windows_vtable.rs`; `oxvba-symbol/src/provider.rs`;
`oxvba-bind`; `oxvba-oxir/src/com.rs`; VM3 `dispatch_early_com_method`;
JIT `emit_project_member_call_to_slot` and `invoke_foreign_member_to_slot_for_jit`.
Detailed evidence: [early_scalar.md](../evidence/programs/ideal-2026-07/windows-x64/WIN-4/early_scalar.md).

Object/Variant receivers and CallByName remain genuinely dynamic. A pure dispinterface
uses an explicit compiler-bound IID/DISPID/accessor Invoke plan, with no runtime name or
type-info lookup; it has no custom slots. This is separately tested on both backends.
Dual members never retry through that path. Missing/unsafe dual signatures fail before
entry. Custom IUnknown-only interfaces, general inherited/proxy transport, ByRef
aliasing/error-edge and named-argument breadth, nominal interface arrays, callbacks and JIT-backed COM exports remain open.
Portable registered objects can explicitly implement `PortableDispatch::invoke_early_bound`;
the default refuses typed calls. Legacy synthetic token arithmetic is not an early-binding target.
Unregistered headless projections still synthesize object/member numbers for dynamic
calls; replace that legacy path with explicit host services under `bd-59co.2.10.7`.
General JIT bodies still use Variant frames/helpers. A run-6 scalar typed primary
family now bypasses those frames on both runtimes; its safe Rust host API checks
owner/generation before entry. A general native host ABI remains open. Native COM
transport alone does not establish worksheet-to-JIT cost or performance parity.

## What public Office contracts establish

[MS-VBAL projects](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/4cd406c7-1ade-4522-8d7b-933183059ac7)
defines source/library/host projects, ordered references and host-introduced public entities.
[MS-OVBA](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc)
defines project storage, references and modules; it is not an executable hosting ABI.
[Automation dual interfaces](https://learn.microsoft.com/en-us/windows/win32/com/dual-interfaces)
permit either direct vtable calls or Invoke. [TYPEKIND](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ne-oaidl-typekind)
distinguishes static module functions/data from interfaces and coclasses.
[ITypeInfo::AddressOfMember](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-itypeinfo-addressofmember)
provides a COM type-information pattern for static callable addresses whose lifetime
is tied to the type description; it does not require inventing an instance receiver.
These contracts do not prove that Excel wraps each standard-module procedure in a COM
object, or reveal Excel's internal registration/call path. We must not invent that path.

Excel owns workbook/macro discovery, formula lookup, dependencies and argument/result
adaptation around VBA. Its documented [Application.Run](https://learn.microsoft.com/en-us/office/vba/api/excel.application.run)
accepts positional Variant arguments and returns the procedure result; it is a dynamic
host facility, not a required fast engine entry or a compile check.
The public [xlUDF C API](https://learn.microsoft.com/en-us/office/client-developer/excel/xludf)
also invokes VBA functions through Excel's function-reference/name adapter. Neither
API documents Excel's internal call ABI or proves per-procedure COM wrapping.
[Option Private Module](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-private-statement)
restricts cross-project references while retaining within-project visibility.

Black-box observations on this machine use source injection into an owned workbook,
VBE Debug Compile command 578 before execution, and an armed PID-scoped UIA guardian.
No Office binary inspection, disassembly, security/Trust Center or registry changes.
The retained [probe](../evidence/programs/ideal-2026-07/windows-x64/WIN-4/hosting-public.bas)
observes same-workbook cell calls to Public and Private Functions; a Sub yields cell `0`
and Run yields Empty. Long parameters round `2.5` to `2`, accept text `"3"`, and reject
`"bad"` as error 2015 (#VALUE!). CVErr(2042) remains error 2042. A referenced cell arrives
as Range (TypeName), while VarType reflects its default numeric value (5); a literal is Double.
An explicit unhandled Err.Raise produced a captured runtime dialog and aborted the
instrumented run, so ordinary worksheet exception mapping is not established by that case.
The Option Private probe retains those within-workbook results. Public/Private Functions
also return 2/73 through explicitly workbook-qualified formulas in another workbook;
that host lookup is distinct from VBA project-reference visibility. Broader coercions,
class-module entry and ordinary unhandled-error mapping remain open observations.
The formulas were calculated while RunProbe was active; idle recalculation and
function-discovery UI visibility were not tested.

## Required host-neutral interface

1. Load code/projects plus ordered references and host-provided type libraries/root objects;
   verify the artifact and report compile diagnostics before publishing any call handles.
2. Enumerate public module procedures with stable project/module/procedure identity, kind
   (Function/Sub/accessor), visibility/Option Private, exact return/parameter types,
   ByVal/ByRef, Optional/default/Missing, ParamArray, array/UDT/interface identity and ABI.
   The host chooses which procedures its UI or formula language exposes.
3. Bind once to an opaque callable owned by a project session. Invoke the selected VM3
   procedure or typed JIT primary entry directly; retain code and state for the handle's
   lifetime. No per-call name search, source wrapper, COM-object fabrication or backend retry.
4. Define a versioned COM-compatible value/ownership/error convention and typed native
   entry signatures. Keep the checked universal Variant thunk as an explicit dynamic adapter.
   Standard-module procedures have no invented receiver, IUnknown identity or connection point.
5. Retain/reset/dispose sessions with globals, Static, Err and objects; define reload
   invalidation, thread/apartment ownership, budgets and reentry. Host errors and CVErr values
   stay distinct. Hosts own scheduling, formula conversion and registration lifetimes.

This mirrors COM's typed metadata, stable callable identity, HRESULT/error and ownership
discipline without pretending module procedures are COM objects. Excel's private glue ABI
is unknown; reproducing undocumented internals is unnecessary.

## Run-6 neutral source-module API

`Module::load_project` consumes the compiler manifest plus the host's typelib resolver;
`with_host` retains explicit HAL services/root ownership. `procedures()` publishes
compiler-owned project/module/name, Function/Sub/accessor kind, visibility, Option
Private Module, exact OxTy/array layout, passing mode, Optional/default and ParamArray.
COM interface IDs refer to the retained compiler interface table. Private procedures
are listed for inspection but cannot bind as public host entries. This catalogue is
separate from the cross-project export list, which excludes Option Private Module.

`bind(module, name, kind)` resolves once. `call_bound` uses its numeric procedure ID
with exact, fully prepared VBA Variant arguments/results and no per-call name search.
It checks argument count and carrier tags. Host ByRef inputs are private copies;
typed native ByRef and complete nominal array validation remain open. Null, CVErr and
unpaired UTF-16 stay exact engine values; no formula conversion occurs in this API.
Host-provided COM objects can be arguments: both runtimes use the preserved IID/slot
plan, with no Invoke/name/type-info recovery in the controlled native host test.

`bind_typed::<(f64, f64), f64>("Math", "Add2")` selects a native primary JIT entry or
the matching VM3 scalar interpreter. The current family accepts zero to four required
ByVal Long/i32 or Double/f64 arguments/results, numeric locals, add/subtract/multiply,
division, comparisons, truthiness and branches. Long operands/literals promote
to Double when a Double operand proves that numeric lane. Physical admission records exact
slots and proven Double-only Variant temporaries; it refuses user Variant locals,
ByRef, Optional, ParamArray, strings, arrays, globals/Static, general calls, handlers,
line/Erl operations and other instructions. It bounds physical slots at 128. Host Double inputs must be finite; non-finite values return error 6. No
unsupported typed call silently falls back to a dynamic backend or Variant frame.

Typed native convention v2 is `(TypedContext*, typed arguments..., R*) -> i32` using
the target C ABI. Context owns a cooperative block budget; 0 means success, -1 host
budget interruption, positive statuses are VBA error numbers. Results are written
only on success. Both universal value routes adapt admitted scalar procedures to the same primary
entries; the dynamic JIT thunk boxes/unboxes around its native entry;
image run-entry diagnostic snapshots keep their original body. Warm typed calls use
no heap storage. This is a bounded Rust-facing family, not the final general C ABI.

Source-module programs have no implicit run entry; hosts select bound procedures.
A session initializes storage before publishing handles, never runs Main, and uses
a separate 1,000,000-block initialization budget. Calls are synchronous/same-thread;
dynamic/stateful same-session reentry is refused. The run-7 typed subset below permits
bounded callback reentry with the active context. Handles retain an opaque owner identity and checked u64 generation,
never a public address. Reload prepares replacement code/state first; failure keeps
old handles valid, success changes generation before old code can be freed. Counter
exhaustion refuses reload. Reset retains code/handles and reinitializes VBA storage.
The shared bounded LRU holds code/metadata while each session owns globals/Static/Err
and HAL; eviction cannot invalidate a live session. The scalar host ABI does not
preempt native calls or prove general recursion/reentry/apartment safety.

## Run-7 host contract and typed object calls

`calcvba-host-api` is the host library dependency. It contains source/project and
ordered reference descriptions, compiler signature projections, exact runtime
Variants for the explicit dynamic adapter, opaque handles and typed-slot contracts.
It links neither compiler nor execution backend nor libffi. Applications compose
`oxvba-embed::provider::Provider`, optionally supplying HAL services and their
existing TypeLibResolver. `RuntimeProvider::info` declares API ABI 1, typed entry
ABI 2 and JIT availability. An incompatible requested API version refuses before
compilation. This is a linked Rust interface: Box/Rc/code/object allocations are
released by their originating owners through Rust drop/vtables. It is not a stable
separately loaded C/plugin ABI.

`Session` lists signatures including accessors/Sub, visibility/Option Private,
scalar/array element/rank metadata, passing modes and defaults. It binds once and
validates the original engine owner/generation on every call. Failed reload keeps
the old module; successful reload invalidates its handles before releasing code.
Hosts can type a bound Property Get with `bind_typed_handle`. Native host objects
still use preserved typelib IID/slots through the existing ComHal boundary; the
API facade does not substitute its portable table for native COM metadata/pointers.

Portable `ObjectLibrary` descriptors name interfaces and fixed slot signatures.
`HostObject` owns immutable callback tables/captures and a runtime IUnknown prefix.
Its typed table has its own C ABI, separate from IDispatch/dual vtables. Current
slots return Double and accept zero arguments or one required ByVal Long. Function
entries additionally accept borrowed ObjectArg alongside the existing Long/Double
arguments. Objects are retained through same-thread Rc and their IUnknown prefix;
ObjectArg borrows cannot outlive the owner. A kept object may outlive an invocation,
but its callback must resolve any sheet/backend context against the active host
call. The engine stores no invocation pointer. Typed object results, persistent
VBA owning object storage and broader signatures remain outside this subset.
No object-count limit of 4,096 is introduced.

Shared physical admission consumes the original OxFunc and retained descriptor,
proves Double-producing host temporaries, and refuses conflicting writers or
unsupported shapes. Both backends check IID, table bounds and native slot
signature. JIT emits call_indirect with native arguments and typed output storage;
VM3 uses the same native signature with unboxed local slots. No Variant argument
vector/frame, runtime member lookup, Invoke retry or libffi occurs on this path.
The API's uniform slot adapter calls that same primary JIT entry directly.

Callbacks may call `call_typed_in_context`/`call_slots` with the borrowed active
context. Nested calls consume the outer budget. A fresh context cannot reenter the
same executing session; dynamic calls and reset also refuse during that activation.
The thread permits at most 64 typed activations across its sessions (error 28).
Guards restore depth/context on success, faults and budget interruption. Host
callback panics become error 5 at the C boundary; failed calls do not read unwritten
results. This proves a declared synchronous typed subset, not arbitrary COM event,
ByRef, general VBA recursion or stateful VM/JIT callback safety.
Reload requires exclusive `&mut Session`; a safe caller cannot acquire it while
the shared session/code borrow spans an activation or callback. Dynamic reset
additionally checks the active execution borrow. Object comparisons/conditions
are refused by typed admission until their value-context semantics are supported;
the object pointer is never treated as a scalar condition or comparison value.

Generic numeric arithmetic now resolves each object's default member through the
existing project/foreign dispatch before entering the pure arithmetic kernel.
The original object remains in its argument/local; no spreadsheet-aware entry
conversion occurs. Default chains are bounded at 64; broader context-specific
coercion and exact deep-chain behavior remain in CORE-3. See the
[run-7 proof and limits](../evidence/programs/ideal-2026-07/core/CORE-7/typed-host-callbacks.md).

## Spreadsheet extraction and DNA Calc glue handoff

All `Value::Range`, SpreadsheetContext, grid limits, synthetic Excel types,
Application/Range/WorksheetFunction callbacks and spreadsheet value/error conversions
have moved out of oxvba-embed to `adapters/spreadsheet-compat`. This explicitly legacy
adapter reuses the neutral compiler/session/cache owners. It preserves the earlier
bounded tests and browser example; it is not a new engine object library. The
DnaCalc ProgID belongs only to that adapter. No DnaCalc repository files were edited.
Existing DnaCalc consumers must pin their prior engine revision until their separately
authorized glue migration uses the new API or the compatibility package.

DNA Calc glue will need:

- its real typelib and host-owned native objects (and declared typed portable subset);
- its procedure exposure/registration policy, names, conflicts and workbook identity;
- bind-once registration records tied to the engine generation and reload lifecycle;
- exact signature-based argument preparation, default/Missing/ParamArray rules and
  the remaining typed carrier families as CORE-7 delivers them;
- Range identity/default-member behavior, caller context, dependency reads, volatility,
  scheduling and worksheet-function callbacks owned by its object library;
- runtime-error to cell-error conversion distinct from returned CVErr values;
- thread/apartment, reentry, budgets, host-object disposal and cache ownership policy;
- consumer migration, C/native ABI review and integration/call-cost measurements.

Idle Excel evidence now establishes unhandled error13 and division-by-zero -> cell
#VALUE! with no runtime dialog, while an untyped single-cell Range uses its default
value in arithmetic. Numeric text and blanks coerce; bad text/multi-cell arithmetic
produce #VALUE!. The run-7 generic object default-value arithmetic subset is delivered
under bd-59co.2.4.9; broader coercion families remain under bd-59co.2.4.3. Glue supplies the real default member without
erasing the object argument at entry. See [worksheet observations](../evidence/programs/ideal-2026-07/windows-x64/WIN-4/worksheet-errors.md).

`UDF_EMBEDDING_V1.md` remains the compatibility adapter's consumer contract. CORE-7
bd-59co.2.9.9 and CORE-8 bd-59co.2.10.7 remain open for the in-progress capabilities: the scalar family does not
complete general typed procedure bodies/carriers, source-free verified sessions,
full cache/reentry ownership or removal of StandardHost synthetic dynamic projections.
Native COM residuals stay with WIN-4. See the [bounded delivery/evidence](../evidence/programs/ideal-2026-07/core/CORE-7/typed-host-entries.md).
