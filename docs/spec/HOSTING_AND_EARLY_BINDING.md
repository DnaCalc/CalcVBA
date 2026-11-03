# Hosting and early binding

2026-11-02. Owner direction; current code and bounded evidence, not full Windows parity.
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
dynamic and general COM-event same-session reentry is refused. The typed subsets below permit
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
existing TypeLibResolver. `RuntimeProvider::info` declares API ABI 2, typed entry
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
call. The engine stores no borrowed invocation pointer. The run-8 owning frame
below extends this subset to persistent typed object storage and object results.
No object-count limit of 4,096 is introduced.

Shared physical admission consumes the original OxFunc and retained descriptor,
proves Double-producing host temporaries, and refuses conflicting writers or
unsupported shapes. Both backends check IID, table bounds and native slot
signature. JIT emits call_indirect with native arguments and typed output storage;
VM3 uses the same native signature with unboxed local slots. No Variant argument
vector/frame, runtime member lookup, Invoke retry or libffi occurs on this path.
The API's uniform slot adapter calls that same primary JIT entry directly.

## Owning typed invocations and continuation contract

`Session::begin_invocation` accepts a typed bound handle, borrowed input slots,
an `InvocationFrame` and a total block budget. It copies scalars and retains
object arguments before returning an owning, same-thread `Invocation`. Its frame
keeps unboxed local/temporary slots, the exact next physical operation, the host's
invocation id/parent/read-view object, and leases on the original code and project
state. `active_invocation` exposes the retained metadata to synchronous host
callbacks only while that frame is executing; nested typed calls inherit it.
The host supplies an immutable snapshot when its read policy requires one. The
engine neither interprets spreadsheet reads nor converts a cached read into a
fresh demand. P01/P03/P10 measured by DNA Calc on 2026-11-01 inform this boundary.

Required interface-typed arguments check their retained nominal IID before body
entry, including unused parameters. An incompatible portable table returns VBA
error 13 without entering the function. An untyped `Object` parameter accepts any
portable table. The slot API represents a null reference with `Slot::Nothing`;
binding Nothing is valid, and dereferencing it raises error 91. The original
native primary entry checks the same identities before executing its first block.

`resume(operations)` runs at most that many physical operations, pausing only
between completed instructions or control transfers. Zero performs no work.
`Suspended` retains the next operation and owning locals. Host callbacks finish
before control returns; a callback is never resumed in its middle. Both VM3 and
native JIT execute the original OxFunc using the same checked physical admission.
The JIT emits a native switch over saved positions and actual native arithmetic,
loads, stores and typed indirect callbacks; it does not use the interpreter.
This frame is an execution layout, not a second language IR or a Variant frame.

Module and compiler-lowered Static storage use the session's existing canonical
global cells, shared with dynamic calls. Object stores AddRef before releasing
the old value. Local slots, temporary global receivers and object results have
explicit owners; a receiver survives a callback that replaces its last global
reference. The retained portable table is checked before projecting it from the
canonical IUnknown carrier. An arbitrary native COM object is never reinterpreted
as this portable table. Object results use `OwnedValue::Object`; the raw scalar
slot adapter refuses them so a borrowed pointer cannot escape a completed frame.

A suspended activation excludes reset, unrelated calls and a second activation
on that same state. A synchronous callback may reenter a supported typed function
with its active context, sharing budget and the existing depth limit. Reload
prepares and publishes fresh code/state without touching the paused activation's
leases. Old bound handles refuse on the new generation, while the old activation
can finish against its original globals and signature, even after the current
session is dropped. Failed preparation preserves both. Dropping an invocation
cancels it and releases locals/leases; hard budget exhaustion and runtime errors
are terminal. A returned or faulted invocation refuses another resume. A distinct
semantic invocation starts afresh and may repeat effects, as Excel re-entry does.

The run-9 owning family adds internal Function/Property Get/Sub calls, required
ByRef parameters, live repeated aliases and handler/Resume continuations to the
Long/Double/portable-object family. It retains at most 64 boxed VBA frames (error
28 at the bound); each frame keeps the existing four-argument/128-slot bound.
Both backends return call/error transitions to the same retained stack mechanics;
JIT body operations and callbacks remain generated native code. Admission checks
the entire reachable call graph before any body effect. The additive value API supplies Void for public typed Sub entry.
Root ByRef inputs are private host copies. Run 10 extends this family with Optional
parameters and procedural source-project calls, as described below. Arrays/records,
ParamArray and suspension within native/COM callbacks remain
outside this family under `bd-59co.2.9.9`/`bd-59co.2.10.7`.
A synchronous callback finishes before suspension. Source-free verified packages
and DNA Calc glue remain open. See the
[run-9 call-stack proof](../evidence/programs/ideal-2026-07/core/CORE-7/typed-host-calls-run9.md).
The unchanged native primary ABI remains 2; owning-frame layout is private to the
linked provider. API version 1 refuses before compilation. See the
[run-8 proof and limits](../evidence/programs/ideal-2026-07/core/CORE-7/typed-host-continuations.md).

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
additionally checks the active execution borrow. Owning Object comparisons/conditions
use retained default-member plans; scalar entries refuse unsupported object contexts.

Generic numeric arithmetic now resolves each object's default member through the
existing project/foreign dispatch before entering the pure arithmetic kernel.
The original object remains in its argument/local; no spreadsheet-aware entry
conversion occurs. Default chains are bounded at 64; broader context-specific
coercion and exact deep-chain behavior remain in CORE-3. See the
[run-7 proof and limits](../evidence/programs/ideal-2026-07/core/CORE-7/typed-host-callbacks.md).

## Additional owning value capability

The original scalar API remains unchanged. Providers may advertise
`value_abi() == Some(TYPED_VALUE_ABI)` (currently 1). The additive
`load_with_object_libraries` accepts `TypedObjectLibrary`/`TypedObjectInterface`/
`TypedObjectMember` metadata; existing Project and scalar descriptor types retain
all their fields and variants. Use `bind_values` and `begin_values` for retained
invocations, or `call_values` for synchronous calls sharing an active TypedContext.
Default provider/session methods refuse this capability rather than guessing.

`ValueType`, borrowed `ValueSlot` and owned `Value` carry Long, Double, Boolean,
exact BSTR, actual Variant and portable HostObject. `Void` describes a Sub result;
`Interface(IID)` describes nominal metadata over the object carrier. A declared
Variant is stored inline only where required, with no universal Variant frame or
heap-boxed callback result. HostSlot::typed supplies checked required positional
slots; return String/Object/Variant ownership transfers into the retained frame.
Nothing, Null, CVErr and invocation faults remain distinct. Numeric legacy slots
keep their original ABI. No native COM pointer is a portable ObjectHeader.

The fixture covers Value/Value2, Cells, Address and Rows/Columns.Count on VM3/native
JIT. DNA Calc supplies their actual meanings and its Application/WorksheetFunction
library; none of those names is an engine intrinsic. The current four-argument,
128-slot and 64-frame bounds remain. Full typed array/record/ParamArray and native
callback suspension remain open. See the
[typed value proof](../evidence/programs/ideal-2026-07/core/CORE-7/typed-value-slots-run9.md).

Run 10 adds runtime object-default coercion in owning value kernels on VM3/native
JIT, including Object payloads retained in Variant. Explicitly supplied host
libraries retain IID/default-getter plans even without a nominal source declaration.
Arithmetic, numeric/string coercion, concatenation, comparison and conditions use
the active typed context and owned results; boxing/Set/Variant copies retain identity.
Default chains are bounded at 64; broader coercion/native signature coverage remains
under bd-59co.2.9.9 and WIN-4. Static lowered default calls are separate. See
[default coercion evidence](../evidence/programs/ideal-2026-07/core/CORE-7/typed-defaults-run10.md).

## Optional and source-project calls

Run 11 adds owning `Is` expressions on both typed backends. Native tables retain
canonical `QI(IID_IUnknown)` identity before publication, including nominal and
Variant result wrappers. `Is` compares that retained identity without defaults,
QI, Invoke or member lookup. Nothing and non-object Variant behavior is checked
against owned Excel VBA. Foreign runtime IUnknown carriers use retained pointer
identity; native interfaces require native preparation. See
[identity evidence](../evidence/programs/ideal-2026-07/core/CORE-7/typed-identity-run11.md).

Run 10 admits Optional parameters in the owning value family on VM3/native JIT.
`bind_values`/`bind_slots` accept a supplied argument prefix when every remaining
parameter is optional. The provider prepares the declared defaults once per
invocation; source-level named arguments and interior omissions are bound by the
compiler. Omitted Variant parameters without defaults carry DISP_E_PARAMNOTFOUND;
`IsMissing` reads that carrier without object-default coercion. Omitted ByRef
parameters have private writable cells. Actual ByRef lvalues retain live aliases;
parenthesized arguments remain private copies. The direct numeric entry ABI still
uses its existing required-parameter signature family.

`RuntimeProvider::load_graph` accepts an explicit `ProjectGraph` of `ProjectNode`s.
Each node contains its existing `Project` descriptor and ordered `ProjectReference`
entries (`Source` or `TypeLibrary`). A nonempty node reference list requires an
empty `Project.references`; otherwise that older type-library list is used.
Supply typed object libraries separately to `load_graph`; graph nodes refuse
legacy `Project.libraries`. These shared host libraries follow explicit references.
The compiler builds the reachable closure leaf-first, deduplicates diamonds and
rejects missing/cyclic/duplicate identities before publishing a session. Ordinary
visibility and Option Private Module rules apply across project boundaries.

Each unit retains its own original OxProgram and canonical globals. Source import
tokens resolve to unit/procedure indices during preparation; calls execute the
same retained stack and selected backend with no runtime name lookup or dynamic
fallback. Preflight checks all reachable callees before any body effect. Code-cache
identity includes every linked program. `reload_graph` prepares a replacement
closure before publication; suspended invocations retain all old code/stores.
`reload(project)` replaces the root descriptor while preserving graph references.
Reset reinitializes every unit and refuses while an invocation is suspended.

The current four-parameter/128-slot/64-frame bounds apply. Source-free graph
publication, broader class/record/ParamArray behavior and general native callback suspension
remain open under bd-59co.2.9.9/bd-59co.2.10.7. See
[the run-10 project-call evidence](../evidence/programs/ideal-2026-07/core/CORE-7/typed-project-calls-run10.md).

## Imported native members on typed slots

A Windows host can pass the compiler-retained selected member descriptors and its
bound native ObjectRef to unsafe `ComHal::prepare_typed_object`. Accepted descriptors
must come unchanged from trusted compiler-retained type information; arbitrary
in-bounds slot/signature combinations cannot be verified by header checks. HAL checks policy; COM
validates source type/convention/IID/slot bound, creating thread, QI and proxy
alignment, then retains the receiver in a same-thread HostObject table. The owning
value API invokes that table on either backend. Native pointers remain separate
from portable host headers. This is an additive HAL composition facility; no
server implementation enters calcvba-host-api.

The original native family is HRESULT scalar/BSTR getters, void methods and one-Long
methods/setters with scalar/BSTR/void results, optionally with trailing hidden LCID=0.
Run 10 adds descriptor-planned Variant inputs/results, repeated ByRef outputs,
nominal interface results and intrinsic SAFEARRAY arguments/results. The additive
unsafe `prepare_typed_object_with_interfaces` retains the compiler's interface table
and resolves result identities before publication. Native out-cell aliases and
mutations survive HRESULT faults before handler continuation. Only source arguments
occupy the bounded typed frame; omitted native optionals/LCIDs are prepared at the
COM boundary, including long optional tails. Records, custom IUnknown-only results,
nominal interface arrays and callback reentry remain open. See
[run-10 native value evidence](../evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-native-values-run10.md).
Rich synchronous
callbacks may publish TypedFault with their returned number; the owning stack
consumes exact Description/Source/HelpFile/HelpContext before the next operation.
Native IErrorInfo is captured immediately after failure, before receiver cleanup.
Scripting.Dictionary and supervised real Excel development cases are recorded in
the [native typed-slot proof](../evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-native-slots-run9.md).

The installed Excel Range and WorksheetFunction interfaces are pure dispinterfaces
with seven IDispatch slots and no custom slot for Value or Sum. The owner's run-12
ruling permits one compiler-prebound Invoke for these interfaces. The owning
planned path retains IID, DISPID, accessor, parameter/default/wire metadata and
nominal result plans. The table uses the complete compiler interface member order
for its internal indices, separately from native vtable slots. Preparation through
prepare_typed_object_with_interfaces requires that unchanged complete table.
Calls reverse supplied positional arguments, name the property-put value with
DISPID_PROPERTYPUT and omit unsupplied optional tails. Declared Variant arguments
box their actual carriers at this boundary, preserving object identity and avoiding
premature default evaluation. Native cells own strings/arrays/interfaces; repeated
ByRef wrappers share the same cell. EXCEPINFO, deferred fill and immediate
IErrorInfo capture preserve full Err details. No runtime name/type-info lookup,
guessed slot or dual-interface fallback occurs. Application's WorksheetFunction
getter remains a native vtable call. Both typed runtimes have controlled transport
proof and owned Excel/VBA Range/default/Value/Value2 scalar/array and function
observations. See [run-12 evidence](../evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-prebound-dispatch-run12.md).

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


## Run 11 source aggregates and classes

The owning value API adds Record and Array carriers (descriptor-backed VbaRecord
and canonical SAFEARRAY), plus RuntimeObject for a retained source-class IUnknown.
These carriers execute through the selected VM3/native JIT backend. ParamArray
creates a zero-based Variant packet; private/repeated lvalue aliases stay live,
while packet copies detach. Host aggregate schemas are checked before body effects.

The linked Provider explicitly accepts class module names with with_class_modules
for a single source project. Bounded New/initialize, statically declared methods,
primitive/object instance fields, Is, object results and termination execute typed
functions. Class objects and the dynamic adapter share canonical runtime storage.
Session/reset ownership isolates termination queues and rejects stale typed class
inputs. Source-class object results require begin_values, not the older result API.
Cancellation releases frames and permits reset; it does not promise completion of
a cancelled VBA finalizer. Richer class/array/native ABI behavior remains open.
See [the bounded proof](../evidence/programs/ideal-2026-07/core/CORE-7/typed-aggregates-classes-run11.md).


## Run 11 synchronous native callbacks

Prepared scalar and descriptor-planned native COM calls establish an active
with_native_call scope. A host callback or native event sink can use
with_native_callback_context and call_typed_in_context/call_values to execute a
bound VBA procedure against that same budget. The scope is same-thread, rejects
overlapping borrows and restores on failure/unwinding. Session generation and
reentry checks remain in force. A callback completes before the native call
returns; suspension inside the foreign call remains unsupported.

Hosts own event subscriptions, target lifetime and native HRESULT/error mapping.
A native connection-point/IDispatch scalar sink now has dual-runtime development
proof, including stale-handler rejection and Unadvise. VBA WithEvents, incoming
ByRef cancellation and broader native event policy remain with bd-59co.3.6.4.
See [native callback evidence](../evidence/programs/ideal-2026-07/core/CORE-7/typed-native-callbacks-run11.md).


Run 12 synchronizes exact native ByRef Long cells with their retained caller
lvalues before synchronous callback entry and refreshes those same native cells
after callback return or unwind. This applies to native slots and prebound Invoke;
repeated aliases retain their original pointer and nested calls use their own
scope. Private arguments remain private. Both typed backends now observe 13:13
for the retained global-alias regression, matching owned Excel VBA. See the
[delivery evidence](../evidence/programs/ideal-2026-07/core/CORE-7/typed-native-long-callbacks-run12.md).
General native ByRef/reentry, other carriers, post-return cleanup callbacks,
incoming event cancellation and certification remain open with the broader owners.


## Run 13 native roots and invocation locale

Typed native views can resolve an explicit destination IID through their retained
library without changing the source object table or searching member names. The
owning store retains that QI'd view. Generic IDispatch results retain dynamic
services; an explicit nominal assignment prepares a separate exact interface view.
TYPEDESC now reads its 32-bit HREFTYPE union member
without consulting unused pointer-sized upper bytes.

prepare_typed_object_with_locale retains the host-selected invocation LCID through
returned views, prebound Invoke and native hidden LCID parameters. Existing default
preparation keeps its prior locale policy. The owned Excel proof selects 1033 to
match VBA's format interpretation; user-default $0.00 writes on this machine mean
a literal dollar and legitimately yield Double on Value. Cross-locale reads retain
the setter provenance. This is an explicit host configuration, not an engine
member-name or format-string rule. Automatic locale selection and broader locale
parity remain open.

Owning VM3/native JIT and VBE-compiled VBA agree on the bounded ordinary range,
formula/text, application option/calculation/Evaluate source. VarType/CStr imports
share the exact value library and active default-call context. Live Range,
Worksheets and Workbooks enumeration uses the retained -4 descriptor and
IEnumVARIANT without runtime name lookup or pre-materializing the collection.
Exhausted object loop variables become Nothing; Exit For retains the current
item. An uninitialized Range source raises Excel's observed error 424. The compiler
retains complete interface metadata for typed property/method results as well as
declared types. Genuine Variant-result member chains remain open. See
[run-13 evidence](../evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-excel-surface-run13.md).

## Run 13 synchronous incoming events

Host-registered typed handlers can subscribe from a retained native object using
its exact imported source-interface descriptor. The synchronous sink owns object
arguments, permits one incoming Boolean ByRef, and copies that Boolean before
Invoke returns. A same-thread check and retained Invoke reference cover callback
reentry and subscription removal. No member or event name is searched at delivery.

The additive call_values_with_writebacks API returns primitive root ByRef changes
on success and VBA errors; structural root aliases are refused before effects.
Owned Excel Worksheet_Change EnableEvents toggling/nested writes and canceled
Workbook_BeforeClose agree with both owning runtimes. This is host registration,
not automatic document/WithEvents wiring. Full handler-error policy, structural
ByRef arguments, apartments and pinned certification remain open. See
[event evidence](../evidence/programs/ideal-2026-07/windows-x64/WIN-5/typed-excel-events-run13.md).

Native synchronous callback synchronization now covers declared ByRef Long, Double
and canonical Automation Boolean cells. Each repeated alias projects one native
authority, publishes its actual carrier to the owning caller before reentry and
refreshes that authority from the caller after the callback. Other carriers and
direct native-pointer identity with physical caller storage remain open.

## Run 14 Object and Variant chains

Owning VM3 and native JIT now execute genuine untyped ComCallLate operations
through retained object-library services. Native services resolve member/argument
names with IDispatch and invoke the selected member once; default calls use
DISPID_VALUE. Returned dispatches retain identity, native ownership, locale and
those services, so arbitrary-length chains do not require nominal casts. Known
interface calls continue to use compiler-retained plans without runtime recovery.
Repeated dynamic ByRef lvalues share one native VARIANT cell; named arguments,
omissions, error copyback and full EXCEPINFO use the common Automation boundary.
The source-class call path remains separate and uses its original owning stack.

Owned VBE-compiled Excel and both owning runtimes agree on Range/Offset/Cells
Object and Variant chains, named Offset, implicit defaults, formulas, returned
strings, all seven touched cell values, and errors 91/424/438. Portable owning
regressions also check aliases, omissions and error copyback. Dynamic native
ByRef callback synchronization, default-member named arguments and broader
signature/error/lifecycle certification remain open. This is bounded development
evidence under bd-59co.3.5.12; see
[chain evidence](../evidence/programs/ideal-2026-07/windows-x64/WIN-4/object-chains-run14.md).
