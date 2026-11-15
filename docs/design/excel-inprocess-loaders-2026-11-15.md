# CalcVBA inside Excel: XLL and VBA loaders (plan, 2026-11-15)

## Goal

Run CalcVBA (VM3 and the JIT) inside `EXCEL.EXE`, on Excel's main thread, with the hosting
Excel's object model as the project's host objects. Then:

- CalcVBA code calls Excel in-process: direct vtables, no marshaling proxies, the same
  objects and identities real VBA sees.
- Excel events reach CalcVBA `WithEvents` sinks synchronously, as they reach VBA's.
- The same VBA source runs in real VBA and in CalcVBA in one Excel process, against the same
  workbook, so the two can be compared case by case.

Today every Excel comparison drives Excel out of process. Calls cross the Automation
marshaler, and events arrive through proxies. That checks the call path (step 8 of the VARIANT
layout programme), but not the in-process cases. Those are an object passed between real VBA
and CalcVBA, a sink that outlives the run, a mix of real VBA and CalcVBA sinks on one source,
and release counts without a proxy in between.

## What exists

- `calcvba-host`: the embedding API (`Module::load_project`, `Session`, VM3/JIT backends, a
  host's `HostServices`).
- `calcvba-com::automation`: the one Automation call path over the runtime cell. It covers
  late-bound and early-bound calls, `For Each`, `Is` and WithEvents sinks, with events
  delivered through the thread's executor (`calcvba_runtime::events`).
- `calcvba-host/tests/support/owned_excel.rs` finds an owned Excel's object model from its
  window (`AccessibleObjectFromWindow(OBJID_NATIVEOM)`).
- The `bd-p5bb` work for XLL shims: identify the hosting Excel by `xlGetHwnd` and its window,
  never the ROT.
- `examples/native/xll` and `tools/calcvba-udf`: XLLs compiled from VBA source by the native
  lane. They show the C API conventions; they do not host the interpreter.
- The Excel oracle tooling (`DnaCalc/tools/excel-session/on-private-desktop.ps1`,
  `scripts/run-excel-probe-session.ps1`). It covers owned processes, the UIA dialog guardian,
  VBOM module injection and type-library references.

What is missing is a DLL that hosts a `calcvba-host` session inside Excel, a way to hand it
Excel's `Application`, and application objects that use it instead of activating a new Excel.

## Design

### One in-process core, two loaders

`tools/calcvba-excel` is a Windows x64 `cdylib`: a test and embedding tool, not a compiler
target (`docs/ARCHITECTURE.md`: no XLL target belongs to the compiler). It owns:

- **Sessions.** One `calcvba-host` session per loaded project (source `.basproj`, or modules
  given as text), on VM3 or the JIT.
- **The host.** `StandardHostServices` with the application objects the loader registers
  (`Application` as the instance of `Excel.Global`).
- **A C ABI.** It is the only contract, and both loaders use it:
  - `calcvba_open(project: BSTR, backend: i32, out handle) -> HRESULT`
  - `calcvba_set_app_object(handle, clsid: BSTR, object: IDispatch*) -> HRESULT`
  - `calcvba_run(handle, entry: BSTR, args: SAFEARRAY(VARIANT)*, result: VARIANT*) -> HRESULT`
  - `calcvba_close(handle) -> HRESULT`
  - `calcvba_last_error(handle, number: i32*, description: BSTR*) -> HRESULT`

  None of it names Excel; the loaders supply Excel's objects.

  Arguments are Automation types, so real VBA calls them with `Declare PtrSafe` and its own
  cells (`ByRef result As Variant` is a `VARIANT*`). No registration is needed.
- **Owning-thread checks.** The first caller's thread owns every session. A call from another
  thread fails with `RPC_E_WRONG_THREAD` and touches nothing, as the native template does.

The two loaders:

1. **The VBA loader.** A standard module injected into the test workbook (VBOM, as the probe
   harness already does). It loads the DLL by full path, opens a session with `Application`
   and `ThisWorkbook`, runs entries, and closes the session in `Workbook_BeforeClose`. Real VBA
   and CalcVBA then share the process, the workbook and the objects. A real VBA procedure can
   pass an object into CalcVBA and compare identity, and both languages can attach sinks to
   the same source.
2. **The XLL loader.** The same DLL also exports `xlAutoOpen`, `xlAutoClose` and
   `xlAutoRegister12`, through `MdCallBack12` from the host process. It registers command
   macros `CALCVBA.OPEN`, `CALCVBA.RUN` and `CALCVBA.CLOSE`, which take the same arguments as
   the C ABI. It finds `Application` from `xlGetHwnd`: the `XLMAIN` window, its `EXCEL7`
   child, `AccessibleObjectFromWindow(OBJID_NATIVEOM)`, then `.Application`. That is the
   `bd-p5bb` identity rule. No VBA project is involved, so this path checks CalcVBA alone in
   Excel, as an embedding host would run it. The harness reaches the commands with
   `Application.Run`.

### The object model: application objects and foreign names

*Revised 2026-11-15 after the Excel probe `temp/variant-layout-probe/global10` (runs 1-6; the main module is `tools/calcvba-excel/probes/GlobProbe.bas`).*

**What Excel's VBA does.** There is no special "host root". Excel's type library has one
*application object* (appobject): the coclass `Global` (`{00020812-0000-0000-C000-000000000046}`,
`TYPEFLAG_FAPPOBJECT | FCANCREATE | FPREDECLID`), whose default interface `_Global`
(`{000208D9-...}`, dual, **non-extensible**) has 57 members. The probe shows:

- An unqualified name VBA cannot resolve in the project or an earlier library is looked up in
  `_Global`: `Range`, `Cells`, `ActiveSheet`, `Workbooks`, and `Application` and `ThisWorkbook`
  themselves. `ThisWorkbook Is Application.ThisWorkbook`, and `Application Is Excel.Application`.
  The library name qualifies them (`Excel.Range("A1")`, `Excel.ThisWorkbook`), and a project
  `Range` function or `ThisWorkbook` variable shadows them, as any library name.
- A name that is not in `_Global` is not found: with `Option Explicit`, bare `MyDog` (a
  workbook name) is "Variable not defined"; without it, an implicit Empty variable.
- `Application.MyDog` and `CallByName(Application, "MyDog")` raise 438. `Application` is
  `_Application`, which is extensible, so the call compiles late-bound and fails when Excel's
  `GetIDsOfNames` does not know the name. `Application.Sum(1, 2)` is 3 the same way.
- A *foreign name* `[text]` is the host's evaluation. Unqualified, it is first an ordinary
  name (`[abc]` is the local `abc`, `[ActiveSheet]` the global, `[Helper]` a project function).
  Otherwise it calls the appobject's `DISPID_EVALUATE` (-5) member (`_Global._Evaluate`) with
  the text: `[MyDog]` is the Range, `[A1+1]` 43, `[{1,2,3}]` an array, `[NoSuchName]` the error
  value 2029 (no run-time error), `[A1] = 5` assigns the Range's value. It compiles under
  `Option Explicit`.
- After a dot, `obj.[text]` is the member named `text` when the static interface has one
  (`Application.[Name]`, `ThisWorkbook.[Name]`). Otherwise it is `DISPID_EVALUATE` on `obj`:
  `Application.[MyDog]` 42, `ActiveSheet.[A1]`. On a non-extensible interface without it,
  `ThisWorkbook.[A1]` is the compile error "Method or data member not found". Late-bound, it
  is always `DISPID_EVALUATE` (`o.[Name]` gives `#NAME?`, a Dictionary raises 438).
- `_Global` and `_Application` give these 57 members the same DISPIDs and vtable offsets, so
  Excel's `Application` pointer serves as the `Global` instance.

**CalcVBA, host-independent.** Nothing here names Excel:

- *Metadata.* A library-level typelib blob lists its appobject coclasses (name, CLSID, default
  interface). Every interface blob records whether it is extensible. The symbol layer loads each
  appobject's default interface with the library.
- *Binding.* Library providers resolve an unqualified name, after the project and in reference
  order, against their appobject's members, and `Lib.member` the same way. The binder lowers
  such a name to a member access on `AppObject(clsid)`, typed as the appobject's interface, so
  it binds early like any typed receiver. `CreateObject` returns to meaning activation; the
  `ComObjectRoot` lowering to `CreateObject` is removed.
- *Unknown members.* A member missing from an extensible early-bound interface is a late-bound
  call by name (438 at run time if the object does not know it). A member missing from a
  non-extensible one stays the compile error "Method or data member not found".
- *Foreign names.* `[text]` is an ordinary name when it resolves. Otherwise it is
  `EvaluateForeign(receiver, text)`: an `IDispatch::Invoke` of `DISPID_EVALUATE` with the
  text. The receiver is the object after the dot, or for an unqualified name the first
  appobject (in reference order) whose interface has a `DISPID_EVALUATE` member. With no such
  appobject, an unresolved unqualified foreign name is the compile error "Variable not defined",
  as a plain name is. The result is assigned and indexed like any call result, so `[A1] = 5`
  is a default-member assignment.
- *Instances.* `ComHal::app_object(clsid)` returns the session's instance. The host may
  register one per CLSID (`HostServices`, the embedding API, the C ABI). With none
  registered, the first use creates one with `CoCreateInstance` when the coclass is creatable
  and the host permits COM activation, and keeps it for the session, as VBA does for a
  library appobject such as DAO's `DBEngine`; out-of-process tests get a new Excel this way.
  Otherwise the use raises 429.
- *Excel hosts* register `Application` as the instance of `Excel.Global`. `ThisWorkbook` then
  comes from Excel: in the VBA loader it is the loader's workbook, as it is for the calling
  macro; from the XLL loader it is what Excel answers outside a VBA macro.
- *Document modules.* CalcVBA cannot be a workbook's `ThisWorkbook` or `Sheet1` module.
  Document events are reached through `WithEvents` on the workbook and sheets instead.

### Events in the process

- **During a run.** As built in step 8c: the run registers as the thread's executor, the sink
  runs the handler inside Excel's `Invoke`, and `End` or an unhandled error ends the run.
- **After a run (new).** In VBA a sink kept in a module variable keeps handling events after
  the macro that created it has returned; this is the classic application-events class. Step
  8c drops an event that arrives with no executor, which is right for an out-of-process run
  that is over, but not for a resident session. A session therefore registers an idle executor
  for its lifetime, at the bottom of the thread's executor stack. An event with no running
  VBA starts a new top-level activation of the handler on the session, the way a host call
  does. An unhandled error there is a top-level run-time error. VBA shows its dialog; the
  session records the error for `calcvba_last_error`, and the loader reports it. `End`
  resets the session's project.
- **Mixed sinks.** Real VBA and CalcVBA sinks on one source are separate connections, and
  Excel calls them in its own order. The comparison reads that order from Excel rather than
  assuming it.

### Lifetime

`calcvba_close` and `xlAutoClose` disconnect every sink and release every object before
the session is dropped. While any session is open, the DLL holds one extra `LoadLibrary`
reference to itself, which the last close releases, so neither `FreeLibrary` from VBA nor an
add-in unload can remove code a sink still points into. Each test run copies the DLL to a new
path, because Excel keeps a loaded DLL locked.

## Tests

Tests run in an owned Excel through the existing session tooling: a private desktop, the
dialog guardian, cleanup of owned processes only, and VBE Debug → Compile before real VBA runs.
A new case kind in `run-excel-probe-session.ps1` handles in-process cases:

- **Load.** Load the core DLL through the VBA loader, or `RegisterXLL` the same DLL.
- **Run.** Run each scenario three ways: real VBA, then CalcVBA on VM3, then on the JIT, all
  in the same process and workbook.
- **Record.** Write each result line and event log into `results.json`, in the format of the
  `events8` probes, so the rows compare one to one.

First cases:

1. **Smoke.** `Application.Version`, a cell write, `ThisWorkbook.Name`, unqualified `Range`
   and `Cells`, `Application.Sum`, and the foreign names of the global10 probe.
2. **The events8 scenarios in-process.** Synchronous delivery, `Cancel` written back,
   re-entrant `SheetChange`, `Set x = Nothing`, the unhandled-error abort and `End`.
3. **Idle delivery.** A sink kept in a module variable after its run, then the harness changes
   a cell and the handler logs it. The idle handler error is also checked.
4. **Mixed sinks and identity.** A real VBA sink and a CalcVBA sink on one `Application`, and
   an object passed from real VBA into CalcVBA (`Is`, `ObjPtr`).
5. **Release counts.** `Set x = Nothing` releases at once with no proxy in between, read
   through `ObjPtr` and `IUnknown::Release` return values from a Declare.
6. **Early-bound in-process.** Omitted optionals, `[lcid]`, ByRef `Cancel`, and objects
   returned through `[out, retval]`.
7. **The XLL loader.** The same cases with no VBA project in the workbook.

## Order

1. **Application objects and foreign names** in the compiler and the HAL, then **the core DLL**
   (C ABI, VBA loader). Done when the smoke cases pass on VM3 and the JIT.
2. **Events during a run.** In-process, compared with real VBA.
3. **Idle delivery.** The session's idle executor, mixed sinks, identity and release counts.
4. **The XLL loader.** `xlGetHwnd` identity and command registration, with the same cases.
5. **The corpus.** The broader corpus run in-process.

*As built 2026-11-15 (items 1 and 4).*

- **Compiler and HAL.** As designed in the object-model section. Appobject members are called
  through `IDispatch` by DISPID, never a vtable: the host's instance need not answer the
  interface's `QueryInterface`, and Excel's `Application` does not answer it for `_Global`.
  A library name qualifies only when it is not itself a bound name.
- **Declares into the runtime.** `Lib "VBE7" Alias "VarPtr"` (and `vbe6`, `msvbvm60`) is the
  runtime's own identity export.
- **The DLL.** `tools/calcvba-excel` builds `calcvba_excel.dll`. Its C ABI:
  - `calcvba_new(out handle)`
  - `calcvba_add_module(handle, name, kind, source)`
  - `calcvba_add_reference(handle, name, libid, major, minor)`
  - `calcvba_set_app_object(handle, class, IDispatch*)`
  - `calcvba_add_document(handle, name, host_type, source, IDispatch*)`: a document module
    (`ThisWorkbook`, `Sheet1`) and its host object, of type `host_type` as the references
    name it (`Excel.Workbook`)
  - `calcvba_compile(handle, backend)`
  - `calcvba_run(handle, module, entry, VARIANT* result)`
  - `calcvba_last_error(handle, number*, VARIANT* description)`
  - `calcvba_close(handle)`
  
  Strings are BSTRs (`StrPtr` from VBA). A run executes the whole project on `calcvba-engine`
  (VM3 or the JIT). A driver module's `Sub Main` stores the entry function's value. The
  session is therefore not resident yet (items 2 and 3).
- **Loaders.**
  - `loaders/CalcVBALoader.bas` copies the workbook's modules into a project: standard,
    class and document modules. Each document module goes in with its workbook, sheet or
    chart (found by code name) and the type `"Excel." & TypeName(object)`.
  - `loaders/CalcVBAXllCompare.bas` drives the XLL. The same DLL's `xlAutoOpen` registers
    `CALCVBA.RUN(source, module, entry, backend)`.
  - In the owned test Excel, `xlGetHwnd` gave 0, so the XLL takes the `XLMAIN` window of its
    own thread (the same process; never the ROT). From that window it reaches `EXCEL7` and
    then `AccessibleObjectFromWindow`.
- **Result.** `scripts/run-inproc-probe.ps1` runs the global10 probe in one owned Excel: real
  VBA, then CalcVBA on VM3, then on the JIT. That is 25 lines covering `ThisWorkbook`,
  `Application`, unqualified and `Excel.`-qualified globals, `Application.Sum`, foreign names
  and their assignment, and the 438 cases. Through both loaders, VM3 and the JIT give every
  line exactly as VBA does.

*As built 2026-11-15 (document modules).*

- **Host-neutral.** `ComHal::register_document_object` / `document_object_variant` keep a host
  object per document module. When VM3 or compiled code first creates a document module's
  default instance, `calcvba_rt_abi::attach_host_document` puts that object in its host field.
  A document instance assigned to its host type (`Set wb = ThisWorkbook`, `wb As Workbook`)
  is the host object; a foreign name on a document instance is its host's.
- **Result.** Probe tw11 (18 lines, one workbook each, `run-inproc-probe.ps1 -PerModule`): the
  five VBA refuses to compile, CalcVBA refuses too (`excel_foreign_names.rs`). Of the 13 that
  run, 11 match VBA on VM3 and the JIT. `v = ws.[A1]` and `v = Sheet1.[A1]` store the Range,
  where VBA's Let takes its value: a general Let gap, bead dc-6hobz.
- **Not yet.** The XLL loader compiles one module's source without a VBA project, so it has
  no document modules.

## Risks

- **A defect crashes Excel.** It takes the session and the run with it. The guardian records
  the crash and the harness reports the case. Tests only use owned Excel processes.
- **Unsigned XLLs.** Trust Center policy can refuse an unsigned XLL. The harness detects the
  refusal and reports it, and never changes the settings. The VBA loader needs VBOM access,
  which the probe harness already relies on.
- **Executable memory.** The JIT maps executable memory inside Excel. Excel does not enable
  arbitrary-code guard, but a policy that does would refuse the JIT. In that case the JIT fails
  with its own error; it never falls back to VM3.
- **Stack.** VM3 and the JIT run on Excel's main thread stack, so deep recursion reaches VBA's
  "Out of stack space" (28) on CalcVBA's own depth guard first.
