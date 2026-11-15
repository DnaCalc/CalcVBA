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
Excel's `Application`, and host roots that use that object instead of activating a new Excel.

## Design

### One in-process core, two loaders

`tools/calcvba-excel` is a Windows x64 `cdylib`: a test and embedding tool, not a compiler
target (`docs/ARCHITECTURE.md`: no XLL target belongs to the compiler). It owns:

- **Sessions.** One `calcvba-host` session per loaded project (source `.basproj`, or modules
  given as text), on VM3 or the JIT.
- **The host.** `StandardHostServices` plus a host-root registry: `Application`, and optionally
  `ThisWorkbook`, as `ObjectRef`s over the interface pointers the loader supplies.
- **A C ABI.** It is the only contract, and both loaders use it:
  - `calcvba_excel_open(project: BSTR, application: IDispatch*, workbook: IDispatch*, backend: i32, out handle) -> HRESULT`
  - `calcvba_excel_run(handle, entry: BSTR, args: SAFEARRAY(VARIANT)*, result: VARIANT*) -> HRESULT`
  - `calcvba_excel_close(handle) -> HRESULT`
  - `calcvba_excel_last_error(handle, number: i32*, description: BSTR*) -> HRESULT`

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

### The object model

- **Host roots.** Today a host root (`DispatchRoute::ComObjectRoot`, for example
  `Application`) is lowered to `CreateObject(progid)` at each use. In-process, that would
  start or attach to the wrong Excel. The root becomes its own native,
  `HostRoot(type_name)`, resolved by a new `ComHal::host_root(name) -> Option<ObjectRef>`
  over the registry. A host with no root of that name falls back to activation, which is
  today's behaviour for out-of-process tests.
- **Globals.** In Excel's VBA, unqualified `Range`, `Cells`, `ActiveSheet`, `Worksheets` and
  the like are members of the type library's `appobject` (`Application`, through `_Global`).
  The binder resolves those names to members of the appobject root. Which names VBA resolves
  that way is checked with an Excel probe first.
- **Calls.** Calls use the type library Excel registers (`{00020813-…}` 1.9) through the
  existing catalog. With direct objects there is no proxy, so early-bound calls always take the
  vtable (`automation::early`), with no `PSOAInterface` check.
- **Document modules.** CalcVBA cannot be a workbook's `ThisWorkbook` or `Sheet1` module.
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
  session records the error for `calcvba_excel_last_error`, and the loader reports it. `End`
  resets the session's project.
- **Mixed sinks.** Real VBA and CalcVBA sinks on one source are separate connections, and
  Excel calls them in its own order. The comparison reads that order from Excel rather than
  assuming it.

### Lifetime

`calcvba_excel_close` and `xlAutoClose` disconnect every sink and release every object before
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

1. **Smoke.** `Application.Version`, a cell write, and `ThisWorkbook.Name` through the host
   roots.
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

1. **The core DLL.** C ABI, VBA loader and the `Application` host root. Done when smoke cases
   pass on VM3 and the JIT.
2. **Events during a run.** In-process, compared with real VBA.
3. **Idle delivery.** The session's idle executor, mixed sinks, identity and release counts.
4. **The XLL loader.** `xlGetHwnd` identity and command registration, with the same cases.
5. **Globals and the corpus.** Appobject globals, the `ThisWorkbook` root, and the broader
   corpus run in-process.

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
