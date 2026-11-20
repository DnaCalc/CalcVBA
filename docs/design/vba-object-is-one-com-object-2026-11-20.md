# A VBA object is one COM object (dc-69djs)

Owner decision, 2026-11-20: a VBA class instance is one COM object with no wrappers. This is the rest
of step 10 of the VARIANT migration (epic dc-yqq0, plan `runtime-variant-automation-layout-2026-11-13.md`).
This note records what the code does today, what was measured, the recommended object layout and the
follow-up beads.

## 1. What VBA does

A VBA class instance is one COM object. Its one identity answers `IUnknown`, `IDispatch`, the class's
default dual interface `_ClassName`, every `Implements` interface, and `IConnectionPointContainer` and
`ISupportErrorInfo` where they apply. `QueryInterface(IID_IUnknown)` always returns the same pointer, so
`Is` compares pointers. Foreign code calls it late-bound through `IDispatch` or early-bound through the
vtable described by the project's type information. The last `Release` runs `Class_Terminate` on the
thread that owns the project.

## 2. What CalcVBA does today (master 2026-11-20, checked)

- **One layout, no COM surface.** VM3 (`new_project_instance`), the JIT (`rt_project_new_object`, rt-abi
  lib.rs:2705) and the native template (`cvb_class_new`, classes.rs:171) all allocate the same
  `CompatObjectBase` (calcvba-runtime object_ref.rs:641). Its vtable has three `IUnknown` slots, and
  `compat_query_interface` answers `IID_IUnknown` only (dc-yqq0.15), because answering anything else with
  that vtable would let foreign code call slots that are not there.
- **Three descriptor builders.** `RuntimeClassDescriptor` (object_ref.rs:291) is built separately by VM3
  (lib.rs:674-750), rt-abi for the JIT (lib.rs:870-941) and the native template (classes.rs:147).
- **DISPIDs and slots exist only with an export surface.** `RuntimeMemberDescriptor` carries
  `dispatch_id`, `vtable_slot` and `dual_dispatch` (object_ref.rs:230-281), but the binder fills real
  values only when the project has an export surface (calcvba-bind ids.rs:806-830). VM3
  (`dispatch_project_method`, lib.rs:4485) and the JIT (helpers.rs:3933, 4455) find members by name.
- **Wrappers at every crossing.**
  - `calcvba_com::automation::project_dispatch` makes a `RuntimeObjectDispatch`
    (windows_runtime_object.rs) whose `GetIDsOfNames` answers `DISP_E_UNKNOWNNAME` and `Invoke`
    `DISP_E_MEMBERNOTFOUND`. The Declare path uses it.
  - The bridge keeps a second wrapper cache, `runtime_object_dispatch_wrappers`
    (windows_runtime_state.rs:397-409), and `bind_native_dispatch_result` unwraps a wrapper that
    comes back (windows_runtime_state.rs:474-500).
  - The in-process server (calcvba-comhost) has its own `DispatchObject`, with a `WRAPPERS` map
    since 45a3a0c1 so the same object gets the same wrapper.
- **The right model exists once.** The native COM EXE server (native-template server.rs) serves a
  compiled vtable with its type information through `DispInvoke`, refuses calls off the owner thread
  (`E_WRONG_THREAD`), and posts a final release from a foreign thread to the owner thread (:168-182).

## 3. Measured on Windows (2026-11-20)

Probe: a `Widget` class (`Value`, `Twice`, `Class_Terminate` counting into a module variable) and the
COM test server `OxVba.TestEventServer`. Three new members were added to the server for this note:
`CallMember(target, name)` is a .NET late-bound call, and `Hold`/`Drop` keep and release a foreign
reference. Excel ran on a private desktop through `run-excel-probe-session.ps1`; CalcVBA ran
`calcvba-cli run-project`. Probe sources are in appendix A.

| Case | Excel VBA | CalcVBA VM3 | CalcVBA JIT |
|---|---|---|---|
| Instance stored in `Scripting.Dictionary`, read back: `Is` the original | True | True | True |
| Member call on the instance read back | 42 | 42 | 42 |
| Foreign late-bound call (`CallMember(w, "Twice")`) | 42 | error 5 | error 5 |
| `Class_Terminate` while a foreign holder keeps it | not run (0) | not run (0) | not run (0) |
| `Class_Terminate` after the holder releases it | run (1) | run (1) | run (1) |

The round trip works today only because the bridge unwraps its own wrapper. Foreign code cannot call a
project object at all. Terminate timing already matches.

## 4. Recommended object layout

The project object itself carries its COM interfaces. Every interface pointer is a field of the object;
no separate allocation stands for it.

```text
#[repr(C)] ProjectObject            (CompatObjectBase, grown)
  primary: *const ClassDualVtbl     identity: QI(IUnknown), QI(IDispatch), QI(class IID) return &primary
  error_info: *const SupportErrorInfoVtbl          QI(ISupportErrorInfo) returns &error_info
  implements: [*const ImplementsVtbl; n]  (later)  QI(interface IID) returns &implements[i]
  connection_points: ...                  (later)
  ref_count: AtomicU32              one count for every interface
  owner_thread: u32                 the project's thread
  invoker: *const dyn MemberInvoker the executor that runs this object's members (VM3, JIT, AOT)
  identity, class: &'static RuntimeClassDescriptor, fields, native_state, termination flags (as today)
```

- **Identity** is the address of `primary`. Each interface's slot 0 to 2 recover the object from the
  interface pointer by a fixed offset, as the EXE server does.
- **`ClassDualVtbl`** is `IUnknown` (3 slots), `IDispatch` (4 slots), then one entry per member in
  `vtable_slot` order. It is built once per class from the shared descriptor, so it is the same off
  Windows (the `system` ABI is the C ABI on x64 Linux and macOS).
- **`GetIDsOfNames` and `Invoke` are table-driven** from the descriptor: name to DISPID, and DISPID plus
  `wFlags` to member. Named arguments work as comhost now does them (parameter DISPID = position).
  Arguments are the caller's VARIANTs, borrowed, with ByRef bound to `pvarVal`, because the runtime
  Variant is the Automation VARIANT. No type library is needed to call an object; one is needed only
  to register a COM export (comhost, the EXE server).
- **Member entries past slot 7** serve early-bound foreign clients (a type library user). The interpreted
  executors fill them with generic slot thunks: one small function per slot index, which reads the
  member's wire signature from the descriptor and calls `invoker`. AOT images fill them with their
  compiled entries, as the EXE server does today.
- **`MemberInvoker`** runs a member by DISPID or slot on the owning executor. A call that arrives while
  VBA is suspended in a host call (a COM event a client answers by calling back, a Dictionary callback)
  runs on the active run. This is the scope the JIT and VM3 already register for COM events
  (`calcvba_runtime::events::ExecutorScope`), widened from event delivery to member calls. Between runs
  (comhost between client calls), the session holder runs it.
- **Threads.** `Invoke` off the owner thread answers `RPC_E_WRONG_THREAD`; a correct client is given
  a proxy by COM marshalling. A final `Release` on a foreign thread does not run `Class_Terminate`
  there: it moves the object to the owner's termination queue. The queue becomes multi-producer and is
  drained at the owner's next statement boundary or message, which is the EXE server's
  `PostThreadMessage` without needing a message loop.

## 5. What else changes

- **VM3 and JIT calls go by DISPID or slot.** Early-bound calls on a class variable use the slot;
  late-bound calls use the plan cache (`RuntimeDispatchPlanCache`, object_ref.rs:482) keyed by name.
  Name scans in `dispatch_project_method` and the JIT helpers go.
- **comhost.** `DispatchObject` and `WRAPPERS` go for project objects. comhost becomes the class
  factory, registration and type library for the project. `CreateInstance` returns the project object
  itself, and connection points move onto the object (later). The comhost-on-JIT work from
  2026-11-20 (not merged) carries over in part:
  - Kept: the JIT session calls for New and member-with-Me, the JIT `RaiseEvent` to host-sink fix,
    the JIT `ProjectRuntimeSession` backend, `RuntimeFlavor=Jit` for WrappedComServer, and Optional
    defaults for Missing arguments.
  - Changed: the session's member call becomes the JIT's `MemberInvoker`.
  - Closed: the re-entry gap (a client calling back during an event got an error on the JIT) closes
    with the active-run invoker.
- **Deleted.**
  - `RuntimeObjectDispatch` and `project_dispatch`.
  - The bridge's `runtime_object_dispatch_wrappers`, and the unwrapping in
    `bind_native_dispatch_result`. A project object that comes back is recognised by its vtable pointer.
  - comhost `DispatchObject`/`WRAPPERS` for project objects, and the Declare path's wrapper.
- **Guarded and DNA Calc hosting** (DnaCalc docs/architecture/CODE_EXECUTION_POLICY.md 2.1).
  - A project object has no origin: it is the project's own code.
  - Today a project object that comes back from a vouched object (c), such as the OS
    `Scripting.Dictionary`, is recognised only through the bridge's wrapper map. Otherwise it would
    take the origin of the object that returned it.
  - With one identity, recognition is a vtable-pointer check on the object itself, the same in every
    host and on every platform, so a project object never takes another origin.
  - Host-injected and host-offered objects (`HostObject`) are the host's own and are not changed here.

## 6. Risks

- **Threads.** Foreign threads calling an in-process object directly is a client bug that VBA tolerates
  only through COM's proxies. Refusing it (as the EXE server does) is safer than running VBA there.
  The final-release path is the one real cross-thread entry and needs a test with a foreign thread.
- **Reference counts.** One count across interfaces replaces several wrapper counts; a mistake now leaks
  or frees the object itself.
  - Coverage: the live counters, Miri on the portable vtable and QI code, and the foreign
    holder/terminate probe on Windows.
- **Terminate timing.** It must stay as measured (section 3): not while a foreign holder keeps the object,
  then once on the owner thread. Terminate on foreign release must not run during another member's body.
- **Re-entry.** Running a foreign call on the active run needs the same care as event handlers: the
  executor holds no borrow across a host call (true for the JIT helpers and VM3's event path today).
- **Performance.** Slot calls are faster than today's name scans. `Invoke` costs only foreign callers.
  The object grows by a few pointers per class interface.
- **DISPID values.** Foreign code can cache DISPIDs, and Excel's VBA assigns specific values. They should
  be measured with `ITypeInfo` on an Excel VBA class, not invented.

## 7. Follow-up beads (proposed order, under dc-yqq0)

1. **Object-identity tests first.** Commit the section 3 probe as Windows tests, with Excel's answers
   as expectations: Dictionary round trip, foreign late-bound call, foreign early-bound call through the
   type library, `Is` after crossing, Terminate with a foreign holder, and final release on a foreign
   thread. Record today's results; the foreign call is expected to fail until bead 4.
   Done: tests in CalcVBA, failing ones marked with the bead that fixes them.
2. **One class descriptor for the three executors,** with DISPIDs and slots for every class, export
   surface or not. Match Excel VBA's values, measured with `ITypeInfo` on an Excel class. Done: VM3,
   rt-abi and the native template share one builder; descriptor tests on Linux and Windows.
3. **VM3 and JIT calls by DISPID or slot,** name lookups through the plan cache only. Done: no class-member
   name scan in VM3 or the JIT helpers; differential suites green.
4. **The object answers IDispatch and its class IID.** Add the `ClassDualVtbl` with table-driven
   `GetIDsOfNames`/`Invoke`, the `MemberInvoker` on VM3, the JIT and AOT, and the active-run scope.
   Done: the foreign late-bound call in bead 1 gives 42 on both executors and AOT; portable Invoke tests
   on Linux; Miri clean.
5. **Member slots for early-bound foreign clients:** generic slot thunks for VM3/JIT and compiled
   entries for AOT. Done: the early-bound foreign case in bead 1 matches Excel.
6. **Final release off the owner thread** queues termination for the owner. Done: the foreign-thread test
   in bead 1 runs `Class_Terminate` once on the owner thread.
7. **Delete the wrappers.** `RuntimeObjectDispatch`/`project_dispatch`, the bridge wrapper cache and
   unwrapping, comhost `DispatchObject`/`WRAPPERS` for project objects, the Declare wrapper. comhost
   returns the object itself. Done: those types are gone; the COM, comhost, Declare and native suites
   are green.
8. **Later: `Implements` interfaces and connection points on the object,** and `ISupportErrorInfo`.
   Done: QI for an implemented interface returns its face on the same object; events reach a connection
   point client.

Beads 2 and 3 are portable and can run on Linux; beads 1, 5 and 6 need Windows for their COM evidence.

## Appendix A: probe sources

`Widget` (class module):

```vb
Public Value As Long
Public Function Twice() As Long
    Twice = Value * 2
End Function
Private Sub Class_Terminate()
    gTerm = gTerm + 1
End Sub
```

`Probe` (standard module):

```vb
Public gTerm As Long
Public Function Run() As String
    Dim w As New Widget, d As Object, back As Object, s As Object, r As String
    w.Value = 21
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "k", w
    Set back = d("k")
    r = "dict-is=" & CStr(back Is w) & ";dict-call=" & CStr(back.Twice())
    Set s = CreateObject("OxVba.TestEventServer")
    On Error Resume Next
    r = r & ";foreign-call=" & CStr(s.CallMember(w, "Twice"))
    If Err.Number <> 0 Then r = r & "ERR" & Err.Number
    On Error GoTo 0
    Dim h As Widget
    Set h = New Widget
    gTerm = 0
    s.Hold h
    Set h = Nothing
    r = r & ";term-held=" & CStr(gTerm)
    s.Drop
    r = r & ";term-dropped=" & CStr(gTerm)
    Run = r
End Function
```

Excel: `dict-is=True;dict-call=42;foreign-call=42;term-held=0;term-dropped=1`.
CalcVBA VM3 and JIT: `dict-is=True;dict-call=42ERR5;term-held=0;term-dropped=1` (the failing call
ends the concatenation, so the `foreign-call=` label is lost with it).
