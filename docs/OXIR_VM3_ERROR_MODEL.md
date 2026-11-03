# VBA error handling

OxIR retains a fault edge for each fallible statement. Its FaultDispatch
terminator names the faulting statement and the next statement, so VM3 and
the JIT can apply On Error and Resume without reconstructing source text.
The runtime tests cover the supported Err/handler behaviour.

## 2. Runtime state (per VM and per activation)

| State | Scope | Meaning |
|---|---|---|
| `error_mode: ErrorMode` | per-activation (VM cell, saved/restored per frame) | the active **handler policy**: `None` (Default), `ResumeNext`, `Goto(BlockId)` |
| `active_error: Option<ResumePoint>` | per-activation (VM cell, saved/restored per frame) | the spec's **"active error"** latch + the captured `(resume, resume_next)` seeds. `Some` ⇔ an error is currently being handled in this activation |
| `err: ErrState` | global (the singleton `Err` object) | `number`, `description`, `source` (+ `helpfile`/`helpcontext` once M2-c-2 lands); `last_dll_error` |
| `gosub_stack: Vec<BlockId>` | per-frame | the **GoSub Resumption List** (LIFO) |

`ErrorMode` is the spec's *handler policy*; `active_error` is the spec's *active error /
fault statement*. **vm2 conflates the two (it has only the policy), which is the root
cause of most divergences in [§7](#7-confirmed-vm2-vs-spec-divergences).**

Per-activation cells are saved on the callee `Frame` at call time and restored on
return / unwind (alongside `error_mode`), so each procedure has its own handler + active
error (MS-VBAL §5.4.4 ll. 2793–2794: each invocation has its own policy, born Default).

---

## 3. The OxIR contract (what the elaboration emits)

| Construct | Form | Notes |
|---|---|---|
| `OxInst::SetErrorHandler(ErrorHandler)` | `ResumeNext \| Goto0 \| GotoLabel(BlockId)` *(+ `ClearActiveError` for `On Error GoTo -1`, M2-c-4)* | sets `error_mode`; **also resets `Err`** (Rule R5) |
| `OxTerminator::FaultDispatch { resume, resume_next }` | the landing pad | dispatches on `error_mode` (Rule R4) |
| `OxTerminator::Resume` / `ResumeNext` / `ResumeLabel(BlockId)` | terminators | the three `Resume` forms (Rules R6–R8) |
| `OxTerminator::Raise { code }` / `RaiseValue(op)` | terminators | `Err.Raise` / `Error n`; route through the block's `fault_target` so `On Error` catches them |
| `OxInst::ErrFieldGet { dst, field }` | `Number/Description/Source/LastDllError` | read `Err` |
| `OxInst::ErrFieldSet { field, src }` | `Number/Description/Source` | write mutable `Err` fields; `LastDllError` is read-only |
| `OxInst::ClearErr` | — | `Err.Clear` → reset `Err` |
| `OxTerminator::GoSub { target, ret }` / `GoSubReturn` | terminators | LIFO resumption list (Rules R12) |
| `OxInst::StmtBoundary { stmt }` | statement index | finalization timing (M3); **not used by the error model** (resume seeds live on the pad) |

A block ending in `Raise`/`RaiseValue` is given a `fault_target` because
`OxTerminator::is_fallible()` is true for them and `finish_to` honours it (committed in
M2-c prep `478df517`); the verifier requires it.

---

## 4. Execution rules (spec-grounded)

Each rule cites its MS-VBAL basis. "R*" ids are referenced from the code.

- **R1 — Per-activation policy, born Default.** Each activation's `error_mode` starts
  `None` (Default) and is saved/restored across calls. The base frame surfacing a
  `Fault` out of the run **is** the host "Terminate" boundary. *(§5.4.4 ll. 2793–2794.)*

- **R2 — Default propagation across calls.** An unhandled fault discards the activation,
  re-applies the *caller's* policy, and re-seats the fault at the **call site**: a later
  `Resume` re-runs the calling statement, `Resume Next` the statement after the call.
  *(§5.4.4 l. 2799; ll. 2801, 2832.)*

- **R3 — The `active_error` latch is the spec's "active error".** It is set when a live
  fault is routed into a handler (`ResumeNext` continuation or `Goto` handler), and
  cleared by a successful `Resume`/`Resume Next`/`Resume <label>`, by `Exit
  Sub/Function/Property`, and by `On Error GoTo -1`. *(§5.4.4 ll. 2803, 2805, 2807.)*

- **R4 — `FaultDispatch` dispatch.** With `error_mode`:
  - `None` (Default) → **propagate** the fault to the caller (early unwind).
  - `ResumeNext` → set `active_error = Some((resume, resume_next))`, then continue at
    `resume_next`. **⚠ Unverified discrepancy (bd-ivaha.31):** the vm3 code
    (`oxvba-rt-abi` `FaultDispatch`) currently clears `active_error` (`None`) on the
    implicit ResumeNext skip, contradicting this rule. `On Error Resume Next` does not
    establish a handler, so `None` may be correct — but neither behavior is
    live-oracle-confirmed. Do not rely on `active_error` after an implicit ResumeNext
    skip until a probe resolves this; the code and this rule must then be reconciled.
  - `Goto(h)` → **demote `error_mode` to `None`** (Rule R9), set `active_error =
    Some((resume, resume_next))`, then transfer to `h`.

- **R5 — Every `On Error` resets `Err`.** All four forms reset the `Err` object
  *unconditionally*, before applying their policy change. *(§5.4.4.1 l. 2821; §6.1.3.2.1.1
  l. 9052.)*

- **R6/R7/R8 — `Resume` semantics.** First, **if `active_error` is empty → raise
  runtime error 20 ("Resume without error")** (R7-precondition, §5.4.4.2 l. 2830). Else
  reset `Err` and clear `active_error`, then transfer: `Resume` → `resume`; `Resume Next`
  → `resume_next`; `Resume <label>` → the label block. *(§5.4.4.2 ll. 2830–2832;
  §6.1.3.2.1.1 l. 9050.)*

- **R9 — `On Error GoTo <label>` is single-shot.** When the handler **fires**, the
  activation policy is demoted to Default *before* control reaches the handler label. A
  re-raise inside the handler therefore **propagates to the caller**, not back into the
  same handler. The handler is re-armed only by an explicit `On Error` inside it, or by
  `Resume*`/`On Error GoTo -1` clearing the active error. *(§5.4.4 l. 2803.)*

- **R10 — `Exit Sub/Function/Property` resets `Err`** (and clears `active_error`). A
  *normal* `End Sub` / fall-off-the-end does **not**. *(§6.1.3.2.1.1 ll. 9049–9051.)*

- **R11 — `Err.Raise(Number[, Source][, Description][, HelpFile][, HelpContext])`.** All
  supplied args honoured; omitted args inherit the current inheritable `Err` fields; omitted
  `Source` defaults to the project name; omitted `Description` is mapped from `Number`
  (else "Application-defined or object-defined error"). Raised errors and
  `Err.Description`/`Err.Source` writes make fields inheritable; `Err.Clear` resets
  that state; `Err.Number` writes do not create it but also do not clear it. `Error <n>`
  uses the same default-description mapping, but does not inherit prior `Err` fields.
  *(§6.1.3.2.1.2 ll. 9055–9071; §5.4.4.3; Excel/VBA oracle in
  `docs/evidence/conformance/vm3_default_error_message_oracle_20260701T1410Z/`
  and `docs/evidence/conformance/vm3_err_property_writes_oracle_20260701T1442Z/`.)*
  **(M2-c-2/3.)**

- **R12 — GoSub is a per-activation LIFO list.** `GoSub` pushes its `ret`; `Return` pops
  the most recent and branches there; `Return` on an empty list raises runtime error 3
  ("Return without GoSub"). *(§5.4.2.14 l. 2523; §5.4.2.15 ll. 2527–2528.)*

- **R13 — `On Error GoTo -1`** clears the `active_error` latch and resets `Err`, but
  leaves the handler label/policy intact (so the same handler can re-catch). The spec
  *grammar* admits `GoTo -1` (§5.4.4.1 l. 2815) but the *runtime prose is silent*; the
  effect here is documented Office behaviour and **must be confirmed by a live oracle
  run before locking** (see [§8](#8-needs-live-excel-oracle)). **(M2-c-4.)**

- **R14 — `End`** terminates the whole run immediately with **no** `Class_Terminate`
  drain (it closes files / clears variables, but does not fire finalizers). vm3's `Halt`
  (truncate to base, no finalization) is spec-consistent. *(§5.4.2.22.)*

---

## 5. Cross-call propagation

A callee fault with no enabled handler unwinds one activation (`propagate_fault`),
restores the caller's `error_mode`/`active_error`, and routes at the caller's call-site
block. Because that block's `FaultDispatch` was seeded by the elaboration with the
*call-site* statement's `(resume, resume_next)`, a `Resume` in the caller's handler
re-runs the calling statement and `Resume Next` continues after the call — Rule R2,
**for free, with no special cross-frame logic**. The `active_error` latch is the
*caller's* (set when the caller's pad catches the propagated fault), never leaked from
the unwound callee.

---

## 6. Implementation map (vm3)

`crates/oxvba-vm3/src/lib.rs`:
- `Vm3` cells: `error_mode`, `active_error: Option<ResumePoint>`, `err: ErrState`,
  `last_dll_error`.
- `Frame`: `saved_error_mode`, `saved_active_error`, `gosub_stack`.
- `run_loop` terminator arms: `FaultDispatch` (R4/R5/R9), `Resume*` (R6/R7/R8),
  `Raise`/`RaiseValue` (route via `fault_target`, R11-runtime), `GoSub`/`GoSubReturn`
  (R12).
- `exec`: `SetErrorHandler` (R5), `ErrFieldGet`, `ErrFieldSet`, `ClearErr`.
- `call_proc`/`do_return`/`propagate_fault`: save/restore `active_error` (R3) +
  per-frame `gosub_stack`.

Elaboration (`crates/oxvba-oxir`): `Raise`/`RaiseValue` fault edge (done, `478df517`);
`Exit` path emits a `ClearErr` (R10); `On Error GoTo -1` → `ErrorHandler::ClearActiveError`
(R13, M2-c-4).

---
