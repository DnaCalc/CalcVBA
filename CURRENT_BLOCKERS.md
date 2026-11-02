# Current Blockers

Date: 2026-11-02
Status: current; no all-path blocker
Authority: blockers only; ordinary unfinished capability work belongs in the three current worksets and canonical matrices

## BLK-WORKSHEETFUNCTION-VTABLE-20261101 — Excel Range and WorksheetFunction have no custom vtable

Status: resolved by the owner's 2026-11-02 ruling; bounded delivery is bd-59co.3.5.9.
Owner: bd-59co.3.5.3 retains broader native client and certification scope.

The installed Excel type library reports `Excel.WorksheetFunction.Sum` as
`SourceTypeKind::Dispatch`, `is_dual=false`, IID `00020845-0000-0000-C000-000000000046`,
DISPID 16388, no vtable slot, and vtable bound 7. Its 30 Variant arguments retain
one required parameter and 29 optional parameters. There is no direct Sum vtable
entry to call. This is checked by `installed_excel_typelib_retains_native_function_and_range_plans`
in typed_native_com; rejected native Excel run calcvba10-native-excel-initial owns
PID 27992 and records owned-process-zero, without passing runtime credit.

The owner permits what VBA does for a pure dispinterface: resolve the DISPID from
the typelib at compile time and issue one prebound Invoke with typed arguments,
without runtime name lookup. Optimizations require measurements. The owning
planned member path now retains that explicit transport; dual calls do not retry
through it. No custom slot, guessed ABI or registration change is introduced.
Range is also `SourceTypeKind::Dispatch`, `is_dual=false`, IID
`00020846-0000-0000-C000-000000000046`, bound 7, with no custom slot for Value
(DISPID 6) or its default member (DISPID 0). Microsoft's public
[Range interface](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.range?view=excel-pia)
and [WorksheetFunction interface](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.worksheetfunction?view=excel-pia)
both declare InterfaceType(2), agreeing with the installed metadata. The second
owned Excel attempt records the Range refusal rather than claiming it passed.
The run-12 controlled fixture and owned Excel/VBA comparison cover both typed
backends. See [dispatch-only delivery evidence](docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-prebound-dispatch-run12.md).
This is prebound IDispatch evidence, never Range vtable evidence. Broader Windows
signatures, native event/ByRef behavior and pinned certification remain in-progress.

## BLK-EXCEL-COMPILE-20261031 — Negative compile oracle owned dismissal

Status: resolved on the development/oracle host by the 2026-10-31 run 2 supervisor repair. Owner: `bd-59co.3.15.6`.

The native button path requires the exact recorded owned dialog, unchanged title and normalized complete message, and one visible enabled matching Button. Both native and UIA actions recheck that proof and require dialog disappearance. Immutable token/full-line capture still precedes dismissal. The worker now waits for the asynchronous compile outcome, publishes controls atomically, and records a complete helper identity only after its executable path is observable.

The final default five-case self-test and the separate unhandled-runtime modal diagnostic pass in Safe Mode without changing Office security settings. The negative compile transcript contains `MissingOracleSymbol`, its full source line and linked successful dismissal; the intrinsic-name probe selects `Fix` in its declaration. Every accepted case reports owned-process-zero. The pre-existing Excel process remains untouched. Rejected intermediate runs are documented and confer no credit.

Evidence: `docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/oracle-supervisor-2026-10-31/README.md`, byte-preserving raw archive and hash manifest. The earlier `vbe-selection-2026-10-31.md` remains historical failed-run evidence. Pinned clean certification is still open under WIN-14 and receives no capability, matrix or release credit from this development observation.

## BLK-ORDINAL-SELECTION-20261031 — VBE ordinal compile errors expose no selected token

Status: open for negative ordinal oracle authority; compiler work can continue from public MS-VBAL authority. Owner: `bd-59co.2.4.5`, with the next supervisor contract extension described below.

The repaired default supervisor passes its five scoped cases and the separate runtime modal. Fresh source-specific ordinal probes show a distinct shape: `#32768`, `#65535`, `#65536` and `#1x` display Invalid ordinal with the full declaration line but an empty UIA selection. The token-required ledger correctly rejects them and never dismisses them on an invented selection. Exact raw attempts remain rejected in `docs/evidence/programs/ideal-2026-07/core/CORE-3/declare-ordinals-2026-10-31/`. A separate `#01` probe is rewritten to `#1` during injection and fails the exact-source precheck.

MS-VBAL 5.2.3.5 independently specifies ordinal integer literals in 0..32767, so the source validator can enforce that rule without promoting these rejected attempts. Accepted fresh positive boundaries and malformed-token failures are listed in `core/CORE-3/c_decl.md`.

Next: capture and prove a genuinely empty selection, its exact caret/line/source identity and an explicit case expectation in a separate compile-error evidence contract; add hostile missing/stale/foreign selection tests. Preserve immutable capture-before-dismiss, exact owned buttons and the existing nonempty-token contract. Do not infer a token from the Alias text or use a post-dismiss snapshot as authority. Broader declaration and pinned certification work remains open; no user setting change is needed.

## BLK-BASELINE-001 — Linux baseline and cross-platform reconciliation remain open

Status: open, progress available
Owner: core workset CORE-1

Impact:

- the aggregate Rust baseline is green: format, strict all-target Clippy,
  ordinary workspace tests, parallel/serial differentials, isolated carrier
  balance, current host/JIT diagnostics and governance pass;
- the Windows x64 development transcript is complete under
  `bd-59co.2.2.10`; the pinned Linux x64 CI transcript remains open; and
- the five CORE-1 canonical rows remain planned until terminal reconciliation.

Unblocking outcome:

- execute and retain the pinned Linux x64 CI baseline under
  `bd-59co.2.2.11`; and
- reconcile both platform results and the five canonical rows under
  `bd-59co.2.2.12`.

This does not block architecture/compiler/package work, but no Core release or
profile closure can pass until the platform transcripts and terminal
reconciliation are complete.

The Windows-hosted plan uses the completed Windows baseline and exact shared
Core producer dependencies for development. It does not execute or certify
the Linux baseline. Detailed delivery owners are recorded in
`docs/worksets/WINDOWS_HOSTED_EXECUTION_PLAN_2026-09-21.md` and the bead graph.

## BLK-COM-EVENT-BYREF-001 — Native COM event writeback requires synchronous reentry

Status: open, design and implementation work available
Owner: Windows workset WIN-5
Clauses: `COM-EVENT-001`, `WIN-PLAN-001`

Impact:

- current queue-backed COM callback transport returns before the VBA handler runs;
- source-owned `VT_BYREF` arguments cannot be mutated before Invoke returns;
- Excel-style cancellable events such as `Workbook.BeforeClose(Cancel)` are not conforming;
- VM3/JIT COM-event profile cannot close.

Unblocking outcome:

- authoritative typed source-interface/parameter metadata;
- synchronous scoped callback into the owning project session;
- handler mutation copied to raw `VT_BYREF` storage before return;
- explicit same-thread/cross-apartment/out-of-proc policy and evidence;
- success/error/reentry/lifecycle cleanup under both backends.

Delivery path: shared VM3 reentry remains `bd-59co.2.7.4`; shared JIT reentry
is `bd-59co.2.7.7`; Windows apartment/reentry policy is `bd-59co.3.3.9`;
incoming event execution remains `bd-59co.3.6.2` and `bd-59co.3.6.3`.
These are open implementation owners, not evidence that synchronous
native-event parity is already achieved.

## BLK-WINDOWS-CERT-001 — Mandatory Windows/Office certification environments

Status: provisioning prerequisite; not yet demonstrated for the current stack
Owner: Windows workset WIN-14 and core workset CORE-9

Required:

- pinned supported Windows x64 environment with actual 64-bit Excel;
- x64 native/COM fixture runners;
- controlled in-proc/out-of-proc COM fixtures;
- non-default locale profile;
- owned-process UIA/VBE compile-oracle automation and cleanup.

Unblocking outcome: checked-in environment manifest and reproducible current-stack x64 runs with artifact/build/fixture hashes. Non-x64 Windows targets are outside the accepted profile and are not blockers.

## Historical blocker migration

Resolved and superseded March-July blocker narration was removed from this current register. Its provenance remains in git history and evidence artifacts. Valid unfinished residuals must be represented in a current workset/matrix row; a historical blocker ID does not remain active merely because an old document cites it.
