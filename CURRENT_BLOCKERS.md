# Current Blockers

Date: 2026-09-21
Status: current; no all-path blocker
Authority: blockers only; ordinary unfinished capability work belongs in the three current worksets and canonical matrices

## BLK-EXCEL-COMPILE-20261031 — Negative compile oracle lacks reliable pre-dismiss selection

Status: open; runtime and other delivery work can continue. Owner: `bd-59co.3.15.6`.

The development host starts normal Excel with TestCrashGuard add-in crash/recovery prompts. Explicit `/safe` permits the scalar runtime probe without changing add-ins, Trust Center, macro security or registry settings. `calcvba-scalar-verified-20261031` compiled through VBE command 578 and returned the expected scalar COM/Declare/Add2 result, with owned cleanup.

The deliberately broken-module probe still fails to retain a reliable UIA selected token/full line before dismissal. Some attempts also expose guardian-ledger identity failures. No compile-diagnostic certification is claimed. Raw attempts are under `artifacts/windows-x64/excel-vba-oracle/calcvba-compile-*`; the retained run summary identifies the exact limitation.

Next: repair the PID-scoped VBE selection/guardian path under WIN-14, then run `scripts/run-excel-vba-oracle.ps1 -EnvironmentId win-x64-dev-oracle-2026-07 -NoMatrixUpdate -DiagnosticCaseId compile-failure -SafeMode` and require the modal text, observed token `MissingOracleSymbol`, complete selected source line and owned cleanup. Do not lower the oracle's evidence requirements or modify Office security to obtain a pass. The pinned clean certification environment remains a separate existing prerequisite.

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
