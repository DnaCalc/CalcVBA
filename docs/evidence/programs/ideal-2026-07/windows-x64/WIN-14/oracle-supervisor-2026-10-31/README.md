# Owned Excel/VBE supervisor development observation

2026-10-31. Support owner `bd-59co.3.15.6` is closed for the scope below. The default five cases and separate
unhandled-runtime modal diagnostic pass on the characterized development host.
This resolves `BLK-EXCEL-COMPILE-20261031` for that host. Pinned certification,
Windows capability rows and release evidence remain with their existing owners.

The owner authorized these runs as the Windows follow-up in CalcVba run 2.
That instruction supersedes the earlier evidence document's stopped/no-retry
state. No Trust Center, macro-security, add-in or registry setting was changed.
Safe Mode avoids the existing add-in crash prompts. AutomationSecurity was read
as 1; the worker never assigns it. The pre-existing Excel PID 3144, with its
original start time, remained running after all owned instances were cleaned.

## Result and commands

Office 16.0 build 20527, 64-bit Windows development/oracle environment
`win-x64-dev-oracle-2026-07`. Compile uses visible VBE command 578. The guardian
is armed before compile and runtime invocation; Application.Run never decides
compile success.

```powershell
./scripts/run-excel-vba-oracle.ps1 -Suite HarnessSelfTest -EnvironmentId win-x64-dev-oracle-2026-07 -NoMatrixUpdate -SafeMode -RunId calcvba2_full_native5
./scripts/run-excel-vba-oracle.ps1 -Suite HarnessSelfTest -EnvironmentId win-x64-dev-oracle-2026-07 -NoMatrixUpdate -SafeMode -DiagnosticCaseId runtime-unhandled-modal -RunId calcvba2_runtime_modal1
./scripts/test-excel-vba-oracle.ps1
```

Both live commands used separate output roots under `C:/Work/DnaCalc/codex-runs`
and bounded supervisor timeouts, recorded in their raw plans. Reproduction must
use new run IDs and output directories; existing run claims reject reuse.

| Case | Compile | Runtime | Result |
|---|---|---|---|
| success | ok | ok | pass |
| compile-failure | compile-error | not-run | pass |
| ambiguous-macro-failure | ok | missing-macro | pass |
| intrinsic-shadow | compile-error | not-run | pass |
| runtime-full-err | ok | runtime-err-captured | pass |
| runtime-unhandled-modal | ok | runtime-error-modal | pass |

Every case reports `owned-process-zero`, with no cleanup authority errors.
The full Err observation is number 513, source `OracleRuntimeSource`, description
`oracle-runtime-error`, help file `oracle-help.chm`, help context 42 and Erl 100.
The unhandled modal is captured before its owned End button is invoked.

## Repairs and rejection checks

The native fallback handles Office dialog fragments that expose no UIA button
descendants. Both action paths require the captured owned HWND/PID, unchanged
title and complete normalized message, and exactly one matching visible enabled
native Button beneath that dialog. Ownership, parent, class, label, enabled and
visible state are rechecked before action. Success requires dialog disappearance.
Unknown/security prompts still block. Capture-before-dismiss and the linked
successful dismissal remain mandatory for accepted compile errors.

The isolated Windows tests compile the production native class and show only
their own helper dialogs. Wrong PID/title/message/button, disabled or duplicate
OK buttons, and appended message content cannot authorize a click. A valid
observation closes its own dialog. No Excel process is started by these tests.

Execute can return before Excel creates a compile-error modal. The worker now
requires the disabled Compile command or a captured modal with linked successful
dismissal and a later heartbeat, in addition to existing operation coverage.
Pure tests reject heartbeat-only, missing/failed dismissal and premature flush.
Atomic control replacement uses File.Replace with NullString.Value; readers
share deletion and retain one complete old file instance. A production-publisher
test holds that reader open and checks both old and replacement JSON sequences.
A bounded startup wait prevents publishing a blank guardian executable path.

Fresh Excel rejected `ByVal Fix As Double` at the declaration, with token `Fix`
and `Expected: identifier`. The previous fixture incorrectly expected the later
call line. Its sealed expected line now reflects this observation; no broader
compile-evidence requirement was weakened.

Rejected intermediate runs remain failures. Their causes and result hashes are
in the manifest. In particular, the passing intermediate native4 run did not
justify keeping File.Move overwrite: the held-reader test rejects that operation
on Windows. The final native5 run uses the corrected File.Replace implementation.
One concurrent offline test attempt observed our live Excel process changing;
the final complete contract suite was rerun with our live runs stopped and passes.

## Retained evidence

`raw.zip` preserves all 98 original files from the two accepted final runs,
including results, supervisor transcripts, immutable guardian events, ownership,
source, selection descriptors, environment and bootstrap bytes. Text stays inside
the archive to preserve its exact original hashes under the repository LF-only
text policy. `manifest.json` records every member's length and SHA-256, the
archive hash, original paths, implementation hashes and rejected-run summaries.
Archive bytes were checked against every member hash after creation.

The complete `scripts/meta-check.ps1 -NoArtifacts` passes: governance, path
stability, format, strict all-target Clippy, 3,643 tests with 179 ignored, and
workspace check. No modeled runtime behavior changed in this support repair.

Fresh-eyes review inspected both action paths, hostile tests, asynchronous flush,
publication, ownership and the accepted raw observations. The wider Windows
profile remains in-progress; next work is declaration legality, structural and
native-string aliases, foreign Boolean payload 1, early COM and callbacks.
