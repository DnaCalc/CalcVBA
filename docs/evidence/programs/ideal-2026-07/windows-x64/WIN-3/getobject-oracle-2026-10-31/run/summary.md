# Excel/VBA Oracle Harness Self-Test

- Run: `calcvba-getobject-3-20261031`
- Environment: `win-x64-dev-oracle-2026-07`
- Result: **TARGETED PROBE CAPTURED (case `windows-scalar-interop`; intrinsic diagnostic=True; expectation pass=True)**
- Authority: development/oracle characterization only; noncertifying
- Credit: no canonical matrix, release, certification, or capability credit
- Ownership: newly created Excel HWNDs and guardian processes were sealed by PID, process start, name, and executable; fallback cleanup required the complete identity
- Compile authority: VBE Debug -> Compile VBAProject command ID 578; runtime invocation was never used as a compile check
- Modal safety: the PID-scoped UIA guardian was ready before each compile and runtime invocation

| Case | Compile | Runtime | Pass |
|---|---|---|---|
| windows-scalar-interop | ok | ok | True |

Raw evidence: `results.json`, `transcript.json`, `owned-processes.jsonl`, and per-case guardian/module artifacts in this run directory.
