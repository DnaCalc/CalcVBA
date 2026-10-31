# Scalar late COM client development evidence

Measured 2026-10-31 local time on Windows x64. Owner: `bd-59co.3.4.2`.
The canonical WCC-PLAN-LATE row remains **in-progress**. This is the bounded
scalar delivery, not Windows profile certification.

`cargo test -p oxvba-differential --test jit_windows_vm3_parity` executes actual
Windows COM under both VM3 and JIT. Interactive policy explicitly selects
IDispatch, so the Dictionary proof cannot silently take a vtable shortcut.
The controlled OxVba.TestDispatch fixture checks activation, SetValue, Value
get, Ping, Count and an unhandled exception including its full Err state.
The ordinary Scripting.Dictionary ProgID is supplied through a VBA variable;
Add, Item put/get and Count return 42 and 1. Matching declines fail.
Successful native IDispatch counts and zero current-thread handle deltas are
asserted after teardown. The existing transport counter counts successful
invocations, so the exception case is not presented as a successful call.

The shared native VARIANT adapter now allocates copied BSTRs with the OLE
allocator directly. VariantClear owns those copies; runtime BSTR accounting
therefore no longer counts a copy that OLE later frees outside the runtime.
The copy preserves exact byte length, including odd-byte BSTRs.

The separate Excel oracle diagnostic `windows-scalar-interop` compiled through
VBE Debug -> Compile and returned `42|42|1234|42|1|42`. Retained source, result,
environment and guardian events are in `scalar-oracle-2026-10-31/`. Excel 16.0,
build 20527, Windows 64-bit, Safe Mode requested, observed AutomationSecurity=1.
Security settings were only read. Normal startup encountered an existing add-in
crash/recovery prompt; the owned Safe Mode process avoided that add-in. The
recovery files were left untouched. Only the run's recorded PIDs were cleaned.
The historical environment routing ID is not a claim that the July pinned
Office fingerprint still matches this machine.

Residuals: named/omitted/ByRef/array/record/default-member/GetObject and wider
activation behavior remain `bd-59co.3.4.3/.4/.5`; Excel object-model certification
remains WIN-14. The negative compile diagnostic still fails closed when UIA
cannot read both selected token and full code line (`BLK-EXCEL-COMPILE-20261031`,
`bd-59co.3.15.6`). The positive runtime probe does not settle that residual.

Reproduce the Excel probe without changing Office security:

```powershell
pwsh -NoProfile -STA -File scripts/run-excel-vba-oracle.ps1 -EnvironmentId win-x64-dev-oracle-2026-07 -NoMatrixUpdate -DiagnosticCaseId windows-scalar-interop -SafeMode -RunId scalar-review -TimeoutSeconds 120
```
