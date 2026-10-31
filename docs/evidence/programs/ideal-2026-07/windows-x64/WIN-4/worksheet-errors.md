# Idle worksheet errors and untyped Range arithmetic

2026-10-31. Excel 16.0 build 20527, 64-bit, Windows development/oracle host.
Source: [worksheet-errors.bas](worksheet-errors.bas).
Runner: `scripts/run-worksheet-errors-oracle.ps1` with the existing process-owned
supervisor, explicit VBE Debug Compile command 578 and PID-scoped UIA guardian.

The worker hides the VBE after compile, sets workbook formulas from COM and calls
CalculateFullRebuild while no VBA macro is active. This differs from the earlier
run-5 experiment calculating inside RunProbe. No security/Trust Center/registry
setting changes occur. AutomationSecurity is observed, not assigned. Only owned
Excel processes are cleaned up; the accepted receipt reports owned-process-zero.

Accepted observation: `calcvba6-worksheet-text-verified`, compile ok, run ok, exact
expected tuple, healthy guardian, no compile/runtime modal observations, no owned
process residue. Input carriers are retained: A1 System.Double 41, A2
System.String "3" with text format, A3 null/blank, A4 System.String "bad". The
runner asserts the numeric text and blank before calculation. This is development
observation, not release certification.

| Formula/probe | Inputs | Result |
|---|---|---|
| RaiseMismatch | unhandled Err.Raise 13 | #VALUE!, no runtime dialog |
| DivideByZero(0) | unhandled VBA division by zero | #VALUE!, no runtime dialog |
| UntypedPlus(A1) | A1 = numeric 41 | 42 |
| UntypedPlus(41) | numeric literal | 42 |
| UntypedTimes(A1) | A1 = numeric 41 | 82 |
| UntypedIdentity(A1) | same reference | Range:5 |
| UntypedPlus(A2) | A2 = text "3" | 4 |
| UntypedPlus(A3) | blank cell | 1 |
| UntypedPlus(A4) | text "bad" | #VALUE! |
| UntypedPlus(A1:A2) | multi-cell reference | #VALUE! |
| DoubleOverflowCode | handled variable Double multiplication overflow | VBA Err.Number 6 |
| DoubleOverflowCompare | Resume Next, Err.Clear, constant multiplication | error 6; x=x True; x>0 True; stored bits 7FF0000000000000 (+infinity) |

All worksheet error cells expose Value2 `-2146826273` (Automation 0x800A07DF,
worksheet error 2015) through COM, with rendered text #VALUE!. Error numbers
13/11 are unhandled VBA exceptions, not returned CVErr values; the cell error
does not preserve those VBA numbers. Existing run-5 CVErr(2042) evidence remains
separate.

The argument remains a Range object (TypeName), while arithmetic obtains its
default value. DNA Calc glue/object-library work should provide that default
member/coercion route and map unhandled runtime errors separately from CVErr.
Both generic object-default arithmetic resolution and Double Resume Next
fault/store ordering remain engine gaps tracked by
`bd-59co.2.4.3`; replacing the object with a scalar at entry would lose TypeName
and IsObject behavior. No DNA Calc implementation was changed and no general Excel object-model claim
is made. Exact error information for other hosts and broader Range shapes remain
their owning work.

The first attempt observed the same ten worksheet results but emitted an extra
ClearContents return value into the worker result stream; the supervisor rejected
that envelope. It receives no passing credit. The fixed runner discards that
return value. A second accepted run confirmed the ten rows; the final accepted
run adds the Double-overflow raw bits. Intermediate descriptor/source/envelope
failures receive no passing credit. The reused worker label
`qualified-entry-returned` describes the substituted COM worksheet action in
this runner; no RunProbe macro executes. Raw accepted receipts/source hashes
are retained in the run-6 archive linked from the CORE-7 evidence.
