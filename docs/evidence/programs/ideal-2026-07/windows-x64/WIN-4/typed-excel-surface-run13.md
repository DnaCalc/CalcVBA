# Typed Excel surface and format provenance, run 13

Owners: bd-59co.3.5.10 (format classification), bd-59co.3.5.11 (ordinary surface).
Clauses: COM-CLIENT-001, HOST-SESSION-001, SYS-DUAL-001.
Status: format classification resolved for explicit invocation locale; ordinary
surface remains in-progress while collection enumeration and additional shapes
are being delivered. WCC-EARLY-COMPLEX and pinned certification stay planned.

The shared excel_surface.bas compiles in VBE command 578 under the owned PID/HWND
supervisor, then runs unchanged in VBA and on owning VM3/native JIT. The native
entry is AccessibleObjectFromWindow on the exact recorded owned EXCEL7 child.
Window.Application uses its compiler-prebound DISPID, followed by native dual
Application/Workbook/Worksheet/Sheets calls. No dynamic setup macro or runtime
member/type-info lookup supplies the engine's root or typed calls.

Delivered observations: typed Range Cells indexing via a declared Range local,
Offset, Resize, Rows/Columns/Count, relative/absolute Address, Formula,
FormulaR1C1, NumberFormat, Text, Application.Calculate, ScreenUpdating,
EnableEvents and Evaluate. Each helper asserts its exact native/Invoke count.
The pure-dispatch fixture rejects/counts every name/type-info read and additionally
checks a generic native object result assigned to a declared interface. Nominal
stores retain a QI'd destination view; the source table/identity remain intact.
Unknown IDispatch results remain opaque until that explicit typed assignment.

A real metadata defect was exposed by using several Excel types in one module:
HREFTYPE is a DWORD in TYPEDESC's pointer-sized union. Reading the unused upper
bytes as usize and falling back to zero imported Range.Cells as Adjustments.
The loader now reads the actual union member, including valid token zero. A dirty
upper-byte regression and the combined installed-typelib case retain this proof.
The adjacent COM build test now also asserts that ext_ConnectMode imports as a
Long Automation value rather than the erroneous interface pointer.

## NumberFormat classification

On this Windows user's regional settings, the same setter text $0.00 denotes a
literal dollar under the user-default locale and currency under LCID 1033.
Cross-reading independent native views of the same owned cells establishes:

| Setter LCID | Reader 0x400 | Reader 1033 |
|---|---|---|
| 0x400 | 5:5, $0.00, 1.234567 | 5:5, \$0.00, 1.234567 |
| 1033 | 6:5, R0.00, 1.2346 | 6:5, $0.00, 1.2346 |

The two numbers are VarType(Value):VarType(Value2). Neutral LCID 0 also produced
Double in the retained minimization; changing only the reader did not fix it.
VBA agrees with the 1033 setter: Currency 1.2346 and Double Value2 1.234567.
This resolves the run-12 observation as setter locale provenance, not a lost
Currency carrier or a name-specific format conversion. Microsoft documents VBA's
English/invariant object-model formatting and explicit COM locale behavior:
https://learn.microsoft.com/en-us/visualstudio/vsto/globalization-and-localization-of-excel-solutions

ComHal::prepare_typed_object_with_locale retains an explicit caller-selected LCID
through all returned interface views, prebound Invoke, and hidden native LCID
parameters. The host in this Excel proof selects 1033; existing callers keep their
previous defaults (dispatch 0x400, native hidden LCID 0). No Excel names or format
strings enter the engine policy. Automatic host/program locale selection remains
with the broader WIN-4 owner; default dynamic calls have not gained blanket parity.

VarType and CStr imports now run against their original exact carrier and the
shared value library on both owning backends, sharing the active context for
object defaults. No frame/global borrow spans such a callback.

Commands: cargo test -p oxvba-embed --test typed_excel_surface --test typed_prebound_dispatch
--test typed_native_com; scripts/run-typed-surface-oracle.ps1 -RuntimeTestExecutable
<fresh Cargo artifact>; scripts/meta-check.ps1 -Fast -NoArtifacts -NoRegistry.
Source/executable/observation hashes and cleanup receipts: run13-excel-surface.json/.zip.
Repository check: 3,769 passed in 203 suites, 182 ignored; strict Clippy, format,
path and governance checks green. Four registry-writing tests are excluded without passing credit. Installed Lean
4.29.1 RefCount/TypedContinuation/TypedReentry/ProcedureHandles abstractions pass;
concrete COM/pointer/locale refinement is unproved and receives no formal credit.

Remaining source shapes include genuine Variant-result member chains such as
Cells(2,1).Address on the owning path; a declared Range local exercises the same
Excel members with static plans. Richer signatures, native event policy and pinned
certification remain with the existing WIN-4/CORE-7/WIN-5 owners. No adapter member,
security/Trust Center/registration/HKLM change or DnaCalc-win/XLL work is included.
