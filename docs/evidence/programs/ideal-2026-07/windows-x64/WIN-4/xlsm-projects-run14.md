# Saved XLSM macro projects, run 14

Owner: bd-59co.3.5.13. Seven-project comparisons pass on both owning backends.
Broader workbook compatibility remains in-progress.

The harness builds and saves real format-52 XLSM packages, closes them, and reopens
their VBA projects. It reads every standard/class/document module, code name,
unchanged source body, role and ordered type-library reference through public
VBOM. Both owning runtimes receive that actual project. Extra host setup code only
obtains and checks the exact owned workbook identity and its document roots.
The packages retain xl/vbaProject.bin and their original source.

Excel visibly compiles the reopened project using VBE command 578 under the
existing PID-scoped UIA guardian. It runs the reference macro on the saved
baseline. Each backend opens a byte-identical disk copy; the harness verifies its
entire source/reference snapshot, removes native VBA code only in memory to
prevent duplicate native handlers, then runs the unchanged project through
CalcVBA. No edited execution copy is saved over the original package.

Each returned String is compared by UTF-16 units. Every cell in every worksheet's
entire used rectangle is compared by Value2 (including numeric/error carriers),
Formula, displayed Text, NumberFormat, Interior.Color, Font.Bold and MergeCells.
Blank cells inside that rectangle are included. The fixtures write no cell
outside those rectangles. Formula errors and independent format changes are
included, so a matching scalar return alone cannot pass.

| Workbook | Compared cells per backend | Behavior |
| --- | ---: | --- |
| chains.xlsm | 24 | Object/Variant chains, names/defaults, 91/424/438, formulas, styles, formula error |
| events.xlsm | 5 | Automatic documents/WithEvents, nesting, rebind, release/Terminate, cancellation |
| mixed.xlsm | 40 | Unchanged event and member-chain modules in one actual project |
| documents.xlsm | 2 | Document code names through Object/Variant and native Is identity |
| arrays.xlsm | 10 | Fixed rank-2 Variant arrays, bulk Value2 writes/reads, Resize and formula |
| bounds.xlsm | 10 | The same bulk operations with nonzero/negative Long bound expressions |
| application.xlsm | 1 | Application WithEvents workbook argument, document order, identity and ByRef cancellation |

The workbook comparisons found and corrected engine differences: internal
Object-to-ByVal-Variant calls lacked a private copy; untyped document members and
Is did not share retained native identity; multidimensional owning array indexing
transposed off-diagonal cells; typed ByVal object parameters lacked private
Variant unboxing and declared interface conversion. The application event probe
now returns Excel's `1:1:DA:True` on both backends. Owning fixed Variant-array allocation is now
admitted, with checked bounds/product/allocation. The ordinary JIT also accepts a
fixed allocation when constant bound expressions retain a conservative dynamic
shape. Exact positive and negative-bound regressions run on both owning backends.
The host loader also preserves the saved project name rather than inventing one.
The allocation dimension limit follows [Microsoft's VBA Dim reference](https://learn.microsoft.com/office/vba/language/reference/user-interface-help/dim-statement).

Broader compatibility remains in-progress: independent MS-OVBA source loading,
arbitrary existing workbooks/project references/forms, document custom members
through untyped receivers, general ReDim/typed array allocation, small scalar
physical carriers (including unsuffixed negative Integer bound temporaries),
general array-return legacy JIT admission, event startup/reload/errors/apartments,
native dynamic ByRef callback synchronization and pinned certification. The next
general carrier/admission delivery path remains bd-59co.2.9.9; native client scope
remains bd-59co.3.4, and automatic event residuals remain bd-59co.3.6.10/.5.
No adapter members or security/COM-registration changes are involved.

Final source, seven actual XLSM packages, complete project/cell/return snapshots,
native executable hashes, VBE/guardian/cleanup receipts and the repository check
are retained in [the shared manifest](../WIN-5/run14-automatic-workbooks.json)
and [archive](../WIN-5/run14-automatic-workbooks.zip). All 92 used cells per backend
match (184 cell comparisons), and all 14 macro returns match. The application
document handler explicitly clears Cancel; only the Application handler cancels,
so that observation checks its Object-plus-Boolean copyback path independently.
The NoRegistry check passes 3,786 tests in 210 suites, with 186 ignored and four
registry-writing record fixtures excluded without credit. Strict Clippy, format,
path stability and governance pass. The oracle contract self-test also passes.
