# GetObject and owned Excel workbook values

2026-10-31. Bounded delivery: `bd-59co.3.4.9`. WCC-LATE-OUTPROC-ERROR
remains **in-progress**; wider residual owner `bd-59co.3.4` is unchanged.

JIT GetObject now admits one/two positional pathname/class operands, including
an omitted pathname. Omission is materialized as Empty, distinct from a present
empty string, through the existing HAL. It does not gain a VM fallback.
Portable tests compare both engines for running-instance/file failures, invalid
empty arguments and new-instance activation. A real Scripting.Dictionary test
executes GetObject("", className), Add, Item and Count on both engines and checks
actual IDispatch traffic, full Err and zero runtime handle drift.

The optional owned-workbook test first requires a live supervisor-supplied
PID/window, then binds one unique bootstrap path and checks Application.Hwnd
again before any writes. It has no explicit Excel launch/quit operation; the
supervisor must hold the workbook through the test. It reads/writes A1 (42, 19, restored 42), reads
Range.Address and Application.Version, then checks a two-dimensional Value2
matrix containing a number, text, Boolean and worksheet error. Every cell's
value/type and all four bounds are checked, rather than comparing only an
opaque array tag. Error values are also written back to B2 and reread.

The first successful owned Excel baseline returned:

```text
42|A1|1|alpha|True|Error 2042|1|2|1|2
```

VM3/JIT initially agreed with each other but returned the raw error payload
0x800A07FA. The native VARIANT reader now maps the standard VBA error facility
to its logical error number; the writer applies the inverse mapping. Other
HRESULTs remain unchanged, including DISP_E_PARAMNOTFOUND. Known wire vectors
cover all seven classic worksheet errors, standard-number bounds and unrelated
HRESULTs. This follows Microsoft's [DLL/Variant error encoding documentation](https://learn.microsoft.com/en-us/office/client-developer/excel/how-to-access-dlls-in-excel)
and the [MS-VBAL Error domain](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/c86480b2-aef2-4488-b177-f55e13cc51f2).

After repair, the exact ignored test passed on both backends against a newly
owned Excel 16.0 build 20527 Safe Mode process. VBE Debug -> Compile command 578
succeeded before Application.Run; the runtime guardian remained armed while
the VBA fixture yielded to the external reader. The Excel result above passed
again after the engine writes. The supervisor recorded owned-process zero.
Macro security was observed (1), never changed; Trust Center and registry
settings were untouched. Another pre-existing Excel process was left alone.

Commands:

```powershell
cargo test -p oxvba-com standard_error_numbers_use_automation_scode
cargo test -p oxvba-differential --test getobject_vm3 --test jit_windows_vm3_parity
# With path/HWND/PID supplied by a live owned supervisor holding the workbook:
cargo test -p oxvba-differential --test jit_windows_vm3_parity getobject_reads_owned_excel_workbook_values_on_both_backends -- --ignored --exact --nocapture
```

The retained owned run and external test output are under
`getobject-oracle-2026-10-31/`. Controller copies override only the diagnostic
source/expected value; worker, guardian, sealed descriptors and cleanup remain
the repository implementation. An earlier controller used the wrong entry
name and failed; the valid baseline then exposed the error conversion defect.
Those failed attempts remain outside the repository in the original run folder.

Broader GetObject coercions/named arguments, locale/error/reentry families,
identity/session producers, other Office hosts and pinned Windows certification
remain `.3.4.4/.8` and their existing producers. No formal model proves these
native wire/COM paths. This test does not implement Range objects in the portable
UDF facade or establish the complete Excel object model.

Final workspace verification passed: 3,643 tests, 179 ignored, strict Clippy,
formatting and governance, with default test concurrency. The final meta-check
used `-Fast -NoArtifacts -SkipPathStability`; the earlier complete default run
also passed the path-stability checks. No golden expectations were changed.
External log: `C:/Work/DnaCalc/codex-runs/calcvba-final-meta.log`, SHA-256
`eedcea3dc6f3e0202abc56deafa4cb216c505058af3da893f7e9877cade3724b`.

Fresh-eyes review added the live PID/window precondition. Its final exact test
passed again in owned run `calcvba-getobject-4-20261031`, with the same Excel
result and owned-process-zero cleanup. A stale-window probe deliberately failed
before GetObject; its retained failure transcript is an expected refusal. The
normal workspace build, ten interpreter-only tests and fresh wasm/V8 probes
also passed after the final source changes.

Implementation checkpoint: d6648120891fbffe829468c9c5bf4926dffcdb8b
