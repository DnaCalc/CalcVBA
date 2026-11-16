# Microsoft Open Specifications used by CalcVBA

Copies of the Microsoft specifications that CalcVBA implements or reads, each as the original
download (`.docx` and the current `.pdf`) plus an extracted plain-text version for searching
and citing. They were copied from the DnaCalc Foundation mirror
(`Foundation/reference/downloads`, downloaded 2026-03-23 to 2026-06-24, hashes in
`Foundation/reference/index.csv`). The documents are Microsoft's, covered by the Intellectual
Property Rights Notice at the start of each; they are kept here to develop and document an
implementation, as that notice allows.

| Spec | What it covers | Used by |
|---|---|---|
| [MS-VBAL](MS-VBAL/) | VBA Language Specification: lexical rules, declarations, types and conversions (Let-coercion), expressions and operators, statements, procedure invocation, the standard library (§6) | the binder, OxIR elaboration, VM3, the JIT, the runtime library |
| [MS-OVBA](MS-OVBA/) | Office VBA File Format: the project storage in `vbaProject.bin` (dir stream, module streams, compression, project info and references) | `calcvba-office` (project loading and writing) |
| [MS-OAUT](MS-OAUT/) | OLE Automation Protocol: VARIANT, SAFEARRAY, BSTR, IDispatch, type libraries (ITypeInfo, ITypeLib) | `calcvba-runtime` value layout, `calcvba-com`, early and late binding |
| [MS-DTYP](MS-DTYP/) | Windows Data Types: base types, HRESULT, GUID, DECIMAL, CURRENCY, LCID | the runtime value types, COM interop |
| [MS-OFORMS](MS-OFORMS/) | Office Forms Binary File Formats: UserForm and control storage (`.frx` and the form streams) | `calcvba-office` form import/export |
| [MS-OFFMACRO](MS-OFFMACRO/) | Office Macro-Enabled File Format: how a VBA project is embedded in an Office Open XML package | `calcvba-office` workbook-container I/O |
| [MS-OFFMACRO2](MS-OFFMACRO2/) | Office Macro-Enabled File Format Version 2: the same parts, specified against ISO/IEC 29500-1:2016 | `calcvba-office` workbook-container I/O |

`learn-pages/` holds three learn.microsoft.com index pages of the VBA language reference
(the function list and the language reference overview), as mirrored. Each spec folder's
`landing-page.md` is the learn.microsoft.com page that links to the downloads.

## Files in each spec folder

- `[MS-XXXX]-<revision>.docx`: the original, the source of the text extraction.
- `[MS-XXXX].pdf`: the current PDF of the same document, for reading.
- `[MS-XXXX]-<revision>.txt`: one line per non-empty paragraph of the docx. Line numbers in
  CalcVBA documents (for example "MS-VBAL §5.4.4 ll. 2793-2794" in
  [OXIR_VM3_ERROR_MODEL.md](../../OXIR_VM3_ERROR_MODEL.md)) refer to these files.
- `[MS-XXXX]-<revision>.images.tsv`: some rules are pictures (formulas, diagrams) whose
  meaning is only in their alt text, which the `.txt` drops. This lists each one with the
  `.txt` line it sits on or follows.

The MS-VBAL, MS-OAUT and MS-DTYP text files are identical to the earlier extractions in
`Foundation/.tmp/spec_scan` (apart from LF line ends); the others are new. To regenerate:

```bash
python -I extract_docx_text.py "MS-VBAL/[MS-VBAL]-250520.docx" "MS-VBAL/[MS-VBAL]-250520.txt" "MS-VBAL/[MS-VBAL]-250520.images.tsv"
```

The spec-pack extraction in `Foundation/reference/runs/20260301-ms-vbal-pass07` was made from
the PDF, has broken text, and is not copied. MS-CFB (Compound File Binary format, under
MS-OVBA) is not in the mirror yet.

Where Excel's VBA (VBE7) differs from these documents, CalcVBA follows Excel; the differences
are measured under `conformance/excel-measurements/`.
