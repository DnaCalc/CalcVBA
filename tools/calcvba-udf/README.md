# VBA UDF wrapper generator

`calcvba-udf` reads the compiler's bound procedure catalogue and parameter types,
then writes ordinary VBA shims, `registration.tsv`, a `.basproj` containing
`NativeExport` entries, and a native `.xll`. It is a source transformation tool;
the compiler has no Excel registration or worksheet-function special cases.

```powershell
cargo run -p calcvba-udf -- --project MyFunctions.basproj --out generated --prefix Native_
```

The output directory must be new. Use repeated `--udf Module.Function` options
to select functions, or omit them to select all public functions in public
standard modules. `--emit-only` writes the sources without building. The prefix
allows comparison with the original VBA functions in the same workbook.

Scalars are assigned through variables of their declared types. `As Range`
parameters receive real Excel COM Range objects; reference arguments passed to
`As Variant` also remain Range objects. Other Variant arguments preserve
values, arrays, missing arguments, and errors. Optional defaults come from the
bound signature. ParamArray calls expand into ordinary VBA calls with a fixed
maximum number of worksheet arguments: `--paramarray-slots 16` by default, at
most 32 total native arguments. Trailing missing slots are omitted; interior
missing slots are retained. Functions returning Range objects return Value2.
Unhandled errors become `#VALUE!`; returned `CVErr` values retain their codes.

The generated `Application` property obtains Excel's native object model from
the C API window handle. It never activates a new instance. Functions register
for the owning Excel calculation thread. The tool refuses document modules,
forms, events, WithEvents, project references, unsupported signature types, and
duplicate worksheet names before writing output. COM references are retained.
The native compiler reports unsupported bodies without an interpreter fallback.

The VBA support library uses the Windows x64 Excel C API and Automation layouts.
The comparison fixture is in `tests/projects`; real Excel checks must reuse an
owned process and perform VBE Debug → Compile before executing the VBA oracle.
