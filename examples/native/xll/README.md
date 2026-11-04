# Plain VBA XLL example

Build `PlainXll.basproj` as `NativeLibrary`, then copy `PlainXll.dll` to
`PlainXll.xll`. Load it with Excel's `Application.RegisterXLL`.
`=CalcVbaAdd(2.25,3.5)` returns 5.75; `=CalcVbaTwice(4.5)` returns 9.

`XlValues.bas` supplies the x64 XLOPER12 UDT and conversions. `PlainXll.bas`
owns the registration table, callback through `DispCallFunc`, and
`xlAutoOpen`/`xlAutoClose`/`xlAutoFree12`.
The compiler uses ordinary pointer, Declare, UDT and module-state operations.
Registration omits the thread-safe `$` flag. Calls belong on the owning thread.

`XlFromVariant` and `XlToVariant` handle numeric Variant types, counted UTF-16
strings (including embedded NULs), Boolean, worksheet errors, missing, nil and
one- or two-dimensional arrays. One-dimensional arrays become a single row;
returned matrices are one-based Variant arrays. Null, unsupported values,
nested arrays and arrays with more than two dimensions become `#VALUE!`.

References use a reserved Variant vector: `"#XlReference12"`, the reference
type (8 or 1024), the exact LongLong sheet ID, area count, then four zero-based
coordinates per area. `XlReferenceCount` and `XlReferenceArea` read references.
`CalcVbaCoerce`, `CalcVbaCallerRow` and `CalcVbaName` demonstrate host callbacks.

Treat each value made by `XlFromVariant` or `XlString` as one owner. Release it
with `XlFreeValue`, or transfer it to `XlReturn`; Excel releases the returned
pointer through `xlAutoFree12`. `XlToVariant` borrows its input. `XlOwnedBuffers`
and `OutstandingResults` expose the allocation counts for checks.

`scripts/run-native-xll-clients.ps1` compares C and VM3/JIT Declare clients.
Pass an existing `-OracleSession` created by
`scripts/run-native-export-oracle-session.ps1` to also compile the Excel VBA
clients and the original library source, load the XLL, check 25 worksheet calls,
close its registrations, check release counts, and measure call costs.
`-DnaCalcPath <dna-calc-desktop.exe>` runs the same cases in DNA Calc's XLL
host. Three blank-cell cases remain Excel-only because the current desktop
harness reports `MissingInput` before invoking the add-in. The other 22 cases
compare individually, including matrices, references, missing values and errors.
The C harness accepts a third argument `bench` for QPC batch measurements.
Excel measurements include VBA Declare or Application.Run overhead; concurrent
work affects the results. The general Declare bridge currently builds its FFI
argument descriptors for each call.
