# Plain VBA XLL example

Build `PlainXll.basproj` as `NativeLibrary`, then copy `PlainXll.dll` to
`PlainXll.xll`. Load it with Excel's `Application.RegisterXLL`.
`=CalcVbaAdd(2.25,3.5)` returns 5.75; `=CalcVbaTwice(4.5)` returns 9.

`PlainXll.bas` owns the x64 XLOPER12 UDT, registration table, callback through
`DispCallFunc`, conversions, and `xlAutoOpen`/`xlAutoClose`/`xlAutoFree12`.
The compiler uses ordinary pointer, Declare, UDT and module-state operations.
Registration omits the thread-safe `$` flag. Calls belong on the owning thread.

This starter converts numbers and returns an Excel value error for other
worksheet inputs. Registration strings use counted UTF-16 buffers. Variant,
array, missing/nil and reference conversions remain to be added to this library.

`scripts/run-native-xll-clients.ps1` compares C and VM3/JIT Declare clients.
Pass an existing `-OracleSession` created by
`scripts/run-native-export-oracle-session.ps1` to also compile the Excel VBA
client, load the XLL, check formulas and release counts, and measure call costs.
The C harness accepts a third argument `bench` for QPC batch measurements.
Excel measurements include VBA Declare or Application.Run overhead; concurrent
work affects the results. The general Declare bridge currently builds its FFI
argument descriptors for each call.
