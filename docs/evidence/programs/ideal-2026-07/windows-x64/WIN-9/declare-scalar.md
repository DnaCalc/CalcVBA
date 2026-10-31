# Scalar x64 Declare development evidence

Measured 2026-10-31 local time. Owner: `bd-59co.3.10.2`; WNI-PLAN-DECLARE remains
**in-progress**. Command: `cargo test -p oxvba-differential --test jit_windows_vm3_parity`.

The test compiles an owned Rust cdylib in a unique temporary directory and
executes actual Win64 FFI through the shared verified plan under VM3 and JIT.
ScalarAdd(19,23) returns Long 42; ScalarWrite updates a ByRef Long to 73;
SetLastError is captured as Err.LastDllError=1234. Result, full Err, raised
state, successful native transport and zero current-thread runtime handle
deltas are compared. No time-dependent GetTickCount equality remains.

An absent named export raises VBA error 453 on both backends, consistent with
[Microsoft's documented error](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/specified-dll-function-not-found-error-453).
The HAL retains an owned LoadLibrary reference for its bindings and frees it
when the host is disposed. Successful removal of the fixture directory after
both engines drop proves that the fixture DLL is no longer retained.

The separate Excel probe documented in `../WIN-3/late-scalar-client.md` compiles
and executes msvcrt sqrt, oleaut32 VarR8FromI4(ByRef Double) and kernel32
SetLastError, returning the expected scalar/ByRef/LastDllError results. It
measures the VBA target behavior, not an identical fixture binary. The test
fixture proves the corresponding OxVba execution paths.

Strings/buffers/arrays/UDTs/As Any remain `bd-59co.3.10.3`; loader search,
ordinal and policy behavior remain `.4`; all scalar widths and aliased ByRef
arguments remain `.5`. This evidence claims neither those subsets nor native
DLL/EXE output or pinned-environment certification. Non-Windows/x64 JIT Declare
continues to decline explicitly.
