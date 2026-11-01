# Imported native typed slots — run 9

2026-11-01. Scoped delivery `bd-59co.3.5.7`; COM-CLIENT-001 and WIN-PLAN-001.
WCC-PLAN-EARLY and Windows certification remain **in-progress**.

Unsafe ComHal::prepare_typed_object requires trusted exact compiler descriptors;
it checks capability/policy and the current apartment. Bounds/IID checks cannot
prove an arbitrary forged in-bounds signature matches native code.
The COM bridge validates compiler-retained metadata, the native binding's creating
thread, exact IID via QueryInterface, slot bounds and proxy alignment. It creates
a same-thread host table whose callbacks retain that native interface independently
of the HAL binding. A native pointer is never cast to portable ObjectHeader.
The owning VM3/native JIT paths consume the same typed member plans/table ABI;
there is no source rewrite, server-name/member-name branch, Invoke retry or VM fallback.
The API-only crate still links no COM/compiler/backend implementation.

This bounded generic native family admits HRESULT getters returning Long, Double,
VARIANT_BOOL or BSTR, void methods, and methods/setters with one Long input and scalar/BSTR/void results. A trailing
hidden LCID is supplied as LOCALE_NEUTRAL. Optional/ByRef/Variant/interface/array/
record signatures remain refused by this typed family before body effects; the
existing canonical native transport retains its separately tested broader paths.
All inputs/results stay in their exact scalar/BSTR carriers; no Variant argument
vector, Variant execution frame or libffi call is used by this native family.

The real registered Scripting.Dictionary case runs on both runtimes: CompareMode
put/get, Count and RemoveAll produce four vtable successes and zero Invoke calls.
An invalid mode reaches the VBA handler with the server's IErrorInfo description.
Separate metadata-bound, convention, source-typekind, wrong-IID and wrong-thread
refusals precede native member entry. Releasing the original HAL binding before
invocation proves the table's independently retained QI receiver remains usable.
A richer portable fault fixture checks Description/Source/HelpFile/HelpContext
across one-operation slices and unhandled return on both runtimes.

Real Scripting.TextStream.Read(Long) returns BSTR directly: two one-operation-sliced
invocations consume abc then def from an exclusively created temporary input file.
Close and final release allow removal of that exact owned file. Native reads occur
once each, without extra Invoke calls. Its FileSystemObject.OpenTextFile bootstrap
also repairs an existing canonical transport check: a VBA lvalue passed to a known
ByVal native parameter does not require nonexistent ByRef writebacks. Missing
metadata and actual ByRef wire shapes retain the conservative writeback check.
The bootstrap forces DispatchOnly so both the Invoke and preferred-vtable routes
respect declared ByVal input; typed reads still use native slots. The broader
late-argument owner bd-59co.3.4.6 remains open.

The PID-scoped existing Excel supervisor arms UIA before explicit VBE compile 578.
The imported Excel Application Version/Build getters exercise real out-of-process
PSOA vtables and hidden LCID, returning 16.0 / 20527 on VM3/native JIT, with no
additional Invoke calls. The initial run used owned PID 51012 and cleaned it up.
The final receipt calcvba9-typed-native-excel-complete additionally covers Double Height=786 and Boolean Visible=True on owned PID 27632, also cleaned up. A registration-free dual fixture independently checks zero native/IErrorInfo references after cancel/return/drop and zero Invoke/name/type-info reads. Raw
metadata, values, executable hash and ownership/cleanup are retained in the final
run-9 value archive. These are development observations, not Office certification.
PID 3144 is never owned; no registration/security/Trust Center/HKLM changes occur.

Broader native carriers, writable ByRef alias/writeback, general proxy/deployment
and full COM reentry remain under bd-59co.3.5.2/.3.5.3 and their existing producers.
The installed Lean abstractions do not prove Windows ABI, apartment behavior,
IErrorInfo lifetimes or generated native code. Those limits remain in the formal
backlog. See CORE-7/typed-value-slots-run9.md and run9-values.json/.zip for the
source-bound repository and interpreter-only ownership checks.

Final source receipt calcvba9-final-native-reviewed uses owned Excel PID 36320,
returns the same four typed values on both backends, and records owned-process-zero.
The final repository check passes 3738 tests in 196 suites (181 ignored; four
registry-writing record fixtures excluded without credit). Scope bd-59co.3.5.7
is delivered; the broad capability/certification rows retain their open producers.
