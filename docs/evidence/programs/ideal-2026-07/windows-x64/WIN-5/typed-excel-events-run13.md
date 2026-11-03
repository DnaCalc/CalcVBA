# Typed Excel event reentry and cancellation, run 13

Owner: bd-59co.3.6.9. Clause: COM-EVENT-001.
The bounded host-registered source path matches owned Excel VBA and both owning
VM3/native JIT. WCE-PLAN-INCOMING remains planned; this is development evidence.

The compiler/host selects exact imported Worksheet Change (DISPID 1545, object
input, source IID 00024411) and Workbook BeforeClose (DISPID 1546, Boolean ByRef,
source IID 00024412) descriptors before execution. No per-event name/type lookup
is used. The native root is the validated owned NativeOM window. Outgoing members
use the same native/prebound tables as the ordinary surface proof.

Shared excel_events.bas handler bodies compile unchanged in visible VBE Debug
Compile, driven there by Sheet/ThisWorkbook wrappers whose full injected source
and hashes are retained. Oracle wrappers are inactive during engine execution;
the host advises the same real Worksheet/Workbook source interfaces directly.

Exact VBA/VM3/native JIT results: off=1:True:20;on=2:True:20;
close=1:oracle-bootstrap.xlsx. With EnableEvents disabled around the neighbor
write, only one Change handler runs. With events enabled, the neighbor write
re-enters synchronously and the count is two. BeforeClose changes incoming Cancel
to True before Close returns; the owned workbook remains usable. Explicit
Unadvise prevents further deliveries. Outgoing transport counts (native, Invoke):
suppressed Change (3,8), nested Change (1,10), canceled Close (2,0).

SynchronousEventSinkConfig admits retained object inputs and at most one Boolean
ByRef. Incoming object arguments own references through the callback. Raw Cancel
pointers stay inside Invoke; only a Boolean Cell is exposed, then copied as the
canonical VARIANT_BOOL before returning. Same-thread admission, exact named and
positional ordering, panic containment, and a retained Invoke lease protect the
sink if the callback releases its subscription owner. TypedEventSubscription
retains the source and owns one Unadvise/Release.

Session::call_values_with_writebacks is an additive synchronous host API. Primitive
root ByRef Long/Double/Boolean/String values copy back on success and VBA errors.
ByVal inputs stay private. Structural/object ByRef roots are refused before body
effects; no root object lifetime is extended by scalar snapshots. Portable tests
on both runtimes cover error copyback, BSTR ownership and ByVal isolation; native
sink tests cover errors, wrong-thread/shape rejection, panic and reentrant owner
release. Existing legacy queued subscriptions keep their separate behavior.

Commands: cargo test -p oxvba-embed --test typed_incoming_byref --test typed_excel_events;
cargo test -p oxvba-com --lib windows_connection_point;
scripts/run-typed-events-oracle.ps1; scripts/meta-check.ps1 -Fast -NoArtifacts -NoRegistry.
Source/executable/observation hashes and owned-process cleanup: run13-excel-events.json/.zip.

Automatic document/WithEvents registration, structural ByRef event arguments,
cross-apartment delivery, full handler-error/EXCEPINFO policy, concrete native
formal refinement and pinned certification remain with bd-59co.3.6.2-.8 and the
formal backlog. Handler failures currently return E_FAIL after Boolean copyback;
this bounded path does not claim exact Excel handler-error parity.

Repository check: 3,776 passed in 206 suites; 183 ignored. Strict Clippy,
format, governance and path checks pass. Four registry writers are excluded with
no credit. Installed Lean 4.29.1 lifecycle/reentry abstractions pass; no concrete
native formal or Miri credit. Fresh review corrected duplicate interface casing
and added a sink Invoke lease for reentrant unsubscribe; relevant proofs reran.
