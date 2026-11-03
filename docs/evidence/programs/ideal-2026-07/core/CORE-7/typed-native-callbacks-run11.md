# Synchronous native callbacks into typed VBA - run 11

2026-11-01. Scoped delivery bd-59co.2.9.19 covers the following host callback and
scalar connection-point cases on VM3/native JIT. CORE-JIT-LOWERING remains
in-progress and the Windows event rows remain planned; this is development proof.

calcvba-host-api::with_native_call lends the caller's live typed context for one
synchronous native operation. with_native_callback_context permits one active
same-thread borrow. Private scope guards restore previous context/borrow state
on return and unwinding. The callback reference cannot escape in the generic
result type. Missing scopes, other threads and overlapping borrows reject.
Existing session/generation/64-level recursion and finite-budget checks apply.
Both scalar and descriptor-planned native COM member factories establish this
scope automatically. Native calls complete synchronously within an instruction;
there is no suspension inside a foreign call. Hosts own native error projection.

The registration-free calculator fixture calls a real native function pointer
back into original typed VBA. One-operation outer slices observe shared globals,
invocation ID 71/read view 83 and shared budget. Results are 82 (scalar success),
13 (handled callback error), 94 (native ByRef mutation plus callback), 66 (ByRef
mutation with handled error), then 82 (recovery). Canonical total is 41 after
each case. Reset and unrelated contexts refuse during the callback. Callback
scope is absent afterwards, native references reach zero, comparison preparation
queries do not repeat and calculator Invoke/name/type-info reads remain zero.

The existing native IDispatch sink is exercised through real
IConnectionPointContainer/FindConnectionPoint/Advise and Unadvise. An owner-thread
host target supplies a bound typed handler without putting a non-Send session in
the Send+Sync sink callback. The target snapshot is retained before running VBA.
Events 7 and 9 produce 16; an event outside an active scope rejects. Reload makes
the old handler stale and charges no VBA blocks. A fresh handler receives 5;
Unadvise prevents further delivery and the source's final Release returns zero.
The fixture explicitly uses IDispatch for event-source control and incoming sink
transport. Calculator client calls use their declared native vtable slots.

Follow-up cancellation tests release record/BSTR/SAFEARRAY carriers and duplicate
object ParamArray packets at multiple instruction positions on both backends.
Handle counters and host payload owners return to their starting balances.
Fresh review also checked compiler-expanded optional class defaults; no extra
restriction was retained for those valid calls.

run11-native-callbacks.json/.zip retain exact sources and raw focused/full checks.
Installed TypedReentry and TypedContinuation Lean abstractions pass; concrete
thread-local context provenance and native sink/handler refinement remain in
EXTENDED_TODO.md. Existing offline Miri limitations receive no passing credit.

Residual delivery: bd-59co.2.9.9 and bd-59co.3.6.4 retain VBA WithEvents binding,
incoming ByRef cancellation, complex event arguments, reentrant subscription
mutation, idle/cross-apartment event policy and general callback coverage.
Native record/serving/ABI and certification retain their existing Windows owners.
No canonical row closes. BLK-WORKSHEETFUNCTION-VTABLE-20261101 remains unchanged;
no adapter member, COM registration, macro-security setting or guessed slot changed.


A further native ByRef parity gap is explicitly owned by bd-59co.2.9.20. With a
global Long initially 2, native code adds 10, then a callback increments/returns
the same global. Owned Excel (calcvba11-live-byref-oracle) returns 13:13 through
an exact-Long Declare/AddressOf native DLL. A retained native-slot diagnostic
against the current typed libraries observes callback result 3, final global 12
(encoded 312) on both backends. The native cell's change is invisible to the
callback and its later copyback overwrites the callback's global write. The
correct target is 13:13/1313; this diagnostic earns no parity passing credit.
Its source, dependency-fingerprint-selected build command, raw results and owned
VBA/native DLL receipt are retained in the archive. The next delivery must add
the failing regression and provide shared native alias authority or correct
entry/return synchronization without dangling frame/global references. This
residual prevents any general native ByRef/reentry compatibility claim.

Scoped bd-59co.2.9.19 is closed after 3763 tests in 201 suites, 181 ignored; four registry-writing fixtures excluded without credit. Strict workspace Clippy, formatting and path checks pass. Metadata is rechecked after recording bd-59co.2.9.20. The observed alias mismatch remains open and receives no parity credit.

Run12 follow-up: the historical mismatch above is corrected for exact native Long
by bd-59co.2.9.20; see [current delivery evidence](typed-native-long-callbacks-run12.md).
The original run11 observation and archive remain historical evidence.
