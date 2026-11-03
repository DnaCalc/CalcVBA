# Live typed Excel collections, run 13

Owner: bd-59co.3.5.11. Clauses: COM-CLIENT-001, HOST-SESSION-001, SYS-DUAL-001.
The declared typed helper subset includes Range, Worksheets and Workbooks item
access and live For Each. WCC-EARLY-COMPLEX and pinned certification remain planned;
genuine Variant-result member chains remain with bd-59co.3.5.3 and CORE-7.

The same excel_surface.bas and typed-surface-run13.bas compile under visible VBE
Debug Compile before Application.Run and execute through the retained NativeOM
window in both owning VM3 and native JIT. Range's -4 getter performs one prebound
Invoke. Dual Sheets/Workbooks getters and IEnumVARIANT::Next use native slots.
Immutable plans are retained before execution; no runtime member/type lookup or
dynamic start is used. The binder retains complete metadata for interfaces reached
through typed calls, including their enumerator descriptors.

Excel and both runtimes agree on row-major iteration while the body changes later
cells: F1:1;G1:11;F2:1;G2:11. Exhaustion leaves the typed variable Nothing
(VarType 9); Exit For keeps F1. Uninitialized Range enumeration raises 424 in
Excel, correcting an initial erroneous 91 assumption in the development probe.
Worksheets yields Run13First then Run13Second; workbook enumeration observes one
owned workbook; name-based Item returns the exact requested item.

Exact per-helper transport counts (native, Invoke): RangeIteration (5,22),
RangeFirst (1,3), SheetIteration (7,0), BookIteration (4,0), SheetItem (3,0).
Portable owning tests on both backends cover live mutation, one-operation slices,
early exit, a caught Next error and canceled-frame balance. Iterator ownership is
detached from the activation map before foreign Next; no frame/map borrow spans
a callback. Unknown objects stay opaque until a declared typed assignment.

Commands: cargo test -p oxvba-embed --test typed_enumeration --test typed_excel_surface;
scripts/run-typed-surface-oracle.ps1; scripts/meta-check.ps1 -Fast -NoArtifacts -NoRegistry.
Source/executable hashes, VBE observations and owned-process receipts:
run13-excel-collections.json/.zip. Failed probes are retained as failures. Four
registry-writing tests receive no credit. Native pointer refinement and iterator
lifecycle formalization remain in the formal backlog. Broader event policy,
general dynamic expressions and pinned certification remain open.

Repository check: 3,776 passed in 206 suites; 183 ignored. Strict Clippy,
format, governance and path checks pass. Four registry writers are excluded with
no credit. Installed Lean 4.29.1 lifecycle/reentry abstractions pass; no concrete
native formal or Miri credit. Fresh review corrected duplicate interface casing
and added a sink Invoke lease for reentrant unsubscribe; relevant proofs reran.
