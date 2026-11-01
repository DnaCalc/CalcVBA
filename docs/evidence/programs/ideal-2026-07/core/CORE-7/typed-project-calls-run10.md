# Typed project calls â€” run 10

2026-11-01. Scoped delivery bd-59co.2.9.16 is closed on this bounded Optional/procedural
source-call subset. Broader CORE-7/CORE-8 and ParamArray/class/record families remain
with bd-59co.2.9.9 and bd-59co.2.10.7.

The additive ProjectGraph/load_graph/reload_graph API retains a leaf-first source
closure, with ordered references, original per-project OxPrograms and separate
canonical global stores. Source imports become retained unit/procedure indices;
neither backend rewrites the source or falls back to dynamic procedure execution.
Owning Optional calls use compiler-bound source defaults or host-prefix defaults,
private omitted ByRef cells, the Missing Variant marker and typed IsMissing.
Graph-wide code-cache identity and old code/store leases span suspension/reload.

Five embedding cases pass on VM3/native JIT using one-operation slices: typed host
interfaces/defaults and nominal object results cross project boundaries; defaults, named calls, live/private
aliases and Missing survive module/project calls; globals remain separate; output
mutation precedes a cross-project error handler; old suspended code/stores remain
usable after reload; reset reinitializes the closure. Diamond references share one
store, ordered references select once, cache keys include dependencies, and
missing/cyclic/duplicate/private/unsupported paths fail before body effects.

Fresh review/testing corrected four concrete gaps: an Empty literal passed by
value needed lazily allocated Variant scratch in a caller with no Variant slots;
omitted Optional ByRef Variant parameters needed private Missing cells in JIT
admission/frame setup; trimmed trailing omissions on cross-project calls needed
matching VM3/JIT preparation; the linked-program reference vector must stay in the
session instead of allocating once per resume. The existing allocation check
again proves zero allocations around the measured Variant getter operation.
Dynamic neighboring calls also exercise the optional-cell fixes.

The supervised Excel probe calcvba10-optional-initial compiled explicitly through
VBE command 578 and produced
`default=42|alias=42:42|private=44:42|missing=True|present=False|zero=0|named=6.5`.
Owned PID 8496 was cleaned up with owned-process-zero. This is development evidence,
not certification. The repeatable script is run-typed-optional-oracle.ps1.

Bounds remain four formal parameters, 128 physical slots and 64 retained frames.
Only procedural source graphs are admitted; legacy object-library descriptors on
graph nodes, source-free publication, classes/records, ParamArray and native
callback suspension are outside this bounded delivery. No broad parity or formal
proof claim follows from these cases.

Final repository check: 3747 passed in 198 suites, 181 ignored; four registry-writing fixtures excluded without credit. Formatting, strict workspace/all-target
Clippy, governance and path stability pass. The existing measured getter resume
allocates zero heap storage on both backends after retaining the linked-unit vector
in the session. Final Optional oracle replay calcvba10-optional-final passes and
cleans owned PID 47652. Source/check/oracle hashes are in [run10-projects.json](run10-projects.json)
and the shared [source archive](../../windows-x64/WIN-4/run10-native.zip). General
owning Is/native identity remains a delivery path under bd-59co.2.9.9/bd-59co.3.5.3;
no broad CORE-7/CORE-8 row closes.
