# Exact native Long visibility during synchronous callbacks

Delivery: bd-59co.2.9.20. Broader CORE-JIT-LOWERING and Windows rows remain in-progress/planned.

Source-bound [manifest](run12-native-long-callbacks.json) and
[archive](run12-native-long-callbacks.zip) retain sources, checks and oracle receipts.

The retained run11 counterexample starts a global Long at 2, adds 10 through a
native ByRef pointer, then calls VBA, which increments and returns that global.
The pre-fix retained test reports 3:12 (312); the corrected VM3 and native JIT
report 13:13 (1313). The original failure log remains in the source-bound archive.

The owned Excel run `calcvba12-native-long-callback` explicitly compiled the
injected module through VBE command 578 and returned 13:13. Its guardian,
compile selection/authority, source hashes, process ownership and cleanup
receipts are archived. This is development oracle evidence, not certification.
Reproduce with `pwsh -NoProfile -File scripts/run-native-long-callback-oracle.ps1`.
The script builds a local callback DLL; it changes no registration or security setting.

The native marshaller derives each canonical ByRef pointer once. Before callback
entry it publishes each exact Long into the retained caller lvalue; after return
or unwind it reads that lvalue back into the same native cell. Repeated aliases
reuse that pointer. Nested native calls clear the outer hook until they return.
Hook faults remain visible even if native code ignores callback failure. Native
HRESULT/error information takes precedence and final writeback still runs.

`cargo test -p oxvba-embed --test typed_native_com --test typed_prebound_dispatch`
covers both backends and both native-slot/prebound-Invoke transports. It retains
local argument cases, success/error global visibility, private parenthesized
arguments, nested native callbacks, shared budgets/view, scope restoration,
zero name/type-info lookups and zero final native reference balance.
`cargo test -p calcvba-host-api native_callback` additionally checks repeated
callbacks, nested calls without a hook, synchronization failure and unwind.

Strict Clippy and the repository NoRegistry check are recorded in the archive.
Four registry-writing fixtures are excluded without passing credit. The four
installed Lean abstractions pass; concrete TLS/caller/raw-pointer refinement
remains unproved in FTODO-TYPED-NATIVE-LONG-SYNC-20261103. No Miri/native formal
passing credit is claimed. Fresh review is retained with the source-bound logs.

The exact-Long delivery does not settle other native carriers, native pointer
identity with caller storage, post-return cleanup callbacks, general event ByRef
cancellation or pinned certification. Existing bd-59co.3.5.3 and bd-59co.3.6.4
retain that scope. Historical run11 mismatch evidence remains unchanged.
