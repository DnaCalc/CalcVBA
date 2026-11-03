# Native Double and Boolean visibility during callbacks

Owner: bd-59co.2.9.21. Clauses: JIT-CORE-001, JIT-PARITY-001,
RUNTIME-ABI-001, SYS-DUAL-001. Broader CORE-JIT-LOWERING remains in-progress.

The shared native-scalar-callback-bodies-run13.bas is compiled unchanged in owned
Excel/VBA and in owning VM3/native JIT. A local registration-free oracle DLL
mutates actual repeated ByRef Double/Boolean arguments before calling AddressOf
VBA. Both repeated arguments must have the same pointer. The engine fixture
performs the matching operations through native slots and prebound Invoke.

VBA observes Double 12.5:12.5 and Boolean False:False. Retained pre-fix engine
counterexamples are 2.5:12 and True:True: the callback read stale globals, then
final native writeback replaced its mutation. The existing exact-Long mechanism
now retains typed i32/f64/i16 native authorities and publishes each supported
canonical alias to its caller before reentry, then refreshes the same native
cell from that caller after return. Both the native-cell and owning-frame ends
select the exact declared carrier. No native Cell is reborrowed during the call.

The shared Boolean callback also exposed a missing owning Not instruction.
Boolean operands with Boolean/Variant destinations now use the existing shared
Boolean operator body, through the same single-operation helper on VM3/JIT.
The shared source checks True-to-False during the native callback and the two
subsequent False-to-True-to-False calls. Numeric/Variant/array Not shapes outside
this admission remain with the broader owner.

All four backend/transport combinations match. Each native member uses one
native slot or one compiler-prebound Invoke; repeated aliases reuse one native
pointer, callback execution consumes the shared budget, runtime name/type-info
reads remain zero, and final native references balance to zero. Existing Long,
error, nested callback, array/interface and pure-dispatch suites remain green in
the focused check. Native pointer equality with the physical caller storage,
other carriers and complete error/cleanup policies remain open.

Commands: cargo test -p oxvba-embed --test typed_native_com --test typed_prebound_dispatch;
scripts/run-native-scalar-callback-oracle.ps1;
scripts/meta-check.ps1 -Fast -NoArtifacts -NoRegistry. Source, DLL/executable
hashes, pre-fix failures, VBE observations and cleanup: run13-native-scalar-callbacks.json/.zip.
Repository validation: 3,777 passed in 206 suites; 183 ignored. Strict Clippy,
format, governance and path checks pass. Four registry-writing fixtures are
excluded without passing credit. All 30 owned Excel instances from this run are
gone; PID reuse is distinguished by the recorded process start time. No pinned certification or concrete
native/Miri/formal refinement is claimed.
