# Typed host value slots — run 9

2026-11-01. Scoped delivery `bd-59co.2.9.14`; JIT-CORE-001,
JIT-PARITY-001, RUNTIME-ABI-001 and SYS-DUAL-001. CORE-JIT-LOWERING stays **in-progress**.

The original host API/scalar ABI 2, scalar enums, descriptors and required trait
methods remain unchanged. The additive value capability reports TYPED_VALUE_ABI 1.
A host supplies TypedObjectLibrary metadata and HostSlot::typed callbacks, then
uses load_with_object_libraries, bind_values, begin_values/call_values. Providers
without that capability refuse through default methods before execution.

ValueType/ValueSlot/Value cover Long, Double, Boolean, BSTR, actual Variant and
portable objects; Void permits typed Sub entry. Interface(IID) adds nominal
metadata while retaining the object carrier. BSTR preserves exact byte/UTF-16
payloads; Variant preserves Null and CVErr independently of invocation faults.
The member getter/let/set/method descriptors use required positional arguments
(up to four). Slot signatures are checked before entry; interface results are
checked before assignment and released on mismatch. The old numeric signatures
retain their original calling convention and allocation-free path.

Both runtimes use typed physical slots over original OxFunc. Actual Variant cells
exist only where declared/required, with bounded conversion scratch, rather than
a universal Variant frame or heap-boxed callback result. BSTR and object locals own
raw carrier references; aliases share stable ancestor cells. The native JIT emits
actual indirect slot calls. Selected string/Variant value operations call shared
value kernels, without a CFG interpreter or backend fallback. A synchronous host
callback completes before suspension; no asynchronous preemption is added.

Eight dual-runtime fixture tests cover Value/Value2 (declared Variant), Cells,
Address and Rows/Columns.Count, exact callback order at slices 1/3/unlimited,
String/Variant ByRef internal calls and repeated Variant aliases whose writes
survive child faults, typed setters/Sub entry, Null error 94 with retained
Err.Description, Nothing error 91, incompatible returned IID error 13, cancellation,
and reload/drop while paused. The fixtures define host-owned Range semantics;
no spreadsheet member or engine-specific glue was added to the compatibility adapter.
The allocation suite records zero heap allocations while resuming a declared
Variant getter followed by Double conversion on both runtimes; old numeric paths
also retain zero heap/Variant allocations. This is a mechanism measurement, not
a DNA Calc integration benchmark.

The minimal exact BSTR/inline Variant storage/reload and repeated Variant
alias/fault cases pass strict-provenance
Miri in interpreter-only mode. The rich fault/handler case also passes default-mode
Miri, retaining the canonical Variant exposed-provenance warning. An earlier rejected run exposed an unnecessary
String-to-String conversion through canonical Variant; identity conversion now
keeps its direct BSTR carrier. Canonical Variant string/object exposed-provenance
limitations remain documented in the formal backlog. Miri does not execute native
JIT code. The existing Lean abstractions prove no new concrete carrier/native ABI.
Final repository results and raw logs are retained in run9-values.json/.zip.

Arrays/records as typed values, Optional/ParamArray, cross-project calls, runtime
object-default coercion (including Object payloads in Variant), writable host ByRef ABI, native callback
preemption and source-free verified packages retain their existing broader owners.
The bounded imported native scalar/BSTR transport is described in
WIN-4/typed-native-slots-run9.md; broader native signatures retain WIN-4.

Fresh review records the object-default gap explicitly under bd-59co.2.9.9:
if an Object reaches the current pure numeric/string value kernel, error 13 can
replace the required default-member dispatch. That result is not a compatibility
target. Retained generic default-member plans and dual-runtime callback/order
observations are the next delivery path; static planned member calls are separate.

The final repository check passes 3738 tests in 196 suites (181 ignored; four
registry-writing fixtures excluded without credit). Interpreter-only embedding:
58 passed, two ignored; wasm32 compile passes with existing HAL warnings.
The scoped value/member delivery is verified; broader capability rows remain open.
