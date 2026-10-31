# Native Variant cell aliases

2026-10-31. Bounded delivery owner `bd-59co.3.10.9` is closed for the scope below; broader structural owner
`bd-59co.3.10.7` remains open behind its existing carrier prerequisites. Native
String/As Any aliases remain with `.3.10.6` and `.3.10.7`. This leaf covers
declared ByRef Variant arguments containing converter-admitted nonowning scalar
or BSTR payloads, including scalar and BSTR replacements. No canonical capability
row, pinned certification or complete structural/native-string parity is claimed.

The alias classifier now admits those VARIANT cells. Repeated caller identity
reuses the original FfiArg pointer, with one boxed storage owner and one readback
and cleanup. The existing converter clones returned BSTR bytes before the native
owner is cleared. Repeated alias groups with actual array/object/record/procedure-reference
payloads or incompatible declarations remain refused before entering the native
callee; existing nonaliased calls retain their separate behavior.

## Independent fixture and fresh Excel

The final independent fixture uses a 24-byte Windows x64 VARIANT with a typed
union, retaining BSTR pointer provenance. VariantClear releases the prior payload
before each replacement. The numeric export writes I4=19, observes sharing, then
writes R8=42.5. The string export writes an owned BSTR, observes sharing, clears
the prior value, and returns `final` + NUL + `payload`, length 13. Allocation
failures leave initialized caller-owned cells. VARIANT-cell pointers are not
retained; fresh BSTR payload ownership transfers through the returned cells.

Eight source-specific probes froze expectations before starting owned Excel:

| Fixture version / case | Exact result |
|---|---|
| primitive prototype / same cell | `42\|Double\|True\|4321` |
| primitive prototype / distinct cells | `-1\|Long\|Double\|True\|True\|4321` |
| prior owned-payload fixture / same cell | `42\|Double\|True\|4321` |
| prior owned-payload fixture / distinct cells | `-1\|Long\|Double\|True\|True\|4321` |
| prior owned-payload fixture / parenthesized copy | `-1\|Double\|True\|4321` |
| prior owned-payload fixture / forwarding | `42\|Double\|True\|4321` |
| final fixture / returned BSTR, same cell | `42\|13\|final:payload\|4321` |
| final fixture / returned BSTR, distinct cells | `-1\|first\|13\|final:payload\|4321` |

The colon replaces the embedded NUL only in the printable VBA report. Raw value
bytes are also checked by the native regression. Earlier fixture versions remain
separate; their measurements do not claim to execute the later refactored file.
The Variant checkpoint `5933403c` contains that exact fixture source. The later
String characterization adds exports to the same file without changing the
Variant bodies.

All eight probes pass the unchanged owned supervisor contract: visible VBE
Compile command 578, exact module/source authority, guardian before invocation,
then runtime invocation and owned-process-zero cleanup. Excel 16.0 build 20527,
Windows x64, AMD Ryzen 7 4800H. Safe Mode and AutomationSecurity=1 were observed;
no Office security or registry setting changed. Pre-existing Excel PID 3144 and
its original start time survived. Runtime Err.LastDllError was saved before any
conversion helper. The DLL sets it to 4321 immediately before returning.

`variant-alias-2026-10-31/raw.zip` retains all three fixture sources/DLL versions,
source-specific wrappers, input modules, sealed descriptors, plans, environment,
guardian records, ownership, results and the failing pre-fix regression log.
The manifest records original paths, build argv, compiler/CPU, source/DLL hashes,
all archive member hashes and current implementation hashes. Debug PDBs and
import libraries are omitted; they are not execution inputs. Archive text remains
byte-exact under the repository LF-only policy.

## Execution checks and limits

Before the fix, both positive regressions fail with Err 5 before native execution.
After it, scalar payloads, initial BSTR/embedded NUL, forwarding, explicit copies,
returned BSTR bytes, distinct cells, VARTYPE changes and LastDllError agree across
VM3 and JIT. Runtime owned-handle balance returns to its initial value after each
engine is dropped. Array/object Variant payloads, mixed Variant/Any declarations
in either order and repeated Any retain explicit refusal tests with native call
delta zero. At the Variant checkpoint, repeated native String calls were still refused.
The later [String temporary repair](native_strings.md) follows their different
per-argument conversion rules.
These protective refusals are open compatibility gaps, not VBA target behavior.

The complete Windows parity suite passes 11 tests with the opt-in owned-workbook
case ignored. The 19 existing native-string controls pass. Full
`scripts/meta-check.ps1 -NoArtifacts` passes: 3,652 tests, 179 ignored, format,
strict Clippy, governance, path stability and workspace check. Independent review checked native
layout, pointer provenance, VariantClear/allocation failure paths and byte cloning.

Runtime handle counters do not count copied OLE BSTR allocations. Single-owner
cleanup is supported by the implementation and measured values, not an OLE heap
allocation counter. No formal model proves this native pointer boundary; the
existing `FTODO-NATIVE-ALIAS-20261031` records that nonblocking refinement work.
Existing internal ByRef control-flow models are unchanged by this native staging
admission. No zero-selected proof or matching refusal earns capability credit.
