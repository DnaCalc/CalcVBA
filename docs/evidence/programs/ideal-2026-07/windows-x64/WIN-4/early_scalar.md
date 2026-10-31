# Preserved-descriptor early COM, run 5

2026-10-31. Delivery bead `bd-59co.3.5.6`; broader WIN-4 and `.3.5.2` remain open.
This is Windows x64 development evidence, not pinned Office certification.
Raw Excel receipts: [archive](run5-excel-observations.zip) and
[SHA-256 manifest](run5-excel-observations.json), including three passing supervised
runs and the separately labelled aborted runtime-dialog observation.

The compiler's TypeLibMemberMetadata now crosses VM3 and JIT execution into
ComHal::invoke_early_bound. The native adapter consumes its IID, vtable slot,
slot bound, accessor and wire signature. It does not look up a member by ProgID,
GetTypeInfo or GetIDsOfNames. Dual members never use Invoke; a pure dispinterface
has an explicit fixed-IID/DISPID/accessor Invoke plan because it has no custom slots.
Host DispatchOnly preferences do not change typed dual transport. Unsupported dual
metadata fails before entry without a retry.

The shared COM marshaller supplies private cells for non-lvalue ByRef arguments,
shares native cells for repeated caller aliases, keeps equal-valued distinct
variables separate, and writes caller aliases back before execution resumes.
IErrorInfo is retrieved immediately after a failing HRESULT, before interface
cleanup can replace it; Source/Description/HelpFile/HelpContext cross HAL without
transport text being appended to VBA's Err.Description. Diagnostic text remains separate.
JIT helper ABI identity advances to private-helpers=3.
The existing portable embedding projection selects its frozen tokens directly,
without forwarding typed calls to dynamic name dispatch. Its fake spreadsheet
types remain extraction debt under `bd-59co.2.10.7`, with no native COM credit.
Registered portable objects now opt into a separate typed adapter; its default
refuses without late retry. The old portable comparison test now supplies a real
host fixture, checks exact results and counts distinct dynamic/typed calls on
both backends. A late-only fixture proves zero dynamic retries after typed refusal.
No synthetic token-arithmetic fallback is retained as an early-binding target.
The older dynamic headless synthetic projection remains a gap for the neutral
host-services replacement `bd-59co.2.10.7`; it is not VBA/COM compatibility evidence.

`cargo test -p oxvba-differential --test jit_windows_vm3_parity early_com -- --test-threads=1`:
seven tests pass on both backends. The registration-free dual fixture refuses all
GetTypeInfo/GetIDsOfNames/Invoke member recovery and supplies unrelated DISPIDs.
Its scalar/string/Boolean/Variant property/method scenario produces seven vtable
successes and zero Invoke calls. Other cases prove HRESULT/IErrorInfo propagation,
invalid-slot refusal, ByRef forwarding and real alias identity versus equal values.
A real registered Scripting.Dictionary typelib/activation also produces seven
vtable successes and zero Invoke calls, including Item get/put and ByRef Variant
temporaries. Error 457 and missing-item insertion are compared with Excel.
Runtime owned-value counters return to baseline after every fixture run. Independent
per-object native counters prove zero Invoke attempts (including failed attempts),
zero name/type-info reads, and zero remaining COM/QI references on success/error/refusal.
Review also corrected the shared fixture's creator reference after
[SetErrorInfo](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-seterrorinfo).
Its separate native error-object reference count must now return to zero as well;
the rich server description is asserted exactly, without appended transport text.
Missing declared IID support raises 430. A separate pure-dispinterface test supplies
a compiler member name unknown to the server and proves one explicit fixed-DISPID Invoke.

`scripts/run-hosting-early-binding-oracle.ps1` reuses the existing process-owned
supervisor and PID-scoped UIA guardian. It explicitly invokes VBE Compile 578.
The dictionary case adds only a workbook-local reference to the already installed
Scripting typelib before compile. No registry, macro-security or Trust Center
settings change; no Office binaries are inspected. PID 3144 is never owned.

Public and Option Private Module probes both compile and run. Within the probe
workbook they produce: Public Function 2; Private Function 73; Sub cell result 0;
Long text coercion 3; bad text error 2015; CVErr error 2042; Range:5 versus Double:5.
Run returns Long:2, Long:73, Empty for the Sub, and an Error value for CVErr.
The dictionary oracle gives value42/count1/ExistsTrue/missingEmpty/after-count2;
duplicate Add raises 457 with description "This key is already associated with an
element of this collection" and project-default source VBAProject.
An explicit unhandled Err.Raise13 produces a captured runtime dialog, so that
aborted probe is not evidence of ordinary worksheet exception-to-error mapping.
Workbook-qualified formulas in a second workbook return 2/73 for the Public/Private
Functions in both module modes. This host lookup is distinct from project references.
These calculations run inside RunProbe; they do not establish idle recalculation
or function-discovery UI visibility.

Bootstrap verification and the post-run hash now tolerate Excel's retained write
handle; verification hashes and parses one exact shared-read snapshot. The worker
and supervisor also share null-versus-empty, case-sensitive expected-value comparison.
Final probes assert the tuples frozen after initial readings. Earlier wrapper attempts
with an empty expected-value string did not pass supervisor reconciliation and are
not counted as passing evidence. The oracle harness self-test covers both regressions.

Residuals: this does not supply typed primary host-to-JIT entries, neutral public
procedure handles, arbitrary custom IUnknown interfaces, complete dispinterface
signature breadth, general out-of-process/apartment transport, named ByRef alias mapping,
all optional/LCID/array/record signatures, synchronous callbacks/events, or JIT COM
serving. ByRef writes on a failing HRESULT and general typed receiver-assignment/QI
timing remain open. Existing broad signature/proxy/formal/verified-image owners remain active.
Native Kani is unavailable; Lean abstractions do not prove Rust/COM ownership or
actual native transport. These are explicitly retained as open evidence work.

Final repository check: `CARGO_BUILD_JOBS=2` and
`./scripts/meta-check.ps1 -Fast -NoArtifacts -NoRegistry -RunId calcvba5-final` pass:
format, strict workspace/all-target Clippy, governance/path checks and 3,671 Rust
tests; 179 existing ignored tests and four registry-writing record fixtures receive
no passing credit. CI's default registry-test behavior is unchanged.
The final no-default-feature embedding run passes 19 tests; the WebAssembly build
and Add2/mixed-matrix/noncatchable-budget probes pass. Existing cross-target build
warnings are retained. Oracle harness self-tests and five installed Lean abstractions pass.
Raw validation logs and source hashes: [archive](run5-validation.zip) and
[SHA-256 manifest](run5-validation.json). Earlier failed/interrupted checks receive no credit.

Fresh-eyes review corrected accidental changes to dynamic alias IDs, added explicit
COM initialization, replaced portable name forwarding with fixed-token adapters,
and corrected/measured the fixture's error-object reference. Review reread descriptor
lifetimes, QI/Release paths, aliases, error fields and refusal paths; no further defect
was found in this bounded delivery. The residuals above remain open.
