# Windows-hosted execution plan

Date: 2026-09-21
Status: detailed planning accepted; capability work remains in progress
Scope: Windows x64 compatibility and its required shared Core contracts, including runtime-backed wrappers.
Authority: CHARTER.md, OPERATIONS.md, the system/Windows specifications and the accepted Windows workset.
Execution state and user stop conditions remain solely in docs/AUTORUN_STATE.md.

## Intended result and boundaries

Deliver general Windows COM clients, incoming/outgoing events, COM serving, VBA7 native imports, pointer/callback behavior and reusable VM3/JIT sessions on x64. Each capability must preserve real VBA compilation, results, errors, mutation, transport and lifetime behavior. Controlled development fixtures establish implementation evidence; pinned Windows/Excel64 certification remains a separate prerequisite for release claims.

This plan excludes full IDE/language-service delivery and Linux baseline execution. Existing Linux baseline and Core terminal owners remain open dependencies of the broader programme. Raw authoritative metadata is handed to future IDE consumers without adding an IDE implementation stream.

The user-directed pause on CORE-7 remaining ideal architecture (`bd-59co.2.9.9`) and genuine WIN-12 DLL/EXE outputs remains in force. Genuine outputs are a separate accepted extension, not part of the requested compatibility delivery. Their existing beads and full-programme terminal rows remain open. Completing this plan's compatibility work cannot close WIN-12, WIN-15, the complete Windows workset or the umbrella programme.

## Current implementation and why these steps are needed

The WIN-2 first slice is already closed: shared plan types, the VM3 migration and the JIT adapter (`bd-59co.3.3.2` through `.4`). `crates/oxvba-runtime/src/interop_plan.rs` currently models LateDispatch and DeclareX64, with no reentry in that slice. Later plan families require explicit work; they cannot be assumed from the type name VerifiedInteropPlan.

`crates/oxvba-host/src/engine.rs` still constructs VM3-backed ProjectRuntimeSession values. Persistent JIT sessions, cache lifetime and shared native reentry therefore appear as producer beads before Windows consumers. `crates/oxvba-differential/tests/jit_windows_vm3_parity.rs` has only two first-slice cases, including GetTickCount. Sequential time-dependent results must be replaced with a controlled deterministic native fixture for exact parity assertions.

The Windows development baseline `bd-59co.2.2.10` is closed. The pinned Linux baseline `.11` and Core baseline reconciliation `.12` remain open. Most Windows matrices are planned; only WAC-TARGET-DEV-ENV is verified. Planning changes no capability status.

## Execution sequence

1. Begin the existing scalar COM (`bd-59co.3.4.2`) and Declare (`bd-59co.3.10.2`) delivery beads. Their WIN-3/WIN-9 planning rollouts were accepted in this planning cycle; the closed WIN-2 adapter prerequisite is sufficient for the bounded fixture slices.
2. In separate resource slots, finish the owned Excel/VBE supervisor, Windows control handoff and controlled fixture inventory. Provision and qualify the clean certification environment through its existing owners.
3. Build authoritative metadata/carrier evidence alongside exact Core binding, image verification, helper and structural-observation contracts. Connect consumers to specific producer leaves rather than waiting for whole unrelated epics.
4. Add persistent JIT sessions/cache, stable reentry and the missing early/event/serving/callback plan families. Every adapter consumes one immutable plan identity.
5. Broaden clients and Declare; add early-vtable clients, incoming events, served objects and outgoing events. Pointer and callback work follows the native-call and reentry contracts.
6. Package JIT wrappers and prove cleanup under failures. Run each WIN-14 certification family when its exact producers and pinned environment are ready; do not wait for all capability work to begin collecting evidence.
7. Publish a compatibility-subset report with exact remaining rows. Keep genuine-output and full-programme terminal work open.

## Scheduling and acceptance rules

- The live claim source is `br ready -l ideal-2026-07 -t task`; this document is a dated delivery map, not a replacement for AUTORUN_STATE.
- At most two Rust writers may run concurrently. Workspace-wide Cargo, Excel/VBE, registry changes, certification provisioning and large JIT/VM3/differential/rt-abi work are serialized according to their resource labels.
- Use small/medium/large complexity labels without elapsed-time estimates. A large leaf still owns one outcome; split unrelated behavior before claiming it.
- Each delivery leaf includes success, error and lifecycle evidence. Keep exact result, full Err, side effects, lifecycle order, transport and balance, with reasons for inapplicable axes. Matching declines are not native execution evidence.
- New detailed leaves advance existing matrix rows without broadening verified claims. A broad row remains planned or in-progress until every required implementation and certification producer has passed.
- Start the owned modal supervisor before Excel. Compile through VBE Debug Compile, capture dialog/token/full line, and clean up only recorded owned PIDs. Reuse `scripts/run-excel-vba-oracle.ps1`; do not drive Application.Run as a compile check.
- Do not silently update a dependency-specific fallback into a general completion claim. Preserve public-spec/oracle uncertainty as an explicit question with a reproducible fixture and a delivery owner.


## Bead map

This plan adds 13 Core beads and 53 Windows beads: 65 delivery outcomes and one non-owning compatibility report. Existing scalar, fixture, certification and broader remaining-scope owners are reused. The separate planning support bead records this breakdown.

| Work area | New beads | Bead IDs |
|---|---:|---|
| CORE-3 Compiler AnalysisResultV1, types, calls and references | 3 | bd-59co.2.4.5, bd-59co.2.4.6, bd-59co.2.4.7 |
| CORE-4 Verified OxIR and OxImage realization | 4 | bd-59co.2.6.2, bd-59co.2.6.3, bd-59co.2.6.4, bd-59co.2.6.5 |
| CORE-5 Runtime, helper ABI and session ownership | 3 | bd-59co.2.7.5, bd-59co.2.7.6, bd-59co.2.7.7 |
| CORE-8 JIT sessions, cache and native continuity | 2 | bd-59co.2.10.2, bd-59co.2.10.3 |
| CORE-9 Structural parity, safety and Excel/VBA certification | 1 | bd-59co.2.11.6 |
| WIN-1 Authoritative metadata and exact Windows carriers | 7 | bd-59co.3.2.2, bd-59co.3.2.3, bd-59co.3.2.4, bd-59co.3.2.5, bd-59co.3.2.6, bd-59co.3.2.7, bd-59co.3.2.8 |
| WIN-2 Shared interop plan and Windows session attachment | 7 | bd-59co.3.3.6, bd-59co.3.3.7, bd-59co.3.3.8, bd-59co.3.3.9, bd-59co.3.3.10, bd-59co.3.3.11, bd-59co.3.3.12 |
| WIN-3 Late-bound COM client and activation | 5 | bd-59co.3.4.4, bd-59co.3.4.5, bd-59co.3.4.6, bd-59co.3.4.7, bd-59co.3.4.8 |
| WIN-4 Early-bound native-vtable COM client | 4 | bd-59co.3.5.2, bd-59co.3.5.3, bd-59co.3.5.4, bd-59co.3.5.5 |
| WIN-5 Incoming COM connection-point events | 4 | bd-59co.3.6.5, bd-59co.3.6.6, bd-59co.3.6.7, bd-59co.3.6.8 |
| WIN-6 Late-bound COM serving foundation | 5 | bd-59co.3.7.2, bd-59co.3.7.3, bd-59co.3.7.4, bd-59co.3.7.5, bd-59co.3.7.6 |
| WIN-7 Early/dual serving and Implements | 3 | bd-59co.3.8.2, bd-59co.3.8.3, bd-59co.3.8.4 |
| WIN-8 Outgoing COM events | 2 | bd-59co.3.9.2, bd-59co.3.9.3 |
| WIN-9 VBA7 Declare import | 4 | bd-59co.3.10.4, bd-59co.3.10.5, bd-59co.3.10.6, bd-59co.3.10.7 |
| WIN-10 Pointer helpers and AddressOf callbacks | 3 | bd-59co.3.11.5, bd-59co.3.11.6, bd-59co.3.11.7 |
| WIN-11 JIT-backed wrapped outputs | 4 | bd-59co.3.12.2, bd-59co.3.12.3, bd-59co.3.12.4, bd-59co.3.12.5 |
| WIN-13 Native-boundary safety and lifecycle | 4 | bd-59co.3.14.3, bd-59co.3.14.4, bd-59co.3.14.5, bd-59co.3.14.6 |
| WIN-14 Windows, native-client and Excel/VBA certification | 1 | bd-59co.3.15.37 |

## Detailed beads

Each card is also stored as a self-contained bead. File paths below are starting points confirmed in the current checkout; implementers must inspect the current code before editing. Listed suite commands exist now and must be extended with the card's exact cases.

### bd-59co.2.4.5 — CORE-3 preserve VBA7 Declare and AddressOf compile contracts

Outcome: CORE-3 preserve VBA7 Declare and AddressOf compile contracts.

Effect: delivery. Parent epic: bd-59co.2.4 (CORE-3).

Clauses: COMP-ANALYSIS-001|COMP-BIND-001|COMP-DIAG-001|IR-CORE-001|LS-FACT-001|PROJ-REF-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.2.10. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Retain PtrSafe, LongPtr, LongLong, aliases, As Any, ByVal/ByRef and target-conditional facts through binding.

2. Validate AddressOf module/procedure/signature eligibility and distinguish compile errors from missing DLL or entry runtime errors.

3. Add positive and negative source fixtures with exact diagnostic spans; use the owned VBE supervisor to resolve uncertain VBA legality.

Starting code/truth surfaces: crates/oxvba-bind/src, crates/oxvba-symbol/src, crates/oxvba-host/tests/native_declare_lane_vm3.rs.

Canonical matrix rows: docs/validation/CORE_COMPILER_VM_JIT_READINESS_MATRIX_V1.csv#CORE-TYPED-BINDING.

Row contract CORE-TYPED-BINDING: semantic_subset=ByVal ByRef Optional named omitted ParamArray returns arrays UDTs and objects.

Acceptance commands: command:cargo test -p oxvba-bind; command:cargo test -p oxvba-host --test native_declare_lane_vm3.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-3/c_decl.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.4; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.2.4.6 — CORE-3 bind imported interfaces and source-event signatures

Outcome: CORE-3 bind imported interfaces and source-event signatures.

Effect: delivery. Parent epic: bd-59co.2.4 (CORE-3).

Clauses: COMP-ANALYSIS-001|COMP-BIND-001|COMP-DIAG-001|IR-CORE-001|LS-FACT-001|PROJ-REF-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.2.10, bd-59co.3.2.3. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Carry authoritative interface/member/parameter identities into typed calls without name-only reconstruction.

2. Bind properties, optional/named arguments, default/source interfaces and WithEvents handler signatures.

3. Reject incompatible handlers, ambiguous references and illegal ByRef conversions with source locations; preserve imported result types.

Starting code/truth surfaces: crates/oxvba-symbol/src, crates/oxvba-bind/src, crates/oxvba-com/src/typelib.rs.

Canonical matrix rows: docs/validation/CORE_COMPILER_VM_JIT_READINESS_MATRIX_V1.csv#CORE-TYPED-BINDING.

Row contract CORE-TYPED-BINDING: semantic_subset=ByVal ByRef Optional named omitted ParamArray returns arrays UDTs and objects.

Acceptance commands: command:cargo test -p oxvba-symbol; command:cargo test -p oxvba-bind.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-3/c_com.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.4; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.2.4.7 — CORE-3 preserve exported class and Implements descriptors

Outcome: CORE-3 preserve exported class and Implements descriptors.

Effect: delivery. Parent epic: bd-59co.2.4 (CORE-3).

Clauses: COMP-ANALYSIS-001|COMP-BIND-001|COMP-DIAG-001|IR-CORE-001|LS-FACT-001|PROJ-REF-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.4.6. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Emit stable exported class/interface/property/event identities and complete calling signatures.

2. Validate imported Implements obligations and distinguish public/default/source members.

3. Round-trip descriptors through packages and compare source versus imported signatures; no serving-specific rebinding.

Starting code/truth surfaces: crates/oxvba-bind/src, crates/oxvba-oxir/src, crates/oxvba-build/src.

Canonical matrix rows: docs/validation/CORE_COMPILER_VM_JIT_READINESS_MATRIX_V1.csv#CORE-TYPED-BINDING.

Row contract CORE-TYPED-BINDING: semantic_subset=ByVal ByRef Optional named omitted ParamArray returns arrays UDTs and objects.

Acceptance commands: command:cargo test -p oxvba-bind; command:cargo test -p oxvba-build.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-3/c_exports.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.4; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.2.6.2 — CORE-4 bound OxImage decoding for Windows packages

Outcome: CORE-4 bound OxImage decoding for Windows packages.

Effect: delivery. Parent epic: bd-59co.2.6 (CORE-4).

Clauses: CONF-QUALITY-001|IMAGE-VERIFY-001|SEC-BOUNDARY-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.2.10. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Bound input bytes, collection counts, nesting and allocation before materializing untrusted package metadata.

2. Reject malformed lengths and excess type/interface/record graphs with deterministic diagnostics.

3. Exercise truncated, recursive and resource-exhaustion packages without executing or loading dependencies.

Starting code/truth surfaces: crates/oxvba-oxir/src, crates/oxvba-host/src/engine.rs.

Canonical matrix rows: docs/validation/OXIMAGE_PACKAGE_CONTRACT_MATRIX_V1.csv#OXI-BOUNDED-DECODE; docs/validation/OXIMAGE_PACKAGE_CONTRACT_MATRIX_V1.csv#OXI-CORRUPTION-RESOURCE.

Row contract OXI-BOUNDED-DECODE: semantic_subset=length count nesting allocation and version limits.

Row contract OXI-CORRUPTION-RESOURCE: semantic_subset=truncation bit flips duplicated IDs bad CFG types ranks descriptors links and zip-bomb shapes.

Acceptance commands: command:cargo test -p oxvba-oxir; command:cargo test -p oxvba-host.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-4/c_decode.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.6; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.2.6.3 — CORE-4 verify external signatures and seal image admission

Outcome: CORE-4 verify external signatures and seal image admission.

Effect: delivery. Parent epic: bd-59co.2.6 (CORE-4).

Clauses: CONF-QUALITY-001|IMAGE-VERIFY-001|IR-OXIR-001|SEC-BOUNDARY-001|SYS-ART-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.6.2. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Verify external member IDs, types, arity, directions, ABI, event/export descriptors and target requirements before planning.

2. Produce sealed verified handles that cannot be constructed from raw external data through safe public APIs.

3. Add a rejection case for each descriptor invariant and ensure both backends reject before side effects.

Starting code/truth surfaces: crates/oxvba-oxir/src, crates/oxvba-runtime/src/interop_plan.rs.

Canonical matrix rows: docs/validation/OXIMAGE_PACKAGE_CONTRACT_MATRIX_V1.csv#OXI-SEALED-VERIFY; docs/validation/OXIR_BACKEND_SUPPORT_MATRIX_V1.csv#OXIR-VERIFIER-REJECT.

Row contract OXI-SEALED-VERIFY: semantic_subset=VerifiedOxProgram and VerifiedOxImage are owning unforgeable product inputs.

Row contract OXIR-VERIFIER-REJECT: semantic_subset=unknown malformed out-of-range type effect descriptor and capability cases.

Acceptance commands: command:cargo test -p oxvba-oxir; command:cargo test -p oxvba-runtime interop_plan.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-4/c_verify.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.6; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.2.6.4 — CORE-4 bind package compatibility to provenance and ABI identities

Outcome: CORE-4 bind package compatibility to provenance and ABI identities.

Effect: delivery. Parent epic: bd-59co.2.6 (CORE-4).

Clauses: DEBUG-MAP-001|IMAGE-ABI-001|IMAGE-VERIFY-001|PROJ-REF-001|SRC-ID-001|SYS-ART-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.6.3. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Bind image content, helper/carrier ABI, target/profile and imported reference identities into deterministic compatibility metadata.

2. Retain source/debug and registered/file reference provenance and reject conflicting case-folded link identities.

3. Mutate each identity independently and prove deterministic rejection; preserve authoritative image entry.

Starting code/truth surfaces: crates/oxvba-oxir/src, crates/oxvba-project/src.

Canonical matrix rows: docs/validation/OXIMAGE_PACKAGE_CONTRACT_MATRIX_V1.csv#OXI-IDENTITY-ABI; docs/validation/OXIMAGE_PACKAGE_CONTRACT_MATRIX_V1.csv#OXI-EXPORT-REFERENCES; docs/validation/OXIMAGE_PACKAGE_CONTRACT_MATRIX_V1.csv#OXI-PROVENANCE-DEBUG.

Row contract OXI-IDENTITY-ABI: semantic_subset=deterministic artifact identity includes helper carrier and target contracts.

Row contract OXI-EXPORT-REFERENCES: semantic_subset=procedures properties public data types constants events and provider identities.

Row contract OXI-PROVENANCE-DEBUG: semantic_subset=original virtual generated spans statement maps and procedure identities.

Acceptance commands: command:cargo test -p oxvba-oxir; command:cargo test -p oxvba-project.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-4/c_identity.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.6; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.2.6.5 — CORE-4 require sealed image inputs at Windows execution boundaries

Outcome: CORE-4 require sealed image inputs at Windows execution boundaries.

Effect: delivery. Parent epic: bd-59co.2.6 (CORE-4).

Clauses: BUILD-PACKAGE-001|HOST-SESSION-001|IMAGE-VERIFY-001|SYS-ART-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.6.4. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Route host, VM3, JIT and comhost Windows package admission through sealed image handles.

2. Use the image entry and validated descriptors instead of last-program or raw-program conventions.

3. Prove malformed images cannot initialize COM, resolve DLLs, allocate executable memory or create sessions.

Starting code/truth surfaces: crates/oxvba-host/src/engine.rs, crates/oxvba-jit/src, crates/oxvba-vm3/src, crates/oxvba-comhost/src/lib.rs.

Canonical matrix rows: docs/validation/OXIMAGE_PACKAGE_CONTRACT_MATRIX_V1.csv#OXI-VERIFIED-CONSUMERS.

Row contract OXI-VERIFIED-CONSUMERS: semantic_subset=VM3 JIT host build cache and language service accept verified handles only.

Acceptance commands: command:cargo test -p oxvba-host; command:cargo test -p oxvba-comhost.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-4/c_consumers.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.6; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.2.7.5 — CORE-5 generate interop helpers from one versioned ABI catalog

Outcome: CORE-5 generate interop helpers from one versioned ABI catalog.

Effect: delivery. Parent epic: bd-59co.2.7 (CORE-5).

Clauses: HOST-HAL-001|HOST-SESSION-001|IR-OXIR-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|RUNTIME-VALUE-001|SEC-BOUNDARY-001|SYS-OWN-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.2.26. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Define catalog records for signatures, ownership, error, reentry, apartment and unwind behavior of Windows helper calls.

2. Generate registration and lookup for VM3/JIT adapters from the same catalog.

3. Validate catalog identity and negative signatures; contain panics at foreign/generated boundaries and prove cleanup on helper failure.

Starting code/truth surfaces: crates/oxvba-rt-abi/src, crates/oxvba-jit/src, crates/oxvba-runtime/src.

Canonical matrix rows: docs/validation/CORE_COMPILER_VM_JIT_READINESS_MATRIX_V1.csv#CORE-RUNTIME-HELPER-SESSION; docs/validation/OXIR_BACKEND_SUPPORT_MATRIX_V1.csv#OXIR-RUNTIME-HELPERS.

Row contract CORE-RUNTIME-HELPER-SESSION: semantic_subset=versioned descriptors typed wrappers lifecycle error and reset.

Row contract OXIR-RUNTIME-HELPERS: semantic_subset=every admitted operation names versioned helpers ownership effects fault seating and backend-neutral session behavior.

Acceptance commands: command:cargo test -p oxvba-rt-abi; command:cargo test -p oxvba-jit.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-5/c_helpers.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.7; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.2.7.6 — CORE-5 own image metadata and runtime state in a shared session

Outcome: CORE-5 own image metadata and runtime state in a shared session.

Effect: delivery. Parent epic: bd-59co.2.7 (CORE-5).

Clauses: HOST-HAL-001|HOST-SESSION-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|RUNTIME-VALUE-001|SEC-BOUNDARY-001|SYS-OWN-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.6.5, bd-59co.2.7.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Define backend-neutral initialize/invoke/reset/dispose state and ownership over globals, Err, objects, events and immutable descriptors.

2. Replace process-lifetime descriptor retention in the touched session paths with owned or safely shared storage.

3. Exercise repeat create/invoke/reset/drop, initialization failure and outstanding handles without leaks or stale-session access.

Starting code/truth surfaces: crates/oxvba-host/src/engine.rs, crates/oxvba-rt-abi/src, crates/oxvba-runtime/src.

Canonical matrix rows: docs/validation/CORE_COMPILER_VM_JIT_READINESS_MATRIX_V1.csv#CORE-RUNTIME-HELPER-SESSION.

Row contract CORE-RUNTIME-HELPER-SESSION: semantic_subset=versioned descriptors typed wrappers lifecycle error and reset.

Acceptance commands: command:cargo test -p oxvba-host; command:cargo test -p oxvba-rt-abi.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-5/c_session.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.7; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.2.10.2 — CORE-8 retain compiled code in persistent Windows-capable JIT sessions

Outcome: CORE-8 retain compiled code in persistent Windows-capable JIT sessions.

Effect: delivery. Parent epic: bd-59co.2.10 (CORE-8).

Clauses: HOST-SESSION-001|JIT-AOT-001|JIT-CACHE-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.7.6, bd-59co.2.9.8. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Implement the common session interface for compiled images while retaining code and relocation owners for the whole session lifetime.

2. Preserve global/class/Err state across calls and implement reset/disposal without recompiling each invocation.

3. Compare repeated invocations against VM3; explicitly decline unsupported inputs before execution and never fall back to VM3.

Starting code/truth surfaces: crates/oxvba-host/src/engine.rs, crates/oxvba-jit/src.

Canonical matrix rows: docs/validation/CORE_COMPILER_VM_JIT_READINESS_MATRIX_V1.csv#CORE-JIT-CACHE-OBJECT.

Row contract CORE-JIT-CACHE-OBJECT: semantic_subset=source-free verified image sessions deterministic key bounded cache persistent objects and object blobs.

Acceptance commands: command:cargo test -p oxvba-host; command:cargo test -p oxvba-jit.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-8/c_jit_session.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.10; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.2.10.3 — CORE-8 validate cache identity and compiled-code lifetime

Outcome: CORE-8 validate cache identity and compiled-code lifetime.

Effect: delivery. Parent epic: bd-59co.2.10 (CORE-8).

Clauses: HOST-SESSION-001|JIT-AOT-001|JIT-CACHE-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.10.2, bd-59co.2.6.4. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Key the shared cache by image, target ISA/CPU features, helper/carrier ABI and host capability profile.

2. Separate immutable code sharing from mutable project state and bound eviction while live sessions retain code ownership.

3. Prove cold versus warm behavior, every key mismatch, concurrent lookup, eviction and stale-handle rejection with deterministic identity evidence.

Starting code/truth surfaces: crates/oxvba-jit/src, crates/oxvba-host/src/engine.rs.

Canonical matrix rows: docs/validation/CORE_COMPILER_VM_JIT_READINESS_MATRIX_V1.csv#CORE-JIT-CACHE-OBJECT.

Row contract CORE-JIT-CACHE-OBJECT: semantic_subset=source-free verified image sessions deterministic key bounded cache persistent objects and object blobs.

Acceptance commands: command:cargo test -p oxvba-jit; command:cargo test -p oxvba-host.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-8/c_cache.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.10; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.2.7.7 — CORE-5 preserve nested JIT execution and Err state during native reentry

Outcome: CORE-5 preserve nested JIT execution and Err state during native reentry.

Effect: delivery. Parent epic: bd-59co.2.7 (CORE-5).

Clauses: HOST-HAL-001|HOST-SESSION-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|RUNTIME-VALUE-001|SEC-BOUNDARY-001|SYS-OWN-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.10.2, bd-59co.2.7.4. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Add a stable session-owned JIT invocation root analogous to the existing VM3 reentry owner.

2. Preserve suspended frames, aliases, pending errors and object lifetimes while a native callback invokes VBA synchronously.

3. Test nested success, handler failure, disposal attempt and recursion exhaustion; native boundaries never observe a Rust unwind.

Starting code/truth surfaces: crates/oxvba-jit/src, crates/oxvba-rt-abi/src, crates/oxvba-host/src/engine.rs.

Canonical matrix rows: docs/validation/CORE_COMPILER_VM_JIT_READINESS_MATRIX_V1.csv#CORE-RUNTIME-HELPER-SESSION.

Row contract CORE-RUNTIME-HELPER-SESSION: semantic_subset=versioned descriptors typed wrappers lifecycle error and reset.

Acceptance commands: command:cargo test -p oxvba-jit; command:cargo test -p oxvba-rt-abi.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-5/c_jit_reentry.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.7; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.2.11.6 — CORE-9 compare structural Windows interop observations

Outcome: CORE-9 compare structural Windows interop observations.

Effect: delivery. Parent epic: bd-59co.2.11 (CORE-9).

Clauses: CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001|JIT-PARITY-001|SYS-DUAL-001.

Scope: shared platform-neutral Core behavior; this plan requires Windows x64 evidence and actual 64-bit Excel for VBA authority where relevant. Linux baseline execution remains with its existing owner; do not introduce Windows-only Core semantics.

Prerequisites: bd-59co.2.2.10. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Extend the shared observable with structural arrays/records, normalized object identity, complete Err/LastDllError, ByRef mutations and ordered side effects.

2. Record plan identities, dispatch/vtable/native transport counters and AddRef/Release/pin/callback balance.

3. Use deterministic fixtures instead of sequential GetTickCount equality; fail rather than classify an expected native success as a matching decline.

Starting code/truth surfaces: crates/oxvba-differential/src, crates/oxvba-differential/tests/jit_windows_vm3_parity.rs.

Canonical matrix rows: docs/validation/CORE_COMPILER_VM_JIT_READINESS_MATRIX_V1.csv#CORE-DIFF-STRUCTURAL.

Row contract CORE-DIFF-STRUCTURAL: semantic_subset=arrays records objects ProcRefs identities errors effects lifecycle and balance.

Acceptance commands: command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-9/c_observe.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.2.11; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.2.2 — WIN-1 resolve registered and file-backed x64 type libraries

Outcome: WIN-1 resolve registered and file-backed x64 type libraries.

Effect: delivery. Parent epic: bd-59co.3.2 (WIN-1).

Clauses: COMP-BIND-001|PROJ-REF-001|WIN-META-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.2.2.10, bd-59co.3.1.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Resolve exact GUID/version/LCID/platform/reference order through Registry64 and file provenance.

2. Retain broken/missing/version-conflict diagnostics and host policy; fixture catalogs remain limited to owned fixtures.

3. Test multiple versions, aliasing, reference priority and file/registry disagreement with scoped registry cleanup.

Starting code/truth surfaces: crates/oxvba-com/src/windows_typelib_loader.rs, crates/oxvba-com/src/typelib_catalog.rs.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-TYPELIB-METADATA.

Row contract WAC-TYPELIB-METADATA: semantic_subset=registered/file typelib selection; GUID version LCID platform and reference order; stable library type member event inheritance default source coclass and broken-reference facts; resolver digest and compiler package language-service handoff; process_shape=development-oracle-host; apartment_shape=n/a; exact_signature=versioned RawTypeLibMetadataV1 identity graph and resolver digest; fixture_id=typelib-metadata-identity-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-com.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-1/w_resolve.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.2; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.2.3 — WIN-1 preserve authoritative member event and type identities

Outcome: WIN-1 preserve authoritative member event and type identities.

Effect: delivery. Parent epic: bd-59co.3.2 (WIN-1).

Clauses: COMP-BIND-001|PROJ-REF-001|WIN-META-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.2.2. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Project stable library/type/member/event IDs, inherited vtable slots, DISPIDs, defaults, coclasses and source interfaces.

2. Include enums, aliases, record/array shapes and resolver revision digest without dependency-name special cases.

3. Compare a second independent synthetic typelib and round-trip metadata to prove general discovery.

Starting code/truth surfaces: crates/oxvba-com/src/typelib.rs, crates/oxvba-com/src/typelib_cache.rs, crates/oxvba-symbol/src.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-TYPELIB-METADATA.

Row contract WAC-TYPELIB-METADATA: semantic_subset=registered/file typelib selection; GUID version LCID platform and reference order; stable library type member event inheritance default source coclass and broken-reference facts; resolver digest and compiler package language-service handoff; process_shape=development-oracle-host; apartment_shape=n/a; exact_signature=versioned RawTypeLibMetadataV1 identity graph and resolver digest; fixture_id=typelib-metadata-identity-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-symbol.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-1/w_meta.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.2; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.2.4 — WIN-1 hand authoritative descriptors to compiler packages and hosts

Outcome: WIN-1 hand authoritative descriptors to compiler packages and hosts.

Effect: delivery. Parent epic: bd-59co.3.2 (WIN-1).

Clauses: COMP-BIND-001|PROJ-REF-001|WIN-META-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.2.3, bd-59co.2.4.6, bd-59co.2.6.4. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Carry the same resolver identities and exact signatures through reference loading, binding, packages and build metadata.

2. Expose raw metadata and broken-reference provenance for future language-service consumers without implementing IDE queries.

3. Compare source and packaged import descriptors and invalidate consumers on changed typelib identity.

Starting code/truth surfaces: crates/oxvba-project/src, crates/oxvba-symbol/src, crates/oxvba-oxir/src, crates/oxvba-build/src.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-TYPELIB-METADATA.

Row contract WAC-TYPELIB-METADATA: semantic_subset=registered/file typelib selection; GUID version LCID platform and reference order; stable library type member event inheritance default source coclass and broken-reference facts; resolver digest and compiler package language-service handoff; process_shape=development-oracle-host; apartment_shape=n/a; exact_signature=versioned RawTypeLibMetadataV1 identity graph and resolver digest; fixture_id=typelib-metadata-identity-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-project; command:cargo test -p oxvba-build.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-1/w_handoff.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.2; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.2.5 — WIN-1 prove x64 BSTR VARIANT and numeric carrier layout

Outcome: WIN-1 prove x64 BSTR VARIANT and numeric carrier layout.

Effect: delivery. Parent epic: bd-59co.3.2 (WIN-1).

Clauses: IMAGE-ABI-001|RUNTIME-VALUE-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.2.2.14, bd-59co.3.15.9. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Probe size, alignment, offsets and native read/write behavior against controlled x64 fixtures.

2. Exercise embedded-NUL strings, Boolean/Currency/Date/Decimal, LongPtr and signed-width boundaries.

3. Prove copy/clear/borrow/error ownership; classify existing carrier work as reused only when the exact probe passes.

Starting code/truth surfaces: crates/oxvba-runtime/src, crates/oxvba-com/src.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-BSTR-LAYOUT; docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-VARIANT-LAYOUT; docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-NUMERIC-LONGPTR.

Row contract WAC-BSTR-LAYOUT: semantic_subset=BSTR layout ownership copy move ByRef and cleanup; process_shape=OxVba-process; apartment_shape=n/a; exact_signature=x64 BSTR pointer length and ownership; fixture_id=carrier-bstr-layout-v1; environment_id=win-x64-dev-oracle-2026-07.

Row contract WAC-VARIANT-LAYOUT: semantic_subset=VARIANT subtype union ownership ByRef copy and clear; process_shape=OxVba-process; apartment_shape=n/a; exact_signature=x64 VARIANT and VARIANTARG; fixture_id=carrier-variant-layout-v1; environment_id=win-x64-dev-oracle-2026-07.

Row contract WAC-NUMERIC-LONGPTR: semantic_subset=numeric primitives LongLong LongPtr Currency Date Decimal and Boolean layouts; process_shape=OxVba-process; apartment_shape=n/a; exact_signature=x64 primitive and pointer-width carriers; fixture_id=carrier-numeric-longptr-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-runtime; command:cargo test -p oxvba-com.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-1/w_layout.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.2; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.2.6 — WIN-1 preserve SAFEARRAY shape and nominal interface ownership

Outcome: WIN-1 preserve SAFEARRAY shape and nominal interface ownership.

Effect: delivery. Parent epic: bd-59co.3.2 (WIN-1).

Clauses: IMAGE-ABI-001|RUNTIME-VALUE-001|WIN-META-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.2.7.2, bd-59co.2.7.3, bd-59co.3.2.3, bd-59co.3.15.9. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Preserve rank, signed lower bounds, element stride, VT_DISPATCH/VT_UNKNOWN and nominal interface identity.

2. Integrate object-carrier thread ownership and native SAFEARRAY dimension/allocation contracts from the existing Core beads.

3. Test zero/empty/nonzero-bound arrays, QI failures, clone/clear/ByRef writeback and every partial-construction cleanup edge.

Starting code/truth surfaces: crates/oxvba-runtime/src, crates/oxvba-com/src.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-SAFEARRAY-LAYOUT; docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-INTERFACE-ARRAY.

Row contract WAC-SAFEARRAY-LAYOUT: semantic_subset=SAFEARRAY ranks bounds elements ownership redim and ByRef; process_shape=OxVba-process; apartment_shape=n/a; exact_signature=x64 SAFEARRAY descriptor and element carriers; fixture_id=carrier-safearray-layout-v1; environment_id=win-x64-dev-oracle-2026-07.

Row contract WAC-INTERFACE-ARRAY: semantic_subset=nominal interface and object arrays with VT_DISPATCH VT_UNKNOWN; process_shape=OxVba-process; apartment_shape=n/a; exact_signature=SAFEARRAY of nominal interface pointers; fixture_id=carrier-interface-array-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-runtime; command:cargo test -p oxvba-com.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-1/w_arrays.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.2; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.2.7 — WIN-1 transport nominal VT_RECORD scalars and arrays

Outcome: WIN-1 transport nominal VT_RECORD scalars and arrays.

Effect: delivery. Parent epic: bd-59co.3.2 (WIN-1).

Clauses: IMAGE-ABI-001|RUNTIME-VALUE-001|WIN-META-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.2.3, bd-59co.2.4.4, bd-59co.3.15.9. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Retain record GUID/typeinfo, size, alignment, field layout and IRecordInfo ownership.

2. Marshal record scalars and arrays using canonical runtime storage and reject structurally similar but nominally different records.

3. Test nested fields, strings/interfaces, copy/clear/ByRef and partial failure; compiler size/depth limits remain with their existing bead.

Starting code/truth surfaces: crates/oxvba-runtime/src, crates/oxvba-com/src, crates/oxvba-build/src/oleaut_typelib.rs.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-VT-RECORD.

Row contract WAC-VT-RECORD: semantic_subset=nominal VT_RECORD descriptors scalar arrays copy and clear; process_shape=OxVba-process; apartment_shape=n/a; exact_signature=IRecordInfo-backed VT_RECORD carriers; fixture_id=carrier-vt-record-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-runtime; command:cargo test -p oxvba-com.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-1/w_record.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.2; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.2.8 — WIN-1 preserve IUnknown identity and carrier lifetime across interfaces

Outcome: WIN-1 preserve IUnknown identity and carrier lifetime across interfaces.

Effect: delivery. Parent epic: bd-59co.3.2 (WIN-1).

Clauses: IMAGE-ABI-001|RUNTIME-VALUE-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.2.3, bd-59co.2.7.2, bd-59co.3.15.9. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Normalize identity through IUnknown while retaining exact interface-pointer ownership and apartment affinity.

2. Cover host-returned objects, cross-interface aliases, AddRef/Release and failed QI without fabricated token identities.

3. Test cyclic references only to the extent VBA ownership specifies, explicit release, clone/drop and failure balance.

Starting code/truth surfaces: crates/oxvba-runtime/src, crates/oxvba-com/src.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-IUNKNOWN-IDENTITY.

Row contract WAC-IUNKNOWN-IDENTITY: semantic_subset=IUnknown identity QI AddRef Release and object carrier mapping; process_shape=OxVba-process; apartment_shape=n/a; exact_signature=x64 IUnknown-compatible identity carrier; fixture_id=carrier-iunknown-identity-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-runtime.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-1/w_identity.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.2; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.3.6 — WIN-2 admit Windows plans only from sealed package descriptors

Outcome: WIN-2 admit Windows plans only from sealed package descriptors.

Effect: delivery. Parent epic: bd-59co.3.3 (WIN-2).

Clauses: IMAGE-VERIFY-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.2.4, bd-59co.2.6.5, bd-59co.2.6.3. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Require sealed package/reference provenance when constructing general Windows interop plans.

2. Reject raw or mutated descriptors, mismatched target/carrier/helper ABI and missing metadata identities before elaboration.

3. Prove native services and executable thunks remain untouched on admission failure through both backends.

Starting code/truth surfaces: crates/oxvba-runtime/src/interop_plan.rs, crates/oxvba-oxir/src, crates/oxvba-host/src/engine.rs.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-WINDOWS-DESCRIPTORS.

Row contract WAC-WINDOWS-DESCRIPTORS: semantic_subset=verified OxImage Windows descriptors only; raw metadata or unverified decoded artifacts cannot elaborate a plan; process_shape=OxVba-process; apartment_shape=declared-by-plan; exact_signature=versioned x64 descriptor signature ownership writeback cleanup and capability record; fixture_id=windows-descriptor-boundary-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-runtime interop_plan; command:cargo test -p oxvba-host.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-2/w_verifyplan.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.3; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.3.7 — WIN-2 attach Windows policy apartment and services to shared sessions

Outcome: WIN-2 attach Windows policy apartment and services to shared sessions.

Effect: delivery. Parent epic: bd-59co.3.3 (WIN-2).

Clauses: HOST-SESSION-001|JIT-CACHE-001|JIT-CORE-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|SYS-OWN-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.2.10.3, bd-59co.3.2.4, bd-59co.3.3.4. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Attach COM apartment, native-loader capability and validated interop services to the Core session.

2. Balance COM initialization on the owning thread and reject incompatible apartment/thread use before a call.

3. Test reset/dispose, denied capabilities and multiple concurrent sessions without a private Windows session/cache implementation.

Starting code/truth surfaces: crates/oxvba-host/src/engine.rs, crates/oxvba-hal/src, crates/oxvba-com/src.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-VERIFIED-INTEROP-PLAN.

Row contract WAC-VERIFIED-INTEROP-PLAN: semantic_subset=one plan identity unchanged VM3/JIT adapters common Core session/cache and no private Windows helper catalog mutable state or cache; process_shape=OxVba-process; apartment_shape=declared-by-plan; exact_signature=versioned interop plan descriptor and verifier; fixture_id=interop-plan-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-host; command:cargo test -p oxvba-hal.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-2/w_session.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.3; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.3.8 — WIN-2 verify shared early-vtable call plans

Outcome: WIN-2 verify shared early-vtable call plans.

Effect: delivery. Parent epic: bd-59co.3.3 (WIN-2).

Clauses: HOST-SESSION-001|JIT-CACHE-001|JIT-CORE-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|SYS-OWN-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.3.6, bd-59co.2.7.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Represent exact interface identity, slot, calling convention, parameter directions and HRESULT/retval contract.

2. Seal validated early-call plans and expose identical immutable plan identity to VM3 and JIT.

3. Reject slot/signature/target/ownership mutations before loading or invoking a vtable.

Starting code/truth surfaces: crates/oxvba-runtime/src/interop_plan.rs, crates/oxvba-com/src.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-VERIFIED-INTEROP-PLAN.

Row contract WAC-VERIFIED-INTEROP-PLAN: semantic_subset=one plan identity unchanged VM3/JIT adapters common Core session/cache and no private Windows helper catalog mutable state or cache; process_shape=OxVba-process; apartment_shape=declared-by-plan; exact_signature=versioned interop plan descriptor and verifier; fixture_id=interop-plan-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-runtime interop_plan; command:cargo test -p oxvba-com.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-2/w_earlyplan.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.3; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.3.9 — WIN-2 enforce apartment and synchronous interop reentry policy

Outcome: WIN-2 enforce apartment and synchronous interop reentry policy.

Effect: delivery. Parent epic: bd-59co.3.3 (WIN-2).

Clauses: HOST-SESSION-001|JIT-CACHE-001|JIT-CORE-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|SYS-OWN-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.3.7, bd-59co.2.7.7, bd-59co.2.7.4. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Route callbacks into the owning live session with explicit same-thread, cross-apartment and out-of-process policy.

2. Preserve synchronous return/writeback while using supported COM marshaling for foreign apartment calls.

3. Test nested entry, prohibited thread use, disposal while active and error propagation under both backends.

Starting code/truth surfaces: crates/oxvba-host/src/engine.rs, crates/oxvba-com/src, crates/oxvba-runtime/src/interop_plan.rs.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-VERIFIED-INTEROP-PLAN.

Row contract WAC-VERIFIED-INTEROP-PLAN: semantic_subset=one plan identity unchanged VM3/JIT adapters common Core session/cache and no private Windows helper catalog mutable state or cache; process_shape=OxVba-process; apartment_shape=declared-by-plan; exact_signature=versioned interop plan descriptor and verifier; fixture_id=interop-plan-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-host; command:cargo test -p oxvba-com.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-2/w_reentry.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.3; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.3.10 — WIN-2 verify incoming and outgoing event plan descriptors

Outcome: WIN-2 verify incoming and outgoing event plan descriptors.

Effect: delivery. Parent epic: bd-59co.3.3 (WIN-2).

Clauses: HOST-SESSION-001|JIT-CACHE-001|JIT-CORE-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|SYS-OWN-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.3.6, bd-59co.2.7.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Describe source interface, handler/sink identity, parameter directions and synchronous ByRef copyback order.

2. Define immutable incoming/outgoing descriptors consumed unchanged by both backends.

3. Reject invalid sink signatures and illegal temporary lifetimes; represent cleanup and handler-error edges explicitly.

Starting code/truth surfaces: crates/oxvba-runtime/src/interop_plan.rs, crates/oxvba-com/src.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-VERIFIED-INTEROP-PLAN.

Row contract WAC-VERIFIED-INTEROP-PLAN: semantic_subset=one plan identity unchanged VM3/JIT adapters common Core session/cache and no private Windows helper catalog mutable state or cache; process_shape=OxVba-process; apartment_shape=declared-by-plan; exact_signature=versioned interop plan descriptor and verifier; fixture_id=interop-plan-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-runtime interop_plan; command:cargo test -p oxvba-com.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-2/w_eventplan.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.3; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.3.11 — WIN-2 verify shared late dual and custom serving plans

Outcome: WIN-2 verify shared late dual and custom serving plans.

Effect: delivery. Parent epic: bd-59co.3.3 (WIN-2).

Clauses: HOST-SESSION-001|JIT-CACHE-001|JIT-CORE-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|SYS-OWN-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.2.4.7, bd-59co.3.3.6, bd-59co.2.7.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Bind class factory, exposed interface, vtable/dispatch signature and session entry facts into a serving plan.

2. Preserve one identity across late/early calls and define native-to-VBA argument ownership and error cleanup.

3. Reject incompatible exports before activation and compare plan identity in VM3/JIT serving adapters.

Starting code/truth surfaces: crates/oxvba-runtime/src/interop_plan.rs, crates/oxvba-comhost/src/lib.rs.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-VERIFIED-INTEROP-PLAN.

Row contract WAC-VERIFIED-INTEROP-PLAN: semantic_subset=one plan identity unchanged VM3/JIT adapters common Core session/cache and no private Windows helper catalog mutable state or cache; process_shape=OxVba-process; apartment_shape=declared-by-plan; exact_signature=versioned interop plan descriptor and verifier; fixture_id=interop-plan-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-runtime interop_plan; command:cargo test -p oxvba-comhost.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-2/w_serveplan.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.3; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.3.12 — WIN-2 verify typed callback and retained registration plans

Outcome: WIN-2 verify typed callback and retained registration plans.

Effect: delivery. Parent epic: bd-59co.3.3 (WIN-2).

Clauses: HOST-SESSION-001|JIT-CACHE-001|JIT-CORE-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|SYS-OWN-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.2.4.5, bd-59co.3.3.6, bd-59co.2.7.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Bind callback signature, target procedure, owning session and synchronous/retained lifetime token.

2. Define argument/result/ByRef transport and disposal/error behavior independent of execution backend.

3. Reject signature mismatch, stale generation and unavailable reentry before thunk invocation; reuse the retained-lifetime fixture contract.

Starting code/truth surfaces: crates/oxvba-runtime/src/interop_plan.rs, crates/oxvba-rt-abi/src.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-VERIFIED-INTEROP-PLAN.

Row contract WAC-VERIFIED-INTEROP-PLAN: semantic_subset=one plan identity unchanged VM3/JIT adapters common Core session/cache and no private Windows helper catalog mutable state or cache; process_shape=OxVba-process; apartment_shape=declared-by-plan; exact_signature=versioned interop plan descriptor and verifier; fixture_id=interop-plan-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-runtime interop_plan; command:cargo test -p oxvba-rt-abi.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-2/w_callbackplan.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.3; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.4.4 — WIN-3 resolve activation GetObject and host-returned objects

Outcome: WIN-3 resolve activation GetObject and host-returned objects.

Effect: delivery. Parent epic: bd-59co.3.4 (WIN-3).

Clauses: COM-CLIENT-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.4.2, bd-59co.3.2.4, bd-59co.3.2.8, bd-59co.3.3.7. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Implement general ProgID/CLSID/CLSCTX activation with Registry64 and policy enforcement.

2. Add ROT, moniker and file GetObject routes plus host-returned interface objects.

3. Test in/out-of-proc success, unavailable object, reference reuse and deterministic owned-resource cleanup.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-hal/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_CLIENT_MATRIX_V1.csv#WCC-PLAN-LATE; docs/validation/WINDOWS_JIT_COM_CLIENT_MATRIX_V1.csv#WCC-LATE-OUTPROC-ERROR.

Row contract WCC-PLAN-LATE: semantic_subset=scalar activation invocation and property access; process_shape=inproc-client-and-server; apartment_shape=STA; exact_signature=IDispatch scalar method and property; fixture_id=com-late-scalar-v1; environment_id=win-x64-cert-vm-pending-v1.

Row contract WCC-LATE-OUTPROC-ERROR: semantic_subset=GetObject out-of-proc locale error and reentry behavior; process_shape=out-of-proc-COM-server; apartment_shape=STA-to-proxy; exact_signature=IDispatch error locale and reentrant calls; fixture_id=com-late-outproc-error-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-3/late_activation.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.4; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.4.5 — WIN-3 implement named omitted default and property argument rules

Outcome: WIN-3 implement named omitted default and property argument rules.

Effect: delivery. Parent epic: bd-59co.3.4 (WIN-3).

Clauses: COM-CLIENT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.4.2, bd-59co.2.4.6, bd-59co.3.2.4, bd-59co.3.3.7. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Preserve named/omitted/Optional/ParamArray arguments and exact reverse DISPPARAMS ordering.

2. Handle default members, indexed properties, property-put named DISPID and put versus putref intent.

3. Use a recording IDispatch fixture to verify wire order, argument identities, error positions and cache invalidation.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-jit/src, crates/oxvba-vm3/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_CLIENT_MATRIX_V1.csv#WCC-LATE-ARGS.

Row contract WCC-LATE-ARGS: semantic_subset=named omitted Optional ParamArray put putref and ByRef arguments; process_shape=inproc-client-and-server; apartment_shape=STA; exact_signature=IDispatch mixed argument and property signatures; fixture_id=com-late-arguments-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-3/late_args.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.4; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.4.6 — WIN-3 implement aliased ByRef dispatch writeback

Outcome: WIN-3 implement aliased ByRef dispatch writeback.

Effect: delivery. Parent epic: bd-59co.3.4 (WIN-3).

Clauses: COM-CLIENT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.4.5, bd-59co.3.2.5, bd-59co.3.2.8. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Map in/out/inout and coerced temporaries to canonical addressable carriers.

2. Apply writeback in specified order before returning to VBA, including repeated aliases and callee failure.

3. Probe scalar/string/object ByRef success, type mismatch, duplicate aliases and pin/temporary balance.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-runtime/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_CLIENT_MATRIX_V1.csv#WCC-LATE-ARGS.

Row contract WCC-LATE-ARGS: semantic_subset=named omitted Optional ParamArray put putref and ByRef arguments; process_shape=inproc-client-and-server; apartment_shape=STA; exact_signature=IDispatch mixed argument and property signatures; fixture_id=com-late-arguments-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-3/late_byref.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.4; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.4.7 — WIN-3 support structural values chaining and enumeration

Outcome: WIN-3 support structural values chaining and enumeration.

Effect: delivery. Parent epic: bd-59co.3.4 (WIN-3).

Clauses: COM-CLIENT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.4.6, bd-59co.3.2.6, bd-59co.3.2.7, bd-59co.3.2.8, bd-59co.2.11.6. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Pass and return arrays, records and interfaces through the shared dispatch plan.

2. Preserve identity through chained/default-member results and implement IEnumVARIANT For Each ownership.

3. Test multidimensional/nonzero-bound arrays, nominal records, early loop exit and enumerator error cleanup with structural comparisons.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-vm3/src, crates/oxvba-jit/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_CLIENT_MATRIX_V1.csv#WCC-LATE-STRUCTURAL.

Row contract WCC-LATE-STRUCTURAL: semantic_subset=object identity arrays records chaining and enumeration; process_shape=inproc-client-and-server; apartment_shape=STA; exact_signature=IDispatch structural values and IEnumVARIANT; fixture_id=com-late-structural-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-3/late_struct.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.4; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.4.8 — WIN-3 preserve COM errors LCID and dispatch lifetime

Outcome: WIN-3 preserve COM errors LCID and dispatch lifetime.

Effect: delivery. Parent epic: bd-59co.3.4 (WIN-3).

Clauses: COM-CLIENT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.4.4, bd-59co.3.4.6, bd-59co.3.3.7, bd-59co.2.11.6. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Map HRESULT, deferred EXCEPINFO and IErrorInfo into full VBA Err without confusing host denial with VBA failure.

2. Thread LCID and locale-dependent conversions through lookup/invoke under explicit policy.

3. Test a failing out-of-proc server, repeated activate/invoke/release and stale DISPID invalidation with zero resource drift.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-host/src/engine.rs.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_CLIENT_MATRIX_V1.csv#WCC-LATE-OUTPROC-ERROR.

Row contract WCC-LATE-OUTPROC-ERROR: semantic_subset=GetObject out-of-proc locale error and reentry behavior; process_shape=out-of-proc-COM-server; apartment_shape=STA-to-proxy; exact_signature=IDispatch error locale and reentrant calls; fixture_id=com-late-outproc-error-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-3/late_error.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.4; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.5.2 — WIN-4 execute scalar and property calls through native vtables

Outcome: WIN-4 execute scalar and property calls through native vtables.

Effect: delivery. Parent epic: bd-59co.3.5 (WIN-4).

Clauses: COM-CLIENT-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.3.8, bd-59co.3.2.8, bd-59co.2.4.6, bd-59co.3.15.9, bd-59co.2.11.6, bd-59co.3.3.7. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Lower ComCallEarly to exact inherited vtable slots using the shared plan and x64 ABI.

2. Implement scalar methods/properties, HRESULT/out-retval and optional/default semantics.

3. Instrument a dual fixture and assert vtable traffic with zero IDispatch fallback under both backends.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-vm3/src, crates/oxvba-jit/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_CLIENT_MATRIX_V1.csv#WCC-PLAN-EARLY.

Row contract WCC-PLAN-EARLY: semantic_subset=scalar typed native-vtable invocation; process_shape=inproc-client-and-server; apartment_shape=STA; exact_signature=typed scalar native vtable method and property; fixture_id=com-early-scalar-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/early_scalar.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.5; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.5.3 — WIN-4 marshal complex and ByRef early-bound signatures

Outcome: WIN-4 marshal complex and ByRef early-bound signatures.

Effect: delivery. Parent epic: bd-59co.3.5 (WIN-4).

Clauses: COM-CLIENT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.5.2, bd-59co.3.2.6, bd-59co.3.2.7, bd-59co.3.2.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Implement BSTR/VARIANT/SAFEARRAY/record/interface and in/out/inout transport from authoritative descriptors.

2. Preserve aliased ByRef storage, retval ownership and exact QI identity.

3. Test conversion failure, partial initialization and returned interface release while proving native-vtable transport.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-runtime/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_CLIENT_MATRIX_V1.csv#WCC-EARLY-COMPLEX.

Row contract WCC-EARLY-COMPLEX: semantic_subset=ByRef arrays records interfaces inheritance and QI identity; process_shape=inproc-client-and-server; apartment_shape=STA; exact_signature=typed Automation vtable complex signatures; fixture_id=com-early-complex-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/early_complex.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.5; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.5.4 — WIN-4 execute custom inherited interfaces without dispatch fallback

Outcome: WIN-4 execute custom inherited interfaces without dispatch fallback.

Effect: delivery. Parent epic: bd-59co.3.5 (WIN-4).

Clauses: COM-CLIENT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.5.3. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Support general non-Automation interface descriptors, inheritance and custom typed returns.

2. Validate ABI shape and interface selection before native invocation.

3. Use independently described custom interfaces to prove the path is signature-driven, including failed QI and native error cleanup.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-jit/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_CLIENT_MATRIX_V1.csv#WCC-EARLY-CUSTOM.

Row contract WCC-EARLY-CUSTOM: semantic_subset=custom interface signatures HRESULT out retval and inheritance; process_shape=inproc-client-and-server; apartment_shape=STA; exact_signature=custom interface native vtable signature; fixture_id=com-early-custom-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/early_custom.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.5; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.5.5 — WIN-4 preserve early transport through apartment and process proxies

Outcome: WIN-4 preserve early transport through apartment and process proxies.

Effect: delivery. Parent epic: bd-59co.3.5 (WIN-4).

Clauses: COM-CLIENT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.5.4, bd-59co.3.3.9. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Use Automation or declared proxy/stub marshaling appropriate to the interface without local-pointer shortcuts.

2. Preserve inherited interface identity and synchronous outputs across apartment/process boundaries.

3. Test STA/MTA and local-server cases with owned proxy/server lifecycle and explicit transport evidence.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-host/src/engine.rs.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_CLIENT_MATRIX_V1.csv#WCC-EARLY-OUTPROC.

Row contract WCC-EARLY-OUTPROC: semantic_subset=out-of-proc typed proxy marshalling and apartment behavior; process_shape=out-of-proc-COM-server; apartment_shape=STA-to-proxy; exact_signature=typed proxy Automation and custom signatures; fixture_id=com-early-outproc-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/early_proxy.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.5; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.6.5 — WIN-5 bind source interfaces and own Automation event subscriptions

Outcome: WIN-5 bind source interfaces and own Automation event subscriptions.

Effect: delivery. Parent epic: bd-59co.3.6 (WIN-5).

Clauses: COM-EVENT-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.2.4.6, bd-59co.3.3.10, bd-59co.3.4.2, bd-59co.3.3.9. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Discover authoritative connection points and bind typed WithEvents handlers.

2. Build IDispatch sinks with owned Advise cookies and deterministic replacement/unsubscribe/termination.

3. Test scalar ByVal event order, multiple listeners, failed Advise, handler replacement and teardown during dispatch.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-host/src/engine.rs.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_EVENTS_MATRIX_V1.csv#WCE-PLAN-INCOMING; docs/validation/WINDOWS_JIT_COM_EVENTS_MATRIX_V1.csv#WCE-INCOMING-LIFECYCLE.

Row contract WCE-PLAN-INCOMING: semantic_subset=synchronous cancellable scalar ByRef event; process_shape=inproc-source; apartment_shape=STA; exact_signature=source interface scalar ByRef cancellation; fixture_id=com-event-incoming-byref-v1; environment_id=win-x64-cert-vm-pending-v1.

Row contract WCE-INCOMING-LIFECYCLE: semantic_subset=fan-out replacement unsubscribe termination and handler errors; process_shape=inproc-source; apartment_shape=STA; exact_signature=connection-point lifecycle and handler error paths; fixture_id=com-event-incoming-lifecycle-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-host.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-5/event_sink.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.6; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.6.6 — WIN-5 preserve structural incoming event arguments and ByRef values

Outcome: WIN-5 preserve structural incoming event arguments and ByRef values.

Effect: delivery. Parent epic: bd-59co.3.6 (WIN-5).

Clauses: COM-EVENT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.6.2, bd-59co.3.6.3, bd-59co.3.2.6, bd-59co.3.2.7. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Extend synchronous event delivery to object/interface/array/record arguments and aliased ByRef cells.

2. Keep source-owned storage alive until handler completion and copy back before native Invoke returns.

3. Compare mutation journals and full Err on successful and failing nested handlers with resource balance.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-vm3/src, crates/oxvba-jit/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_EVENTS_MATRIX_V1.csv#WCE-INCOMING-COMPLEX.

Row contract WCE-INCOMING-COMPLEX: semantic_subset=object interface array and record event arguments; process_shape=inproc-source; apartment_shape=STA; exact_signature=typed source interface structural arguments; fixture_id=com-event-incoming-complex-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-5/event_complex.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.6; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.6.7 — WIN-5 receive custom-vtable source events

Outcome: WIN-5 receive custom-vtable source events.

Effect: delivery. Parent epic: bd-59co.3.6 (WIN-5).

Clauses: COM-EVENT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.6.5, bd-59co.3.3.10, bd-59co.3.5.4, bd-59co.3.3.9. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Generate signature-driven custom-vtable sink entries from source-interface metadata.

2. Route sink entries through the same synchronous session/event plan and contain all foreign-boundary unwinds.

3. Prove scalar and complex custom event arguments, QI identity and failed-handler cleanup without IDispatch substitution.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-rt-abi/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_EVENTS_MATRIX_V1.csv#WCE-INCOMING-CUSTOM.

Row contract WCE-INCOMING-CUSTOM: semantic_subset=custom native-vtable source-interface sink; process_shape=inproc-source; apartment_shape=STA; exact_signature=custom source-interface vtable event; fixture_id=com-event-incoming-custom-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-5/event_custom.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.6; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.6.8 — WIN-5 preserve synchronous events across apartments and teardown

Outcome: WIN-5 preserve synchronous events across apartments and teardown.

Effect: delivery. Parent epic: bd-59co.3.6 (WIN-5).

Clauses: COM-EVENT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.6.6, bd-59co.3.6.7, bd-59co.3.3.9. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Implement required same-STA, cross-apartment and out-of-proc source delivery with synchronous cancellation.

2. Preserve ordering/fan-out and define handler failure, reentrant unsubscribe and disposal while callback is active.

3. Prove native caller observes final ByRef values and no owned sink, cookie or callback survives disposal.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-host/src/engine.rs.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_EVENTS_MATRIX_V1.csv#WCE-INCOMING-APARTMENT; docs/validation/WINDOWS_JIT_COM_EVENTS_MATRIX_V1.csv#WCE-INCOMING-LIFECYCLE.

Row contract WCE-INCOMING-APARTMENT: semantic_subset=cross-apartment and out-of-proc synchronous reentry; process_shape=out-of-proc-event-source; apartment_shape=STA-cross-apartment-proxy; exact_signature=proxy-delivered cancellable event; fixture_id=com-event-incoming-apartment-v1; environment_id=win-x64-cert-vm-pending-v1.

Row contract WCE-INCOMING-LIFECYCLE: semantic_subset=fan-out replacement unsubscribe termination and handler errors; process_shape=inproc-source; apartment_shape=STA; exact_signature=connection-point lifecycle and handler error paths; fixture_id=com-event-incoming-lifecycle-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-5/event_apartment.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.6; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.7.2 — WIN-6 activate scalar COM classes through VM3 and JIT sessions

Outcome: WIN-6 activate scalar COM classes through VM3 and JIT sessions.

Effect: delivery. Parent epic: bd-59co.3.7 (WIN-6).

Clauses: COM-SERVE-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.3.11, bd-59co.3.3.7, bd-59co.2.4.7, bd-59co.3.15.9. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Implement stable CLSID/ProgID/version identity, class factories and IUnknown/IDispatch/typeinfo over selected backend sessions.

2. Route exported scalar methods/properties through verified serving plans without project-specific Rust code.

3. Test native client activation, Class_Initialize/Terminate, explicit JIT selection and correct AddRef/Release.

Starting code/truth surfaces: crates/oxvba-comhost/src/lib.rs, crates/oxvba-com/src, crates/oxvba-build/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_SERVER_MATRIX_V1.csv#WCS-LATE-INPROC.

Row contract WCS-LATE-INPROC: semantic_subset=same verified serving plan with late dispatch and VM3/JIT plan-digest equality; process_shape=inproc-COM-server; apartment_shape=STA; exact_signature=IDispatch methods properties defaults and enumeration; fixture_id=com-server-late-inproc-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-comhost; command:cargo test -p oxvba-build.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-6/serve_scalar.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.7; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.7.3 — WIN-6 expose default named optional and enumeration dispatch contracts

Outcome: WIN-6 expose default named optional and enumeration dispatch contracts.

Effect: delivery. Parent epic: bd-59co.3.7 (WIN-6).

Clauses: COM-SERVE-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.7.2, bd-59co.3.4.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Expose property groups, defaults, named/omitted/Optional/ParamArray and enumeration through general IDispatch.

2. Map incoming parameter order, flags and argument errors into typed VBA invocation.

3. Compare native and VBA clients, typeinfo signatures and resulting side effects under both served backends.

Starting code/truth surfaces: crates/oxvba-comhost/src/lib.rs, crates/oxvba-com/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_SERVER_MATRIX_V1.csv#WCS-LATE-INPROC; docs/validation/WINDOWS_JIT_COM_SERVER_MATRIX_V1.csv#WCS-LATE-COMPLEX.

Row contract WCS-LATE-INPROC: semantic_subset=same verified serving plan with late dispatch and VM3/JIT plan-digest equality; process_shape=inproc-COM-server; apartment_shape=STA; exact_signature=IDispatch methods properties defaults and enumeration; fixture_id=com-server-late-inproc-v1; environment_id=win-x64-cert-vm-pending-v1.

Row contract WCS-LATE-COMPLEX: semantic_subset=Optional named ParamArray ByRef arrays records interfaces errors and lifecycle; process_shape=inproc-and-localserver; apartment_shape=STA-and-proxy; exact_signature=IDispatch complex Automation signatures; fixture_id=com-server-late-complex-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-comhost; command:cargo test -p oxvba-com.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-6/serve_args.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.7; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.7.4 — WIN-6 serve structural values ByRef and full COM error state

Outcome: WIN-6 serve structural values ByRef and full COM error state.

Effect: delivery. Parent epic: bd-59co.3.7 (WIN-6).

Clauses: COM-SERVE-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.7.3, bd-59co.3.2.6, bd-59co.3.2.7, bd-59co.3.2.8, bd-59co.3.3.9. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Support object/interface/array/record inputs and outputs plus synchronous ByRef writeback.

2. Map VBA errors to HRESULT/EXCEPINFO/IErrorInfo and preserve source/description/help details.

3. Test partial conversion failure, repeated aliases and nested client calls with exact resource and state ownership.

Starting code/truth surfaces: crates/oxvba-comhost/src/lib.rs, crates/oxvba-com/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_SERVER_MATRIX_V1.csv#WCS-LATE-COMPLEX.

Row contract WCS-LATE-COMPLEX: semantic_subset=Optional named ParamArray ByRef arrays records interfaces errors and lifecycle; process_shape=inproc-and-localserver; apartment_shape=STA-and-proxy; exact_signature=IDispatch complex Automation signatures; fixture_id=com-server-late-complex-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-comhost; command:cargo test -p oxvba-com.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-6/serve_complex.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.7; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.7.5 — WIN-6 run COM LocalServer activation marshaling and shutdown

Outcome: WIN-6 run COM LocalServer activation marshaling and shutdown.

Effect: delivery. Parent epic: bd-59co.3.7 (WIN-6).

Clauses: COM-SERVE-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.7.2, bd-59co.3.3.9. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Implement owned LocalServer32 registration, class-object registration, message loop, locks and shutdown.

2. Use the required marshaling strategy and retain session state across external activation/calls.

3. Test multiple clients, idle shutdown, failed startup and restart without leaking processes or altering unrelated registry values.

Starting code/truth surfaces: crates/oxvba-comhost/src/lib.rs, crates/oxvba-build/src, crates/oxvba-com/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_SERVER_MATRIX_V1.csv#WCS-LATE-LOCALSERVER.

Row contract WCS-LATE-LOCALSERVER: semantic_subset=out-of-process LocalServer activation apartment and shutdown; process_shape=out-of-proc-LocalServer32; apartment_shape=STA-message-loop-and-proxy; exact_signature=IDispatch proxy methods properties and lifecycle; fixture_id=com-server-late-local-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-comhost; command:cargo test -p oxvba-build.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-6/serve_local.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.7; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.7.6 — WIN-6 prove class factory unload and registration lifetime

Outcome: WIN-6 prove class factory unload and registration lifetime.

Effect: delivery. Parent epic: bd-59co.3.7 (WIN-6).

Clauses: COM-SERVE-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.7.4, bd-59co.3.7.5, bd-59co.3.1.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Tie class locks, objects, callbacks and metadata to explicit unload eligibility.

2. Prove repeated register/activate/invoke/release/unregister, DllCanUnloadNow and failed initialization behavior.

3. Inspect loader-lock paths and keep compilation, COM initialization and blocking work outside DllMain.

Starting code/truth surfaces: crates/oxvba-comhost/src/lib.rs, crates/oxvba-build/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_SERVER_MATRIX_V1.csv#WCS-SERVER-SAFETY.

Row contract WCS-SERVER-SAFETY: semantic_subset=class factories locks registration unload errors and hostile clients; process_shape=inproc-and-out-of-proc; apartment_shape=STA-and-proxy; exact_signature=class factory and server lifecycle contracts; fixture_id=com-server-lifecycle-safety-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-comhost; command:cargo test -p oxvba-build.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-6/serve_lifetime.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.7; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.8.2 — WIN-7 generate dual vtables sharing the served object identity

Outcome: WIN-7 generate dual vtables sharing the served object identity.

Effect: delivery. Parent epic: bd-59co.3.8 (WIN-7).

Clauses: COM-SERVE-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.7.2, bd-59co.3.3.11, bd-59co.3.5.2, bd-59co.2.7.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Generate native entry thunks from verified exported signatures and select the requested backend.

2. Share IUnknown identity, project state and lifecycle between late and early interfaces.

3. Test mixed early/late clients and instrument vtable calls to prove no hidden IDispatch fallback.

Starting code/truth surfaces: crates/oxvba-comhost/src/lib.rs, crates/oxvba-com/src, crates/oxvba-build/src/oleaut_typelib.rs.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_SERVER_MATRIX_V1.csv#WCS-DUAL-INPROC.

Row contract WCS-DUAL-INPROC: semantic_subset=same verified serving plan with early/dual vtable VM3/JIT plan-digest equality and no dispatch fallback; process_shape=inproc-COM-server; apartment_shape=STA; exact_signature=generated dual interface scalar and property vtables; fixture_id=com-server-dual-inproc-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-comhost; command:cargo test -p oxvba-build.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-7/dual_scalar.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.8; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.8.3 — WIN-7 serve imported Implements and complex custom interfaces

Outcome: WIN-7 serve imported Implements and complex custom interfaces.

Effect: delivery. Parent epic: bd-59co.3.8 (WIN-7).

Clauses: COM-SERVE-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.8.2, bd-59co.2.4.7, bd-59co.3.7.4, bd-59co.3.5.4. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Implement imported custom interface obligations and inherited QI relations from exact compiler facts.

2. Generate complex array/record/interface/ByRef entries with correct HRESULT and ownership handling.

3. Compare emitted typelib signatures to exposed thunks and test failure paths from independently compiled native clients.

Starting code/truth surfaces: crates/oxvba-comhost/src/lib.rs, crates/oxvba-build/src, crates/oxvba-com/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_SERVER_MATRIX_V1.csv#WCS-IMPLEMENTS-CUSTOM.

Row contract WCS-IMPLEMENTS-CUSTOM: semantic_subset=imported Implements and custom interface vtables; process_shape=inproc-COM-server; apartment_shape=STA; exact_signature=imported custom interface generated entry thunks; fixture_id=com-server-implements-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-comhost; command:cargo test -p oxvba-build.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-7/dual_custom.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.8; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.8.4 — WIN-7 package and prove early serving across process boundaries

Outcome: WIN-7 package and prove early serving across process boundaries.

Effect: delivery. Parent epic: bd-59co.3.8 (WIN-7).

Clauses: COM-SERVE-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.8.3, bd-59co.3.7.5, bd-59co.3.5.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Supply matching typelib and required proxy/stub artifacts for dual/custom signatures.

2. Register only owned x64 artifacts and preserve one served identity through remote interfaces.

3. Run early-bound native and Excel clients with process restart, error, ByRef and clean uninstall evidence.

Starting code/truth surfaces: crates/oxvba-build/src, crates/oxvba-comhost/src/lib.rs.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_SERVER_MATRIX_V1.csv#WCS-EARLY-OUTPROC.

Row contract WCS-EARLY-OUTPROC: semantic_subset=out-of-process early proxy marshalling arrays records interfaces and ByRef; process_shape=out-of-proc-LocalServer32; apartment_shape=STA-to-proxy; exact_signature=dual and custom proxy signatures; fixture_id=com-server-early-outproc-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-build; command:cargo test -p oxvba-comhost.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-7/dual_proxy.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.8; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.9.2 — WIN-8 expose connection points and invoke native event sinks

Outcome: WIN-8 expose connection points and invoke native event sinks.

Effect: delivery. Parent epic: bd-59co.3.9 (WIN-8).

Clauses: COM-EVENT-001|COM-SERVE-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.7.2, bd-59co.3.3.10, bd-59co.3.3.9, bd-59co.3.6.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Expose source interfaces and connection-point containers with owned sink cookies.

2. Route RaiseEvent through the verified outgoing plan for VM3/JIT served sessions.

3. Test ordered multi-sink delivery, sink errors, unsubscribe during notification and zero cookie/refcount drift.

Starting code/truth surfaces: crates/oxvba-comhost/src/lib.rs, crates/oxvba-com/src, crates/oxvba-jit/src, crates/oxvba-vm3/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_EVENTS_MATRIX_V1.csv#WCE-PLAN-OUTGOING.

Row contract WCE-PLAN-OUTGOING: semantic_subset=served source-interface scalar event fan-out; process_shape=inproc-served-object; apartment_shape=STA; exact_signature=served source-interface scalar event; fixture_id=com-event-outgoing-scalar-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-comhost; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-8/outgoing_scalar.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.9; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.9.3 — WIN-8 preserve outgoing structural and synchronous ByRef events

Outcome: WIN-8 preserve outgoing structural and synchronous ByRef events.

Effect: delivery. Parent epic: bd-59co.3.9 (WIN-8).

Clauses: COM-EVENT-001|COM-SERVE-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.9.2, bd-59co.3.7.4, bd-59co.3.8.3, bd-59co.3.6.6. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Marshal object/array/record/interface and ByRef event arguments to native and VBA sinks.

2. Preserve sink ordering, alias/writeback semantics and handler error disposition across apartments.

3. Test native and Excel WithEvents consumers, cancellation, nested RaiseEvent and teardown through both backends.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-comhost/src/lib.rs.

Canonical matrix rows: docs/validation/WINDOWS_JIT_COM_EVENTS_MATRIX_V1.csv#WCE-OUTGOING-COMPLEX.

Row contract WCE-OUTGOING-COMPLEX: semantic_subset=object array ByRef errors lifecycle and Excel consumption; process_shape=inproc-and-localserver-served-object; apartment_shape=STA-and-proxy; exact_signature=served structural and ByRef source-interface events; fixture_id=com-event-outgoing-complex-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-comhost; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-8/outgoing_complex.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.9; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.10.4 — WIN-9 resolve DLL entries ordinals and policy failures exactly

Outcome: WIN-9 resolve DLL entries ordinals and policy failures exactly.

Effect: delivery. Parent epic: bd-59co.3.10 (WIN-9).

Clauses: NATIVE-IMPORT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.10.2, bd-59co.2.4.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Separate native ABI/invocation ownership from HAL capability policy and use explicit secure DLL search.

2. Resolve aliases, named exports and ordinals while distinguishing missing module, missing entry and denied capability.

3. Capture LastDllError immediately after invocation before cleanup; test duplicate load/release and hostile search-path neighbors.

Starting code/truth surfaces: crates/oxvba-hal/src, crates/oxvba-com/src, crates/oxvba-host/tests/native_declare_lane_vm3.rs.

Canonical matrix rows: docs/validation/WINDOWS_JIT_NATIVE_IMPORT_MATRIX_V1.csv#WNI-DECLARE-LOADER-ERROR.

Row contract WNI-DECLARE-LOADER-ERROR: semantic_subset=ordinal resolution missing library export policy and LastDllError; process_shape=inproc-native-DLL; apartment_shape=STA; exact_signature=named ordinal missing and denied native entries; fixture_id=native-declare-loader-error-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-hal; command:cargo test -p oxvba-host --test native_declare_lane_vm3.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-9/native_loader.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.10; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.10.5 — WIN-9 preserve scalar ABI returns and aliased ByRef arguments

Outcome: WIN-9 preserve scalar ABI returns and aliased ByRef arguments.

Effect: delivery. Parent epic: bd-59co.3.10 (WIN-9).

Clauses: NATIVE-IMPORT-001|WIN-PLAN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.10.2, bd-59co.2.4.5, bd-59co.3.2.5, bd-59co.2.11.6, bd-59co.3.3.7, bd-59co.3.10.4. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Exercise x64 register/stack placement and scalar return classes with deterministic native fixtures.

2. Preserve ByVal/ByRef, repeated aliases and copyback order including Boolean/Currency/Date/LongPtr edge cases.

3. Compare full Err/LastDllError and native mutation journals, not just scalar tags or success status.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-hal/src, crates/oxvba-jit/src, crates/oxvba-vm3/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_NATIVE_IMPORT_MATRIX_V1.csv#WNI-PLAN-DECLARE.

Row contract WNI-PLAN-DECLARE: semantic_subset=scalar named-symbol x64 call; process_shape=inproc-native-DLL; apartment_shape=STA; exact_signature=PtrSafe scalar function and sub signatures; fixture_id=native-declare-scalar-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-9/native_byref.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.10; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.10.6 — WIN-9 marshal ANSI Wide BSTR and mutable native buffers

Outcome: WIN-9 marshal ANSI Wide BSTR and mutable native buffers.

Effect: delivery. Parent epic: bd-59co.3.10 (WIN-9).

Clauses: NATIVE-IMPORT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.10.4, bd-59co.3.10.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Implement declaration-specific string/buffer conversions with explicit pin/temporary ownership.

2. Preserve embedded NUL, input length, fixed buffers, terminators and ByRef copyback without assuming UTF-8.

3. Test locale-dependent ANSI behavior, truncation, error cleanup and LastDllError capture before conversion helpers.

Starting code/truth surfaces: crates/oxvba-host/tests/native_declare_string_marshalling_end_to_end.rs, crates/oxvba-hal/src, crates/oxvba-runtime/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_NATIVE_IMPORT_MATRIX_V1.csv#WNI-DECLARE-STRINGS.

Row contract WNI-DECLARE-STRINGS: semantic_subset=ANSI Wide BSTR and mutable string-buffer calls; process_shape=inproc-native-DLL; apartment_shape=STA; exact_signature=PtrSafe string pointer and buffer signatures; fixture_id=native-declare-strings-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-host --test native_declare_string_marshalling_end_to_end; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-9/native_strings.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.10; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.10.7 — WIN-9 marshal arrays UDTs and As Any using exact x64 storage

Outcome: WIN-9 marshal arrays UDTs and As Any using exact x64 storage.

Effect: delivery. Parent epic: bd-59co.3.10 (WIN-9).

Clauses: NATIVE-IMPORT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.10.5, bd-59co.3.2.6, bd-59co.3.2.7, bd-59co.2.4.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Distinguish array descriptors, data pointers, UDT layout and As Any addressability according to declaration semantics.

2. Preserve nested layout and ByRef ownership while rejecting illegal or unsupported source shapes at the correct phase.

3. Use native layout/mutation probes for nonzero bounds, records with owned fields, aliasing and failure cleanup.

Starting code/truth surfaces: crates/oxvba-hal/src, crates/oxvba-runtime/src, crates/oxvba-bind/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_NATIVE_IMPORT_MATRIX_V1.csv#WNI-DECLARE-STRUCTURAL.

Row contract WNI-DECLARE-STRUCTURAL: semantic_subset=arrays UDT As Any aliases and structural ByRef calls; process_shape=inproc-native-DLL; apartment_shape=STA; exact_signature=PtrSafe array UDT As Any and alias signatures; fixture_id=native-declare-structural-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-hal; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-9/native_struct.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.10; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.11.5 — WIN-10 expose exact VarPtr StrPtr and ObjPtr addressable storage

Outcome: WIN-10 expose exact VarPtr StrPtr and ObjPtr addressable storage.

Effect: delivery. Parent epic: bd-59co.3.11 (WIN-10).

Clauses: NATIVE-IMPORT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.2.5, bd-59co.3.2.6, bd-59co.3.2.8, bd-59co.2.4.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Return addresses of the exact requested variable, BSTR payload or interface storage with VBA-compatible null behavior.

2. Keep addressability and lifetimes explicit across move, resize, reentrant call and temporary coercion.

3. Probe native reads/writes against VM3/JIT storage, including empty strings, object identity and invalidated pointers without dereferencing stale storage.

Starting code/truth surfaces: crates/oxvba-runtime/src, crates/oxvba-jit/src, crates/oxvba-vm3/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_NATIVE_IMPORT_MATRIX_V1.csv#WNI-POINTER-HELPERS.

Row contract WNI-POINTER-HELPERS: semantic_subset=VarPtr StrPtr ObjPtr addressability and LongPtr results; process_shape=OxVba-process; apartment_shape=STA; exact_signature=x64 pointer helper intrinsic signatures; fixture_id=native-pointer-helpers-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-runtime; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-10/ptr_storage.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.11; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.11.6 — WIN-10 invoke typed synchronous AddressOf callbacks

Outcome: WIN-10 invoke typed synchronous AddressOf callbacks.

Effect: delivery. Parent epic: bd-59co.3.11 (WIN-10).

Clauses: NATIVE-IMPORT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.3.12, bd-59co.3.3.9, bd-59co.3.10.5, bd-59co.2.4.5. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Generate x64 callback thunks for legal procedure signatures and invoke the owning live session.

2. Preserve arguments, return values and ByRef cells through native-to-VBA-to-native calls.

3. Test scalar, pointer and structural signature breadth, handler errors and callback return before native caller resumes.

Starting code/truth surfaces: crates/oxvba-rt-abi/src, crates/oxvba-jit/src, crates/oxvba-vm3/src.

Canonical matrix rows: docs/validation/WINDOWS_JIT_NATIVE_IMPORT_MATRIX_V1.csv#WNI-CALLBACK-SYNC.

Row contract WNI-CALLBACK-SYNC: semantic_subset=typed synchronous scalar and ByRef callback; process_shape=inproc-native-DLL; apartment_shape=STA; exact_signature=typed synchronous callback thunk; fixture_id=native-callback-sync-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-rt-abi; command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-10/callback_sync.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.11; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.11.7 — WIN-10 preserve nested native COM and VBA callback cycles

Outcome: WIN-10 preserve nested native COM and VBA callback cycles.

Effect: delivery. Parent epic: bd-59co.3.11 (WIN-10).

Clauses: NATIVE-IMPORT-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.11.6, bd-59co.3.11.2, bd-59co.3.11.3, bd-59co.3.6.8. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Exercise native-to-VBA-to-COM-to-VBA nesting using the same session/reentry contract.

2. Preserve each invocation frame, Err state, aliases and retained callback generation without global mutable routing.

3. Test permitted apartment transitions, forbidden-thread entry, recursion exhaustion and disposal with an active callback.

Starting code/truth surfaces: crates/oxvba-rt-abi/src, crates/oxvba-com/src, crates/oxvba-host/src/engine.rs.

Canonical matrix rows: docs/validation/WINDOWS_JIT_NATIVE_IMPORT_MATRIX_V1.csv#WNI-CALLBACK-NESTED.

Row contract WNI-CALLBACK-NESTED: semantic_subset=nested native COM VBA reentry errors threads and disposed-session safety; process_shape=inproc-native-and-COM-cycle; apartment_shape=STA-and-rejected-wrong-thread; exact_signature=typed nested callback and error signatures; fixture_id=native-callback-nested-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-differential --test jit_windows_vm3_parity.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-10/callback_nested.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.11; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.12.2 — WIN-11 run verified images in JIT WrapperExe artifacts

Outcome: WIN-11 run verified images in JIT WrapperExe artifacts.

Effect: delivery. Parent epic: bd-59co.3.12 (WIN-11).

Clauses: BUILD-CLASS-001|BUILD-PACKAGE-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.2.6.5, bd-59co.2.10.3, bd-59co.3.3.7. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Package a verified embedded or deployed OxImage with explicit entry, target, ABI and backend manifest.

2. Initialize a persistent JIT session and retain source-map/runtime error behavior in the wrapper.

3. Run a separate process without development compiler dependencies and reject corrupted or incompatible images before execution.

Starting code/truth surfaces: crates/oxvba-build/src, crates/oxvba-host/src/engine.rs.

Canonical matrix rows: docs/validation/WINDOWS_NATIVE_EXPORT_AND_PACKAGING_MATRIX_V1.csv#WNE-WRAPPER-EXE.

Row contract WNE-WRAPPER-EXE: semantic_subset=standalone WrapperExe with embedded or deployed verified OxImage; process_shape=standalone; apartment_shape=n/a; exact_signature=generated WrapperExe entry to JIT session; fixture_id=output-wrapper-exe-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-build; command:cargo test -p oxvba-host.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-11/wrapper_exe.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.12; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.12.3 — WIN-11 initialize and unload JIT WrapperLibrary outside loader lock

Outcome: WIN-11 initialize and unload JIT WrapperLibrary outside loader lock.

Effect: delivery. Parent epic: bd-59co.3.12 (WIN-11).

Clauses: BUILD-CLASS-001|BUILD-PACKAGE-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.12.2, bd-59co.2.7.6. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Expose explicit or safe lazy initialization and versioned wrapper entry contracts.

2. Ensure DllMain performs no JIT compilation, COM initialization or blocking work.

3. Instrument load/invoke/reset/free cycles, concurrent entry policy and initialization failure from native clients.

Starting code/truth surfaces: crates/oxvba-build/src, crates/oxvba-host/src/engine.rs.

Canonical matrix rows: docs/validation/WINDOWS_NATIVE_EXPORT_AND_PACKAGING_MATRIX_V1.csv#WNE-WRAPPER-LIBRARY.

Row contract WNE-WRAPPER-LIBRARY: semantic_subset=WrapperLibrary external ABI and repeated session lifecycle; process_shape=inproc-wrapper-library; apartment_shape=caller-declared; exact_signature=generated wrapper library entry ABI; fixture_id=output-wrapper-library-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-build; command:cargo test -p oxvba-host.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-11/wrapper_lib.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.12; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.12.4 — WIN-11 select JIT sessions in reusable WrappedComServer artifacts

Outcome: WIN-11 select JIT sessions in reusable WrappedComServer artifacts.

Effect: delivery. Parent epic: bd-59co.3.12 (WIN-11).

Clauses: BUILD-CLASS-001|BUILD-PACKAGE-001|PROFILE-TOOL-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.12.3, bd-59co.3.7.6, bd-59co.3.8.3, bd-59co.3.9.3. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Select VM3/JIT from verified package policy and use the common session/cache instead of a private comhost runtime.

2. Package matching class/interface/event metadata and keep wrappers explicitly labeled.

3. Prove early/late clients, outgoing events, DllCanUnloadNow and repeated backend-specific sessions without fallback.

Starting code/truth surfaces: crates/oxvba-comhost/src/lib.rs, crates/oxvba-build/src.

Canonical matrix rows: docs/validation/WINDOWS_NATIVE_EXPORT_AND_PACKAGING_MATRIX_V1.csv#WNE-PLAN-WRAPPED.

Row contract WNE-PLAN-WRAPPED: semantic_subset=JIT WrappedComServer with verified package and Excel clients; process_shape=inproc-and-localserver-wrapped-COM; apartment_shape=STA-and-proxy; exact_signature=generated COM host entries and class metadata; fixture_id=output-wrapped-comserver-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-comhost; command:cargo test -p oxvba-build.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-11/wrapper_com.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.12; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.12.5 — WIN-11 verify clean wrapper deployment and version rejection

Outcome: WIN-11 verify clean wrapper deployment and version rejection.

Effect: delivery. Parent epic: bd-59co.3.12 (WIN-11).

Clauses: BUILD-CLASS-001|BUILD-PACKAGE-001|PROFILE-TOOL-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.12.4, bd-59co.3.15.11. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Exercise deploy/load/invoke/reset/unload/uninstall for all three wrapper classes on a resettable environment.

2. Prove manifests identify compiler-free runtime dependencies, ABI mismatch diagnostics and retained source maps.

3. Capture client and artifact hashes, startup/warm-call measurements and zero owned registration/process/file residue.

Starting code/truth surfaces: crates/oxvba-build/src, docs/validation/WINDOWS_NATIVE_EXPORT_AND_PACKAGING_MATRIX_V1.csv.

Canonical matrix rows: docs/validation/WINDOWS_NATIVE_EXPORT_AND_PACKAGING_MATRIX_V1.csv#WNE-WRAPPER-EXE; docs/validation/WINDOWS_NATIVE_EXPORT_AND_PACKAGING_MATRIX_V1.csv#WNE-WRAPPER-LIBRARY; docs/validation/WINDOWS_NATIVE_EXPORT_AND_PACKAGING_MATRIX_V1.csv#WNE-PLAN-WRAPPED.

Row contract WNE-WRAPPER-EXE: semantic_subset=standalone WrapperExe with embedded or deployed verified OxImage; process_shape=standalone; apartment_shape=n/a; exact_signature=generated WrapperExe entry to JIT session; fixture_id=output-wrapper-exe-v1; environment_id=win-x64-cert-vm-pending-v1.

Row contract WNE-WRAPPER-LIBRARY: semantic_subset=WrapperLibrary external ABI and repeated session lifecycle; process_shape=inproc-wrapper-library; apartment_shape=caller-declared; exact_signature=generated wrapper library entry ABI; fixture_id=output-wrapper-library-v1; environment_id=win-x64-cert-vm-pending-v1.

Row contract WNE-PLAN-WRAPPED: semantic_subset=JIT WrappedComServer with verified package and Excel clients; process_shape=inproc-and-localserver-wrapped-COM; apartment_shape=STA-and-proxy; exact_signature=generated COM host entries and class metadata; fixture_id=output-wrapped-comserver-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:cargo test -p oxvba-build; command:cargo test -p oxvba-comhost.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-11/wrapper_deploy.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.12; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.14.3 — WIN-13 reject mutated Windows boundary descriptors before side effects

Outcome: WIN-13 reject mutated Windows boundary descriptors before side effects.

Effect: delivery. Parent epic: bd-59co.3.14 (WIN-13).

Clauses: CONF-QUALITY-001|SEC-BOUNDARY-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.2.6.3, bd-59co.3.3.8, bd-59co.3.3.10, bd-59co.3.3.11, bd-59co.3.3.12. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Generate malformed identities, slots, signatures, lengths, record graphs and callback lifetime descriptors.

2. Assert rejection before COM activation, DLL load or executable thunk publication.

3. Minimize failures into durable regressions and run the applicable modeled/formal lanes for touched verifier and ownership invariants.

Starting code/truth surfaces: crates/oxvba-runtime/src/interop_plan.rs, crates/oxvba-oxir/src, crates/oxvba-com/src.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-SAFETY-MUTATION.

Row contract WAC-SAFETY-MUTATION: semantic_subset=malformed descriptors fault injection repeated lifecycle and balance; process_shape=OxVba-process; apartment_shape=n/a; exact_signature=hostile carrier descriptor and lifecycle inputs; fixture_id=carrier-safety-mutation-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-runtime interop_plan; command:cargo test -p oxvba-oxir.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-13/safety_mutate.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.14; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.14.4 — WIN-13 inject COM native and allocation failures into owned lifetimes

Outcome: WIN-13 inject COM native and allocation failures into owned lifetimes.

Effect: delivery. Parent epic: bd-59co.3.14 (WIN-13).

Clauses: CONF-QUALITY-001|SEC-BOUNDARY-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.4.8, bd-59co.3.5.3, bd-59co.3.7.4, bd-59co.3.11.6. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Inject failing QI/Invoke/Advise, allocation, conversion, native-loader and callback registration operations.

2. Prove every acquired reference, record, pin, sink and callback token is released on partial failure.

3. Run repeated lifetime stress and appropriate sanitizer/Miri-compatible pure ownership fixtures; record unsupported tool/host combinations explicitly.

Starting code/truth surfaces: crates/oxvba-com/src, crates/oxvba-runtime/src, crates/oxvba-comhost/src/lib.rs.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-SAFETY-MUTATION.

Row contract WAC-SAFETY-MUTATION: semantic_subset=malformed descriptors fault injection repeated lifecycle and balance; process_shape=OxVba-process; apartment_shape=n/a; exact_signature=hostile carrier descriptor and lifecycle inputs; fixture_id=carrier-safety-mutation-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-com; command:cargo test -p oxvba-runtime; command:cargo test -p oxvba-comhost.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-13/safety_fault.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.14; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: large.

### bd-59co.3.14.5 — WIN-13 enforce capability policy and isolate all native cleanup

Outcome: WIN-13 enforce capability policy and isolate all native cleanup.

Effect: delivery. Parent epic: bd-59co.3.14 (WIN-13).

Clauses: CONF-QUALITY-001|SEC-BOUNDARY-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.1.5, bd-59co.3.10.4, bd-59co.3.7.5, bd-59co.3.3.9. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Exercise denied COM/native/file capabilities separately from missing implementation and VBA runtime errors.

2. Apply owned process/file/Registry64 journals to failure, timeout and cancellation paths.

3. Prove unrelated Excel processes, dialogs, registry keys and files survive hostile fixture cleanup; never use blanket termination or dialog dismissal.

Starting code/truth surfaces: crates/oxvba-hal/src, crates/oxvba-com/src, crates/oxvba-build/src.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-SAFETY-MUTATION.

Row contract WAC-SAFETY-MUTATION: semantic_subset=malformed descriptors fault injection repeated lifecycle and balance; process_shape=OxVba-process; apartment_shape=n/a; exact_signature=hostile carrier descriptor and lifecycle inputs; fixture_id=carrier-safety-mutation-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-hal; command:cargo test -p oxvba-com.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-13/safety_policy.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.14; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.14.6 — WIN-13 prove safety across the Windows compatibility subset

Outcome: WIN-13 prove safety across the Windows compatibility subset.

Effect: delivery. Parent epic: bd-59co.3.14 (WIN-13).

Clauses: CONF-QUALITY-001|SEC-BOUNDARY-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.14.3, bd-59co.3.14.4, bd-59co.3.14.5, bd-59co.3.12.5, bd-59co.3.11.7, bd-59co.3.6.8. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Aggregate mutation, fault, apartment, reentry and lifecycle evidence for clients, events, serving, imports, callbacks and wrappers.

2. Check native-boundary panic containment, unrelated-resource survival and repeated session balance.

3. Publish an explicitly bounded compatibility report; full native-output safety remains with existing bd-59co.3.14.2 and is not closed by this bead.

Starting code/truth surfaces: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv, crates/oxvba-differential/tests/jit_windows_vm3_parity.rs.

Canonical matrix rows: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-SAFETY-MUTATION.

Row contract WAC-SAFETY-MUTATION: semantic_subset=malformed descriptors fault injection repeated lifecycle and balance; process_shape=OxVba-process; apartment_shape=n/a; exact_signature=hostile carrier descriptor and lifecycle inputs; fixture_id=carrier-safety-mutation-v1; environment_id=win-x64-dev-oracle-2026-07.

Acceptance commands: command:cargo test -p oxvba-differential --test jit_windows_vm3_parity; command:./scripts/run-truth-reconciliation.ps1.

Acceptance scope: extend the named existing suites with positive, negative and cleanup cases for the behavior above; a zero-test filter or matching backend decline is not passing evidence.

Expected observable: exact result, full Err and LastDllError where relevant, ordered side effects/ByRef mutation, lifecycle order, actual transport and zero owned-resource balance drift. State any inapplicable axis with a reason.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-13/safety_compat.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: remaining accepted scope stays under bd-59co.3.14; any discovered gap gets an exact delivery successor before closure. This leaf advances only its named subset and cannot certify its whole row without all required producer and WIN-14 evidence.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

### bd-59co.3.15.37 — WIN-14 publish the Windows compatibility subset evidence report

Outcome: WIN-14 publish the Windows compatibility subset evidence report.

Effect: support. Parent epic: bd-59co.3.15 (WIN-14).

Clauses: AUTH-CLEAN-001|AUTH-SPEC-001|AUTH-VBA-001|CONF-DIFF-001|CONF-ORACLE-001|PROFILE-WIN-001.

Scope: Windows x64 and actual 64-bit Excel where VBA authority is required; development evidence is noncertifying.

Prerequisites: bd-59co.3.15.3, bd-59co.3.15.32, bd-59co.3.15.24, bd-59co.3.15.25, bd-59co.3.15.26, bd-59co.3.15.27, bd-59co.3.14.6. Read each producer acceptance artifact before claiming.

Why: this provides the exact reusable contract needed by downstream Windows consumers; a fixture-specific workaround cannot satisfy a general boundary.

Implementation steps:

1. Assemble exact implementation and pinned-environment evidence for COM clients/events/serving, Declare, pointers/callbacks and wrappers.

2. Retain an immutable replay index for this subset and list every one of the 57 Windows rows with its actual status and remaining owner.

3. Leave genuine outputs, native/full-profile aggregates, bd-59co.3.15.33/.34, WIN-15 and programme closure open; this report is a named compatibility subset and supplies no full-release completion credit.

Starting code/truth surfaces: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv, docs/validation/WINDOWS_NATIVE_EXPORT_AND_PACKAGING_MATRIX_V1.csv, docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14.

Read-only reconciliation references: docs/validation/WINDOWS_ABI_CARRIER_MATRIX_V1.csv#WAC-RELEASE-CERT.

Row contract WAC-RELEASE-CERT: semantic_subset=all mandatory COM native carrier lifecycle locale deployment and cleanup rows; process_shape=clean-certification-VM; apartment_shape=all-required; exact_signature=profile report over exact signatures and controlled fixtures; fixture_id=windows-release-cert-v1; environment_id=win-x64-cert-vm-pending-v1.

Acceptance commands: command:./scripts/run-truth-reconciliation.ps1; command:./scripts/check-governance.ps1.

Acceptance scope: verify every cited producer and certification artifact, immutable replay identity and remaining owner. Missing evidence stays open with a delivery owner; this report does not invent test results.

Expected observable: a reproducible compatibility-subset index classifies all 57 rows against actual source evidence, explicitly retains paused native/full-programme rows and changes no status without its required proof.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/compat_report.md. Retain source/fixture/target/environment/command hashes and per-case transcripts; resolve uncertain VBA behavior with public authority or owned VBE compile/runtime observations.

Fresh-eyes review: reread changed paths and actively probe missed errors, aliases, reentry, thread ownership and teardown; rerun relevant checks after repairs. Run applicable existing formal lanes for changed modeled behavior and record unresolved results.

Residual behavior: this report neither owns nor advances WAC-RELEASE-CERT or any full-programme aggregate. Missing capability evidence retains its delivery or certification owner; reporting cannot complete an outstanding producer.

Scheduling: serialize registry/Excel/provisioning and large runtime writers as applicable; respect the two Rust-writer limit. Relative complexity: medium.

## Existing work reused and preserved

The following existing open leaves remain authoritative. They are retained rather than duplicated. Their current scope, prerequisites and acceptance are reproduced without historical estimates. The graph contains the full dependencies; parent-child ownership does not by itself prove a prerequisite has passed.

### bd-59co.2.4.4 — CORE-3 enforce VBA UDT declaration size and depth limits

Status at planning snapshot: open. Prerequisites: no direct blocking edge; inspect parent prerequisites.

Outcome: enforce VBA's declaration-context size limits for user-defined types and fixed/static data in the compiler, complementing the runtime's 64 KiB sealed VbaRecord payload bound without inventing context in the runtime. Effect: delivery. Clauses: COMP-ANALYSIS-001|COMP-BIND-001|COMP-DIAG-001|IR-CORE-001|LS-FACT-001|PROJ-REF-001. Canonical matrix row: CORE-READINESS/CORE-TYPED-BINDING, with Excel/VBA evidence in EXCEL-ORACLE. The bd-59co.2.2.15 runtime layout review intentionally left compiler-owned rules open: the published VBA diagnostic distinguishes maximum UDT size, procedure-local fixed/static data (including the 32 KiB rule), and module/project aggregate/static-data limits. Build exact byte sizes from sealed type facts; reject at the correct declaration phase and source span with stable diagnostic facts before CoreProgram emission. Use the Microsoft VBA documentation plus controlled Excel/VBA compile oracle boundary probes to adjudicate exact inclusive/exclusive thresholds, nested UDTs, fixed strings, fixed arrays, multiple locals/module fields, and conditional declarations. Acceptance commands: command:cargo test -p oxvba-bind udt_declaration_context_size_limits -- --nocapture; command:cargo test -p oxvba-symbol udt_size_diagnostics -- --nocapture; command:./scripts/run-excel-oracle.ps1 -Lane udt-size-limits; command:cargo test -p oxvba-differential udt_size_limits -- --nocapture. Expected observable: accepted boundary shapes compile and rejected boundary+1/aggregate shapes produce VBA-compatible compile diagnostics at the exact declaration; no oversized executable CoreProgram is emitted; strict/editor facts agree for valid source. Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-3/udt-declaration-size-limits.md. Residual behavior: runtime layout safety remains bd-59co.2.2.15; Windows external ABI layout remains WIN-1; uncertainty stays in-progress until Excel/VBA evidence adjudicates it. Fresh-eyes non-author review is required. Runtime safety addendum: bd-59co.2.2.15 introduces the named non-VBA-claim security cap MAX_VBA_RECORD_LAYOUT_GRAPH_DEPTH=64 to prevent recursive record-shape stack exhaustion. This compiler/oracle bead must also determine Excel/VBA behavior for deeply nested acyclic UDT graphs at and beyond that boundary. If VBA admits a shape the cap rejects, create an exact stack-safe runtime/verifier successor instead of treating the security cap as compatibility authority.

Acceptance: Microsoft authority and controlled Excel/VBA compile probes establish exact UDT/local/module size thresholds and diagnostics; boundary shapes compile, boundary+1 and aggregate violations fail at the exact declaration; nested UDT/fixed-string/fixed-array/conditional cases are covered; strict/editor facts agree for valid source; no rejected program reaches CoreProgram; evidence and fresh-eyes review are clean. Nested UDT graph-depth probes cover the runtime security boundary and either establish a matching VBA compile limit or leave an exact stack-safe runtime/verifier successor for any VBA-admitted deeper shape.

### bd-59co.2.7.2 — CORE-5 enforce SafeArray object-carrier thread ownership

Status at planning snapshot: open. Prerequisites: bd-59co.2.2.10.

Outcome: decide and enforce the sound thread-ownership model for SafeArray, Variant, and VbaRecord carriers that can own VT_DISPATCH or VT_UNKNOWN ObjectRef elements. Effect: delivery. Clauses: HOST-HAL-001|HOST-SESSION-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|RUNTIME-VALUE-001|SEC-BOUNDARY-001|SYS-OWN-001. Canonical matrix row: CORE-READINESS/CORE-RUNTIME-HELPER-SESSION. Audit the unconditional SafeArray Send/Sync contract together with VbaRecord auto-Send/Sync caused by hidden Variant payloads and all Variant ObjectRef transit, compat-owner RefCell/Cell mutation, final Release, thread-local termination queues and foreign COM apartment affinity; choose thread confinement or an explicit enforced owner-thread carrier and encode the invariant in types and session boundaries rather than comments. Acceptance commands: command:cargo test -p oxvba-runtime safe_array_thread_ownership; command:cargo test -p oxvba-runtime object_ref; command:cargo clippy -p oxvba-runtime --all-targets -- -D warnings. Expected observable: compile-time and runtime tests reject unsound cross-thread ownership, valid numeric/value carriers retain intended mobility, final release/termination order is deterministic and object/array/ref balances return to zero. Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-5/safearray-object-thread-ownership.md. Residual behavior: Windows COM apartment attachment/reentry implementation is routed to WIN-2 rollout producer bd-59co.3.3.1 and apartment/lifecycle evidence is routed to WIN-13 rollout producer bd-59co.3.14.1; the core ownership contract cannot remain comment-only. Fresh-eyes Sol review is required.

Acceptance: The chosen ownership model is explicit and type-enforced; unsafe Send/Sync is removed or backed by an enforced owner-thread carrier; ObjectRef/Variant/VbaRecord/COM transit, final Release and termination-queue behavior have focused compile-time/runtime/lifecycle tests; strict Clippy passes; Windows apartment implementation/evidence residuals are explicitly related to bd-59co.3.3.1 and bd-59co.3.14.1.

### bd-59co.2.7.3 — CORE-5 enforce exact SafeArray dimension and allocation layout

Status at planning snapshot: open. Prerequisites: bd-59co.2.2.10.

Outcome: make every OxVba-owned SAFEARRAY descriptor carry a destruction-time dimension count that exactly reconstructs its allocation layout, and reject or represent unsupported ranks without truncation. Effect: delivery. Clauses: HOST-HAL-001|HOST-SESSION-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|RUNTIME-VALUE-001|SEC-BOUNDARY-001|SYS-OWN-001. Canonical matrix row: CORE-READINESS/CORE-RUNTIME-HELPER-SESSION. The bd-2cjy fresh-eyes audit found that alloc_header stores bounds.len() as u16 while the public dimensions() projection truncates to u8 and Drop recomputes owner_layout from that truncated value; rank 256 leaks and rank 257 can deallocate with a different layout. Define the supported VBA/product rank contract, preserve the exact allocation rank for destruction, reject hostile/oversized construction before allocation, and audit clone/raw-adoption/bounds paths for the same invariant. Acceptance commands: command:cargo test -p oxvba-runtime safe_array_dimension_layout; command:cargo test -p oxvba-runtime safe_array; command:cargo clippy -p oxvba-runtime --all-targets -- -D warnings. Expected observable: boundary and hostile-rank tests prove no truncation, no mismatched deallocation, and zero SafeArray counter drift while supported arrays retain their exact dimensions and bounds. Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-5/safearray-dimension-layout.md. Residual behavior: any external COM SAFEARRAY rank/adoption compatibility question is routed to exact WIN-2 rollout producer bd-59co.3.3.1, but no safe Core constructor or Drop path may retain an allocation-contract hazard. Fresh-eyes Sol review is required.

Acceptance: The supported rank contract is explicit; all constructors reject or faithfully represent boundary ranks before allocation; Drop reconstructs the exact owner layout without narrowing; hostile rank 0/255/256/257/u16-boundary tests, clone/raw adoption and balance checks pass; strict Clippy passes; external COM rank compatibility has exact downstream route bd-59co.3.3.1.

### bd-59co.2.7.4 — CORE-5 establish a stable VM3 synchronous reentry root

Status at planning snapshot: open. Prerequisites: bd-59co.2.2.10, bd-59co.2.2.26, bd-59co.2.7.6.

Outcome: establish one stable raw VM3/session execution root so future synchronous host, COM event, native callback, and DoEvents re-entry can enter the same VM3 session without any outer typed `&mut Vm3`, `ExecState`, frame, or LibContext borrow spanning the callback. Effect: delivery. Clauses: HOST-HAL-001|HOST-SESSION-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|RUNTIME-VALUE-001|SEC-BOUNDARY-001|SYS-OWN-001. Canonical matrix row: CORE-READINESS/CORE-RUNTIME-HELPER-SESSION. The CORE-VM3-REFERENCE consumer row already declares CORE-5 as a producer dependency and remains owned by CORE-6. Current StandardHost queues event-pump callbacks and VM3 dispatches them at a later StmtBoundary; bd-59co.2.2.26 proves that current queue-only behavior and removes the generic safe LibContext+host wrapper, but it intentionally does not create synchronous same-VM re-entry. Refactor the full `run_entry -> run_loop -> exec` boundary and nested `run_proc_with_values`, `run_proc_with_me`, `run_callback_proc_with_values`, native Declare callback and event producers around a stable session-owned raw root with short pre/post borrows, explicit same-thread/non-unwind/lifetime authority, bounded recursion, full Err/Erl/frame/ParamArray restoration, and deterministic late/disposed-session rejection. Integrate the verified Windows interop plan consumers rather than introducing a Linux-only shortcut. Acceptance commands: command:cargo fmt --all -- --check; command:cargo clippy -p oxvba-vm3 --all-targets -- -D warnings; command:cargo test -p oxvba-vm3; command:cargo test -p oxvba-differential --test class_lifecycle_vm3 -- --nocapture; command:cargo test -p oxvba-differential --test jit_project_objects -- --nocapture. Expected observable: synchronous nested VM3 entry through host/event/native callback executes on one stable root without overlapping typed borrows; results, full Err/Erl, side effects, callback order, frame/alias lifecycle, transport, late-call rejection, and balances match JIT/VBA authority. Evidence: artifact:docs/evidence/programs/ideal-2026-07/core/CORE-5/vm3-stable-reentry-root.md. Residual behavior: apartment/COM/native fixture certification remains in Windows producers and WIN-14 consumers; this bead owns the portable VM3 session mechanism and cannot claim those rows.

Acceptance: The complete VM3 run-loop and nested-entry caller chain uses one stable session root with no typed VM/state/context borrow spanning arbitrary synchronous callbacks; queue-only and synchronous paths are distinct and tested; nested success/failure restores full Err/Erl, frames and real ParamArray aliases; late/disposed/cross-thread calls fail deterministically; focused VM3, strict Clippy and differential checks pass; Sol fresh-eyes review is clean.

### bd-59co.3.1.8 — WIN-0 certify the x64 control plane and successor queue

Status at planning snapshot: open. Prerequisites: bd-59co.3.1.1, bd-59co.3.1.7.

Outcome: certify WIN-0 support truth and hand off only current Windows delivery work. Effect: support. Clauses: CONF-MATRIX-001|CONF-ORACLE-001|DOC-AUTH-001|DOC-TRACE-001|PROFILE-WIN-001. Canonical row: WIN-ABI-CARRIER/WAC-TARGET-DEV-ENV plus the six-matrix control report. Acceptance commands: command:./scripts/run-truth-reconciliation.ps1; command:./scripts/check-governance.ps1; command:./scripts/test-path-stability.ps1; command:br lint --json; command:br dep cycles. Expected observable: WAC-TARGET-DEV-ENV is verified with six axes and no residual; the other 56 required rows retain exact downstream owners; only current successor work is ready. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-0/terminal.md. Residual behavior: any control inconsistency or missing successor keeps WIN-0 open. Fresh-eyes non-author review is required.

Acceptance: Terminal commands and review pass; dev-host fixture/environment hashes are concrete but noncertifying; all other required rows remain truthful; exact ready queue and three-agent capacity snapshot are recorded.

### bd-59co.3.10.2 — WIN-9 x64 PtrSafe scalar Declare import

Status at planning snapshot: open. Prerequisites: bd-59co.3.10.1, bd-59co.3.3.4.

Outcome: deliver the first WIN-9 x64 PtrSafe scalar Declare import under VM3 and JIT using the shared verified interop plan. Effect: delivery. Clauses: COMP-BIND-001|NATIVE-IMPORT-001|WIN-PLAN-001. Canonical matrix rows: WIN-NATIVE-IMPORT/WNI-PLAN-DECLARE. Controlled x64 scalar fixture first; Excel rows stay planned under WIN-14. Acceptance commands: command:cargo test -p oxvba-differential --test jit_windows_vm3_parity -- --nocapture. Expected observable: VM3 and JIT match on result full Err LastDllError raised and handle-balance for the scalar Declare fixture, or an exact owned residual is named. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-9/declare-scalar.md. Residual behavior: strings/buffers/arrays/UDTs/As Any/loader-policy stay with bd-59co.3.10.3. Fresh-eyes review is required. Canonical matrix truth surface: the exact row named above in WINDOWS_JIT_NATIVE_IMPORT_MATRIX_V1.csv.

Detailed acceptance addition: Replace sequential GetTickCount equality with a deterministic owned x64 scalar native fixture. Prove Function/Sub returns, mutation, full Err/LastDllError and Win64 transport; matching declines are failures. Keep this bounded first slice independent of broader session delivery.

Acceptance: x64 scalar Declare fixture matches under VM3 and JIT through the shared plan; Excel rows stay planned. Fresh-eyes review is clean.

### bd-59co.3.10.3 — WIN-9 realize remaining VBA7 Declare import rows

Status at planning snapshot: open. Prerequisites: bd-59co.3.10.2, bd-59co.3.10.4, bd-59co.3.10.5, bd-59co.3.10.6, bd-59co.3.10.7.

Outcome: integrate and prove the remaining behavior owned by WIN-9 realize remaining VBA7 Declare import rows.

Effect: delivery. Clauses: COMP-BIND-001|NATIVE-IMPORT-001|WIN-PLAN-001.

Prerequisites: bd-59co.3.10.2, bd-59co.3.10.4, bd-59co.3.10.5, bd-59co.3.10.6, bd-59co.3.10.7. These leaves own the component implementations; this bead owns their combined behavior through one production route.

Steps: validate the producer artifacts; execute combined success/error/alias/reentry/lifecycle cases under VM3 and JIT; verify the same immutable plan and session identity at each required boundary; correct any integration defect and add its regression.

Canonical matrix rows: WIN-NATIVE-IMPORT/WNI-PLAN-DECLARE.

Acceptance commands: command:cargo test -p oxvba-differential --test jit_windows_vm3_parity -- --nocapture; command:./scripts/run-truth-reconciliation.ps1.

Expected observable: exact structural results, full Err/LastDllError where relevant, side effects, lifecycle ordering, native transport and zero balance drift across the combined producer paths. A matching decline, fixture-specific shortcut or paper split does not satisfy this delivery outcome.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-9/remaining-declare.md.

Fresh-eyes review must probe missed cross-feature aliases, errors, threading and cleanup and rerun relevant checks after repairs.

Residual behavior: this bead stays open until its scoped integrated behavior and required producer evidence pass. Newly discovered gaps require exact open delivery successors and keep the broader area in-progress. Full canonical-row certification remains with the applicable WIN-14 evidence owners; no planning-only closure is permitted.

Acceptance: All named producer leaves pass and their combined native behavior is proven under VM3/JIT with exact structural/error/transport/lifetime evidence and a clean fresh-eyes review. Creating successor beads alone cannot close this delivery item.

### bd-59co.3.11.2 — WIN-10 execute retained callback lifetime in VM3

Status at planning snapshot: open. Prerequisites: bd-59co.2.7.4, bd-59co.3.11.1, bd-59co.3.3.12, bd-59co.3.3.9, bd-9sed.17.

Bounded outcome: Execute one externally owned retained callback record in VM3 with deterministic release behavior. Clauses: NATIVE-IMPORT-001|WIN-PLAN-001. Matrix: WIN-NATIVE-IMPORT. Acceptance references: command:cargo test -p oxvba-vm3 vm3_retained_native_callback_lifetime -- --nocapture. Expected observable: live entry succeeds release rejects late entry and thunk/session ownership balances. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-10/retained-callback-vm3.md. Residual behavior: JIT and terminal certification remain siblings. Fresh-eyes non-author review is required.

Detailed Windows plan handoff: bd-59co.3.3.12, bd-59co.3.3.9, bd-59co.2.7.4. Consume the owned shared-session and immutable plan contracts; preserve the existing fixture producer.

Acceptance: The named command/environment references pass or resolve; expected observable is recorded in artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-10/retained-callback-vm3.md; residual scope remains explicitly owned.

### bd-59co.3.11.3 — WIN-10 execute retained callback lifetime in JIT

Status at planning snapshot: open. Prerequisites: bd-59co.2.7.7, bd-59co.3.11.1, bd-59co.3.3.12, bd-59co.3.3.9, bd-9sed.17.

Bounded outcome: Execute the same retained callback record and typed trampoline in the JIT with panic containment. Clauses: NATIVE-IMPORT-001|WIN-PLAN-001|RUNTIME-ABI-001. Matrix: WIN-NATIVE-IMPORT. Acceptance references: command:cargo test -p oxvba-jit jit_retained_native_callback_lifetime -- --nocapture. Expected observable: JIT equals VM3 for live/released/reentrant states with no AV/unwind and balanced ownership. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-10/retained-callback-jit.md. Residual behavior: terminal certification remains a sibling. Fresh-eyes non-author review is required.

Detailed Windows plan handoff: bd-59co.3.3.12, bd-59co.3.3.9, bd-59co.2.7.7. Consume the owned shared-session and immutable plan contracts; preserve the existing fixture producer.

Acceptance: The named command/environment references pass or resolve; expected observable is recorded in artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-10/retained-callback-jit.md; residual scope remains explicitly owned.

### bd-59co.3.11.4 — WIN-10 certify retained callback disposal exhaustion and reentry

Status at planning snapshot: open. Prerequisites: bd-59co.3.11.1, bd-59co.3.11.2, bd-59co.3.11.3, bd-59co.3.15.3, bd-59co.3.15.6.

Bounded outcome: Certify actual native/Excel callback disposal exhaustion late-entry and reentry behavior. Clauses: NATIVE-IMPORT-001|CONF-ORACLE-001|CONF-QUALITY-001. Matrix: WIN-NATIVE-IMPORT. Acceptance references: command:cargo test -p oxvba-host retained_callback_lifecycle_windows -- --ignored --nocapture|command:./scripts/run-excel-vba-oracle.ps1|environment:win-x64-cert-vm-pending-v1. Expected observable: no AV or stale entry and complete result/Err/order/transport/balance evidence. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-10/retained-callback-certification.md. Residual behavior: other callback shapes keep separate rows. Fresh-eyes non-author review is required.

Detailed Windows plan handoff: bd-59co.3.15.3, bd-59co.3.15.6. Consume the owned shared-session and immutable plan contracts; preserve the existing fixture producer. The certification environment placeholder must be resolved through the sealed environment identity produced by bd-59co.3.15.3 before execution. Capture the actual environment ID/hash in the transcript; never certify on the development host.

Acceptance: The named command/environment references pass or resolve; expected observable is recorded in artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-10/retained-callback-certification.md; residual scope remains explicitly owned.

### bd-59co.3.13.1 — Roll out WIN-12 executable child beads

Status at planning snapshot: open. Prerequisites: no direct blocking edge; inspect parent prerequisites.

Outcome: create or refresh the executable child bead set for WIN-12 without substituting planning for delivery. The parent contract is reproduced below so future workers do not need the workset document.

Type: delivery
Clauses: `BUILD-NATIVE-001`, `JIT-AOT-001`

Deliver project export manifest/signature eligibility; versioned external ownership/error/concurrency ABI; Cranelift object/blob and PE/COFF relocation/import format; x64 entries/names/ordinals; runtime/helper initialization outside loader lock; DLL/EXE global/session/unload rules; source/debug maps; ASLR/clean-machine/reproducible builds; native C/Rust clients.

First beads: ABI/export contract; object/blob prototype; scalar x64 DLL; signature breadth; native EXE; reloc/import/ASLR; reproducibility/debug maps.

Close: mandatory x64 native targets are real program-specific outputs, not wrappers.

Every child must identify one reviewable outcome, delivery/support effect, contract clauses, exact dependencies, canonical matrix rows/truth surfaces, targeted tests or evidence, relative complexity, fresh-eyes review, and blocker/residual behavior.

Executable rollout acceptance: command:./scripts/validate-workset-rollout.ps1; command:./scripts/run-truth-reconciliation.ps1; artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-12/rollout-acceptance.md. Expected observable: bounded delivery/support children exist with exact row traces, producer dependencies, typed acceptance commands and residual owners; no capability closes on this support bead.

Tooling-profile clause: PROFILE-TOOL-001; wrapped and genuine native outputs remain distinct and terminal tooling claims require their own verified rows.

Exact trace-contract addendum:
Clauses: `BUILD-CLASS-001`, `DEBUG-MAP-001`, `HOST-SESSION-001`, `IMAGE-VERIFY-001`, `JIT-CACHE-001`, `SYS-ART-001`
These clauses are the exact producer/consumer and matrix-boundary union for this leaf; support relationships do not close capability.

Acceptance: A believable ready path exists for WIN-12: first delivery beads are explicit, dependencies reflect engineering truth, tests/evidence are included, no accepted residual is only narrative, and capability closure cannot occur on support beads alone. The two named commands must pass and the rollout artifact must enumerate exact child IDs, matrix rows, commands, observables and remaining owners.

### bd-59co.3.14.2 — WIN-13 run final native boundary safety lifecycle certification

Status at planning snapshot: open. Prerequisites: bd-59co.3.10, bd-59co.3.11, bd-59co.3.12, bd-59co.3.13, bd-59co.3.14.1, bd-59co.3.4, bd-59co.3.5, bd-59co.3.6, bd-59co.3.7, bd-59co.3.8, bd-59co.3.9.

Bounded outcome: Rerun fault repeated-session apartment registry callback and COM lifecycle lanes after WIN-3..WIN-12. Clauses: SEC-BOUNDARY-001|CONF-QUALITY-001. Matrix: WIN-ABI-CARRIER. Acceptance references: command:cargo test --workspace|command:./scripts/run-truth-reconciliation.ps1|command:./scripts/run-windows-safety-gates.ps1. Expected observable: no unwind AV leak stale registration or ownership imbalance. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-13/final-safety.md. Residual behavior: every defect gets a delivery owner. Fresh-eyes non-author review is required.

Acceptance: The named command/environment references pass or resolve; expected observable is recorded in artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-13/final-safety.md; residual scope remains explicitly owned.

### bd-59co.3.15.10 — WIN-14 promote x64 Declare pointer and callback fixtures

Status at planning snapshot: open. Prerequisites: bd-59co.3.1, bd-59co.3.15.1, bd-59co.3.15.8.

Outcome: build and self-test deterministic x64 native-import and callback fixtures. Effect: support. Clauses: CONF-ORACLE-001|CONF-QUALITY-001|PROFILE-WIN-001. Canonical matrix: WIN-NATIVE-IMPORT plus WAC-SAFETY-MUTATION. Include names ordinals ANSI/Wide/BSTR UDT As Any LastDllError pointers synchronous retained and nested callbacks C/Rust callers late-call rejection and balance. Acceptance commands: command:./scripts/build-win14-controlled-fixtures.ps1 -Domain NativeImport -Configuration Release; command:./scripts/validate-windows-fixture-manifest.ps1. Expected observable: reproducible DLL/client hashes and self-tests expose exact x64 ABI callback lifetime/error/cleanup observables with zero drift. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/fixtures/native-import/manifest.json. Residual behavior: fixture success does not certify runtime consumers. Fresh-eyes non-author review is required.

Acceptance: Native-import fixtures build reproducibly, self-tests cover scalar/structural/error/callback shapes, late calls are safely rejected and all handles/callbacks balance.

### bd-59co.3.15.11 — WIN-14 promote wrapper and genuine-native external clients

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.8.

Outcome: build deterministic external clients and Excel consumers for wrapped and genuine-native outputs before product artifacts exist. Effect: support. Clauses: CONF-QUALITY-001|JIT-PARITY-001|PROFILE-WIN-001. Canonical matrix: WIN-NATIVE-EXPORT. Build-only scope: provide reproducible C/Rust load unload restart probes, exact-symbol/ABI validators, negative placeholder artifacts, and Excel consumer sources without executing or certifying a WIN-11/WIN-12 product. Acceptance commands: command:./scripts/build-win14-controlled-fixtures.ps1 -Domain NativeOutputClients -Configuration Release. Expected observable: reproducible clients and validators are ready to consume later artifacts while wrapper versus genuine-native labels remain fail-closed. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/fixtures/native-output-clients/manifest.json. Residual behavior: product validation and capability credit remain blocked on their exact WIN-11/WIN-12 producers and WIN-14 output-certification beads. Fresh-eyes non-author review is required.

Acceptance: Build-only external clients and negative validators compile reproducibly, inspect exact symbols/ABI/lifecycle contracts, and reject mislabeled or placeholder artifacts; no wrapped or genuine-native product capability is executed or credited.

### bd-59co.3.15.12 — WIN-14 provision and pin the clean Windows x64 Excel64 VM

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.4, bd-59co.3.15.5.

Outcome: provision a provider-neutral resettable release-certification VM and emit a candidate immutable identity for later sealing. Effect: support. Clauses: CONF-ORACLE-001|CONF-QUALITY-001|PROFILE-WIN-001. Canonical row: WIN-ABI-CARRIER/WAC-CLEAN-CERT-ENV. Pin actual Windows and Office64 channel/build/license-safe source VM generation/CPU non-default locale image and reset contract; preserve the provisional manifest ID while recording a candidate concrete identity. Only bd-59co.3.15.3 may atomically replace the provisional environment ID after qualification reset and seal evidence pass. Acceptance commands: command:./scripts/provision-win14-certification-vm.ps1 -Plan docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/vm-provision-plan.json -Apply. Expected observable: a concrete candidate VM/image identity and clean snapshot exist with no dev-host substitution and no canonical manifest mutation. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/vm-provision.md. Residual behavior: provider license or provisioning failure is an explicit blocking environment condition. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: The checked-in plan is provider-neutral; Apply yields candidate pinned Windows/Excel64/locale/image identities and a reset snapshot; it does not replace the provisional canonical ID, which remains reserved to bd-59co.3.15.3 after qualification.

### bd-59co.3.15.13 — WIN-14 qualify Excel64 locale VBOM and UIA on the pinned VM

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.12, bd-59co.3.15.6.

Outcome: qualify the pinned VM for authoritative Excel/VBA use. Effect: support. Clauses: AUTH-VBA-001|CONF-ORACLE-001|CONF-QUALITY-001|PROFILE-WIN-001. Canonical rows: WAC-CLEAN-CERT-ENV and WCC-EXCEL-AUTHORITY. Prove x64 Excel process pinned LCID/codepage/calendar VBOM/macro modes forced compile success/error transcripts PID ownership and modal cleanup. Acceptance commands: command:./scripts/capture-ideal-environment.ps1 -Plan docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/vm-provision-plan.json -RequireExcel64 -RequireNonDefaultLocale -RequireUia; command:./scripts/run-excel-vba-oracle.ps1 -Suite HarnessSelfTest. Expected observable: qualification records immutable host facts and deterministic compile/modal/full-Err evidence with zero residue. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/vm-qualification.md. Residual behavior: any unavailable Excel VBOM UIA or locale requirement blocks certification. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: All VM/Excel64/locale/VBOM/UIA requirements and owned cleanup pass; success and failure compile transcripts are complete; no release row is yet verified.

### bd-59co.3.15.14 — WIN-14 prove clean reset deploy and cleanup cycles

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.10, bd-59co.3.15.12, bd-59co.3.15.13, bd-59co.3.15.7, bd-59co.3.15.9.

Outcome: prove repeatable clean-snapshot deployment and owned cleanup. Effect: support. Clauses: CONF-ORACLE-001|CONF-QUALITY-001|PROFILE-WIN-001. Canonical row: WAC-CLEAN-CERT-ENV. Revert before each cycle deploy only hashed fixtures/runner capture pre/post processes registry files COM and dialogs remove owned state and verify no residue. Acceptance commands: command:./scripts/verify-win14-clean-reset.ps1 -Plan docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/vm-provision-plan.json -Cycles 3. Expected observable: three independent restore/deploy/run/cleanup cycles have identical hashes and zero unowned or residual state. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/clean-reset.md. Residual behavior: any drift blocks environment sealing. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: Three reset cycles reproduce environment/fixture hashes, touch only owned resources and end with zero process dialog registry file COM connection callback or handle drift.

### bd-59co.3.15.15 — WIN-14 certify Excel64 compiler reference and typelib authority

Status at planning snapshot: open. Prerequisites: bd-59co.2.4.5, bd-59co.2.4.6, bd-59co.2.4.7, bd-59co.3.15.1, bd-59co.3.15.3, bd-59co.3.15.7, bd-59co.3.15.9, bd-59co.3.2.4.

Outcome: certify compiler/reference/typelib behavior against actual Excel64 and controlled metadata. Effect: delivery. Clauses: AUTH-VBA-001|CONF-ORACLE-001. Canonical rows: WCC-EXCEL-AUTHORITY plus WAC-TYPELIB-METADATA and WAC-EXCEL-COM-CERT evidence. Run valid/invalid compile broken/order registered/file references and VBE diagnostics through the owned supervisor. Acceptance commands: command:./scripts/run-excel-vba-oracle.ps1 -Suite CompilerReferenceTypelib; command:./scripts/run-windows-profile-gates.ps1 -Lane excel-authority. Expected observable: compiler facts and Excel compile/error transcripts agree with exact environment source fixture and cleanup hashes. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/excel-reference-authority.md. Residual behavior: any disagreement stays with the exact compiler metadata or oracle producer. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: Valid and invalid reference cases, ordering/broken/file/registered typelibs and VBE diagnostics are complete; six axes and owned cleanup are recorded; no unrelated COM capability is credited.

### bd-59co.3.15.16 — WIN-14 certify Excel64 carrier roundtrip and locale parity

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.3, bd-59co.3.15.7, bd-59co.3.15.9, bd-59co.3.2.5, bd-59co.3.2.6, bd-59co.3.2.7, bd-59co.3.2.8, bd-59co.3.4.7, bd-59co.3.5.3.

Outcome: certify exact x64 carriers and non-default-locale behavior across VM3 JIT native fixtures and Excel64. Effect: delivery. Clauses: COM-CLIENT-001|CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001|JIT-PARITY-001|SYS-DUAL-001|VM3-REF-001. Canonical rows: WAC-BSTR-LAYOUT WAC-VARIANT-LAYOUT WAC-SAFEARRAY-LAYOUT WAC-IUNKNOWN-IDENTITY WAC-NUMERIC-LONGPTR WAC-INTERFACE-ARRAY WAC-VT-RECORD WAC-CARRIER-EXCEL-ROUNDTRIP and WCC-LATE-OUTPROC-ERROR. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane carrier-locale; command:./scripts/run-excel-vba-oracle.ps1 -Suite CarrierLocale. Expected observable: values full Err LCID/date/decimal/case/codepage transport lifecycle and balance agree exactly. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/carrier-locale.md. Residual behavior: each divergent carrier stays open under WIN-1/3/4 or its authority row. Fresh-eyes non-author review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: All named carrier/locale rows have structural VM3/JIT/native/Excel observations, exact transport and zero lifecycle drift; divergences are not normalized away.

### bd-59co.3.15.17 — WIN-14 certify late-bound COM clients externally and in Excel64

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.3, bd-59co.3.15.7, bd-59co.3.15.9, bd-59co.3.4.7, bd-59co.3.4.8.

Outcome: certify late-dispatch activation invocation properties arguments structures out-of-proc and errors. Effect: delivery. Clauses: AUTH-VBA-001|COM-CLIENT-001|CONF-DIFF-001|CONF-ORACLE-001|JIT-PARITY-001|SYS-DUAL-001|VM3-REF-001. Canonical rows: WCC-PLAN-LATE WCC-LATE-ARGS WCC-LATE-STRUCTURAL and WCC-LATE-OUTPROC-ERROR. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane com-client-late; command:./scripts/run-excel-vba-oracle.ps1 -Suite ComClientLate. Expected observable: VM3 JIT native fixture and Excel values full Err effects dispatch transport reentry lifecycle and balances agree for in/out-proc and GetObject. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/com-client-late.md. Residual behavior: any row remains with WIN-3 and cannot be hidden by fixture success. Fresh-eyes non-author review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: All four late COM rows pass across VM3/JIT/native/Excel with full six-axis evidence, actual dispatch transport and zero reference/process/registry drift.

### bd-59co.3.15.18 — WIN-14 certify early-vtable COM clients externally and in Excel64

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.3, bd-59co.3.15.7, bd-59co.3.15.9, bd-59co.3.5.5.

Outcome: certify typed native-vtable COM client signatures and prove no dispatch fallback. Effect: delivery. Clauses: AUTH-VBA-001|COM-CLIENT-001|CONF-DIFF-001|CONF-ORACLE-001|JIT-PARITY-001|SYS-DUAL-001|VM3-REF-001. Canonical rows: WCC-PLAN-EARLY WCC-EARLY-COMPLEX WCC-EARLY-CUSTOM and WCC-EARLY-OUTPROC. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane com-client-early; command:./scripts/run-excel-vba-oracle.ps1 -Suite ComClientEarly. Expected observable: exact slots/signatures values ByRef arrays records interfaces HRESULT/full Err identity lifecycle and balance agree while dispatch counter remains zero. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/com-client-early.md. Residual behavior: any fallback or mismatch stays with WIN-4. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: All early COM rows pass with native-vtable counter proof, zero dispatch fallback, exact full Err/identity/writeback and zero resource drift.

### bd-59co.3.15.19 — WIN-14 certify incoming COM event and cancellable ByRef behavior

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.3, bd-59co.3.15.7, bd-59co.3.15.9, bd-59co.3.6.4, bd-59co.3.6.6.

Outcome: certify incoming event functional custom and complex argument behavior including synchronous cancellation. Effect: delivery. Clauses: COM-EVENT-001|CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001|JIT-PARITY-001|SYS-DUAL-001|VM3-REF-001. Canonical rows: WCE-PLAN-INCOMING WCE-INCOMING-COMPLEX and WCE-INCOMING-CUSTOM. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane com-events-incoming; command:./scripts/run-excel-vba-oracle.ps1 -Suite EventsIncoming. Expected observable: event results full Err object/array/record arguments order transport and balance agree and cancellable ByRef writeback occurs before native return. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/events-incoming.md. Residual behavior: any transport/writeback mismatch stays with WIN-5. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: All functional incoming-event rows pass with synchronous pre-return ByRef proof, exact event order/full Err and balanced connections/callbacks.

### bd-59co.3.15.2 — WIN-14 certify the complete Windows x64 profile

Status at planning snapshot: open. Prerequisites: bd-59co.3.1, bd-59co.3.10, bd-59co.3.11, bd-59co.3.12, bd-59co.3.13, bd-59co.3.14, bd-59co.3.15.1, bd-59co.3.15.3, bd-59co.3.15.32, bd-59co.3.15.33, bd-59co.3.15.34, bd-59co.3.2, bd-59co.3.3, bd-59co.3.4, bd-59co.3.5, bd-59co.3.6, bd-59co.3.7, bd-59co.3.8, bd-59co.3.9.

Outcome: certify the complete Windows x64 profile after every producer row environment aggregate archive safety and deployment checks is green. Effect: delivery. Clauses: AUTH-CLEAN-001|AUTH-SPEC-001|AUTH-VBA-001|CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001|JIT-PARITY-001|PROFILE-WIN-001|SYS-DUAL-001|VM3-REF-001. Canonical row: WIN-ABI-CARRIER/WAC-RELEASE-CERT. Run the immutable release lane on the sealed VM and confirm the other 54 pre-terminal Windows rows are verified; do not credit the two WIN-15 profile/tool terminal rows. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane release -EnvironmentPlan docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/vm-provision-plan.json; command:./scripts/run-truth-reconciliation.ps1; command:./scripts/check-governance.ps1; command:./scripts/test-path-stability.ps1. Expected observable: complete VM3/JIT/Excel/native results full Err effects lifecycle/event order transport and balance evidence with zero divergence or cleanup debt. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/final-certification.md. Residual behavior: any open producer row unavailable environment evidence mismatch or owned-resource drift keeps WAC-RELEASE-CERT and WIN-14 open. Fresh-eyes Sol terminal review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: All named commands pass on the sealed clean VM; the other 54 pre-terminal Windows rows and both COM/native aggregates are verified; WAC-RELEASE-CERT alone clears its residual; the two WIN-15 terminal rows remain downstream.

### bd-59co.3.15.20 — WIN-14 certify incoming event apartment reentry and lifecycle

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.19, bd-59co.3.15.3, bd-59co.3.15.7, bd-59co.3.15.9, bd-59co.3.6.8.

Outcome: certify cross-apartment/out-of-proc event reentry fanout replacement unsubscribe termination and handler errors. Effect: delivery. Clauses: COM-EVENT-001|CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001|JIT-PARITY-001|SYS-DUAL-001|VM3-REF-001. Canonical rows: WCE-INCOMING-APARTMENT and WCE-INCOMING-LIFECYCLE. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane com-events-incoming-lifecycle; command:./scripts/run-excel-vba-oracle.ps1 -Suite EventsIncomingLifecycle. Expected observable: apartment/reentry/event order full Err and connection/reference/process balance match declared policy. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/events-incoming-lifecycle.md. Residual behavior: any nondeterministic or leaking lifetime stays with WIN-5. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: Cross-apartment/outproc, rebind/fanout/unsubscribe/terminate/error cases pass with exact order and zero connection/callback/process drift.

### bd-59co.3.15.21 — WIN-14 certify late COM serving in and out of process

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.3, bd-59co.3.15.7, bd-59co.3.15.9, bd-59co.3.7.6.

Outcome: certify late-bound served classes through VM3/JIT plans and Excel/native clients. Effect: delivery. Clauses: COM-SERVE-001|CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001|JIT-PARITY-001|SYS-DUAL-001|VM3-REF-001. Canonical rows: WCS-LATE-INPROC WCS-LATE-LOCALSERVER WCS-LATE-COMPLEX and WCS-SERVER-SAFETY. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane com-server-late; command:./scripts/run-excel-vba-oracle.ps1 -Suite ComServerLate. Expected observable: plan digests values full Err class factory lock shutdown transport lifecycle and balances agree for in/out-proc clients. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/com-server-late.md. Residual behavior: any serving divergence stays with WIN-6. Fresh-eyes non-author review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: All late serving rows pass with VM3/JIT plan equality, Excel/native clients, exact class-factory/shutdown order and zero registry/process/reference drift.

### bd-59co.3.15.22 — WIN-14 certify early dual and custom COM serving

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.21, bd-59co.3.15.3, bd-59co.3.15.7, bd-59co.3.15.9, bd-59co.3.8.4.

Outcome: certify early/dual/custom Implements serving with one identity and no dispatch fallback. Effect: delivery. Clauses: COM-SERVE-001|CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001|JIT-PARITY-001|SYS-DUAL-001|VM3-REF-001. Canonical rows: WCS-DUAL-INPROC WCS-IMPLEMENTS-CUSTOM and WCS-EARLY-OUTPROC. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane com-server-dual; command:./scripts/run-excel-vba-oracle.ps1 -Suite ComServerDual. Expected observable: exact vtables/QI/state values full Err transport lifecycle and balance agree with zero dispatch fallback. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/com-server-dual.md. Residual behavior: any identity slot or fallback gap stays with WIN-7. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: All dual/custom serving rows pass with one identity/state, exact vtable/QI proof, zero dispatch fallback and balanced serving lifetime.

### bd-59co.3.15.23 — WIN-14 certify outgoing served events and Excel WithEvents

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.19, bd-59co.3.15.21, bd-59co.3.15.22, bd-59co.3.15.3, bd-59co.3.15.7, bd-59co.3.15.9, bd-59co.3.9.3.

Outcome: certify outgoing source-interface events consumed by Excel64 and external clients. Effect: delivery. Clauses: COM-EVENT-001|CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001|JIT-PARITY-001|SYS-DUAL-001|VM3-REF-001. Canonical rows: WCE-PLAN-OUTGOING and WCE-OUTGOING-COMPLEX. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane com-events-outgoing; command:./scripts/run-excel-vba-oracle.ps1 -Suite EventsOutgoing. Expected observable: fanout object array ByRef error and teardown values/order/transport/lifecycle/balance agree across served VM3/JIT and Excel WithEvents. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/events-outgoing.md. Residual behavior: gaps stay with WIN-8 and upstream serving/event producers. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: Both outgoing-event rows pass with Excel WithEvents, exact fanout/writeback/error/teardown order and zero connection/reference/process drift.

### bd-59co.3.15.24 — WIN-14 certify Declare imports in Excel64 and native clients

Status at planning snapshot: open. Prerequisites: bd-59co.3.10.3, bd-59co.3.15.1, bd-59co.3.15.10, bd-59co.3.15.3, bd-59co.3.15.7.

Outcome: certify VBA7 PtrSafe Declare imports across VM3 JIT Excel64 and controlled native clients. Effect: delivery. Clauses: AUTH-VBA-001|CONF-DIFF-001|CONF-ORACLE-001|JIT-PARITY-001|NATIVE-IMPORT-001|SYS-DUAL-001|VM3-REF-001. Canonical rows: WNI-PLAN-DECLARE WNI-DECLARE-STRINGS WNI-DECLARE-STRUCTURAL and WNI-DECLARE-LOADER-ERROR. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane native-declare; command:./scripts/run-excel-vba-oracle.ps1 -Suite NativeDeclare. Expected observable: x64 ABI names/ordinals strings UDT As Any immediate LastDllError policy/missing errors values full Err effects transport lifecycle and balance agree. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/declare.md. Residual behavior: each gap stays with WIN-9. Fresh-eyes non-author review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: All Declare rows pass across VM3/JIT/Excel/native with exact ABI and error seating, immediate LastDllError and zero module/buffer/session drift.

### bd-59co.3.15.25 — WIN-14 certify pointer helpers and synchronous callbacks

Status at planning snapshot: open. Prerequisites: bd-59co.3.11.5, bd-59co.3.11.6, bd-59co.3.15.1, bd-59co.3.15.10, bd-59co.3.15.3, bd-59co.3.15.7.

Outcome: certify x64 pointer helpers storage addressability and synchronous callback thunks. Effect: delivery. Clauses: CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001|JIT-PARITY-001|NATIVE-IMPORT-001|SYS-DUAL-001|VM3-REF-001. Canonical rows: WNI-POINTER-HELPERS and WNI-CALLBACK-SYNC. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane native-callback-sync; command:./scripts/run-excel-vba-oracle.ps1 -Suite CallbackSync. Expected observable: addresses lifetimes values full Err callback order pre-return ByRef writeback transport and balances agree. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/callback-sync.md. Residual behavior: any pointer or thunk gap stays with WIN-10. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: Pointer/storage and synchronous callback rows pass with pre-return writeback, exact error/order/transport and zero thunk/module/session drift.

### bd-59co.3.15.26 — WIN-14 certify retained nested and disposed-session callbacks

Status at planning snapshot: open. Prerequisites: bd-59co.3.11.4, bd-59co.3.11.7, bd-59co.3.15.1, bd-59co.3.15.10, bd-59co.3.15.25, bd-59co.3.15.3, bd-59co.3.15.7.

Outcome: certify retained and nested callback ownership reentry threads errors release and late-call containment. Effect: delivery. Clauses: CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001|JIT-PARITY-001|NATIVE-IMPORT-001|SYS-DUAL-001|VM3-REF-001. Canonical rows: WNI-PLAN-CALLBACK and WNI-CALLBACK-NESTED. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane native-callback-retained; command:./scripts/run-excel-vba-oracle.ps1 -Suite CallbackRetained. Expected observable: native-to-COM-to-VBA reentry and retained callback values full Err order transport lifetime release/late rejection and balance agree without UAF. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/callback-retained.md. Residual behavior: any unsafe late-call or ownership gap stays with WIN-10. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: Retained/nested callbacks pass through live/release/disposed cases, late calls reject safely, reentry order/full Err is exact and all callback/module/session resources balance.

### bd-59co.3.15.27 — WIN-14 certify JIT wrapper artifacts and WrappedComServer clients

Status at planning snapshot: open. Prerequisites: bd-59co.3.12.5, bd-59co.3.15.1, bd-59co.3.15.11, bd-59co.3.15.3, bd-59co.3.15.7.

Outcome: certify wrapped executable/library/COM-server artifacts and their external clients. Effect: delivery. Clauses: CONF-DIFF-001|CONF-QUALITY-001|JIT-PARITY-001|PROFILE-TOOL-001|PROFILE-WIN-001. Canonical rows: WNE-WRAPPER-EXE WNE-WRAPPER-LIBRARY and WNE-PLAN-WRAPPED. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane outputs-wrapped. Expected observable: compiler-free promise loader-lock-safe initialization values errors reset/unload transports and balances match honest wrapped labels across C/Rust/Excel clients. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/outputs-wrapped.md. Residual behavior: any wrapper gap stays with WIN-11 and is never relabeled genuine native. Fresh-eyes non-author review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: All wrapped rows pass external clients with honest labels, exact lifecycle/error/transport and zero process/module/registry/session drift.

### bd-59co.3.15.28 — WIN-14 certify genuine native DLL and EXE outputs

Status at planning snapshot: open. Prerequisites: bd-59co.3.13, bd-59co.3.15.1, bd-59co.3.15.11, bd-59co.3.15.3, bd-59co.3.15.7.

Outcome: certify real x64 PE DLL/EXE outputs with C and Rust clients. Effect: delivery. Clauses: CONF-DIFF-001|CONF-QUALITY-001|JIT-PARITY-001|PROFILE-TOOL-001|PROFILE-WIN-001. Canonical rows: WNE-PLAN-NATIVE WNE-NATIVE-EXE and WNE-NATIVE-ABI-BREADTH. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane outputs-native. Expected observable: genuine PE outputs export exact names/ordinals/ABI ownership/errors and behave correctly under concurrent external clients without runtime/compiler masquerade. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/outputs-native.md. Residual behavior: unsupported output shapes stay with WIN-12. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: Genuine DLL/EXE rows pass real PE inspection and C/Rust clients with exact ABI/error/concurrency and honest packaging identity.

### bd-59co.3.15.29 — WIN-14 certify native reproducibility ASLR debug maps and clean deployment

Status at planning snapshot: open. Prerequisites: bd-59co.3.13, bd-59co.3.15.1, bd-59co.3.15.11, bd-59co.3.15.14, bd-59co.3.15.28, bd-59co.3.15.3, bd-59co.3.15.7.

Outcome: certify reproducible-policy native artifacts under ASLR and clean-machine deployment. Effect: delivery. Clauses: CONF-DIFF-001|CONF-QUALITY-001|JIT-PARITY-001|PROFILE-TOOL-001|PROFILE-WIN-001. Canonical row: WNE-NATIVE-REPRO-DEPLOY. Inspect two-build hash policy relocations/imports ASLR source/debug maps load/unload and cleanup on the clean VM. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane outputs-native-deploy. Expected observable: reproducibility policy and semantic identity hold under relocation and clean deploy with zero file/module/process residue. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/outputs-native-deploy.md. Residual behavior: any packaging/debug/deploy gap stays with WIN-12. Fresh-eyes non-author review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: Two-build policy, PE relocation/import inspection, ASLR execution, debug/source map and clean deployment pass with zero residual state.

### bd-59co.3.15.3 — WIN-14 seal the pinned Windows x64 Excel64 certification environment

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.10, bd-59co.3.15.12, bd-59co.3.15.13, bd-59co.3.15.14, bd-59co.3.15.5, bd-59co.3.15.8, bd-59co.3.15.9.

Outcome: seal the concrete pinned Windows x64 Excel64 certification environment after provisioning qualification fixture promotion and repeated clean-reset proof. Effect: support. Clauses: CONF-ORACLE-001|CONF-QUALITY-001|PROFILE-WIN-001. Canonical row: WIN-ABI-CARRIER/WAC-CLEAN-CERT-ENV. Replace the provisional environment ID only through an atomic matrix/trace/environment update with concrete OS Office64 locale image snapshot fixture and policy hashes. Acceptance commands: command:./scripts/capture-ideal-environment.ps1 -Plan docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/vm-provision-plan.json -Seal; command:./scripts/validate-environment-manifest.ps1; command:./scripts/verify-win14-clean-reset.ps1 -Cycles 3. Expected observable: the environment row is verified with immutable six-axis evidence and zero cleanup drift; dev-host evidence remains noncertifying. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/certification-vm.md. Residual behavior: any missing pinned identity fixture hash qualification or reset proof keeps the environment and every executable certification lane blocked. Fresh-eyes Sol review is required.

Acceptance: Concrete pinned Windows/Excel64/locale/image/fixture identities, three clean reset cycles and owned cleanup all verify; the provisional ID is atomically replaced and WAC-CLEAN-CERT-ENV alone clears its residual.

### bd-59co.3.15.30 — WIN-14 certify hostile COM faults and repeated lifecycles

Status at planning snapshot: open. Prerequisites: bd-59co.3.14.6, bd-59co.3.15.1, bd-59co.3.15.17, bd-59co.3.15.18, bd-59co.3.15.19, bd-59co.3.15.20, bd-59co.3.15.21, bd-59co.3.15.22, bd-59co.3.15.23, bd-59co.3.15.3, bd-59co.3.15.7.

Outcome: certify COM boundaries under malformed metadata hostile servers fault injection and repeated activate/reset/unload. Effect: delivery. Clauses: COM-EVENT-001|COM-SERVE-001|CONF-DIFF-001|CONF-QUALITY-001|JIT-PARITY-001|SYS-DUAL-001|VM3-REF-001. Canonical evidence rows: WAC-SAFETY-MUTATION WCS-SERVER-SAFETY and WCE-INCOMING-LIFECYCLE without stealing producer ownership. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane safety-com. Expected observable: deterministic errors no corruption and zero reference/connection/callback/process/registry/session drift over repeated cycles. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/safety-com.md. Residual behavior: any safety failure stays with WIN-13 or the exact COM producer. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: Hostile/malformed/fault/repeated COM cases reject or run deterministically without corruption and every owned resource returns to baseline.

### bd-59co.3.15.31 — WIN-14 certify hostile native faults and repeated lifecycles

Status at planning snapshot: open. Prerequisites: bd-59co.3.14, bd-59co.3.15.1, bd-59co.3.15.24, bd-59co.3.15.25, bd-59co.3.15.26, bd-59co.3.15.27, bd-59co.3.15.28, bd-59co.3.15.29, bd-59co.3.15.3, bd-59co.3.15.7.

Outcome: certify native import callback and output boundaries under missing dependencies fault injection stale calls and repeated load/unload. Effect: delivery. Clauses: CONF-DIFF-001|CONF-QUALITY-001|JIT-PARITY-001|NATIVE-IMPORT-001|PROFILE-TOOL-001|SYS-DUAL-001|VM3-REF-001. Canonical evidence rows: WAC-SAFETY-MUTATION WNI-PLAN-CALLBACK and WNE-NATIVE-REPRO-DEPLOY without stealing producer ownership. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane safety-native. Expected observable: deterministic errors UAF containment and zero callback/module/buffer/file/registry/process/session drift. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/safety-native.md. Residual behavior: failures stay with WIN-13 or exact native producer. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: Hostile/missing/fault/stale/repeated native cases reject or run deterministically without UAF/corruption and all owned resources balance.

### bd-59co.3.15.32 — WIN-14 aggregate Excel64 COM authority and certification

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.15, bd-59co.3.15.16, bd-59co.3.15.17, bd-59co.3.15.18, bd-59co.3.15.19, bd-59co.3.15.20, bd-59co.3.15.21, bd-59co.3.15.22, bd-59co.3.15.23, bd-59co.3.15.30.

Outcome: aggregate all certified COM client event serving carrier and authority lanes. Effect: delivery. Clauses: AUTH-CLEAN-001|AUTH-SPEC-001|AUTH-VBA-001|CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001|JIT-PARITY-001|PROFILE-WIN-001|SYS-DUAL-001|VM3-REF-001. Canonical rows: WCC-EXCEL-AUTHORITY and WAC-EXCEL-COM-CERT. Verify only when every required WCC WCE WCS and relevant WAC producer row is verified and environment/fixture/product hashes plus six axes agree. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane com-cert; command:./scripts/run-excel-vba-oracle.ps1 -Suite ComCertification; command:./scripts/run-truth-reconciliation.ps1. Expected observable: aggregate COM results full Err effects order transports and balances are complete with no open input. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/com-certification.md. Residual behavior: either aggregate row remains open for any producer or oracle gap. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: Both COM aggregate rows verify only after all producer rows and immutable evidence pass; exact matrix updates clear no residual prematurely and all six axes agree.

### bd-59co.3.15.33 — WIN-14 aggregate Excel64 native import and output certification

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.24, bd-59co.3.15.25, bd-59co.3.15.26, bd-59co.3.15.27, bd-59co.3.15.28, bd-59co.3.15.29, bd-59co.3.15.31.

Outcome: aggregate all certified Declare pointer callback wrapper genuine-output deployment and safety lanes. Effect: delivery. Clauses: AUTH-CLEAN-001|AUTH-SPEC-001|AUTH-VBA-001|CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001|JIT-PARITY-001|PROFILE-WIN-001|SYS-DUAL-001|VM3-REF-001. Canonical row: WAC-EXCEL-NATIVE-CERT. Verify only when every required WNI WNE and relevant WAC producer row is verified and immutable evidence plus six axes agree. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -Lane native-cert; command:./scripts/run-excel-vba-oracle.ps1 -Suite NativeCertification; command:./scripts/run-truth-reconciliation.ps1. Expected observable: aggregate native results full Err effects order transports deployment and balances are complete with no open input. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/native-certification.md. Residual behavior: the aggregate remains open for any producer environment or oracle gap. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: The native aggregate row verifies only after all WNI/WNE inputs and clean deployment pass; matrix residual is cleared only with immutable six-axis evidence.

### bd-59co.3.15.34 — WIN-14 archive immutable certification transcripts and replay index

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.3, bd-59co.3.15.32, bd-59co.3.15.33.

Outcome: archive an immutable replayable index over every WIN-14 environment source fixture product artifact log VBE/UIA and cleanup transcript. Effect: support. Clauses: CONF-ORACLE-001|CONF-QUALITY-001. Canonical rows: WAC-EXCEL-COM-CERT WAC-EXCEL-NATIVE-CERT and WAC-RELEASE-CERT. Acceptance commands: command:./scripts/archive-win14-certification.ps1 -Verify. Expected observable: every certification claim resolves through hashes to exact inputs/results and mutable latest pointers are never sole evidence. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/archive-manifest.json. Residual behavior: missing or mismatched evidence keeps final certification open. Fresh-eyes non-author review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: Archive verification covers all case manifests and transcripts with immutable hashes, owned cleanup evidence and replay paths; no mutable alias is sole authority.

### bd-59co.3.15.6 — WIN-14 promote the owned Excel64 VBE oracle supervisor

Status at planning snapshot: open. Prerequisites: bd-59co.3.1.2, bd-59co.3.15.1, bd-59co.3.15.35, bd-59co.3.15.36, bd-59co.3.15.4.

Outcome: provide the authoritative owned Excel/VBE compile and runtime supervisor. Effect: support. Clauses: AUTH-VBA-001|CONF-ORACLE-001|CONF-QUALITY-001|PROFILE-WIN-001. Canonical rows: WCC-EXCEL-AUTHORITY and WAC-EXCEL-COM-CERT. Build on existing guardians but start UIA before runs force VBE Debug Compile never use Application.Run as compile check capture modal text selected token expanded line module source and full runtime Err; touch owned PIDs/dialogs only. Acceptance commands: command:./scripts/run-excel-vba-oracle.ps1 -Suite HarnessSelfTest -EnvironmentId win-x64-dev-oracle-2026-07 -NoMatrixUpdate; command:./scripts/test-excel-vba-oracle.ps1. Expected observable: success compile failure ambiguous macro failure intrinsic-shadow and runtime-error self-tests yield deterministic owned transcripts. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/oracle-supervisor.md. Residual behavior: unsafe or ambiguous modal handling blocks Excel evidence. Fresh-eyes Sol review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: The supervisor self-tests all required modal and full-Err shapes, applies PID-scoped cleanup, proves VBE compile invocation and does not mutate matrices or claim release evidence on the dev host.

### bd-59co.3.15.7 — WIN-14 implement the resumable Windows profile checks orchestrator

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.1, bd-59co.3.15.4, bd-59co.3.15.5, bd-59co.3.15.6.

Outcome: implement one fail-closed resumable orchestrator for all Windows certification lanes. Effect: support. Clauses: CONF-DIFF-001|CONF-ORACLE-001|CONF-QUALITY-001. Canonical aggregate rows: WAC-EXCEL-COM-CERT WAC-EXCEL-NATIVE-CERT and WAC-RELEASE-CERT. The runner reads the case manifest refuses unverified environment fixture or producer checks serializes Cargo Excel and registry lanes supports resume/retry without overwriting captures and hashes every result manifest. Acceptance commands: command:./scripts/run-windows-profile-gates.ps1 -List -ValidateOnly; command:./scripts/test-windows-profile-gates.ps1. Expected observable: deterministic plans and self-tests prove refusal serialization timeout failure resume and immutable capture behavior. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/profile-runner.md. Residual behavior: an unproven producer or environment leaves the exact lane blocked. Fresh-eyes non-author review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: The orchestrator validates and lists every case, fails closed on unmet checks, serializes constrained resources, resumes without overwriting evidence and emits deterministic hashed manifests.

### bd-59co.3.15.8 — WIN-14 inventory and freeze controlled certification fixtures

Status at planning snapshot: open. Prerequisites: bd-59co.3.1, bd-59co.3.15.1, bd-59co.3.15.4.

Outcome: freeze the controlled source/build/client inventory consumed by the certification VM. Effect: support. Clauses: CONF-ORACLE-001|CONF-QUALITY-001|PROFILE-WIN-001. Canonical rows: WAC-CLEAN-CERT-ENV and all fixture-bearing Windows rows. Cover x64 carrier DLL Automation dual/custom in/out-proc servers events and cancellation callback/reentry DLL C/Rust clients Excel sources build inputs and owned cleanup with separate recipe and artifact hashes. Acceptance commands: command:./scripts/build-win14-controlled-fixtures.ps1 -ValidateManifest -NoBuild; command:./scripts/validate-windows-fixture-manifest.ps1. Expected observable: every certification case resolves to a frozen recipe and pending or built artifact without historical-binary substitution. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/fixtures-manifest.json. Residual behavior: missing products remain blocked on their producer epic. Fresh-eyes non-author review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: The fixture manifest is complete, source/build and artifact hashes are distinct, controlled cleanup is explicit and validation rejects historical or architecture-mismatched artifacts.

### bd-59co.3.15.9 — WIN-14 promote x64 COM carrier event and serving fixtures

Status at planning snapshot: open. Prerequisites: bd-59co.3.1, bd-59co.3.15.1, bd-59co.3.15.8.

Outcome: build and self-test deterministic x64 COM certification fixtures. Effect: support. Clauses: CONF-ORACLE-001|CONF-QUALITY-001|PROFILE-WIN-001. Canonical matrices: WIN-COM-CLIENT WIN-COM-EVENTS WIN-COM-SERVER and WIN-ABI-CARRIER. Prove typelibs GUID/version/LCID/platform transport counters in/out-proc dual/custom VT_RECORD interface arrays cancellable ByRef event and owned register/unregister. Acceptance commands: command:./scripts/build-win14-controlled-fixtures.ps1 -Domain Com -Configuration Release; command:./scripts/validate-windows-fixture-manifest.ps1. Expected observable: reproducible fixture hashes and self-tests cover exact COM transports and leave zero owned registry process connection or reference drift. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-14/fixtures/com/manifest.json. Residual behavior: fixture success does not certify VM3 JIT or Excel capability. Fresh-eyes non-author review is required.

Canonical matrix truth surface: the exact row or rows named above in the manifest-owned Windows matrices and IDEAL_MATRIX_BEAD_TRACEABILITY_V1.csv.

Acceptance: COM fixtures build reproducibly, self-tests and transport counters pass, owned registration cleans completely, and all product capability rows remain with their producer owners.

### bd-59co.3.16.1 — Roll out WIN-15 executable child beads

Status at planning snapshot: open. Prerequisites: no direct blocking edge; inspect parent prerequisites.

Outcome: create or refresh the executable child bead set for WIN-15 without substituting planning for delivery. The parent contract is reproduced below so future workers do not need the workset document.

Type: support/conformance
Clauses: `CONF-DONE-001`, `DOC-AUTH-001`, `DOC-TRACE-001`

Reconcile system/Windows/JIT/package contracts, current architecture, code, matrices, old COM ladders/blockers, output labels and derived compatibility/native-output reports. Run final security/deployment/user-path fresh-eyes review.

Close: both internal claim checks and every required delivery epic are green.

Every child must identify one reviewable outcome, delivery/support effect, contract clauses, exact dependencies, canonical matrix rows/truth surfaces, targeted tests or evidence, relative complexity, fresh-eyes review, and blocker/residual behavior.

Executable rollout acceptance: command:./scripts/validate-workset-rollout.ps1; command:./scripts/run-truth-reconciliation.ps1; artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-15/rollout-acceptance.md. Expected observable: bounded delivery/support children exist with exact row traces, producer dependencies, typed acceptance commands and residual owners; no capability closes on this support bead.

Tooling-profile clause: PROFILE-TOOL-001; wrapped and genuine native outputs remain distinct and terminal tooling claims require their own verified rows.

Exact trace-contract addendum:
Clauses: `PROFILE-WIN-001`
These clauses are the exact producer/consumer and matrix-boundary union for this leaf; support relationships do not close capability.

Acceptance: A believable ready path exists for WIN-15: first delivery beads are explicit, dependencies reflect engineering truth, tests/evidence are included, no accepted residual is only narrative, and capability closure cannot occur on support beads alone. The two named commands must pass and the rollout artifact must enumerate exact child IDs, matrix rows, commands, observables and remaining owners.

### bd-59co.3.3.5 — WIN-2 realize remaining session cache and broader plan kinds

Status at planning snapshot: open. Prerequisites: bd-59co.3.12.4, bd-59co.3.3.10, bd-59co.3.3.11, bd-59co.3.3.12, bd-59co.3.3.6, bd-59co.3.3.7, bd-59co.3.3.8, bd-59co.3.3.9.

Outcome: integrate and prove the remaining behavior owned by WIN-2 realize remaining session cache and broader plan kinds.

Effect: delivery. Clauses: HOST-SESSION-001|IMAGE-VERIFY-001|JIT-CACHE-001|JIT-CORE-001|RUNTIME-ABI-001|RUNTIME-EVAL-001|SYS-OWN-001|WIN-PLAN-001.

Prerequisites: bd-59co.3.3.6, bd-59co.3.3.7, bd-59co.3.3.9, bd-59co.3.3.8, bd-59co.3.3.10, bd-59co.3.3.11, bd-59co.3.3.12, bd-59co.3.12.4. These leaves own the component implementations; this bead owns their combined behavior through one production route.

Steps: validate the producer artifacts; execute combined success/error/alias/reentry/lifecycle cases under VM3 and JIT; verify the same immutable plan and session identity at each required boundary; correct any integration defect and add its regression.

Canonical matrix rows: WIN-ABI-CARRIER/WAC-VERIFIED-INTEROP-PLAN; WIN-ABI-CARRIER/WAC-WINDOWS-DESCRIPTORS.

Acceptance commands: command:cargo test -p oxvba-differential --test jit_windows_vm3_parity -- --nocapture; command:./scripts/run-truth-reconciliation.ps1.

Expected observable: exact structural results, full Err/LastDllError where relevant, side effects, lifecycle ordering, native transport and zero balance drift across the combined producer paths. A matching decline, fixture-specific shortcut or paper split does not satisfy this delivery outcome.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-2/remaining-session-attachment.md.

Fresh-eyes review must probe missed cross-feature aliases, errors, threading and cleanup and rerun relevant checks after repairs.

Residual behavior: this bead stays open until its scoped integrated behavior and required producer evidence pass. Newly discovered gaps require exact open delivery successors and keep the broader area in-progress. Full canonical-row certification remains with the applicable WIN-14 evidence owners; no planning-only closure is permitted.

Acceptance: All named producer leaves pass and their combined native behavior is proven under VM3/JIT with exact structural/error/transport/lifetime evidence and a clean fresh-eyes review. Creating successor beads alone cannot close this delivery item.

### bd-59co.3.4.2 — WIN-3 late-bound scalar CreateObject method property

Status at planning snapshot: open. Prerequisites: bd-59co.3.3.4, bd-59co.3.4.1.

Outcome: deliver the first WIN-3 late-bound scalar CreateObject/method/property row under VM3 and JIT using the shared verified interop plan. Effect: delivery. Clauses: COM-CLIENT-001|WIN-PLAN-001. Canonical matrix rows: WIN-COM-CLIENT/WCC-PLAN-LATE. Controlled OxVba.TestDispatch fixture first; Excel Application rows stay planned under WIN-14. Acceptance commands: command:cargo test -p oxvba-differential --test jit_windows_vm3_parity -- --nocapture. Expected observable: VM3 and JIT match on result full Err raised and handle-balance for the scalar late COM fixture, or an exact owned residual is named. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-3/late-scalar-client.md. Residual behavior: named/omitted/ByRef/arrays/records/GetObject/out-of-proc stay with bd-59co.3.4.3. Fresh-eyes review is required. Canonical matrix truth surface: the exact row named above in WINDOWS_JIT_COM_CLIENT_MATRIX_V1.csv.

Detailed acceptance addition: Expand the controlled TestDispatch fixture to actual scalar methods plus property get/put. Assert successful native execution and dispatch counters; matching declines are failures. Keep this bounded first slice independent of the later common-session delivery.

Acceptance: Scalar late CreateObject/method/property fixture matches under VM3 and JIT through the shared plan; Excel rows stay planned. Fresh-eyes review is clean.

### bd-59co.3.4.3 — WIN-3 realize remaining late-bound COM client rows

Status at planning snapshot: open. Prerequisites: bd-59co.3.4.2, bd-59co.3.4.4, bd-59co.3.4.5, bd-59co.3.4.6, bd-59co.3.4.7, bd-59co.3.4.8.

Outcome: integrate and prove the remaining behavior owned by WIN-3 realize remaining late-bound COM client rows.

Effect: delivery. Clauses: COM-CLIENT-001|WIN-PLAN-001.

Prerequisites: bd-59co.3.4.2, bd-59co.3.4.4, bd-59co.3.4.5, bd-59co.3.4.6, bd-59co.3.4.7, bd-59co.3.4.8. These leaves own the component implementations; this bead owns their combined behavior through one production route.

Steps: validate the producer artifacts; execute combined success/error/alias/reentry/lifecycle cases under VM3 and JIT; verify the same immutable plan and session identity at each required boundary; correct any integration defect and add its regression.

Canonical matrix rows: WIN-COM-CLIENT/WCC-PLAN-LATE.

Acceptance commands: command:cargo test -p oxvba-differential --test jit_windows_vm3_parity -- --nocapture; command:./scripts/run-truth-reconciliation.ps1.

Expected observable: exact structural results, full Err/LastDllError where relevant, side effects, lifecycle ordering, native transport and zero balance drift across the combined producer paths. A matching decline, fixture-specific shortcut or paper split does not satisfy this delivery outcome.

Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-3/remaining-late-client.md.

Fresh-eyes review must probe missed cross-feature aliases, errors, threading and cleanup and rerun relevant checks after repairs.

Residual behavior: this bead stays open until its scoped integrated behavior and required producer evidence pass. Newly discovered gaps require exact open delivery successors and keep the broader area in-progress. Full canonical-row certification remains with the applicable WIN-14 evidence owners; no planning-only closure is permitted.

Acceptance: All named producer leaves pass and their combined native behavior is proven under VM3/JIT with exact structural/error/transport/lifetime evidence and a clean fresh-eyes review. Creating successor beads alone cannot close this delivery item.

### bd-59co.3.6.2 — WIN-5 execute synchronous ByRef event plan in VM3

Status at planning snapshot: open. Prerequisites: bd-59co.2.7.4, bd-59co.3.3.10, bd-59co.3.3.9, bd-59co.3.6.1, bd-59co.3.6.5, bd-aprs.8.8.9.

Bounded outcome: Execute the verified scalar cancellable event plan in VM3 with same-call handler and copyback. Clauses: COM-EVENT-001|WIN-PLAN-001. Matrix: WIN-COM-EVENTS. Acceptance references: command:cargo test -p oxvba-vm3 vm3_com_event_scalar_byref_copyback -- --nocapture. Expected observable: writeback occurs before native return with exact Err/order and balanced Advise/Unadvise. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-5/byref-event-vm3.md. Residual behavior: JIT and Excel certification remain siblings. Fresh-eyes non-author review is required.

Detailed Windows plan handoff: bd-59co.3.6.5, bd-59co.3.3.9, bd-59co.3.3.10, bd-59co.2.7.4. Consume the owned shared-session and immutable plan contracts; preserve the existing fixture producer.

Acceptance: The named command/environment references pass or resolve; expected observable is recorded in artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-5/byref-event-vm3.md; residual scope remains explicitly owned.

### bd-59co.3.6.3 — WIN-5 execute synchronous ByRef event plan in JIT

Status at planning snapshot: open. Prerequisites: bd-59co.2.7.7, bd-59co.3.3.10, bd-59co.3.3.9, bd-59co.3.6.1, bd-59co.3.6.5, bd-aprs.8.8.9.

Bounded outcome: Execute the unchanged verified scalar cancellable event plan in the persistent JIT session. Clauses: COM-EVENT-001|WIN-PLAN-001|JIT-PARITY-001. Matrix: WIN-COM-EVENTS. Acceptance references: command:cargo test -p oxvba-jit jit_com_event_scalar_byref_copyback -- --nocapture. Expected observable: JIT transport and complete structural observable equal VM3 without queue fallback. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-5/byref-event-jit.md. Residual behavior: Excel certification remains a sibling. Fresh-eyes non-author review is required.

Detailed Windows plan handoff: bd-59co.3.6.5, bd-59co.3.3.9, bd-59co.3.3.10, bd-59co.2.7.7. Consume the owned shared-session and immutable plan contracts; preserve the existing fixture producer.

Acceptance: The named command/environment references pass or resolve; expected observable is recorded in artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-5/byref-event-jit.md; residual scope remains explicitly owned.

### bd-59co.3.6.4 — WIN-5 certify scalar ByRef event in Excel and lifecycle lanes

Status at planning snapshot: open. Prerequisites: bd-59co.3.15.3, bd-59co.3.15.6, bd-59co.3.6.1, bd-59co.3.6.2, bd-59co.3.6.3.

Bounded outcome: Run Excel/VBA and repeated native-client lifecycle certification after both event adapters. Clauses: COM-EVENT-001|CONF-ORACLE-001|CONF-QUALITY-001. Matrix: WIN-COM-EVENTS. Acceptance references: command:./scripts/run-excel-vba-oracle.ps1|command:cargo test -p oxvba-host --test com_matrix_events -- --ignored --nocapture|environment:win-x64-cert-vm-pending-v1. Expected observable: VBA writeback-before-return and repeated result/Err/order/transport/balance agree. Evidence: artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-5/byref-event-certification.md. Residual behavior: richer event rows remain open. Fresh-eyes non-author review is required.

Detailed Windows plan handoff: bd-59co.3.15.3, bd-59co.3.15.6. Consume the owned shared-session and immutable plan contracts; preserve the existing fixture producer. The certification environment placeholder must be resolved through the sealed environment identity produced by bd-59co.3.15.3 before execution. Capture the actual environment ID/hash in the transcript; never certify on the development host.

Acceptance: The named command/environment references pass or resolve; expected observable is recorded in artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-5/byref-event-certification.md; residual scope remains explicitly owned.

### bd-9sed.17 — Native AddressOf callback lifetime residuals beyond CallWindowProc

Status at planning snapshot: open. Prerequisites: bd-59co.3.10.5, bd-59co.3.11.1, bd-59co.3.3.12.

Bounded delivery slice: define the backend-neutral retained x64 callback ownership record and controlled native/oracle fixture for one externally owned callback such as SetTimer, including signature, owner session, release trigger and late-call rejection. Clauses: `NATIVE-IMPORT-001`, `WIN-PLAN-001`, `SEC-BOUNDARY-001`. Matrix row: `WNI-PLAN-CALLBACK`. Acceptance command: `command:cargo test -p oxvba-runtime retained_callback_owner_fixture_contract -- --nocapture`. Expected observable: the fixture records entry while live, deterministic rejection after release, and balanced allocation/drop without AV. Evidence artifacts: `artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-10/retained-callback-fixture.md` and `oracle:win-x64-dev-oracle-2026-07`. Residual behavior: VM3/JIT adapters and disposal/exhaustion/reentry certification are separate siblings; no non-x64 scope is created.

Acceptance: The named fixture test and dev-oracle transcript establish one exact retained-callback lifetime contract with balanced ownership and late-call rejection; backend and terminal certification siblings remain explicit.

### bd-aprs.8.8.9 — FE-7.6.a TestEventServer ByRef COM event writeback parity

Status at planning snapshot: open. Prerequisites: bd-59co.3.3.10, bd-59co.3.6.1.

Bounded delivery slice: add the verified x64 interop-plan shape and controlled native fixture transport for one synchronous cancellable scalar `VT_BYREF` connection-point event, including source-slot ownership and copyback-before-return order. Clauses: `COM-EVENT-001`, `WIN-PLAN-001`, `SEC-BOUNDARY-001`. Matrix row: `WCE-PLAN-INCOMING`. Acceptance command: `command:cargo test -p oxvba-com event_plan_scalar_byref_copyback_contract -- --nocapture`. Expected observable: verifier accepts the exact fixture plan, rejects queued/asynchronous copyback, and the fixture observes cancellation writeback before native return with balanced Advise/Unadvise. Evidence artifact: `artifact:docs/evidence/programs/ideal-2026-07/windows-x64/WIN-5/byref-event-plan-fixture.md`. Residual behavior: VM3 adapter, JIT adapter and Excel/lifecycle certification are separate siblings; richer payload/apartment rows stay in WIN-5 rollout.

Acceptance: The named controlled-fixture test proves verifier-checked source-slot ownership and synchronous writeback order, rejects the queue-backed shape, records transport/balance evidence, and leaves backend/certification siblings open.

## Coverage and remaining programme boundaries

All six Windows matrices retain their original truth states. The table maps detailed new producers and existing certification owners to each row. A deferred native row has an existing future owner; it is not marked complete or silently removed.

| Matrix row | New detailed producers | Evidence owner | Disposition |
|---|---|---|---|
| WCC-PLAN-LATE | bd-59co.3.4.4 | bd-59co.3.4.2 | delivery and/or existing WIN-14 certification required |
| WCC-LATE-ARGS | bd-59co.3.4.5, bd-59co.3.4.6 | bd-59co.3.4.6 | delivery and/or existing WIN-14 certification required |
| WCC-LATE-STRUCTURAL | bd-59co.3.4.7 | bd-59co.3.4.7 | delivery and/or existing WIN-14 certification required |
| WCC-LATE-OUTPROC-ERROR | bd-59co.3.4.4, bd-59co.3.4.8 | bd-59co.3.4.8 | delivery and/or existing WIN-14 certification required |
| WCC-PLAN-EARLY | bd-59co.3.5.2 | bd-59co.3.5.2 | delivery and/or existing WIN-14 certification required |
| WCC-EARLY-COMPLEX | bd-59co.3.5.3 | bd-59co.3.5.3 | delivery and/or existing WIN-14 certification required |
| WCC-EARLY-CUSTOM | bd-59co.3.5.4 | bd-59co.3.5.4 | delivery and/or existing WIN-14 certification required |
| WCC-EARLY-OUTPROC | bd-59co.3.5.5 | bd-59co.3.5.5 | delivery and/or existing WIN-14 certification required |
| WCC-EXCEL-AUTHORITY | existing bead(s) below | bd-59co.3.15.32 | delivery and/or existing WIN-14 certification required |
| WCE-PLAN-INCOMING | bd-59co.3.6.5 | bd-59co.3.6.4 | delivery and/or existing WIN-14 certification required |
| WCE-INCOMING-COMPLEX | bd-59co.3.6.6 | bd-59co.3.6.6 | delivery and/or existing WIN-14 certification required |
| WCE-INCOMING-CUSTOM | bd-59co.3.6.7 | bd-59co.3.6.7 | delivery and/or existing WIN-14 certification required |
| WCE-INCOMING-APARTMENT | bd-59co.3.6.8 | bd-59co.3.6.8 | delivery and/or existing WIN-14 certification required |
| WCE-INCOMING-LIFECYCLE | bd-59co.3.6.5, bd-59co.3.6.8 | bd-59co.3.6.8 | delivery and/or existing WIN-14 certification required |
| WCE-PLAN-OUTGOING | bd-59co.3.9.2 | bd-59co.3.9.2 | delivery and/or existing WIN-14 certification required |
| WCE-OUTGOING-COMPLEX | bd-59co.3.9.3 | bd-59co.3.9.3 | delivery and/or existing WIN-14 certification required |
| WCS-LATE-INPROC | bd-59co.3.7.2, bd-59co.3.7.3 | bd-59co.3.7.3 | delivery and/or existing WIN-14 certification required |
| WCS-LATE-LOCALSERVER | bd-59co.3.7.5 | bd-59co.3.7.5 | delivery and/or existing WIN-14 certification required |
| WCS-LATE-COMPLEX | bd-59co.3.7.3, bd-59co.3.7.4 | bd-59co.3.7.4 | delivery and/or existing WIN-14 certification required |
| WCS-DUAL-INPROC | bd-59co.3.8.2 | bd-59co.3.8.2 | delivery and/or existing WIN-14 certification required |
| WCS-IMPLEMENTS-CUSTOM | bd-59co.3.8.3 | bd-59co.3.8.3 | delivery and/or existing WIN-14 certification required |
| WCS-EARLY-OUTPROC | bd-59co.3.8.4 | bd-59co.3.8.4 | delivery and/or existing WIN-14 certification required |
| WCS-SERVER-SAFETY | bd-59co.3.7.6 | bd-59co.3.7.6 | delivery and/or existing WIN-14 certification required |
| WNI-PLAN-DECLARE | bd-59co.3.10.5 | bd-59co.3.10.5 | delivery and/or existing WIN-14 certification required |
| WNI-DECLARE-STRINGS | bd-59co.3.10.6 | bd-59co.3.10.6 | delivery and/or existing WIN-14 certification required |
| WNI-DECLARE-STRUCTURAL | bd-59co.3.10.7 | bd-59co.3.10.7 | delivery and/or existing WIN-14 certification required |
| WNI-DECLARE-LOADER-ERROR | bd-59co.3.10.4 | bd-59co.3.10.4 | delivery and/or existing WIN-14 certification required |
| WNI-POINTER-HELPERS | bd-59co.3.11.5 | bd-59co.3.11.5 | delivery and/or existing WIN-14 certification required |
| WNI-CALLBACK-SYNC | bd-59co.3.11.6 | bd-59co.3.11.6 | delivery and/or existing WIN-14 certification required |
| WNI-PLAN-CALLBACK | existing bead(s) below | bd-59co.3.11.4 | delivery and/or existing WIN-14 certification required |
| WNI-CALLBACK-NESTED | bd-59co.3.11.7 | bd-59co.3.11.7 | delivery and/or existing WIN-14 certification required |
| WNE-WRAPPER-EXE | bd-59co.3.12.2, bd-59co.3.12.5 | bd-59co.3.12.2 | delivery and/or existing WIN-14 certification required |
| WNE-WRAPPER-LIBRARY | bd-59co.3.12.3, bd-59co.3.12.5 | bd-59co.3.12.3 | delivery and/or existing WIN-14 certification required |
| WNE-PLAN-WRAPPED | bd-59co.3.12.4, bd-59co.3.12.5 | bd-59co.3.12.4 | delivery and/or existing WIN-14 certification required |
| WNE-PLAN-NATIVE | existing bead(s) below | bd-59co.3.13.1 | deferred native/full-programme aggregate; stays open |
| WNE-NATIVE-EXE | existing bead(s) below | bd-59co.3.13.1 | deferred native/full-programme aggregate; stays open |
| WNE-NATIVE-ABI-BREADTH | existing bead(s) below | bd-59co.3.13.1 | deferred native/full-programme aggregate; stays open |
| WNE-NATIVE-REPRO-DEPLOY | existing bead(s) below | bd-59co.3.13.1 | deferred native/full-programme aggregate; stays open |
| WNE-PROFILE-TOOL-TERMINAL | existing bead(s) below | bd-59co.3.16.1 | deferred native/full-programme aggregate; stays open |
| WAC-BSTR-LAYOUT | bd-59co.3.2.5 | bd-59co.3.2.5 | delivery and/or existing WIN-14 certification required |
| WAC-VARIANT-LAYOUT | bd-59co.3.2.5 | bd-59co.3.2.5 | delivery and/or existing WIN-14 certification required |
| WAC-SAFEARRAY-LAYOUT | bd-59co.3.2.6 | bd-59co.3.2.6 | delivery and/or existing WIN-14 certification required |
| WAC-IUNKNOWN-IDENTITY | bd-59co.3.2.8 | bd-59co.3.2.8 | delivery and/or existing WIN-14 certification required |
| WAC-NUMERIC-LONGPTR | bd-59co.3.2.5 | bd-59co.3.2.5 | delivery and/or existing WIN-14 certification required |
| WAC-INTERFACE-ARRAY | bd-59co.3.2.6 | bd-59co.3.2.6 | delivery and/or existing WIN-14 certification required |
| WAC-VT-RECORD | bd-59co.3.2.7 | bd-59co.3.2.7 | delivery and/or existing WIN-14 certification required |
| WAC-CARRIER-EXCEL-ROUNDTRIP | existing bead(s) below | bd-59co.3.15.16 | delivery and/or existing WIN-14 certification required |
| WAC-SAFETY-MUTATION | bd-59co.3.14.3, bd-59co.3.14.4, bd-59co.3.14.5, bd-59co.3.14.6 | bd-59co.3.14.2 | delivery and/or existing WIN-14 certification required |
| WAC-TARGET-DEV-ENV | existing bead(s) below | bd-59co.3.1.2 | development baseline already verified |
| WAC-TYPELIB-METADATA | bd-59co.3.2.2, bd-59co.3.2.3, bd-59co.3.2.4 | bd-59co.3.2.4 | delivery and/or existing WIN-14 certification required |
| WAC-VERIFIED-INTEROP-PLAN | bd-59co.3.3.7, bd-59co.3.3.8, bd-59co.3.3.9, bd-59co.3.3.10, bd-59co.3.3.11, bd-59co.3.3.12 | bd-59co.3.3.5 | delivery and/or existing WIN-14 certification required |
| WAC-WINDOWS-DESCRIPTORS | bd-59co.3.3.6 | bd-59co.3.3.6 | delivery and/or existing WIN-14 certification required |
| WAC-CLEAN-CERT-ENV | existing bead(s) below | bd-59co.3.15.3 | delivery and/or existing WIN-14 certification required |
| WAC-RELEASE-CERT | existing bead(s) below | bd-59co.3.15.2 | deferred native/full-programme aggregate; stays open |
| WAC-EXCEL-COM-CERT | existing bead(s) below | bd-59co.3.15.32 | delivery and/or existing WIN-14 certification required |
| WAC-EXCEL-NATIVE-CERT | existing bead(s) below | bd-59co.3.15.33 | deferred native/full-programme aggregate; stays open |
| WAC-PROFILE-TERMINAL | existing bead(s) below | bd-59co.3.16.1 | deferred native/full-programme aggregate; stays open |

## Planning review and implementation handoff

Planning evidence and review findings: [review.md](../evidence/programs/ideal-2026-07/windows-x64/PLAN-2026-09-21/review.md). Planning finishes after graph, traceability, governance and review checks; it does not start the capability implementation loop.
