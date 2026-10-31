# Formal Extended Todo

- ID: FTODO-TYPED-HOST-20261031-R7
  Profile: bounded host API/typed slots/reentry (`bd-59co.2.10.9`, `bd-59co.2.9.12`, `bd-59co.2.4.9`).
  Summary: Lean TypedReentry models context/depth admission and shared finite budget; existing scalar/handle/refcount abstractions remain applicable. Installed nightly strict-provenance Miri checks the host object's complete-allocation IUnknown retain/QI/release path and bounded VM3 callbacks/reentry/active-view resolution. It found an invalid cached pointer into moved Box storage; immutable Rc table storage and provenance-preserving union slots correct it. These bounded executions do not prove general Rust session aliasing, native function signatures, generated slot loads/call_indirect, all callback destruction paths or general object-default coercion.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: installed Lean 4.29.1 on formal/lean/OxVba/TypedReentry.lean; `cargo +nightly miri test -p calcvba-host-api native_unknown_prefix_retains_the_exact_object_until_last_release`; `cargo test -p oxvba-embed --test typed_host_callbacks`.
  Suggested next action: refine the ABI, code/object leases and callback alias model under CORE-5/CORE-7/CORE-8 before general owning carriers, stateful synchronous callbacks or COM-event reentry claims. Broad owners remain bd-59co.2.7.4, bd-59co.2.9.9, bd-59co.2.10.7 and bd-59co.2.4.3.

- ID: FTODO-EARLY-COM-20261031
  Profile: preserved-descriptor dual COM calls (`bd-59co.3.5.6`; broader WIN-4 ownership `.3.5.2/.3`).
  Summary: installed Lean 4.29.1 VarType/Coerce/Arithmetic/RefCount/ExecutionBudget abstractions pass. They do not prove Rust descriptor provenance, call-local COM/QI ownership, native ByRef alias cells, private temporaries, error-edge cleanup or actual transport. Native cargo-kani is unavailable; no toolchain was installed.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `cargo test -p oxvba-differential --test jit_windows_vm3_parity early_com -- --test-threads=1`; compile VarType.lean first to an output OxVba/VarType.olean and set LEAN_PATH to that output before the dependent models.
  Suggested next action: refine shared native-call ownership and alias/error outcomes under WIN-4/CORE-5; prove success/failure cleanup and retain broader callback/apartment/verified-image obligations. Registry-writing record fixtures are excluded when using meta-check -NoRegistry and receive no passing evidence from that run.

Non-blocking formal issues and follow-up items for later ladder profiles.

## Template
- ID:
- Profile:
- Summary:
- Current status (`todo` / `investigating` / `resolved`):
- Reproduction command:
- Suggested next action:

## Active Items

- ID: FTODO-TYPED-HOST-20261031
  Profile: bounded typed host entries and neutral session handles (`bd-59co.2.9.11`, `bd-59co.2.10.8`; broader CORE-7/8 remain open).
  Summary: installed Lean 4.29.1 passes VarType, Coerce, Arithmetic, RefCount, ExecutionBudget and new ProcedureHandles abstractions. The new model proves stale-generation, different-owner and unloaded rejection plus reload invalidation in an abstract model. Rust/native ABI signature correctness, physical-slot specialization, pointer provenance/code lifetime, generation-identity refinement and error/budget adapter refinement remain unproved. Native Kani is unavailable. The pinned Lean launcher attempted a download; it was canceled, and automatic approval review blocked removal of its uninstalled temporary folder.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: run the installed Lean 4.29.1 binary directly for formal/lean/OxVba/ProcedureHandles.lean; run `cargo test -p oxvba-embed` and `cargo test -p oxvba-jit stale_typed_address`.
  Suggested next action: refine owner/generation and finite-block budget models to the Rust checks, verify generated argument/result signatures and publish nonempty Rust/native proofs on an already approved runner. Keep general carriers, ByRef, recursion, reentry and verified image ownership with existing owners; executable tests and abstraction proofs do not certify native memory safety.
- ID: FTODO-NATIVE-ALIAS-20261031
  Profile: WIN-9 scalar ABI and Variant cells (`bd-59co.3.10.5`, `bd-59co.3.10.9`)
  Summary: Controlled native and Excel fixtures cover scalar ABI classes, caller-variable aliases, scalar/BSTR Variant cells, per-argument ANSI String temporaries and cleanup. The Variant route shares one existing boxed VARIANT, clones its BSTR readback once and clears one native owner. Runtime handle counters do not count OLE copied-BSTR allocations. No existing formal model proves these native paths. Native cargo-kani is unavailable. Fresh-eyes review removed a second mutable projection of shared boxed cells and derives pointers after storing owners; executable tests alone do not prove pointer provenance.
  Current status (`todo` / `investigating` / `resolved`): todo
  Reproduction command: `cargo test -p oxvba-differential --test jit_windows_vm3_parity`; `cargo test -p oxvba-hal legacy_adapter_declines_repeated_aliases_before_execution`
  Suggested next action: extract a portable scalar-cell ownership/alias harness for Miri on the existing approved runner, then connect caller identity, one-cell ownership and positional copyback to the implementation. Keep array/object/record/procedure-reference Variant aliases, native String pointer replacement/nullable storage and mixed As Any aliases, noncanonical Boolean payloads and reentry under their existing delivery owners.
- ID: FTODO-UDF-BUDGET-20261031
  Profile: CORE-8 value embedding (`bd-59co.2.10.6`)
  Summary: The new abstract ExecutionBudget Lean model proves sticky zero, exact decrement and monotonicity (all five Lean modules passed with installed Lean 4.29.1; pinned 4.19 was absent and unchanged). Refinement to the Rust u64 counter and VM3/JIT noncatchable abort propagation remain unproved; executable coverage includes u64 boundaries, loops and recursion. Native `cargo kani --version` reports no such command on this Windows host.
  Current status (`todo` / `investigating` / `resolved`): todo
  Reproduction command: `cargo test -p oxvba-rt-abi execution_budget_exhaustion`; `cargo test -p oxvba-embed budget`
  Suggested next action: connect the abstract step proof to the implementation and model noncatchable abort propagation, then run a nonempty proof lane on the existing approved formal runner. Keep native-call preemption and unbounded native recursion outside the cooperative guarantee. Do not count retired-crate or zero-selected harness runs as evidence.
- ID: FTODO-M4-1-001
  Profile: M4-1 JIT IR-prep
  Summary: The non-blocking formal runner could not start in the current Linux environment because neither `pwsh` nor `powershell` is available on `PATH`.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `pwsh -NoProfile -ExecutionPolicy Bypass -File scripts/run-formal.ps1 -Quiet -NoArtifacts`
  Suggested next action: install PowerShell for this Linux runner or rerun the no-artifact formal lane on a host with PowerShell available.
- ID: FTODO-M4-2-001
  Profile: M4-2 runtime ABI substrate
  Summary: The non-blocking formal runner could not start in the current Linux environment because neither `pwsh` nor `powershell` is available on `PATH`.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `pwsh -NoProfile -ExecutionPolicy Bypass -File scripts/run-formal.ps1 -Quiet -NoArtifacts`
  Suggested next action: install PowerShell for this Linux runner or rerun the no-artifact formal lane on a host with PowerShell available.
- ID: FTODO-M4-3-001
  Profile: M4-3 Cranelift skeleton
  Summary: The non-blocking formal runner could not start in the current Linux environment because neither `pwsh` nor `powershell` is available on `PATH`.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `pwsh -NoProfile -ExecutionPolicy Bypass -File scripts/run-formal.ps1 -Quiet -NoArtifacts`
  Suggested next action: install PowerShell for this Linux runner or rerun the no-artifact formal lane on a host with PowerShell available.
- ID: FTODO-M4-4-001
  Profile: M4-4 control/calls implemented-subset
  Summary: The non-blocking formal runner could not start in the current Linux environment because neither `pwsh` nor `powershell` is available on `PATH`.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `pwsh -NoProfile -ExecutionPolicy Bypass -File scripts/run-formal.ps1 -Quiet -NoArtifacts`
  Suggested next action: install PowerShell for this Linux runner or rerun the no-artifact formal lane on a host with PowerShell available.
- ID: FTODO-V2-001
  Profile: v2-v4
  Summary: Native Windows `cargo-kani` remains unavailable; strict Kani is routed through WSL lanes.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `cargo kani --version` (Windows), `wsl bash -lc 'source $HOME/.cargo/env && cargo kani --version'` (WSL)
  Suggested next action: keep strict lanes on WSL path and avoid relying on native Windows Kani for profile gating.
- ID: FTODO-V72-001
  Profile: v72 reconciliation (`v67..v71` DG foldback)
  Summary: Deferred-gate strict Kani runs `v67-kani` through `v71-kani` completed and were folded during v86 reconciliation.
  Current status (`todo` / `investigating` / `resolved`): resolved
  Reproduction command: `./scripts/run-formal-kani-async.ps1 -Action Status -Name v67-kani` (repeat for `v68-kani`..`v71-kani`)
  Suggested next action: none.
- ID: FTODO-V76-001
  Profile: v76 reconciliation (`v73..v75` DG foldback)
  Summary: `v73-kani`, `v74-kani`, and `v75-kani` have completed and are now folded in `DEFERRED_GATES.md`.
  Current status (`todo` / `investigating` / `resolved`): resolved
  Reproduction command: `./scripts/run-formal-kani-async.ps1 -Action Status -Name v73-kani` (repeat for `v74-kani`, `v75-kani`)
  Suggested next action: none.
- ID: FTODO-V79-001
  Profile: v79 reconciliation (`v77..v78` DG foldback)
  Summary: Deferred-gate strict runs `v77-kani` and `v78-kani` completed with `exit=0` and are folded.
  Current status (`todo` / `investigating` / `resolved`): resolved
  Reproduction command: `./scripts/run-formal-kani-async.ps1 -Action Status -Name v77-kani` (repeat for `v78-kani`)
  Suggested next action: none.
- ID: FTODO-V84-001
  Profile: v84 reconciliation (`v80..v83` DG foldback)
  Summary: Post-crash status sweep shows `v80-kani` completed pass, while `v81-kani`..`v83-kani` are `stale` (stopped runners/watchers without completion markers).
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-async.ps1 -Action Status -Name v80-kani` (repeat for `v81-kani`..`v83-kani`)
  Suggested next action: fold pass for `v80-kani`; rerun `v81..v83` on remote Linux host (no local restart per current policy).
- ID: FTODO-V86-001
  Profile: v86 terminal reconciliation
  Summary: DG rows `DG-V79-001`, `DG-V80-001`, `DG-V81-001`, `DG-V82-001`, and `DG-V83-001` remain unresolved and are marked `dg-deferred`; `DG-V74/75/77/78` are now folded.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-async.ps1 -Action Status -Name <run-name>`
  Suggested next action: execute `DGD-V86-001` steps in `docs/evidence/formal/DG_AUDIT_V86.md` for remaining `v79..v83` lanes only.
- ID: FTODO-V86-002
  Profile: v86 terminal reconciliation
  Summary: `DG-V85-001` failed quickly (`exit_code=1`) with WSL Kani detection error; rerun lane `DG-V85-002` completed pass (`exit_code=0`) in post-crash status sweep.
  Current status (`todo` / `investigating` / `resolved`): resolved
  Reproduction command: `./scripts/run-formal-kani-async.ps1 -Action Status -Name v85-kani-rerun`
  Suggested next action: fold completion into formal summary artifacts.
- ID: FTODO-V94-001
  Profile: v87..v94 formal lanes
  Summary: New strict Kani lanes for the current language-closure tranche were intentionally not started locally after crash/recovery due host resource-risk concerns.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-async.ps1 -Action Start -Name v87-kani -ProfileScope mvp-lang-for-step-v87` (repeat for `v88`, `v89`, `v90`, `v91`, `v93`, `v94`) on remote Linux host.
  Suggested next action: migrate async Kani execution to remote Linux runner; keep local lanes in `dg-not-started` until remote orchestration is validated.
- ID: FTODO-V99-001
  Profile: v95/v96/v99 formal lanes
  Summary: Additional strict Kani lanes (`v95-kani`, `v96-kani`, `v99-kani`) intentionally not started locally per post-crash resource policy.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-async.ps1 -Action Start -Name v95-kani -ProfileScope mvp-lang-resume-full-v95` (repeat for `v96`, `v99`) on remote Linux host.
  Suggested next action: execute these lanes on remote Linux and fold results back via `DEFERRED_GATES.md` + `latest_run` reconciliation.
- ID: FTODO-V106-001
  Profile: v100..v106 formal lanes
  Summary: Language-closure tail strict lanes (`v100-kani`..`v106-kani`) intentionally not started locally per post-crash resource policy.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-async.ps1 -Action Start -Name v100-kani -ProfileScope mvp-lang-udt-fields-v100` (repeat for `v101`, `v102`, `v103`, `v104`, `v105`, `v106`) on remote Linux host.
  Suggested next action: execute remote async Kani runs, update `DEFERRED_GATES.md`, and fold results into `latest_run.*` during post-v106 reconciliation.
- ID: FTODO-V107-001
  Profile: v107 formal lane
  Summary: `v107` strict Kani lane is deferred to remote Linux execution; local formal run completed with non-blocking Kani skips.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-remote.ps1 -Action StartDeferred -DeferredVersions "107" -DeferredMode cumulative`
  Suggested next action: run remote async `v107-kani`, then fold status into `DEFERRED_GATES.md` and `latest_run.*`.
- ID: FTODO-V132-001
  Profile: v120..v134 formal lanes
  Summary: New strict Kani lanes for conversion/introspection/built-in expansion/file-stub subsets were not started locally; they are queued as remote Linux deferred gates (`DG-V120-001`, `DG-V126-001`, `DG-V132-001`, `DG-V134-001`).
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-remote.ps1 -Action StartDeferred -DeferredVersions "120,126,132,134" -DeferredMode cumulative`
  Suggested next action: launch remote async lanes, poll completion, and fold results into `latest_run.*` at the next foldback checkpoint.
- ID: FTODO-V146-001
  Profile: v146 terminal formal lane
  Summary: Terminal strict lane `DG-V146-001` remains deferred pending remote Linux capacity and batching policy.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-remote.ps1 -Action StartDeferred -DeferredVersions "146" -DeferredMode cumulative`
  Suggested next action: run terminal strict lane remotely during final v146 gate reconciliation and fold status into `DEFERRED_GATES.md`.
- ID: FTODO-V175-001
  Profile: v175 formal lane expansion I
  Summary: New strict harnesses (`cverr_tag_encoding_stays_in_reserved_error_band`, `resume_next_clears_err_number_after_raise`) were added but not yet executed in strict Kani mode.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-remote.ps1 -Action StartDeferred -DeferredVersions "175" -DeferredMode cumulative`
  Suggested next action: dispatch lane `v175-kani` on remote Linux and fold completion status into `DEFERRED_GATES.md` + `latest_run.*`.
- ID: FTODO-V176-001
  Profile: v176 formal lane expansion II
  Summary: Strict foldback lane for the v175/v176 formal expansion tranche remains deferred to remote Linux batch execution.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-remote.ps1 -Action StartDeferred -DeferredVersions "176" -DeferredMode cumulative`
  Suggested next action: execute `v176-kani` remotely and reconcile deferred register + formal summaries at the v186 terminal closure sweep.
- ID: FTODO-V287-001
  Profile: v287 PMR/Declare formal lane setup
  Summary: New strict harnesses for PMR project-graph invariants and Declare descriptor-contract checks were executed in pinned remote lane with non-empty selection (`selected_count=3`); host PMR harnesses timed out while HAL dynlink harness passed.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-remote.ps1 -Action StartDeferred -DeferredVersions "287" -DeferredMode exact`
  Suggested next action: perform host-harness timeout remediation for `FO-V287-001/002` (slicing/assumption/bounds), rerun remote `v287` lane, and only then reopen bridge-retirement decision.
- ID: FTODO-KANI-REVIEW-001
  Profile: cross-profile (`v87+` remote deferred lanes)
  Summary: Multiple deferred lanes report `selected_count=0` / `completed:no-op` with `probable-commit-obligation-mismatch`; this is a runner-input mismatch class, not a proof result.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-remote.ps1 -Action Status`
  Suggested next action: add an explicit lane-to-obligation selection preflight and fail fast when selected count is zero unexpectedly.
- ID: FTODO-KANI-REVIEW-002
  Profile: high-signal harnesses (`FO-V2-001`, `FO-V4-001`, `FO-V287-001`, `FO-V287-002`)
  Summary: Strict Kani runs are dominated by timeout/OOM from broad symbolic state (`unicode`/`memchr`/host-heavy harnesses).
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-remote.ps1 -Action Tail -Lane v89-kani -TailLines 120`
  Suggested next action: slice high-signal harnesses into smaller bounded proofs with focused assumptions and/or stubs.
- ID: FTODO-KANI-REVIEW-003
  Profile: remote runner control plane
  Summary: Memory guardrails are now implemented (`soft`/`hard` thresholds + `pause`/`halt-*` actions), but need burn-in and threshold tuning under real queue load.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-remote.ps1 -Action Monitor -MonitorDurationSeconds 600 -MonitorIntervalSeconds 30`
  Suggested next action: run 24h monitor sample, record pressure/action incidence, then tune defaults for this host.
- ID: FTODO-KANI-REVIEW-004
  Profile: remote deferred queue continuity (`v2/v4/v162/v175/v287`)
  Summary: Active deferred dispatch is currently executing on commit `560e5a0` while local head has advanced; this is acceptable for now but requires explicit post-run reconcile/restart on latest commit to avoid prolonged drift.
  Current status (`todo` / `investigating` / `resolved`): investigating
  Reproduction command: `./scripts/run-formal-kani-remote.ps1 -Action Status`
  Suggested next action: keep current run until terminal state, then run `./scripts/run-formal-kani-sync.ps1` to reconcile and restart unresolved lanes against latest head.

- ID: FTODO-UDF-SESSION-20261031
  Profile: bounded UDF persistent invocation (`bd-59co.2.10.6`, general ownership `bd-59co.2.7.6` and JIT sessions `bd-59co.2.10.2`).
  Summary: Existing Lean ExecutionBudget/RefCount abstractions pass on installed Lean 4.29.1; they do not prove Rust session-storage isolation, global alias lifetime, callback lifetime, native code-cache ownership/identity, host default-value reflection or disposal refinement. Native cargo-kani remains unavailable. Execution tests cover independent sessions, retry after failed initialization, As New bindings, private inputs, errors, budget resets and disposal balances; project classes/UDTs/End are refused by this facade.
  Current status: investigating.
  Reproduction command: cargo test -p oxvba-embed; use the installed Lean 4.29.1 binary on formal/lean/OxVba/ExecutionBudget.lean and RefCount.lean.
  Suggested next action: add session-storage/alias/refcount refinement with the shared CORE-5 ownership work; retain the class-disposal and descriptor-lifetime residuals until code and executable evidence establish them.
