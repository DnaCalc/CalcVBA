# Windows-hosted plan review

Plan date: 2026-09-21. Final review: 2026-09-22 (Africa/Johannesburg).
Owner: bd-59co.3.1.9
Effect: support only; no capability or certification completion claimed.

## Scope and evidence inspected

The user chose Windows compatibility plus required Core work. The plan reuses existing certification, scalar-adapter, VM3 reentry, event and retained-callback owners. Grounding included the current bead export, six Windows matrices, accepted Windows workset/specification, Core contracts, interop_plan.rs, engine.rs and jit_windows_vm3_parity.rs.

## Review rounds and corrections

1. Scope and code review: kept the existing late/Declare adapter slice; confirmed missing general session/reentry/plan families and the nondeterministic GetTickCount comparison. Preserved the explicit CORE-7/WIN-12 pause and Linux/IDE boundaries.
2. Contract and self-containment review: every new leaf names code anchors, matrix rows, exact fixture/signature/process/apartment facts, positive/error/cleanup observations, commands, evidence, prerequisites and residual behavior. Shared Core implementation remains platform-neutral. Relative size labels replace new time estimates.
3. Dependency and ownership review: replaced coarse start dependencies with specific producer leaves, added missing reentry dependencies to existing event/callback execution, kept broad residuals open as behavioral integration joins, and separated certification from producer-epic completion to avoid circular completion requirements.
4. Independent alternate-model review: a fresh-context reviewer inspected the draft and graph. Its findings led to sealed-descriptor plan admission, session prerequisites on general consumers, separation of synchronous callbacks from retained fixture readiness, exact native-loader ordering, concrete matrix evidence owners and an explicit compatibility-only report. The report leaves native/full-programme aggregate rows open.
5. Live graph and validator review: removed two imported-fixture dependencies on whole producer epics that created circular completion requirements. Updated 43 exact evidence-owner routes while retaining strict row, target, authority and residual-owner checks. Carrier Excel roundtrip evidence belongs to its exact WIN-14 certification leaf; two mutation cases reject a different owner or a wrong parent.
6. Final fresh review: the independent reviewer found no remaining planning blocker. The compatibility report references WAC-RELEASE-CERT only for read-only reconciliation, neither owns nor advances it, and is excluded from the new-producer coverage column. No capability, full release or native-output claim follows from this plan.

## Validation

- `scripts/check-governance.ps1`: passed, including line-ending checks, documentation consistency, ownership/traceability, environment and certification manifests, and negative validator cases. The governance transcript records 14 Windows route rejection cases, 26 programme validator cases and the relative-size tests.
- `scripts/run-truth-reconciliation.ps1 -Refresh`: passed again after support-bead completion and final handoff updates. All 57 Windows rows retain their prior state: 56 planned and one verified noncertifying development environment. Certification remains blocked across all 57 cases and 342 observation axes.
- `scripts/test-ideal-bead-size.ps1`: four positive and five negative cases passed. New beads use relative size; existing historical estimates remain readable.
- `br dep cycles --json`: zero dependency cycles. The exact live graph remains the scheduling authority; topology reports do not authorize a claim or override paused work.
- Source comparison against `58a3164c5b082857133a38115fbc6167011ce651`: all six Windows matrices changed only their `evidence_owner_bead` fields. Every new code reference exists. The [dependency comparison](dependency-changes.json) records 42 existing dependency sets changed by this plan.
- `br lint --json`: exit success with 15 pre-existing heading warnings on 11 unchanged `bd-ivaha` issues. No warning belongs to a new or updated planning bead. The [lint report](bead-lint.json) retains the exact existing issues; these unrelated records were not rewritten.
- Three existing Rust files were checked out with CRLF despite the repository LF contract. Normalizing their line endings restored bytes identical to HEAD; this plan changes no Rust source behavior. No runtime compatibility or formal verification result is claimed from planning checks.

## Result and handoff

Created 66 detailed open beads: 13 shared Core and 53 Windows, comprising 65 delivery outcomes and one non-owning compatibility report. A separate support bead, `bd-59co.3.1.9`, records this planning work. Each detailed bead includes prerequisites, code/truth references, implementation steps, acceptance commands, observable outcomes and an evidence artifact.

Completed only the planning support bead and 11 rollout support beads for WIN-1, WIN-3 through WIN-11, and WIN-13. Existing scalar, imported event/callback, execution, certification and broader remaining-scope owners are reused. All new delivery/report beads remain open; no capability epic or matrix row is completed by this result.

The first Windows delivery items are `bd-59co.3.4.2` (scalar late COM) and `bd-59co.3.10.2` (x64 scalar Declare). The dated [ready snapshot](../../capacity/WINDOWS_PLAN_READY_2026-09-22.md) lists 13 scoped ready leaves, including independent required Core and environment work. `docs/AUTORUN_STATE.md` remains the sole current execution control surface. The user's current request ends with planning; CORE-7 remaining architecture and genuine WIN-12 outputs stay paused.

Validation transcripts: [governance](governance.txt) and [final truth reconciliation](truth-reconciliation.txt).
