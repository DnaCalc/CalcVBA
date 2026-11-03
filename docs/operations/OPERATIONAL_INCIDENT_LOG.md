# Operational Incident Log

Status: `active`

Purpose:
- record process failures that affected planning, truth, or execution trust,
- capture root cause,
- record the doctrine/tooling change that prevents recurrence.

## Incident OI-001

- Date: 2026-03-29
- Title: `For Each` subset support was allowed to read as full closure
- Impact:
  - array `For Each` support was implemented and evidenced,
  - object-enumerator `For Each` remained unimplemented,
  - some active truth surfaces widened the claim and later work consumed that wording as if full support existed.
- Root cause:
  - subset-support truth was not enforced consistently across all active artifacts,
  - canonical validation matrices did not yet own the truth,
  - summaries could drift broader than evidence.
- Preventive changes:
  - canonical validation matrices introduced,
  - truth ownership map introduced,
  - conformance topics mapped to matrix owners,
  - closure-language doctrine tightened,
  - `For Each` arrays vs object-enumerators made the standing canary.

## Incident OI-002

- Date: 2026-03-29
- Title: Parallel bead graph mutations produced misleading `br create` results
- Impact:
  - multiple `br create` calls in parallel returned confusing ID stdout,
  - the intended epic rollout had to be audited against persisted bead state before further graph work could proceed safely.
- Root cause:
  - bead DB mutations were treated like read-only parallelizable operations,
  - no serialized wrapper or explicit single-writer doctrine existed.
- Preventive changes:
  - bead mutation serialization rule added to doctrine,
  - `scripts/invoke-br-serialized.ps1` introduced,
  - rollout work is now expected to use serialized graph mutations.


## Incident OI-003

- Date: 2026-11-03
- Title: Owned oracle cleanup wait expired before process reconciliation
- Impact: calcvba13-collections-vba failed and receives no passing evidence. Its guardian captured and dismissed an owned runtime error 91; the cleanup wait then expired and the supervisor could not confirm the already-stopping guardian identity.
- Root cause: the exact process-query/wait race is not established. The retained termination wait timed out during a concurrent workspace build. Subsequent read-only checks found both recorded PIDs (Excel 30896 and guardian 39928) absent; cleanup-followup.json retains their original start identities and the later absence observation. No unrelated process was stopped.
- Preventive action: keep the failed capture, require the follow-up ownership receipt, and serialize further live engine oracles after workspace link activity. The fresh calcvba13-collections-vba-reset run compiles, executes and cleans up successfully. Cleanup reliability follow-up remains with the open WIN-14 orchestration work bd-59co.3.15.7; this recovery grants no certification credit.
