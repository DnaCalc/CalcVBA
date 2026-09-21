# WIN-3 detailed delivery rollout

Date: 2026-09-21
Effect: support only; capability remains in progress.
Rollout owner: `bd-59co.3.4.1`. Parent: `bd-59co.3.4`.

This rollout records the detailed Windows-hosted plan and explicit delivery path. It grants no implementation or certification credit. Existing first slices and imported fixture owners are preserved.

Acceptance commands: `./scripts/validate-workset-rollout.ps1` and `./scripts/run-truth-reconciliation.ps1`.
Review and final command outcomes are retained in [planning review](../PLAN-2026-09-21/review.md).

| Child bead | Outcome | Direct prerequisites | Exact matrix rows |
|---|---|---|---|
| bd-59co.3.4.2 | WIN-3 late-bound scalar CreateObject method property | bd-59co.3.3.4, bd-59co.3.4.1 | WIN-COM-CLIENT/WCC-PLAN-LATE |
| bd-59co.3.4.3 | WIN-3 realize remaining late-bound COM client rows | bd-59co.3.4.2, bd-59co.3.4.4, bd-59co.3.4.5, bd-59co.3.4.6, bd-59co.3.4.7, bd-59co.3.4.8 | WIN-COM-CLIENT/WCC-PLAN-LATE |
| bd-59co.3.4.4 | WIN-3 resolve activation GetObject and host-returned objects | bd-59co.3.2.4, bd-59co.3.2.8, bd-59co.3.3.7, bd-59co.3.4.2 | WIN-COM-CLIENT/WCC-LATE-OUTPROC-ERROR, WIN-COM-CLIENT/WCC-PLAN-LATE |
| bd-59co.3.4.5 | WIN-3 implement named omitted default and property argument rules | bd-59co.2.4.6, bd-59co.3.2.4, bd-59co.3.3.7, bd-59co.3.4.2 | WIN-COM-CLIENT/WCC-LATE-ARGS |
| bd-59co.3.4.6 | WIN-3 implement aliased ByRef dispatch writeback | bd-59co.3.2.5, bd-59co.3.2.8, bd-59co.3.4.5 | WIN-COM-CLIENT/WCC-LATE-ARGS |
| bd-59co.3.4.7 | WIN-3 support structural values chaining and enumeration | bd-59co.2.11.6, bd-59co.3.2.6, bd-59co.3.2.7, bd-59co.3.2.8, bd-59co.3.4.6 | WIN-COM-CLIENT/WCC-LATE-STRUCTURAL |
| bd-59co.3.4.8 | WIN-3 preserve COM errors LCID and dispatch lifetime | bd-59co.2.11.6, bd-59co.3.3.7, bd-59co.3.4.4, bd-59co.3.4.6 | WIN-COM-CLIENT/WCC-LATE-OUTPROC-ERROR |

Each child description includes exact commands, expected observables, evidence artifacts, error/lifetime neighbors and fresh-eyes requirements. Matrix rows retain their truth states and remaining-scope owners. The canonical traceability table links every child; the detailed plan supplies the full implementation steps.

Boundaries: CORE-7 remaining ideal architecture and genuine WIN-12 outputs remain deferred. Shared Core work remains platform-neutral. Controlled development fixtures cannot certify Excel or clean deployment. The safety rollout includes a compatibility subset and preserves the full native-output safety successor.
