# WIN-4 detailed delivery rollout

Date: 2026-09-21
Effect: support only; capability remains in progress.
Rollout owner: `bd-59co.3.5.1`. Parent: `bd-59co.3.5`.

This rollout records the detailed Windows-hosted plan and explicit delivery path. It grants no implementation or certification credit. Existing first slices and imported fixture owners are preserved.

Acceptance commands: `./scripts/validate-workset-rollout.ps1` and `./scripts/run-truth-reconciliation.ps1`.
Review and final command outcomes are retained in [planning review](../PLAN-2026-09-21/review.md).

| Child bead | Outcome | Direct prerequisites | Exact matrix rows |
|---|---|---|---|
| bd-59co.3.5.2 | WIN-4 execute scalar and property calls through native vtables | bd-59co.2.11.6, bd-59co.2.4.6, bd-59co.3.15.9, bd-59co.3.2.8, bd-59co.3.3.7, bd-59co.3.3.8 | WIN-COM-CLIENT/WCC-PLAN-EARLY |
| bd-59co.3.5.3 | WIN-4 marshal complex and ByRef early-bound signatures | bd-59co.3.2.5, bd-59co.3.2.6, bd-59co.3.2.7, bd-59co.3.5.2 | WIN-COM-CLIENT/WCC-EARLY-COMPLEX |
| bd-59co.3.5.4 | WIN-4 execute custom inherited interfaces without dispatch fallback | bd-59co.3.5.3 | WIN-COM-CLIENT/WCC-EARLY-CUSTOM |
| bd-59co.3.5.5 | WIN-4 preserve early transport through apartment and process proxies | bd-59co.3.3.9, bd-59co.3.5.4 | WIN-COM-CLIENT/WCC-EARLY-OUTPROC |

Each child description includes exact commands, expected observables, evidence artifacts, error/lifetime neighbors and fresh-eyes requirements. Matrix rows retain their truth states and remaining-scope owners. The canonical traceability table links every child; the detailed plan supplies the full implementation steps.

Boundaries: CORE-7 remaining ideal architecture and genuine WIN-12 outputs remain deferred. Shared Core work remains platform-neutral. Controlled development fixtures cannot certify Excel or clean deployment. The safety rollout includes a compatibility subset and preserves the full native-output safety successor.
