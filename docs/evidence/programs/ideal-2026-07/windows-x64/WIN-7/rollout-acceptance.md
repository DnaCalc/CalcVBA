# WIN-7 detailed delivery rollout

Date: 2026-09-21
Effect: support only; capability remains in progress.
Rollout owner: `bd-59co.3.8.1`. Parent: `bd-59co.3.8`.

This rollout records the detailed Windows-hosted plan and explicit delivery path. It grants no implementation or certification credit. Existing first slices and imported fixture owners are preserved.

Acceptance commands: `./scripts/validate-workset-rollout.ps1` and `./scripts/run-truth-reconciliation.ps1`.
Review and final command outcomes are retained in [planning review](../PLAN-2026-09-21/review.md).

| Child bead | Outcome | Direct prerequisites | Exact matrix rows |
|---|---|---|---|
| bd-59co.3.8.2 | WIN-7 generate dual vtables sharing the served object identity | bd-59co.2.7.5, bd-59co.3.3.11, bd-59co.3.5.2, bd-59co.3.7.2 | WIN-COM-SERVER/WCS-DUAL-INPROC |
| bd-59co.3.8.3 | WIN-7 serve imported Implements and complex custom interfaces | bd-59co.2.4.7, bd-59co.3.5.4, bd-59co.3.7.4, bd-59co.3.8.2 | WIN-COM-SERVER/WCS-IMPLEMENTS-CUSTOM |
| bd-59co.3.8.4 | WIN-7 package and prove early serving across process boundaries | bd-59co.3.5.5, bd-59co.3.7.5, bd-59co.3.8.3 | WIN-COM-SERVER/WCS-EARLY-OUTPROC |

Each child description includes exact commands, expected observables, evidence artifacts, error/lifetime neighbors and fresh-eyes requirements. Matrix rows retain their truth states and remaining-scope owners. The canonical traceability table links every child; the detailed plan supplies the full implementation steps.

Boundaries: CORE-7 remaining ideal architecture and genuine WIN-12 outputs remain deferred. Shared Core work remains platform-neutral. Controlled development fixtures cannot certify Excel or clean deployment. The safety rollout includes a compatibility subset and preserves the full native-output safety successor.
