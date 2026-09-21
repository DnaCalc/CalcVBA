# WIN-6 detailed delivery rollout

Date: 2026-09-21
Effect: support only; capability remains in progress.
Rollout owner: `bd-59co.3.7.1`. Parent: `bd-59co.3.7`.

This rollout records the detailed Windows-hosted plan and explicit delivery path. It grants no implementation or certification credit. Existing first slices and imported fixture owners are preserved.

Acceptance commands: `./scripts/validate-workset-rollout.ps1` and `./scripts/run-truth-reconciliation.ps1`.
Review and final command outcomes are retained in [planning review](../PLAN-2026-09-21/review.md).

| Child bead | Outcome | Direct prerequisites | Exact matrix rows |
|---|---|---|---|
| bd-59co.3.7.2 | WIN-6 activate scalar COM classes through VM3 and JIT sessions | bd-59co.2.4.7, bd-59co.3.15.9, bd-59co.3.3.11, bd-59co.3.3.7 | WIN-COM-SERVER/WCS-LATE-INPROC |
| bd-59co.3.7.3 | WIN-6 expose default named optional and enumeration dispatch contracts | bd-59co.3.4.5, bd-59co.3.7.2 | WIN-COM-SERVER/WCS-LATE-COMPLEX, WIN-COM-SERVER/WCS-LATE-INPROC |
| bd-59co.3.7.4 | WIN-6 serve structural values ByRef and full COM error state | bd-59co.3.2.6, bd-59co.3.2.7, bd-59co.3.2.8, bd-59co.3.3.9, bd-59co.3.7.3 | WIN-COM-SERVER/WCS-LATE-COMPLEX |
| bd-59co.3.7.5 | WIN-6 run COM LocalServer activation marshaling and shutdown | bd-59co.3.3.9, bd-59co.3.7.2 | WIN-COM-SERVER/WCS-LATE-LOCALSERVER |
| bd-59co.3.7.6 | WIN-6 prove class factory unload and registration lifetime | bd-59co.3.1.5, bd-59co.3.7.4, bd-59co.3.7.5 | WIN-COM-SERVER/WCS-SERVER-SAFETY |

Each child description includes exact commands, expected observables, evidence artifacts, error/lifetime neighbors and fresh-eyes requirements. Matrix rows retain their truth states and remaining-scope owners. The canonical traceability table links every child; the detailed plan supplies the full implementation steps.

Boundaries: CORE-7 remaining ideal architecture and genuine WIN-12 outputs remain deferred. Shared Core work remains platform-neutral. Controlled development fixtures cannot certify Excel or clean deployment. The safety rollout includes a compatibility subset and preserves the full native-output safety successor.
