# WIN-13 detailed delivery rollout

Date: 2026-09-21
Effect: support only; capability remains in progress.
Rollout owner: `bd-59co.3.14.1`. Parent: `bd-59co.3.14`.

This rollout records the detailed Windows-hosted plan and explicit delivery path. It grants no implementation or certification credit. Existing first slices and imported fixture owners are preserved.

Acceptance commands: `./scripts/validate-workset-rollout.ps1` and `./scripts/run-truth-reconciliation.ps1`.
Review and final command outcomes are retained in [planning review](../PLAN-2026-09-21/review.md).

| Child bead | Outcome | Direct prerequisites | Exact matrix rows |
|---|---|---|---|
| bd-59co.3.14.2 | WIN-13 run final native boundary safety lifecycle certification | bd-59co.3.10, bd-59co.3.11, bd-59co.3.12, bd-59co.3.13, bd-59co.3.14.1, bd-59co.3.4, bd-59co.3.5, bd-59co.3.6, bd-59co.3.7, bd-59co.3.8, bd-59co.3.9 | WIN-ABI-CARRIER/WAC-SAFETY-MUTATION |
| bd-59co.3.14.3 | WIN-13 reject mutated Windows boundary descriptors before side effects | bd-59co.2.6.3, bd-59co.3.3.10, bd-59co.3.3.11, bd-59co.3.3.12, bd-59co.3.3.8 | WIN-ABI-CARRIER/WAC-SAFETY-MUTATION |
| bd-59co.3.14.4 | WIN-13 inject COM native and allocation failures into owned lifetimes | bd-59co.3.11.6, bd-59co.3.4.8, bd-59co.3.5.3, bd-59co.3.7.4 | WIN-ABI-CARRIER/WAC-SAFETY-MUTATION |
| bd-59co.3.14.5 | WIN-13 enforce capability policy and isolate all native cleanup | bd-59co.3.1.5, bd-59co.3.10.4, bd-59co.3.3.9, bd-59co.3.7.5 | WIN-ABI-CARRIER/WAC-SAFETY-MUTATION |
| bd-59co.3.14.6 | WIN-13 prove safety across the Windows compatibility subset | bd-59co.3.11.7, bd-59co.3.12.5, bd-59co.3.14.3, bd-59co.3.14.4, bd-59co.3.14.5, bd-59co.3.6.8 | WIN-ABI-CARRIER/WAC-SAFETY-MUTATION |

Each child description includes exact commands, expected observables, evidence artifacts, error/lifetime neighbors and fresh-eyes requirements. Matrix rows retain their truth states and remaining-scope owners. The canonical traceability table links every child; the detailed plan supplies the full implementation steps.

Boundaries: CORE-7 remaining ideal architecture and genuine WIN-12 outputs remain deferred. Shared Core work remains platform-neutral. Controlled development fixtures cannot certify Excel or clean deployment. The safety rollout includes a compatibility subset and preserves the full native-output safety successor.
