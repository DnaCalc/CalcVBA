# WIN-10 detailed delivery rollout

Date: 2026-09-21
Effect: support only; capability remains in progress.
Rollout owner: `bd-59co.3.11.1`. Parent: `bd-59co.3.11`.

This rollout records the detailed Windows-hosted plan and explicit delivery path. It grants no implementation or certification credit. Existing first slices and imported fixture owners are preserved.

Acceptance commands: `./scripts/validate-workset-rollout.ps1` and `./scripts/run-truth-reconciliation.ps1`.
Review and final command outcomes are retained in [planning review](../PLAN-2026-09-21/review.md).

| Child bead | Outcome | Direct prerequisites | Exact matrix rows |
|---|---|---|---|
| bd-59co.3.11.2 | WIN-10 execute retained callback lifetime in VM3 | bd-59co.2.7.4, bd-59co.3.11.1, bd-59co.3.3.12, bd-59co.3.3.9, bd-9sed.17 | WIN-NATIVE-IMPORT/WNI-PLAN-CALLBACK |
| bd-59co.3.11.3 | WIN-10 execute retained callback lifetime in JIT | bd-59co.2.7.7, bd-59co.3.11.1, bd-59co.3.3.12, bd-59co.3.3.9, bd-9sed.17 | WIN-NATIVE-IMPORT/WNI-PLAN-CALLBACK |
| bd-59co.3.11.4 | WIN-10 certify retained callback disposal exhaustion and reentry | bd-59co.3.11.1, bd-59co.3.11.2, bd-59co.3.11.3, bd-59co.3.15.3, bd-59co.3.15.6 | WIN-NATIVE-IMPORT/WNI-PLAN-CALLBACK |
| bd-59co.3.11.5 | WIN-10 expose exact VarPtr StrPtr and ObjPtr addressable storage | bd-59co.2.4.5, bd-59co.3.2.5, bd-59co.3.2.6, bd-59co.3.2.8 | WIN-NATIVE-IMPORT/WNI-POINTER-HELPERS |
| bd-59co.3.11.6 | WIN-10 invoke typed synchronous AddressOf callbacks | bd-59co.2.4.5, bd-59co.3.10.5, bd-59co.3.3.12, bd-59co.3.3.9 | WIN-NATIVE-IMPORT/WNI-CALLBACK-SYNC |
| bd-59co.3.11.7 | WIN-10 preserve nested native COM and VBA callback cycles | bd-59co.3.11.2, bd-59co.3.11.3, bd-59co.3.11.6, bd-59co.3.6.8 | WIN-NATIVE-IMPORT/WNI-CALLBACK-NESTED |
| bd-9sed.17 | Native AddressOf callback lifetime residuals beyond CallWindowProc | bd-59co.3.10.5, bd-59co.3.11.1, bd-59co.3.3.12 | WIN-NATIVE-IMPORT/WNI-PLAN-CALLBACK |

Each child description includes exact commands, expected observables, evidence artifacts, error/lifetime neighbors and fresh-eyes requirements. Matrix rows retain their truth states and remaining-scope owners. The canonical traceability table links every child; the detailed plan supplies the full implementation steps.

Boundaries: CORE-7 remaining ideal architecture and genuine WIN-12 outputs remain deferred. Shared Core work remains platform-neutral. Controlled development fixtures cannot certify Excel or clean deployment. The safety rollout includes a compatibility subset and preserves the full native-output safety successor.
