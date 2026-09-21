# WIN-9 detailed delivery rollout

Date: 2026-09-21
Effect: support only; capability remains in progress.
Rollout owner: `bd-59co.3.10.1`. Parent: `bd-59co.3.10`.

This rollout records the detailed Windows-hosted plan and explicit delivery path. It grants no implementation or certification credit. Existing first slices and imported fixture owners are preserved.

Acceptance commands: `./scripts/validate-workset-rollout.ps1` and `./scripts/run-truth-reconciliation.ps1`.
Review and final command outcomes are retained in [planning review](../PLAN-2026-09-21/review.md).

| Child bead | Outcome | Direct prerequisites | Exact matrix rows |
|---|---|---|---|
| bd-59co.3.10.2 | WIN-9 x64 PtrSafe scalar Declare import | bd-59co.3.10.1, bd-59co.3.3.4 | WIN-NATIVE-IMPORT/WNI-PLAN-DECLARE |
| bd-59co.3.10.3 | WIN-9 realize remaining VBA7 Declare import rows | bd-59co.3.10.2, bd-59co.3.10.4, bd-59co.3.10.5, bd-59co.3.10.6, bd-59co.3.10.7 | WIN-NATIVE-IMPORT/WNI-PLAN-DECLARE |
| bd-59co.3.10.4 | WIN-9 resolve DLL entries ordinals and policy failures exactly | bd-59co.2.4.5, bd-59co.3.10.2 | WIN-NATIVE-IMPORT/WNI-DECLARE-LOADER-ERROR |
| bd-59co.3.10.5 | WIN-9 preserve scalar ABI returns and aliased ByRef arguments | bd-59co.2.11.6, bd-59co.2.4.5, bd-59co.3.10.2, bd-59co.3.10.4, bd-59co.3.2.5, bd-59co.3.3.7 | WIN-NATIVE-IMPORT/WNI-PLAN-DECLARE |
| bd-59co.3.10.6 | WIN-9 marshal ANSI Wide BSTR and mutable native buffers | bd-59co.3.10.4, bd-59co.3.10.5 | WIN-NATIVE-IMPORT/WNI-DECLARE-STRINGS |
| bd-59co.3.10.7 | WIN-9 marshal arrays UDTs and As Any using exact x64 storage | bd-59co.2.4.5, bd-59co.3.10.5, bd-59co.3.2.6, bd-59co.3.2.7 | WIN-NATIVE-IMPORT/WNI-DECLARE-STRUCTURAL |

Each child description includes exact commands, expected observables, evidence artifacts, error/lifetime neighbors and fresh-eyes requirements. Matrix rows retain their truth states and remaining-scope owners. The canonical traceability table links every child; the detailed plan supplies the full implementation steps.

Boundaries: CORE-7 remaining ideal architecture and genuine WIN-12 outputs remain deferred. Shared Core work remains platform-neutral. Controlled development fixtures cannot certify Excel or clean deployment. The safety rollout includes a compatibility subset and preserves the full native-output safety successor.
