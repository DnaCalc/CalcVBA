# WIN-11 detailed delivery rollout

Date: 2026-09-21
Effect: support only; capability remains in progress.
Rollout owner: `bd-59co.3.12.1`. Parent: `bd-59co.3.12`.

This rollout records the detailed Windows-hosted plan and explicit delivery path. It grants no implementation or certification credit. Existing first slices and imported fixture owners are preserved.

Acceptance commands: `./scripts/validate-workset-rollout.ps1` and `./scripts/run-truth-reconciliation.ps1`.
Review and final command outcomes are retained in [planning review](../PLAN-2026-09-21/review.md).

| Child bead | Outcome | Direct prerequisites | Exact matrix rows |
|---|---|---|---|
| bd-59co.3.12.2 | WIN-11 run verified images in JIT WrapperExe artifacts | bd-59co.2.10.3, bd-59co.2.6.5, bd-59co.3.3.7 | WIN-NATIVE-EXPORT/WNE-WRAPPER-EXE |
| bd-59co.3.12.3 | WIN-11 initialize and unload JIT WrapperLibrary outside loader lock | bd-59co.2.7.6, bd-59co.3.12.2 | WIN-NATIVE-EXPORT/WNE-WRAPPER-LIBRARY |
| bd-59co.3.12.4 | WIN-11 select JIT sessions in reusable WrappedComServer artifacts | bd-59co.3.12.3, bd-59co.3.7.6, bd-59co.3.8.3, bd-59co.3.9.3 | WIN-NATIVE-EXPORT/WNE-PLAN-WRAPPED |
| bd-59co.3.12.5 | WIN-11 verify clean wrapper deployment and version rejection | bd-59co.3.12.4, bd-59co.3.15.11 | WIN-NATIVE-EXPORT/WNE-PLAN-WRAPPED, WIN-NATIVE-EXPORT/WNE-WRAPPER-EXE, WIN-NATIVE-EXPORT/WNE-WRAPPER-LIBRARY |

Each child description includes exact commands, expected observables, evidence artifacts, error/lifetime neighbors and fresh-eyes requirements. Matrix rows retain their truth states and remaining-scope owners. The canonical traceability table links every child; the detailed plan supplies the full implementation steps.

Boundaries: CORE-7 remaining ideal architecture and genuine WIN-12 outputs remain deferred. Shared Core work remains platform-neutral. Controlled development fixtures cannot certify Excel or clean deployment. The safety rollout includes a compatibility subset and preserves the full native-output safety successor.
