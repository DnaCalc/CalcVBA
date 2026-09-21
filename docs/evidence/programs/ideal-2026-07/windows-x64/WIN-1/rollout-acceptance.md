# WIN-1 detailed delivery rollout

Date: 2026-09-21
Effect: support only; capability remains in progress.
Rollout owner: `bd-59co.3.2.1`. Parent: `bd-59co.3.2`.

This rollout records the detailed Windows-hosted plan and explicit delivery path. It grants no implementation or certification credit. Existing first slices and imported fixture owners are preserved.

Acceptance commands: `./scripts/validate-workset-rollout.ps1` and `./scripts/run-truth-reconciliation.ps1`.
Review and final command outcomes are retained in [planning review](../PLAN-2026-09-21/review.md).

| Child bead | Outcome | Direct prerequisites | Exact matrix rows |
|---|---|---|---|
| bd-59co.3.2.2 | WIN-1 resolve registered and file-backed x64 type libraries | bd-59co.2.2.10, bd-59co.3.1.5 | WIN-ABI-CARRIER/WAC-TYPELIB-METADATA |
| bd-59co.3.2.3 | WIN-1 preserve authoritative member event and type identities | bd-59co.3.2.2 | WIN-ABI-CARRIER/WAC-TYPELIB-METADATA |
| bd-59co.3.2.4 | WIN-1 hand authoritative descriptors to compiler packages and hosts | bd-59co.2.4.6, bd-59co.2.6.4, bd-59co.3.2.3 | WIN-ABI-CARRIER/WAC-TYPELIB-METADATA |
| bd-59co.3.2.5 | WIN-1 prove x64 BSTR VARIANT and numeric carrier layout | bd-59co.2.2.14, bd-59co.3.15.9 | WIN-ABI-CARRIER/WAC-BSTR-LAYOUT, WIN-ABI-CARRIER/WAC-NUMERIC-LONGPTR, WIN-ABI-CARRIER/WAC-VARIANT-LAYOUT |
| bd-59co.3.2.6 | WIN-1 preserve SAFEARRAY shape and nominal interface ownership | bd-59co.2.7.2, bd-59co.2.7.3, bd-59co.3.15.9, bd-59co.3.2.3 | WIN-ABI-CARRIER/WAC-INTERFACE-ARRAY, WIN-ABI-CARRIER/WAC-SAFEARRAY-LAYOUT |
| bd-59co.3.2.7 | WIN-1 transport nominal VT_RECORD scalars and arrays | bd-59co.2.4.4, bd-59co.3.15.9, bd-59co.3.2.3 | WIN-ABI-CARRIER/WAC-VT-RECORD |
| bd-59co.3.2.8 | WIN-1 preserve IUnknown identity and carrier lifetime across interfaces | bd-59co.2.7.2, bd-59co.3.15.9, bd-59co.3.2.3 | WIN-ABI-CARRIER/WAC-IUNKNOWN-IDENTITY |

Each child description includes exact commands, expected observables, evidence artifacts, error/lifetime neighbors and fresh-eyes requirements. Matrix rows retain their truth states and remaining-scope owners. The canonical traceability table links every child; the detailed plan supplies the full implementation steps.

Boundaries: CORE-7 remaining ideal architecture and genuine WIN-12 outputs remain deferred. Shared Core work remains platform-neutral. Controlled development fixtures cannot certify Excel or clean deployment. The safety rollout includes a compatibility subset and preserves the full native-output safety successor.
