# WIN-5 detailed delivery rollout

Date: 2026-09-21
Effect: support only; capability remains in progress.
Rollout owner: `bd-59co.3.6.1`. Parent: `bd-59co.3.6`.

This rollout records the detailed Windows-hosted plan and explicit delivery path. It grants no implementation or certification credit. Existing first slices and imported fixture owners are preserved.

Acceptance commands: `./scripts/validate-workset-rollout.ps1` and `./scripts/run-truth-reconciliation.ps1`.
Review and final command outcomes are retained in [planning review](../PLAN-2026-09-21/review.md).

| Child bead | Outcome | Direct prerequisites | Exact matrix rows |
|---|---|---|---|
| bd-59co.3.6.2 | WIN-5 execute synchronous ByRef event plan in VM3 | bd-59co.2.7.4, bd-59co.3.3.10, bd-59co.3.3.9, bd-59co.3.6.1, bd-59co.3.6.5, bd-aprs.8.8.9 | WIN-COM-EVENTS/WCE-PLAN-INCOMING |
| bd-59co.3.6.3 | WIN-5 execute synchronous ByRef event plan in JIT | bd-59co.2.7.7, bd-59co.3.3.10, bd-59co.3.3.9, bd-59co.3.6.1, bd-59co.3.6.5, bd-aprs.8.8.9 | WIN-COM-EVENTS/WCE-PLAN-INCOMING |
| bd-59co.3.6.4 | WIN-5 certify scalar ByRef event in Excel and lifecycle lanes | bd-59co.3.15.3, bd-59co.3.15.6, bd-59co.3.6.1, bd-59co.3.6.2, bd-59co.3.6.3 | WIN-COM-EVENTS/WCE-PLAN-INCOMING |
| bd-59co.3.6.5 | WIN-5 bind source interfaces and own Automation event subscriptions | bd-59co.2.4.6, bd-59co.3.3.10, bd-59co.3.3.9, bd-59co.3.4.2 | WIN-COM-EVENTS/WCE-INCOMING-LIFECYCLE, WIN-COM-EVENTS/WCE-PLAN-INCOMING |
| bd-59co.3.6.6 | WIN-5 preserve structural incoming event arguments and ByRef values | bd-59co.3.2.6, bd-59co.3.2.7, bd-59co.3.6.2, bd-59co.3.6.3 | WIN-COM-EVENTS/WCE-INCOMING-COMPLEX |
| bd-59co.3.6.7 | WIN-5 receive custom-vtable source events | bd-59co.3.3.10, bd-59co.3.3.9, bd-59co.3.5.4, bd-59co.3.6.5 | WIN-COM-EVENTS/WCE-INCOMING-CUSTOM |
| bd-59co.3.6.8 | WIN-5 preserve synchronous events across apartments and teardown | bd-59co.3.3.9, bd-59co.3.6.6, bd-59co.3.6.7 | WIN-COM-EVENTS/WCE-INCOMING-APARTMENT, WIN-COM-EVENTS/WCE-INCOMING-LIFECYCLE |
| bd-aprs.8.8.9 | FE-7.6.a TestEventServer ByRef COM event writeback parity | bd-59co.3.3.10, bd-59co.3.6.1 | WIN-COM-EVENTS/WCE-PLAN-INCOMING |

Each child description includes exact commands, expected observables, evidence artifacts, error/lifetime neighbors and fresh-eyes requirements. Matrix rows retain their truth states and remaining-scope owners. The canonical traceability table links every child; the detailed plan supplies the full implementation steps.

Boundaries: CORE-7 remaining ideal architecture and genuine WIN-12 outputs remain deferred. Shared Core work remains platform-neutral. Controlled development fixtures cannot certify Excel or clean deployment. The safety rollout includes a compatibility subset and preserves the full native-output safety successor.
