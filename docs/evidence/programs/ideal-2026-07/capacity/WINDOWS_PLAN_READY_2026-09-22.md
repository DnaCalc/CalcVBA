# Windows planning ready snapshot

Date: 2026-09-22
Source: `br ready -l ideal-2026-07 -t task --limit 0 --json`.
Bead export SHA-256: 3c2fbb857f4bdc5d9402e4e34da8e4ce05248dbc7586265c0264795e09bfd08b

This dated snapshot filters the programme queue to the Windows compatibility plan and its named Core prerequisites. Live claims and pause instructions remain in `docs/AUTORUN_STATE.md`. Open work outside this list is not complete; it may be blocked or outside this scope.

| Bead | Ready outcome | Resources |
|---|---|---|
| bd-59co.2.11.6 | CORE-9 compare structural Windows interop observations | resource-large-differential, resource-rust-writer |
| bd-59co.2.4.4 | CORE-3 enforce VBA UDT declaration size and depth limits | resource-rust-writer |
| bd-59co.2.4.5 | CORE-3 preserve VBA7 Declare and AddressOf compile contracts | resource-rust-writer |
| bd-59co.2.6.2 | CORE-4 bound OxImage decoding for Windows packages | resource-rust-writer |
| bd-59co.2.7.2 | CORE-5 enforce SafeArray object-carrier thread ownership | resource-rust-writer |
| bd-59co.2.7.3 | CORE-5 enforce exact SafeArray dimension and allocation layout | resource-rust-writer |
| bd-59co.2.7.5 | CORE-5 generate interop helpers from one versioned ABI catalog | resource-large-jit, resource-rust-writer |
| bd-59co.3.1.8 | WIN-0 certify the x64 control plane and successor queue | resource-none |
| bd-59co.3.10.2 | WIN-9 x64 PtrSafe scalar Declare import | resource-large-differential, resource-rust-writer |
| bd-59co.3.15.12 | WIN-14 provision and pin the clean Windows x64 Excel64 VM | resource-vm-provision |
| bd-59co.3.15.6 | WIN-14 promote the owned Excel64 VBE oracle supervisor | resource-excel-vbe |
| bd-59co.3.2.2 | WIN-1 resolve registered and file-backed x64 type libraries | resource-registry, resource-rust-writer |
| bd-59co.3.4.2 | WIN-3 late-bound scalar CreateObject method property | resource-large-differential, resource-rust-writer |

Prioritize the existing scalar late-COM and x64 Declare slices. Shared image decoding, binding, helper catalog, observation, metadata and environment work may proceed when listed ready and resources permit. Never start paused CORE-7 or genuine WIN-12 from a topology recommendation. Limit Rust writers to two; serialize large Cargo, Excel/VBE, registry and provisioning activity.
