# Owning prebound dispinterface calls, run 12

Owner: bd-59co.3.5.9; COM-CLIENT-001, SYS-DUAL-001, HOST-SESSION-001.
Status: scoped delivery bd-59co.3.5.9 passes; source-bound run12-prebound-dispatch.json/.zip retain the receipt.
Broader WCC-EARLY-COMPLEX/WIN-4 and pinned certification remain in-progress.

The 2026-11-02 owner ruling resolves BLK-WORKSHEETFUNCTION-VTABLE-20261101:
pure dispinterfaces use their typelib-selected IID/DISPID/accessor in one Invoke.
No runtime names or type information are queried, and dual interfaces never retry
through this route. The engine contains no Excel-specific member or substitute.

`typed_slots::plan` admits this distinct native transport. Complete compiler member
order assigns internal typed table indices without inventing a native vtable slot.
Preparation retains the QI'd exact interface and canonical IUnknown identity.
Native cells own values through the call; property puts use DISPID_PROPERTYPUT,
source arguments reverse at the boundary and omitted tails stay omitted. Declared
Variant parameters box actual carriers while preserving object arguments. Both
VM3 and native JIT execute the same immutable plan on their owning physical stack.

`typed_prebound_dispatch` uses a registration-free seven-slot fixture whose
GetIDsOfNames and type-info methods refuse and count every attempt. Both backends
check named argument ordering, Long/Double/BSTR/Variant inputs, omitted 30-parameter
tails with different supplied arities, repeated ByRef aliases (1313), default
arithmetic, native interface identity/results, full deferred EXCEPINFO fields and
zero native references after drop. Malformed native shapes and incomplete dispatch
table preparation refuse before member entry. This fixture is transport evidence,
not evidence for the separately open callback/global alias bug.

The owned Excel supervisor compiles the exact shared `prebound_excel.bas` plus
`typed-prebound-run12.bas` using VBE command 578 with the PID-scoped guardian armed.
The native ignored test attaches only through the recorded PID/HWND. It executes
the same retained API functions against that workbook on both backends. Numeric,
Boolean, text, Empty and CVErr scalars, two-dimensional mixed Variant arrays,
single-cell array extraction, Range implicit defaults, Value/Value2, date/currency
tags, Sum/Average/Match and error 1004 are the scoped observations.
The core array packet uses row-major storage matching VBA's logical cells;
independent A2/B1 reads after each property write establish actual worksheet
orientation, beyond a self-consistent writer/reader array roundtrip. Exact owned
workbook identity is checked before writes. Final owned run:
`calcvba12-prebound-final-cells`, both backends pass.

Commands: `cargo test -p oxvba-embed --test typed_prebound_dispatch --test typed_native_com`;
`scripts/run-typed-prebound-oracle.ps1 -RuntimeTestExecutable <fresh Cargo test artifact>`;
`scripts/meta-check.ps1 -Fast -NoArtifacts -NoRegistry`.
Four registry-writing COM record fixtures are excluded explicitly, without credit.
Final repository check: 3764 passed in 202 suites, 181 ignored. Formatting, strict
workspace Clippy, path and governance checks pass. The final ignored live test
has a fresh focused Clippy/build receipt. Interpreter-only native/dispatch tests
pass (10 passed, 1 ignored); interpreter-only wasm32 compilation passes. Fresh
review corrected the platform-neutral IID predicate and strengthened the array
input/orientation and workbook checks.

Generic dynamic NumberFormat setup remains separately open under bd-59co.3.5.10:
the displayed `$0.00` format did not produce VBA's Currency classification after
that setup, while VBA-prepared D1 reads match exactly. Changing the read LCID did
not resolve that observation and was discarded. The failed setup attempts earn
no passing credit; they remain archived for the explicit next delivery.

Installed Lean RefCount, TypedContinuation, TypedReentry and ProcedureHandles
abstractions pass. Concrete Automation-cell/proxy/native provenance is unproved;
the existing offline Miri dependency/sysroot limitation receives no passing credit.
No canonical row closes. Native callback/global visibility remains bd-59co.2.9.20;
broader signatures/events, nominal interface arrays and certification retain their
existing owners. No adapter member, security/Trust Center, registration or HKLM
change is part of this delivery.
