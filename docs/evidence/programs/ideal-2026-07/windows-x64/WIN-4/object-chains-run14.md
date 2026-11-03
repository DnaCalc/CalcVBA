# Object and Variant chains, run 14

Owner: bd-59co.3.5.12. Clause: COM-CLIENT-001.
Status: scoped development proof; broader WCC-LATE-ARGS remains in-progress.

Unchanged excel_chains.bas compiles using visible VBE command 578 with the
PID-scoped guardian armed. Both owning VM3 and native JIT agree with VBA:
ObjectChains=A2:12:19:38:9:A2; VariantChains=E2:14;
ChainErrors=91:424:438; ChainSnapshot=7:12:19:38:3:12:14.
This covers Object/Variant intermediate results, Offset/Cells chains, named
arguments, implicit defaults, formulas and all seven explicitly touched values.

Dynamic native calls QI IDispatch and bind names/argument IDs through GetIDsOfNames;
they invoke once, preserving native identity, result ownership and invocation locale.
Only opaque IDispatch defaults use runtime dispatch. Declared interfaces keep their
compiler-selected plans. There is no typed-call retry or engine object model.
The common boundary preserves EXCEPINFO and owns its native VARIANT cells.

Owning portable tests check repeated aliases, named arguments, omissions and error
copyback. The registration-free native fixture checks repeated native pointer
identity, chained object results/defaults, deferred EXCEPINFO, zero type-info reads
and zero final references. Fresh review changed native projection to reuse each
cell's original raw authority and added exact owned workbook identity before writes.

Commands: cargo test -p oxvba-embed --test typed_dynamic_members
--test typed_prebound_dispatch --test typed_excel_chains;
scripts/run-typed-chains-oracle.ps1; scripts/meta-check.ps1 -Fast -NoArtifacts -NoRegistry.
Installed Lean 4.29.1 RefCount/TypedContinuation/TypedReentry/ProcedureHandles
abstractions pass; concrete dynamic native pointer/refinement proof remains open.
Dynamic native callback synchronization, named default-member arguments and broader
signature/error/lifecycle certification remain with WIN-3/WIN-4 and the formal backlog.
No adapter members or Office security/registration settings changed.

Repository check: 3,780 passed in 208 suites; 184 ignored. Four registry-writing
fixtures were explicitly excluded without credit. Strict Clippy, formatting,
governance and path checks pass. run14-object-chains.json/.zip retain source,
executable, observation and archive hashes plus all four owned cleanup receipts.
