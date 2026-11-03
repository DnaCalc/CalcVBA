# Legacy spreadsheet adapter

This package contains the earlier spreadsheet-facing API removed from
`calcvba-host`. It retains the synthetic Excel typelib, Range/SpreadsheetContext,
grid limits, callbacks and formula conversions solely for compatibility and tests.
Compiler metadata, executable memory, project state and code caching come from
the neutral `calcvba-host` owner.

New engine hosts should use `calcvba-host`. DNA Calc's separately authorized glue
migration will supply its real object library and own these spreadsheet policies.
No DNA Calc source has been changed here. Existing consumers need a deliberate
dependency/API migration; the engine crate no longer exports the old Value API.

The retained contract is [UDF_EMBEDDING_V1.md](../../docs/spec/UDF_EMBEDDING_V1.md).
The current engine boundary is
[HOSTING_AND_EARLY_BINDING.md](../../docs/spec/HOSTING_AND_EARLY_BINDING.md).
