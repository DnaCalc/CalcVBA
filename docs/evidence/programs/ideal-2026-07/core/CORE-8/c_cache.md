# Bounded embedding code cache

2026-10-31; `bd-59co.2.10.3` remains open. The caller-owned same-thread LRU compares exact verified program/signatures/defaults, backend, host policy/descriptor, target/CPU/codegen settings and helper/carrier ABI. Mutable state and host object identities never enter shared code. Eviction and clear preserve live owners; allocation is released after the last owner, including failed-compile cleanup. [Run 3 evidence](udf-session-embedding.md) records executable checks.

Limits count entries and retained description bytes, not executable arena allocation. No disk persistence, cross-process/cross-thread sharing, general verified package, generation invalidation or broad host reentry claim is made. Those requirements remain with the open owner.
