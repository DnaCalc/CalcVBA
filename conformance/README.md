# VBA test inputs

Source fixtures here feed parser tests, VM3/JIT comparisons, focused runtime
regressions and benchmarks. `crates/oxvba-differential` compares source results
and full Err state; its golden snapshot retains expected VM3 observations.
Captured Excel return values and their VBA probes live under `oracle/`.

- `tests/`: individual language and library examples.
- `integration/`: module and referenced-project source fixtures.
- `jit_v2/`: focused JIT inputs, including shared native-call ABI cases.
- `vm_package/`: package and value-model source inputs.
- `golden/`, `harness/`, `divergences/`: expected values and additional probes.

Run `./scripts/check.ps1` for the ordinary repository check. Live Excel
comparisons are opt-in; see the root README and `scripts/README.md`.
