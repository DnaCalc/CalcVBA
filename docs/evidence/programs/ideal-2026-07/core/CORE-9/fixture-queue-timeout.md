# Differential fixture queue and execution timeout

2026-10-31. Bounded delivery: `bd-59co.2.11.7`; wider differential coverage remains
`bd-59co.2.11.3` and the CORE-DIFF-SEMANTICS-FUZZ row remains planned.

The concurrent workspace run produced a semantic golden TIMEOUT for a fast
binding diagnostic because `run_with_timeout` counted time waiting for the
shared fixture mutex as execution. The same full golden passed alone.

The worker now acquires the existing mutex, signals that execution can start,
then runs the same compiler/runtime path under that lock. Only then does the
caller start the source execution deadline. A separate bounded queue watchdog
reports a stalled fixture queue explicitly; disconnected/panicked workers are
also failures, rather than semantic TIMEOUT values. The execution deadline and
all golden expectations are unchanged. Timed-out execution still has the existing
detached-worker limitation; process-level containment remains wider work.

The regression holds the fixture lock beyond the complete execution allowance,
checks that no result/timeout was published while queued, releases it, and
requires the fast source to finish. Command:

```powershell
cargo test -p oxvba-differential --lib execution_timeout_excludes_fixture_lock_queueing
```

Final acceptance passed: the repository meta-check ran `cargo test --workspace`
with default concurrency: 3,640 passed, 178 ignored. Format, strict workspace
Clippy and governance also passed, without a golden bless or a
RUST_TEST_THREADS override. This repair changes harness scheduling, not VBA
semantics, and earns no broader conformance or certification completion.

Implementation checkpoint: bcd6a25399f97e5e637e0554e92e1ec5868f1455
