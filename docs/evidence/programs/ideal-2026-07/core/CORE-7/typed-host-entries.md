# Typed host entries and neutral hosting — run 6

Run-7 [host API/callback evidence](typed-host-callbacks.md) extends this historical
numeric-only scope with borrowed host-object slots and bounded typed reentry.
Its current contract uses typed-entry ABI 2. The measurements below remain run-6
results and are not performance claims for the added callback path.

2026-10-31. Windows x64 development evidence. Bounded deliveries: CORE-7 `bd-59co.2.9.11`, CORE-8 `bd-59co.2.10.8`. Broader CORE-7/8 capabilities remain
**in-progress**; residual beads `bd-59co.2.9.9` and `bd-59co.2.10.7` stay open. This is a bounded scalar family,
not full typed calling, general package ownership or Windows certification.

`oxvba-rt-abi::host_entry::ScalarLayout` admits the original OxFunc, records
physical types/slots and proven Double-only Variant temporaries, and rejects
unsupported shapes before binding. JIT entries use native i32/f64 arguments and
a typed output pointer. VM3 reads the same OxFunc using unboxed stack storage.
Successful warm typed calls create no Variant frame or heap allocation. Both universal value
routes adapt admitted scalar procedures to the unboxed primary entries; image run-entry bodies keep
their diagnostic local snapshots. Other signatures retain their explicit dynamic
API and are refused by typed binding.

Scope: zero to four required ByVal Long/Double parameters, Long/Double results,
128 local/temporary slots, local arithmetic/comparisons/truthiness/branches and
unhandled scalar faults, including proved Long-to-Double arithmetic/comparison
promotion. Source-module sessions have no implicit program run entry. No ByRef, Optional/ParamArray, object/string/array/record,
global/Static, call, On Error, line/Erl or recursion primary entry is claimed.
The Rust-facing convention is version 1; the general native hosting interface is
still an owner review decision. Private JIT helper identity is 4.

Neutral `oxvba-embed` publishes compiler-owned exact procedure metadata, including
Sub/accessors and Option Private Module, and checks session identity/generation
before bound invocation. Failed reload preserves handles; successful reload
invalidates old generations before releasing code; counter exhaustion refuses
reload. Cache eviction preserves live sessions. Exact Variant values keep Null,
CVErr and UTF-16 units. Native supplied-object tests prove preserved IID/slot
transport on VM3/JIT and zero Invoke/name/type-info recovery/native references.
Spreadsheet types/conversions/callbacks now live in `adapters/spreadsheet-compat`,
which uses the same neutral compiler, session and cache owners.

Warm Add2(19#, 23#), release build, seven samples of 100,000 calls after 1,000
warmups; compile/initialize/bind are outside timing. Baseline is clean source
commit `0ad2ad2867364a9f4b39ba131e68104ae9531ed6`, built from a git archive with
the same 151 registry lock entries as the changed workspace. Two alternating baseline/after passes ran with no Cargo process active at the
start; this is a shared Windows host without thread pinning. The table uses the
first pass; repeat typed medians were 92.64 ns (VM3) and 15.74 ns (JIT). Nanoseconds/call, sample medians:

| Runtime | Before: legacy name/value call | After: bound dynamic adapter | After: typed call |
|---|---:|---:|---:|
| VM3 | 1184.91 | 198.06 | 90.66 |
| JIT | 950.00 | 598.71 | 15.68 |

These are end-to-end warm API costs for this small numeric body, not an Excel
comparison, cold compilation measurement or promise for other signatures.
The retained benchmark examples and raw samples make the comparison reproducible.
Allocation instrumentation separately counts zero allocations over 1,000 typed
calls on each backend, including the owned-session generation check.

Focused tests cover arities 0–4/mixed register positions, signed zero, loops/block
budgets, finite typed-host input checks versus exact dynamic carriers, faults,
signature refusal, distinct same-named procedures across modules, exact carriers,
shared Err, reset, stale/cross-session handles,
failed reload, cache eviction and generation exhaustion. The retained 19 legacy
adapter tests continue to exercise persistent state/defaults/ParamArray/objects.
Final repository meta-check passed: 3690 tests; 179 ignored and four
registry-writing fixtures excluded with no credit. Final focused/default-disabled
suites pass 35/33 checks. Both bounded delivery leaves are recorded complete;
the broader capability rows remain in-progress.

The [run archive](run6-validation.zip) retains raw receipts, benchmark harnesses,
lockfiles and source hashes; [manifest](run6-manifest.json) records its SHA-256.
Failures and canceled attempts receive no passing credit. The registry-writing
record tests excluded by `-NoRegistry` receive no credit either.

Reproduction uses `cargo run --release -p oxvba-embed --example call_cost` and
`cargo test -p oxvba-embed -p oxvba-spreadsheet-compat`. The repository check is
`./scripts/meta-check.ps1 -Fast -NoArtifacts -NoRegistry`, with
`CARGO_BUILD_JOBS=2` and `CARGO_PROFILE_TEST_DEBUG=0` on this Windows host.
The latter avoids this machine's MSVC PDB-linker failure; it does not change code
semantics. The portable interpreter/adapter tests also run with
`--no-default-features`; the retained wasm browser probe exercises the moved
adapter. Lean uses the already installed 4.29.1 binary directly, compiling
VarType.olean into a temporary OxVba directory and setting LEAN_PATH before the
models that import it.

Installed Lean 4.29.1 checks six abstractions, including ProcedureHandles and
ExecutionBudget. These do not prove Rust/native signature, pointer, code-memory,
descriptor or COM ownership. The unavailable pinned launcher attempted a download;
it was canceled. Automatic approval review blocked removal of its uninstalled
temporary folder. Rust/native refinement remains open in the formal backlog.

Fresh-eyes review actively checked code lifetime/identity reuse, generation wrap,
argument ABI positions, refusal paths, scalar specialization, Err/source fields,
initialization versus Main, runtime ownership and preserved adapter behavior.
It corrected cross-project error-source propagation, mixed numeric admission
and source-module implicit-entry coupling, and retained standalone image-entry
local snapshots. Long-only widening and conflicting temporary writers are refused. Typed scalar callee overflow reaches the caller handler on both
backends. Excel also observes error 6 plus a stored positive infinity under
Resume Next; generic Double fault/store ordering remains an explicit CORE-3
`bd-59co.2.4.3` gap. Generic arithmetic and its existing snapshot are unchanged;
they are not claimed as VBA parity.

Remaining delivery is owned by the two open residual beads above: general typed
carriers/bodies/internal calls/recursion/helper catalogue and sealed admission;
verified source-free sessions/general lifetime/reentry; general standalone run-entry
mixed-floating comparison lowering; StandardHost synthetic
dynamic-projection removal; and separately authorized DNA Calc glue migration.
