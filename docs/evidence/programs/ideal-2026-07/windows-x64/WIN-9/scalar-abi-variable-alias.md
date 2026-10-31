# Native scalar ABI and caller-variable alias evidence

2026-10-31 development evidence. Bounded delivery: `bd-59co.3.10.8`.
The wider WNI-PLAN-DECLARE row remains **in-progress**, with residual owner
`bd-59co.3.10`; this does not certify the Windows profile.

The owned Excel probes first exposed shared VM3/JIT defects. Passing one Long
variable twice to a DLL produced `41|2` instead of Excel's `42|42`; equal-valued
distinct variables and parenthesized copies behaved differently. A second probe
returned `True|True|Byte|1|42` in Excel versus `False|False|Long|-1|-1004` in OxVba:
Currency argument/return storage, Date return registers and Byte tagging were
wrong. The Boolean payload observation is an explicit remaining gap below.

The shared native boundary now receives `RuntimeCallArgument` with canonical
caller-slot IDs. VM3 resolves local/global aliases, including forwarded ByRef
parameters; JIT assigns IDs from resolved storage, with no address escaping the
call. The Windows HAL shares one compatible scalar cell per ID. Repeated aliases
reuse the original raw pointer, without a fresh mutable projection of its box.
Scalar cell pointers are derived after their owners enter the retained cell
vector. Distinct identities and temporary copies remain separate.

Currency uses scaled i64 arguments/returns, Date uses f64, and Byte returns and
writebacks retain their subtype. LastError is captured before libffi cleanup.
The VM3 host snapshot also now retains LastDllError on entry faults instead of
reconstructing it as zero. Native missing-entry descriptions contain the alias
and library, without nondeterministic module addresses.

`cargo test -p oxvba-differential --test jit_windows_vm3_parity` covers seven real
Windows tests, including the admitted alias and ABI cases:

- repeated Long variables, distinct equal values, parenthesized copies, nested
  forwarding and an alias reached through both a global and a ByRef parameter;
- Integer, Byte, Boolean 0/-1, Single, Double, LongLong, LongPtr, Currency and Date
  native storage/return widths, with mutation counts proving actual DLL entry;
- ten mixed register/stack arguments, large Currency and 64-bit integer values;
- named missing entry (453), absent valid ordinal 32767 (452), exact descriptions
  and preservation of a prior LastDllError=1234;
- unsupported repeated String carriers refused before the DLL counter changes;
- full Err state, actual transport, zero current-thread runtime handle deltas and
  removal of the owned DLL directory after each engine drops.

The HAL legacy-adapter regression proves that repeated native ByRef IDs are
refused before invoking an adapter that only implements the old value method;
distinct equal-valued arguments still delegate successfully.

## Excel target observations

The exact repository fixture source was copied and built with rustc as an owned
x64 DLL, then called from a freshly started Excel 16.0 build 20527 Safe Mode
process. Debug -> Compile command 578 succeeded before Application.Run. The
result was:

```text
42|42|41|41|2|42|42|41|2|42|42|True|True|True|True|True|True|True|True|True|14|True|True|Byte|42|0|4321||
```

This covers the same admitted alias/width cases and full clean Err state. The
separate named and valid-ordinal observations returned 453 and 452 respectively,
`Can't find DLL entry point <entry> in <library>`, project source `VBAProject`,
and LastDllError 1234 before and after resolution failed. Evidence under
`scalar-abi-oracle-2026-10-31/` retains source, controller, hashes, environment,
result and owned-process journals. Controller copies override only the diagnostic
source/expected value; worker, guardian, sealed descriptor and cleanup behavior
remain the repository implementation. All positive probes record owned-process
zero. No Office security, Trust Center or registry setting was changed.

## Remaining accepted work

- A foreign Boolean payload of 1 still becomes -1 under OxVba `CLng`, while Excel
  returns 1. `bd-59co.3.10.5` and carrier producer `bd-59co.3.2.5` own this.
- Field/element alias identity remains `bd-59co.3.3.7`; repeated structural/native
  string carriers, buffers and broader signatures remain `.3.10.5/.6`.
- Alias `#65535` produced an owned VBE compile modal, `Invalid ordinal`, while
  OxVba admits it. Selected token/line capture failed, so that probe failed closed
  and earns no compile-oracle credit. Declaration legality remains `.2.4.5`, and
  negative-oracle repair remains `.3.15.6`/BLK-EXCEL-COMPILE-20261031.
- Loader search/policy, callbacks, concurrency, initializer-failure error-state
  capture and certification remain with their existing owners.
- No existing formal model proves the native ABI or alias implementation. Kani
  is absent locally; the formal backlog records the proof/Miri follow-up. Native
  fixtures and explicit borrow/cleanup review are executable evidence, not a
  formal proof claim.

Final default-concurrency `scripts/meta-check.ps1 -Fast -NoArtifacts` passed:
3,640 workspace tests, 178 ignored, strict Clippy, formatting and governance.
The complete external log is
`C:/Work/DnaCalc/codex-runs/calcvba-meta-default.log`, SHA-256
`dbc881c05e6932d7bdabe4480346a8ad188d799cc9ff9dc68411561a0532bcca`.
Fresh-eyes review explicitly checked resolved caller identities, raw-pointer
reuse, scalar cell ownership, Byte tags, lookup descriptions and fault snapshots.
It also corrected stale architecture text; initializer faults remain an explicit
residual rather than being credited by the entry-fault test.

Implementation checkpoint: bcd6a25399f97e5e637e0554e92e1ec5868f1455
