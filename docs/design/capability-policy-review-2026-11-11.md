# Capability policy for CalcVBA's host layer: design review (2026-11-11)

Read at CalcVBA `4640a317` and DnaCalc master, both pulled. Read-only: nothing was built or run, and no Excel was consulted, so every Excel error number below is from documented VBA behaviour, not a capture.

**Verdict.** Cheap, and worth a first sprint. Every runtime OS service in VM3 and the JIT already goes through one `&dyn HostServices`, so a single wrapper covers both backends. The existing `HostPolicy` flags are not a Safe mode: reads, `Dir`, `Environ` and `Command` are ungated, denials surface as error 5 with a Rust debug dump, and the "deterministic" mode fabricates values. Safe mode is a policy gate, not a sandbox; it does not make hostile VBA safe on a Windows desktop.

## 1. Does the current host layer make this cheap?

**What carries OS services today (CalcVBA `crates/`):**
- `calcvba-hal/src/traits.rs:92` `HostServices` hands out `FileSystemHal` (:170), `ProcessEnvHal` (:323: Shell, Environ, Command, Get/Save/DeleteSetting, Dir), `ComHal` (:367), `UiInteractionHal` (:144: MsgBox, InputBox, SendKeys, AppActivate), `DynamicLinkHal` (:616: Declare), plus console, events, time and diagnostics.
- All 60-odd library host built-ins funnel through `calcvba-lib/src/host.rs` (dispatched at `lib.rs:455-503`), shared by both backends.
- Declare has exactly two call sites, both through `host.dynlink()`: `calcvba-vm3/src/lib.rs:5044` and `calcvba-jit/src/helpers.rs:4771`. `calcvba-com/src/windows_ffi_bridge.rs` (LoadLibrary/dlopen) is used only by `calcvba-hal/src/adapters/standard/dynlink.rs`.
- `CoCreateInstance`/`GetActiveObject` (`calcvba-com/src/windows_client.rs:256,290`) are reached only via `ComHal::create_object_variant`/`get_object_variant`, called from `calcvba-lib/src/host.rs:499,506`.
- A policy vocabulary exists: `model.rs` `CapabilityId`, `HostPolicy { allow_process_spawn, allow_filesystem_mutation, allow_dynamic_link, allow_com_activation, allow_interaction, deterministic_mode }`, and `error.rs` `HalErrorKind::PolicyDenied` (`HAL-E-POLICY-DENIED`).
- A wrapping host is established practice: DnaCalc's `SpreadsheetHost` (`crates/calc-vba/src/runtime/spreadsheet.rs:769`) and the thread-local `output_tap.rs`.

**What makes the existing flags leaky or wrong as a Safe mode:**
1. **Reads are ungated.** With `deterministic_mode: false`, `Open ... For Input` (`standard/filesystem.rs:486` gates only `mode != 0`), `FileLen`, `GetAttr`, `FileDateTime`, `CurDir`, `Dir`, `Environ` and `Command` (`standard/process.rs:37,258,292`) reach the real OS whatever the `allow_*` flags say. A UDF can return `~/.ssh/id_rsa` into a cell.
2. **Denials are the wrong error.** `PolicyDenied` carries no `host_error_code`, so VBA sees error 5 (`calcvba-rt-abi/src/lib.rs:143`) or, through the library path, error 5 whose `Err.Description` is `format!("{err:?}")` (`calcvba-lib/src/lib.rs:77`).
3. **`deterministic_mode: true` invents results instead of refusing.** `Environ("X")` returns the key length, `Shell` returns 1, `Dir` returns 0/1 (`process.rs:228-232,283-289,345-357`); files live in an in-memory map. This is the default of `Engine::new` (`calcvba-engine/src/engine.rs:473`). Corpus code run this way gives plausible wrong values, not refusals.
4. **Typed COM calls skip the HAL per member.** `calcvba-rt-abi/src/host_entry.rs` and `calcvba-com/src/typed_slots` call vtable slots directly on the object pointer. COM can therefore be gated and audited only at activation.
5. **Compiled outputs are outside any policy.** `calcvba-native-template/src/services.rs:27` hard-wires `HostPolicy::native()`; AOT Declare loads libraries itself (`native-template/src/foreign.rs`); `calcvba-comhost` receives a real Excel `Application`.
6. **Compile time touches the OS before any runtime policy.** Reference resolution loads type libraries by registry and by path (`calcvba-com/src/typelib_catalog.rs:179,223` → `LoadTypeLibEx`). I did not trace whether that path can come from the workbook's reference record; if it can, a UNC path is a network touch at compile time.
7. **Full mode's Shell is not Excel's.** `standard/mod.rs:981-998` runs `cmd /C <text>` or `sh -c <text>`, which interprets shell syntax; Excel's `Shell` starts a process directly. Full mode is more dangerous than Excel here, and semantically different.

**DnaCalc today.** `SpreadsheetHost` sits on `HostBuilder::new().build()` (`spreadsheet.rs:295`), the Null host, so files, Shell, Environ, Declare and UI are already refused (as `CapabilityUnavailable`, error 5). The one live path is Windows `platform_com()` (`spreadsheet.rs:817`): `New X` of any class in a referenced type library activates, including `Scripting.FileSystemObject`, `WScript.Shell` and `WinHttpRequest`. Only literal `CreateObject`/`GetObject` are refused (`runtime/late_binding.rs`). `calc-host-vba/src/trust.rs` `ProjectGrant` is all-or-nothing, with no capability dimension.

## 2. The smallest correct design

**Where the policy lives.** One new file, `calcvba-hal/src/capability.rs`: a `PolicyHost` that implements `HostServices` around any inner `Arc<dyn HostServices>`. Do not edit the ~60 existing `if !self.policy.allow_*` sites. Wrapped outermost, it covers the standard adapter, the CLI and DnaCalc's `SpreadsheetHost` (including `platform_com`) identically, and both backends, because both reach the OS only through `exec.host`. With no wrapper installed, behaviour and cost are exactly today's.
- Forwarding hazard: `ComHal` has about 35 methods, many with default bodies. A wrapper that forgets one silently changes COM behaviour. Either add a test that drives every method through the wrapper against a counting inner host, or gate COM activation at `calcvba-lib/src/host.rs:499,506` instead.

**What a capability is (first sprint: eight, each with a free-text detail).**

| Capability | Gated operations | Detail logged | Safe-mode error |
|---|---|---|---|
| `fs.read` | Open For Input/Binary-read, Dir, FileLen, GetAttr, FileDateTime, CurDir | path | 70 Permission denied |
| `fs.write` | Open for Output/Append/Random/Binary-write, Kill, MkDir, RmDir, Name, FileCopy, SetAttr, ChDir, ChDrive | path(s) | 70 |
| `process` | Shell | command text | 70 |
| `env` | Environ, Command | variable name | 70 (see question 1) |
| `settings` | GetSetting, GetAllSettings, SaveSetting, DeleteSetting | app/section/key | 70 |
| `ffi` | every Declare call | `lib!symbol` | 53 File not found |
| `com.activate` | CreateObject, GetObject, `New` of an external coclass | ProgID | 429 ActiveX component can't create object |
| `ui.keys` | SendKeys, AppActivate | none | 70 |

- Gate file access at `open` and at path operations only. Handle operations (`Print #`, `Get`, `EOF`) pass through unlogged; a denied open means no handle exists.
- Network and clipboard are not capabilities yet. No VBA built-in reaches them; they are reachable only through `com.activate` or `ffi`, which Safe denies. The logged ProgID or symbol lets the report name them (a small table: `MSXML2.XMLHTTP`/`WinHttp` → network, `Scripting.FileSystemObject` → fs, `WScript.Shell` → process and registry).
- Not gated: Now/Timer, Debug.Print, DoEvents, MsgBox/InputBox (already governed by `allow_interaction` and `ui_virtualization`), and the host object model (the embedder's job).

**The error raised.** Set `host_error_code` and `HostErrorInfo.description` on a `PolicyDenied` error so `Err.Number`, `Err.Description` and `On Error` behave as for any run-time error; keep `stable_code = HAL-E-POLICY-DENIED` so a harness can tell a policy 53 from a real one.
- 70 is what Excel raises for a file locked by another process and for a blocked `SendKeys`; 53 is what it raises for a Declare whose DLL is absent; 429 for a class that cannot be created. Those are the nearest real analogues.
- Excel also uses 75 (Path/File access error) for some denied path operations. Excel has no such policy to match, so I recommend one number per family now and an oracle capture of real ACL denials later.

**The audit record.** `{ seq, capability, op, detail (truncated to 260 chars), decision: allowed | denied, err_number }`, where `op` is the `&'static str` the HAL already passes to its error constructors.
- Storage: an `Rc<RefCell<..>>` on the wrapper (the runtime is single-threaded on the owning VBA thread), with exact per-`(capability, detail)` counts and the first N full records, so a Declare in a loop cannot grow it without bound.
- The harness drains it after each call, which gives per-case attribution without the HAL knowing the procedure.

**The predicted set.** Compute it from bound OxIR, not from text. For each function in an `OxProgram`, walk `OxInst` (`calcvba-oxir/src/inst.rs`):
- `CallNative { Builtin(NativeImplId) }` → capability from the same `NativeImplId → Capability` table the wrapper's operations use (one table, an exhaustive `match`, so a new built-in fails to compile until classified).
- `CallNative { Declare { descriptor_id } }` → `ffi`, with library and symbol from the program's `external_calls`.
- Activation with a constant ProgID → `com.activate` plus the ProgID; non-constant → `com.activate`, ProgID unknown.
- `ComCallLate`, `CallByName`, `CallProcRef` → a "dynamic" marker, not a capability.
- When a project does not bind (wasabi does not parse today), fall back to a `calcvba-syntax` scan for `Declare` statements and the built-in names, labelled `lexical`. `late_binding.rs` in DnaCalc already tokenises for this and its `x.CreateObject` exclusion should be reused.
- This replaces the hand-maintained OS entries of `needs` in `crates/calc-vba/tests/corpus/*/project.json`; no code in either repository computes them today.

**Reporting prediction against observation.** Per project and backend, with P = predicted and O = observed (requested, whether allowed or denied):
- `O \ P` is a triage miss. It must be empty; each entry is a triage bug and becomes a test.
- `P ∩ O` is confirmed.
- `P \ O` is **unconfirmed**, not "wrong": the code was not executed, or the prediction is too wide, and the log cannot tell which. Report it beside how many procedures were entered out of how many exist.
- The only accuracy the data supports is recall on exercised code, `|P ∩ O| / |O|`. Precision cannot be claimed.
- Second caveat: in Safe mode a denial changes the path (the error is raised), so requests after the first denial on that path are never observed. Safe-mode observation is a lower bound even for executed procedures; Full mode sees more but only runs on trusted code.

## 3. First sprint (four steps)

1. **`PolicyHost`, the capability table and the error numbers in `calcvba-hal`.** Test in `calcvba-differential`: one small `.bas` per gated operation, run on VM3 and the JIT under Safe with the native standard host and a temp directory. Result: a table of operations × two backends, every row the same `Err.Number`, the same audit record, and no file or process created.
2. **`predicted_capabilities(&OxProgram)` and a `calcvba triage <project>` CLI subcommand printing JSON.** Result: predicted sets for the 10 corpus projects and the 580 conformance fixtures; riff, wasabi (lexical) and vba-web show `ffi`, matching today's hand-written `needs`.
3. **DnaCalc: bump the pin, wrap the host in Safe, drain the log per case in `apps/dna-onecalc/tests/vba_corpus.rs`.** A case with any denied record is classed `refused (policy: <capability>)`, never matched or mismatched. Result: `baseline.json` counts unchanged (module route 208/317 on both backends, file route 175/272), which is the evidence that Safe mode does not alter pure UDFs; and a P/O table with `O \ P` empty.
4. **An own fixture project, one function per capability, each with and without `On Error Resume Next`, plus one function never called.** Run Safe everywhere and Full on firehorse in a temp directory. Result: recall 100% on both backends, the uncalled function reported as unconfirmed, and every swallowed denial still classed as a policy refusal.

**Not now:** path allow-lists or a virtual file system for Safe; a per-ProgID allow-list (`Scripting.Dictionary`); network and clipboard as capabilities; module/line attribution; call-graph reachability; coverage instrumentation; consent UI or tying capabilities to `ProjectGrant`; resource limits; AOT, native-template and comhost outputs; rewriting the existing `allow_*` checks; OS-level sandboxing; correct 53/75/76 numbers for real file errors in the native adapter (a separate correctness bead; most map to 5 today).

## 4. Risks

**Semantics when a call is denied.**
- Pure UDFs make no host call, so the wrapper is never entered; step 3's unchanged baseline is the check.
- `On Error Resume Next` swallows a denial and the function returns a value Excel would not. Scored naively, that is a "mismatch" that pollutes correctness numbers, or a lucky "match". The audit log is what prevents this: classify by the log, not by the value.
- One loaded module keeps globals across calls (VBA_UDFS.md). A denial in one case can leave state that changes later cases; mark the session as tainted after the first denial and flag later cases.
- Do not reuse `deterministic_mode` as Safe: fabricated values are worse than errors for validation.

**JIT and interpreter.** Both go through `exec.host` at the sites cited in section 1, and the JIT needs no codegen change. The step 1 test keeps this true. The exceptions are the AOT/native lane (`calcvba-jit/src/native.rs:450`, native-template), which is outside the policy, and typed COM member calls, which are invisible after activation.

**Performance.** One extra dynamic call on host operations only; nothing on arithmetic, strings or project calls. The only volume risk is a Declare in a hot loop (riff), handled by counting rather than appending.

**Where Safe mode could give a false sense of safety.**
- It is not a sandbox. The JIT emits native code and the runtime has substantial `unsafe`; an engine bug triggered by hostile source bypasses any host policy. Public-spreadsheet VBA should still run only on an unprivileged Linux account with no secrets. The ruling in VBA_VALIDATION.md §5 could move to "Safe mode on Linux", not to "anywhere".
- Denial of service is untouched: infinite loops and allocation. DnaCalc has instruction budgets; the CalcVBA CLI and sweeps do not.
- The embedder's object model is out of sight: `Workbooks.Open`, `SaveAs`, `Application.Run`, `ExecuteExcel4Macro` in whatever host supplies `Application`. With comhost and a real Excel, that is full OS reach.
- Compile-time type-library loading (item 6 in section 1) runs before the policy exists.
- In Full mode on Windows, an activated COM object is unaudited from then on, and Shell is `cmd /C`.

## 5. Owner questions

1. **Environ, Command and GetSetting in Safe mode:** raise 70, or return the empty string Excel gives for an unset variable? I recommend the error: it is loud and cannot be mistaken for a result. The cost is that libraries calling `Environ("TEMP")` during initialisation are refused.
2. **`Scripting.Dictionary` and `VBScript.RegExp` on Windows:** allow in Safe mode? They are pure and common (dna-examples, vba-web). I recommend no for this sprint, so Safe means the same on Linux and Windows.
3. **Is Safe mode meant to license public-corpus VBA on the Windows desktop, or only on Linux?** I recommend Linux only until there is fuzzing evidence for the engine.
4. **Defaults:** should Safe become the default of `Engine::new` and the CLI? I recommend an explicit flag for the CLI, Safe always in the DnaCalc corpus runner, and no change to other defaults in this sprint.
5. **`ProjectGrant`:** does `PersonEnabled` imply Full (Excel's meaning of "enable macros") and `EvidenceManifest` imply Safe? I recommend leaving this out of the sprint, but it decides where the policy is chosen in `calc-host-vba`.
