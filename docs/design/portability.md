# Running VBA on four platforms

CalcVBA runs VBA written for Excel. Two VBA dialects exist in practice, and CalcVBA runs on four kinds of host:

- **Dialects:** the VBA of Excel for Windows, and the VBA of Excel for Mac. They share the language and most of the library, but differ in paths, COM, Declare, settings, file types, AppleScript and conditional-compilation constants.
- **Hosts:** Windows, macOS, Linux, and WebAssembly. WebAssembly comes as two kinds, WASI (a command-line or server sandbox) and the browser (a page or worker, no file system, no blocking).

This note covers four questions:
- which parts of VBA code are portable to which host;
- which need a shim to give the answer the code's dialect gives;
- which only differ in value by machine;
- which cannot work on a host.

It proposes run modes and a static check, so a user knows before running whether a project will behave as it does in the Excel it was written for. A project that won't is either shimmed or rejected with a clear reason.

The [portability matrix](../../crates/calcvba-symbol/src/portability/matrix.rs) covers the VBA library catalogue plus the statements and language features that reach the host. Run `calcvba-cli check --portability --host linux <project|source.bas|workbook>` for a JSON report. Rust hosts can call `calcvba_host::portability::check_project` on the project manifests. For each it gives:
- a class;
- how Mac VBA differs;
- a verdict per host;

The Excel object model (Workbooks, Range, Application.PathSeparator...) is a host concern: OneCalc and DNA Calc provide it. It needs the same treatment separately.

### JSON format

`calcvba-cli check --portability` writes one JSON object to standard output. The same object is available to Rust hosts as `calcvba_host::portability::Report`.

| Field | Version 1 meaning |
| --- | --- |
| `schema_version` | Integer `1`. Readers can ignore added fields. |
| `host` | `windows`, `macos`, `linux`, `wasi` or `browser`. |
| `analysis_complete` | `true` when the input parsed without errors. Examine this field before using the project verdict. |
| `verdict` | Project summary label, one of `portable`, `host-dependent values`, `portable with shims` or `needs Windows`. `needs Windows` is the existing summary label for a rejected finding on the selected host. |
| `findings` | Rows carry `project`, `module`, `procedure`, `line`, `feature`, `class`, `verdict`, `detail` and `mac_dialect`. The last field describes the Mac dialect. Module, procedure, line, detail and Mac dialect can be null. Source lines are one-based. |
| `parse_errors` | Rows carry `project`, `module`, `line` and `message`. |
| `references` | Ordered source type-library references, including unused ones. Rows carry `project`, `name`, `guid` (nullable) and `state` (`implemented`, `types_only` or `unknown`). |
| `effects_complete` | `true` when a bound procedure graph was available. For source, this requires binding and elaboration with the available static type information. |
| `procedures` | Bound rows carry `project`, `module`, `procedure`, `status`, `may`, `must` and `unknown`. Each effect carries `feature`, nullable `detail` and a finding verdict. |

Finding verdicts use the shared vocabulary `native`, `model`, `host`, `shim`, `replace`, `mediated` and `reject`. The static checker uses its catalogue and reports unknown references on every host. Their `findings` rows carry `reject`. In source-less images, `findings` include retained interfaces and `references` is empty. Compiled findings can identify a procedure without a source line.

Procedure statuses are `runs`, `may_stop` and `cannot_run`. `cannot_run` indicates a rejected must-effect. `may_stop` indicates a possible rejection or an unresolved call such as a computed ProgID. Bound effects follow static calls through the project closure and intersect branch effects on normal paths. Procedures with dynamic error control or computed branches can have a conservative may-set and an empty must-set. A source report can retain its findings when bound effects are unavailable.

When the faithful loader cannot resolve a library, procedures with an explicitly qualified declaration or `New` using one of its types refuse on entry and name the library. Other procedures can load and run in the same project.

## 1. References

| Dialect | Reference | State |
|---|---|---|
| Windows VBA | Excel for Windows (VBE7) on this machine, driven by the probe harness on a private desktop | In use: spec-pass, vba-library (177 modules, 4,865 cases) |
| Mac VBA | Excel for Mac on a Mac, driven by AppleScript or osascript instead of UI Automation. The Mac build is sandboxed, so its file probes run inside the container or with granted access | To set up. Until then, every Mac-dialect statement in the matrix says "to measure" |

The same probe modules run on both references. Where the two Excels disagree, the project's dialect decides. One known risk is that Mac Excel's maths library differs in the last bit from Windows Excel's. That would make "bit-exact" a per-dialect target, and the matrix keeps the two apart.

## 2. Classes

Each feature's class says what it depends on, relative to the Windows dialect:

| Class | Meaning | Rows | Examples |
|---|---|---:|---|
| **pure** | The result depends only on the arguments. Portable once CalcVBA's implementation is the same on every host. | 137 | Left, Mid, Val, CInt, DateAdd, Pmt, Collection, Err; also Sin to Log, whose results come from CalcVBA's own model (below) |
| **locale** | Depends on regional settings: date order, separators, names, compare rules. | 29 | CStr of a Date, CDate("1/2/2026"), Format, MonthName, StrComp text compare |
| **shim** | Host facilities that differ but can be emulated: paths, attributes, settings store, code page, environment. | 25 | Open, Dir, GetAttr, SaveSetting, Environ, Chr(200), Print # |
| **host** | The value depends on the machine or a UI host; the shape is portable. | 13 | Now, Timer, Command, MsgBox, UserForms |
| **windows** | Needs Windows itself. | 10 | Declare into Win32, CreateObject/GetObject of COM servers, SendKeys, AppActivate, Shell of a Windows command line, AddressOf |

## 3. Verdicts per host

For each host, the matrix says what CalcVBA must do there to give the dialect's answer:

| Verdict | Meaning |
|---|---|
| **native** | The host's own facility gives it |
| **model** | CalcVBA's own model, the same code on every host: the locale model, the code page, Excel-exact maths |
| **shim** | Emulation: path map, settings store, attribute store, synthetic environment |
| **host** | A machine value or the UI host's answer |
| **replace** | Only through a portable replacement: a portable COM class, an emulated Declare entry point |
| **mediated** | The platform or user grants access to the facility |
| **reject** | The facility is unavailable on the selected host |

Counts for Windows-dialect code (rows: 214):

| | Windows | macOS | Linux | WASI | Browser |
|---|---:|---:|---:|---:|---:|
| native | 158 | 129 | 129 | 128 | 128 |
| model | 42 | 42 | 42 | 42 | 42 |
| shim | 0 | 20 | 19 | 20 | 20 |
| host | 13 | 13 | 13 | 10 | 13 |
| replace | 0 | 5 | 6 | 5 | 5 |
| reject | 1 | 5 | 5 | 9 | 6 |

"Model" appears on Windows too, deliberately. Locale, code page and transcendental maths come from CalcVBA's own implementation on every host, so results don't depend on the OS's maths library or regional APIs. Today Sin's last bit already differs between Linux and Windows (dc-q6brh.112), and date text ignores regional settings on every host (dc-q6brh.6).

## 4. The hosts

### Windows
- **Native today:** paths, attributes, COM, Declare into Win32, Shell and UI.
- **Gaps on every host, Windows included:** settings are in memory (no registry), text files are written as UTF-8, and there is no locale model.
- **Mac-dialect code on Windows:** the reverse problem. Mac Declares name dylibs, which are rejected, and MacID and MacScript need replacements or rejection.

### macOS
- **Profile:** CalcVBA's MacOs profile exists, with file system, process, Declare (dylibs) and UI on and COM off. On it `#If Mac` is True, as in Mac Excel.
- **Mac-dialect code** is native here: POSIX paths, Mac Declares into dylibs, AppleScript (where the sandbox allows) and plist settings. Check each against the Mac reference.
- **Windows-dialect code** needs the same shims as on Linux:
  - path map;
  - attribute store for hidden/system/archive;
  - Windows Environ;
  - no COM, so portable classes or rejection;
  - Win32 Declares, emulated or rejected.
- **The default file system is case-insensitive (APFS),** which helps, but case-sensitive volumes exist, so the shim still resolves names itself.
- **JIT:** on Apple Silicon it needs MAP_JIT and the hardened-runtime entitlement for generated code.
- **Code signing and notarisation** are a packaging concern, not a VBA one.

### Linux
- **Today:**
  - Windows paths go to the OS unchanged; `\` and `C:\` are filename characters.
  - Wildcards are case-sensitive, and only read-only and directory attributes exist.
  - Settings are in memory and text files are UTF-8.
  - Chr, Asc, StrConv and String use a fixed Windows-1252 table.
  - Environ names are case-sensitive, with no TEMP or USERPROFILE.
  - A Declare into Win32 compiles, then fails at run time: dlopen looks for libkernel32.so.
  - COM is limited to Scripting.Dictionary, and Shell runs `sh -c`.
  - Linux is treated as 64-bit Windows for `#If`.
- **Plan:** the full Windows shim set (section 6). The Mac dialect also needs a path map and a plist-like store, but no drive letters.

### WebAssembly: WASI
- **Profile:** CalcVBA's WasmWasiLocal profile turns off the file system, process/environment and dynamic linking. A WASI runtime gives preopened directories, environment variables passed explicitly, clocks, and stdio.
- **File statements:** run through the path map onto preopened directories, the only file access there is.
- **Settings:** a file in a preopened directory.
- **Environ:** whatever the embedder passes, plus the synthesised Windows names.
- **No processes or dynamic linking:** Shell, every Declare without an emulation, COM, AddressOf and UserForms (no window system) are rejected.
- **Execution:** no JIT (no runtime code generation), so VM3 runs.
- **Time:** clocks come from WASI. The time zone is what the embedder gives (often UTC), so Now is a host value with a configurable zone.

### WebAssembly: browser
- **Profile:** CalcVBA's WasmBrowserSandbox profile has no file system, process, dynamic linking or UI.
- **Files:** go through the path map onto a virtual file system (in memory, or the origin-private file system for persistence). Files are imported and exported through the page, never the user's disk directly.
- **Settings:** in the page's storage (IndexedDB), keyed as the registry is.
- **UI and blocking:** MsgBox and InputBox go to the page as a host callback. VBA is synchronous and the browser's main thread cannot block, so VBA runs in a worker. A synchronous channel to the page (shared memory and Atomics.wait) lets MsgBox wait for the user's answer. Without that, a project that uses modal UI is rejected for the browser.
- **DoEvents and long loops** yield to the worker's event loop.
- **Rejected:** Shell, Declare (bar emulations), COM (bar portable classes), SendKeys and AppActivate.
- **Execution:** no JIT, so VM3 runs; maths comes from the model, as everywhere.
- **Time:** the clock is `Date.now`, the time zone is the browser's (`Intl`) unless the project fixes one, and Timer resolution is reduced by the browser.
- **Locale:** from the project, falling back to the browser's language, never silently.

## 5. Modes

### Execution engine

With `BackendPolicy::Automatic`, the host API selects JIT on Windows, Linux
and macOS desktop builds with the `jit` feature. Browser and WASI hosts use VM3,
as do builds of `calcvba-host` with default features disabled. A desktop host
that cannot allocate executable memory, including a hardened macOS deployment
without the JIT entitlement, supplies `BackendPolicy::NoCodeGeneration`.
The host supplies `BackendPolicy::Debugging` for VBA debugger sessions. Both
policies select VM3 before the project is loaded. With JIT selected, a refusal
or compilation failure is reported to the caller.

Guarded mode uses the same default engine. The host's capability policy limits
effects on either engine. An embedder can also apply `NoCodeGeneration` to
Guarded deployments that prohibit code generation.

### VBA policy and dialect

Three settings decide how a project runs:

1. **Policy** (exists): what the code may touch. Safe or Full capabilities; the deterministic or native host policy.
2. **Dialect** (new): which Excel the code was written for: `windows` (the default for workbooks from Windows Excel), or `mac` (workbooks from Mac Excel, or a project setting). The dialect decides:
   - the conditional-compilation constants (`Win64`/`Win32` or `Mac`);
   - path rules (drive letters and `\`, or POSIX);
   - which Declare libraries are native (Win32 DLLs or dylibs);
   - where settings live (registry or plist);
   - which reference CalcVBA is checked against.
3. **Portability strictness** (new):
   - **faithful:** shims and replacements on. The load-time report lists unavailable facilities. The host can refuse a procedure before entry when its effect analysis establishes that it must reach an unavailable facility. A call reached during execution raises the corresponding VBA error.
   - **strict:** for checking, CI and servers. Every windows-class construct, and anything without a native or model verdict on this host, is rejected, and host-class uses are listed.
   - **native:** for code written for the host itself, such as a Linux script. No dialect emulation.

Strict load diagnostics name the construct, location, host, and reason. The static report is available independently of execution through `check --portability` and the host API.

On a non-Windows native host, `HostPolicy::windows_faithful` enables faithful runtime handling for `Shell` and Windows DLL calls. The CLI sets it when selecting `windows-faithful`. An embedder can supply `HostPolicy::windows_shell_commands` for the project. Each entry maps one exact VBA command line to an explicit executable and argument list. An absent mapping raises the file-not-found error. Process policy still applies to mapped commands.

`HostPolicy::with_windows_drive_mount` maps a drive-absolute prefix such as `C:\Temp` to a host directory with `Read` or `ReadWrite` access. The longest case-insensitive prefix wins after Windows path components are resolved, including `..` and drive-relative paths. `CurDir` retains the Windows path. The reverse map uses the deepest matching host directory. Named UNC root bindings remain available through `with_windows_unc_root`. Mapped `Dir` and `Kill` diagnose a wildcard collision with both filenames before `Kill` can remove a match.

Off Windows, the faithful `Environ` table supplies the Windows persona. It reads user and machine identity from the host and resolves file-related aliases through configured mounts. Other host variables require an entry in `HostPolicy::windows_environment_passthrough` or a CLI `--pass-env NAME` option. Repeat the option to set the order of forwarded entries before synthesized entries. Names are matched case-insensitively in VBA, and indexed lookups are one-based. A reverse path must resolve back to the same host location, including when one mount shadows a directory in another mount.

## 6. Shims and models

| Shim or model | Hosts | Design |
|---|---|---|
| **Path map** | macOS, Linux, WASI, browser (Windows dialect); Windows, Linux, WASI, browser (Mac dialect) | Per project. Windows dialect: drive letters to roots, UNC to mounts, `\` to `/`. Mac dialect: POSIX paths under mapped roots. CurDir reports the dialect's form; an unmapped drive or root raises 76. On WASI the roots are preopened directories; in the browser, the virtual file system. |
| **Case-insensitive lookup** | Every non-Windows host | Resolve each component as NTFS (or default APFS) does. Two names differing only in case are an error naming both. |
| **Attribute store** | macOS, Linux, WASI, browser | Hidden, system and archive bits, kept in extended attributes where available, else a sidecar. |
| **Settings store** | All but Windows-dialect on Windows | The registry on Windows, a plist-like file on macOS (Windows dialect: a file; Mac dialect: measured against Mac Excel), a file on Linux and WASI, IndexedDB in the browser. Keys follow `VB and VBA Program Settings\app\section\key`. |
| **Code page** (model) | All | The project's ANSI code page, default 1252 for the Windows dialect; Mac dialect to measure (perhaps Mac Roman). Used for Chr, Asc, String, StrConv, Input, Get/Put and text files, which are written and read in it, not as UTF-8. |
| **Locale** (model) | All | Date order, separators, month and day names, case and compare rules, per project; the default is the regional settings of the reference machine. Checked against Excel under at least two locales on each reference. |
| **Maths** (model) | All | One implementation of Sin, Cos, Tan, Atn, Exp, Log and `^` that reproduces the reference Excel's results. Per dialect if the Windows and Mac references differ. |
| **Environ** | Non-Windows hosts for the Windows dialect | Case-insensitive names, synthesised TEMP, TMP, USERPROFILE, APPDATA, COMPUTERNAME, USERNAME and so on through the path map, and VBA's 1-based order. The Mac dialect gets the Unix names, as on a Mac. |
| **Emulated Declare entry points** | Every host but the dialect's own | A table, chosen by frequency in the workbook corpus. Windows: GetTickCount, Sleep, QueryPerformanceCounter, RtlMoveMemory, lstrlenW, GetTempPathW, GetUserNameW, MultiByteToWideChar... Mac: common libc calls. The rest are rejected. |
| **Portable COM classes** | Every host but Windows | Scripting.Dictionary (exists), Scripting.FileSystemObject, VBScript.RegExp, ADODB.Stream, MSXML2 (parsing), chosen by corpus frequency. Each is checked against the Windows reference by probes. On Mac Excel these classes don't exist, so Mac-dialect code doesn't use them. |
| **Shell** | | No general shim. Faithful mode may take a per-project command map for named programs; everything else is rejected. |
| **UI host** | All | MsgBox and InputBox go to the embedding host: a native dialog, a terminal, scripted answers, or the browser page through the worker channel. |

## 7. Knowing it's the same

- **The same Excel-answered corpora run on every host.** That means spec-pass and the vba-library probes, file-system probes included, through the path map. In faithful mode each line must match the dialect's reference. Host-class lines are compared by shape (type, range, relation), as the probes already print them.
- **A pure line that differs between two hosts is a bug,** like the Sin last bit today.
- **The portability report** (static, section 5) runs over the workbook corpus per host. The headline number is the share of real projects that run unchanged, run with shims, or are rejected, per host. Refused counts as fail.
- **The Mac reference** answers the same probes; where it differs from Windows, the matrix's mac_dialect column gets the measured fact in place of "to measure".

## 8. Bugs on every host

These are wrong whatever the mode:
- **Open and Name create missing parent folders.** VBA raises 76.
- **Text files are written as UTF-8,** where VBA writes ANSI.
- **Settings don't persist,** on Windows too.
- **Environ(n) order differs** from VBA.
- **Maths uses the platform library.**
- **There is no locale model.**

## 9. Defaults chosen here (open for review)

- **Dialect:** `windows` for workbooks from Windows Excel and VBE exports; `mac` when the workbook or project says it comes from Mac Excel.
- **Strictness:** faithful by default; strict for CI and servers.
- **Path map:** `C:` is a per-project root for the Windows dialect, with no other drives mapped. In the browser the root is the project's virtual file system.
- **Case collisions:** an error.
- **Code page and locale:** the reference machine's settings (code page 1252, en-ZA regional settings, until the locale work, dc-7ofm, says otherwise).
- **Pointer width:** in wasm32, VBA keeps 64-bit LongPtr semantics as on Windows Excel x64, so code that does pointer arithmetic sees the sizes it was written for.

The work items are beads under the epic dc-8svfs:
- report: .1
- modes: .2
- path map, case and attributes: .3
- settings store: .4
- code page: .5
- Environ: .6
- emulated Declare: .7
- portable COM classes: .8
- Open/Name parent folders: .9
- verification: .10
- Mac Excel reference: .11
- Mac dialect: .12
- WASI host: .13
- browser host: .14

Locale and maths are dc-q6brh.6 and dc-q6brh.112 in the library check.
