# CalcVBA

CalcVBA, in the DnaCalc/CalcVBA repository, is a Rust implementation of Office VBA,
especially Excel VBA. It supports embedding in DNA Calc and compiling VBA into
other useful forms. The runtime is single threaded. Cross-platform support is
practical where the host services exist; native COM interop is Windows specific.
VBA UserForms belong in CalcForms/OxForms; VB6 forms and old VB6 projects are
outside the project scope.

The implementation is substantial and still incomplete. See the
[current architecture](docs/ARCHITECTURE.md) for the code layout and limits.

## Build and check

Install stable Rust, Git and PowerShell 7. Windows COM/native tests also use
the Windows SDK; opt-in Excel comparisons need an installed Excel.

```powershell
cargo build --workspace
./scripts/check.ps1
```

The check runs fmt, strict Clippy and workspace tests. Four optional COM record
fixtures register type libraries and are excluded, so ordinary checks leave
COM registrations alone. Live Office tests are opt-in. Tests cover compiler and
runtime behaviour, VM3/JIT comparisons, captured VBA observations and controlled
native COM fixtures. Test inputs live with their tests, in `conformance/`, or
in `oracle/`; generated observations go under `temp/`.

Before each commit: `git pull --no-rebase`, run the check, review the diff,
commit and push. Use [AGENTS.md](AGENTS.md) and the [charter](CHARTER.md) for
contributor guidance. Track real behaviour gaps in beads.

## Run VBA

Create `hello.bas`:

```vb
Sub Main()
    Debug.Print "Hello from CalcVBA"
End Sub
```

```powershell
cargo run -p calcvba-cli -- run hello.bas --backend vm3
cargo run -p calcvba-cli -- run hello.bas --backend jit
cargo run -p calcvba-cli -- run-project project.basproj --backend vm3
```

Both backends consume typed OxIR. The JIT compiles with Cranelift and refuses
unsupported programs before execution; it never silently falls back to VM3.
Use `--help` for current CLI options and `--diagnostic-format json` for tools.

## Packages and hosts

```powershell
cargo run -p calcvba-cli -- build project.basproj --target OxImage --out-dir out
cargo run -p calcvba-cli -- run-image out/ProjectName.oxi --backend jit
cargo run -p calcvba-cli -- build project.basproj --target ComServer --out-dir out
```

`.oxi` packages contain the compiled project closure. ComServer emits a
runtime-backed Windows COM DLL and type library, currently using VM3. The
[COM DLL example](examples/comserver) checks source Instancing, class factories
and an early-bound client in an existing owned Excel, with temporary private
per-user registration. Explicit JIT selection is refused for this COM host.
NativeLibrary emits a
Windows x64 DLL for admitted native exports. Native EXE outputs and JIT-backed
COM wrappers remain unfinished.

```powershell
cargo run -p calcvba-cli -- build examples/native/scalars/NativeScalars.basproj --target NativeLibrary --out-dir out
```

The installed CLI patches its embedded template without invoking Rust or a
linker. Each named export calls a compiled procedure directly. The initial
boundary supports ByVal/ByRef Byte, Integer, Long, LongLong/LongPtr, Single,
Double, Currency, Date and Boolean, with direct procedure calls
and typed Abs/Sqr library helpers. CLng/CDbl/CBool specialize numeric operands.
Long integer division/modulo and Boolean/Long bitwise operators are compiled.
Calls may cross standard modules, with up to 32 arguments and 128 active native
frames. BSTR strings, terminated UTF-16/ANSI inputs, native UDTs, typed SAFEARRAYs,
Automation Variants and COM object references have owning boundary operations.
Module and Static state initialize on the first call, outside the loader lock.
The DLL binds to that calling thread; calls from other threads leave caller
storage alone and report error 1444. Unhandled errors return zero, or an HRESULT
with a trailing result pointer under ErrorPolicy=HResult. CalcVbaGetLastError
returns the number and optional caller-owned source/description BSTRs.
CalcVbaGetLastErrorHelp retrieves retained dynamic Err.Raise help file/context.
Call CalcVbaShutdown on that thread before unloading to release module owners.
Some reachable constructs are still refused; see the
[native design](docs/spec/NATIVE_EXPORT_PE_TEMPLATE_V1.md) for limits.
Generated code uses baseline
x64 instructions; unsupported reachable code is refused.
`scripts/build-native-template.ps1` rebuilds the shipped template for developers.
`scripts/run-native-export-clients.ps1 -Excel` compares a single batch through
C, CalcVBA Declare (both backends) and one owned Excel instance. CalcVBA has no
XLL-specific target or runtime support.
`scripts/run-native-interop-clients.ps1` checks the owning boundary cases. Its
`-OracleSession` option submits the entire batch to an existing owned Excel
started by `scripts/run-native-export-oracle-session.ps1`; a `stop` file closes it.
`scripts/run-native-wide-clients.ps1` also checks narrow storage, 64-bit integer
precision, floating/currency/date signatures, explicit casts and overflow.
The [plain VBA XLL example](examples/native/xll) uses general native pointer and
Declare support for its XLOPER12 UDT, registration table and callbacks.
`scripts/run-native-xll-clients.ps1 -OracleSession <session>` checks both
25 worksheet calls in that existing Excel and measures batch call costs.
The ordinary VBA library converts rich Variants, arrays, errors, missing/nil
and references; `-DnaCalcPath <dna-calc-desktop.exe>` also compares the shared
cases in DNA Calc's XLL host.

Hosts use [`calcvba-host-api`](crates/calcvba-host-api) and
[`calcvba-host`](crates/calcvba-host). They supply application objects and type
information; Excel's Range/Application/WorksheetFunction model belongs to the
host. Bound procedures avoid repeated name lookup. Sessions retain globals,
Static variables, errors and objects. DNA Calc uses invocation slices, read
views, execution budgets, code ownership across reload and checked handles;
these remain part of the embedding API. See
[hosting](docs/spec/HOSTING_AND_EARLY_BINDING.md) and
[project files](docs/spec/BASPROJ_SPEC_V1.md).

Criterion benchmarks remain available with
`cargo bench -p calcvba-differential --bench jit_m4_baseline`.

## Excel comparisons

`oracle/` contains VBA probes and captured return values.
`crates/calcvba-host/tests/support/` contains source fixtures used by the native
hosting and saved-XLSM tests. The `run-typed-*-oracle.ps1`,
`run-automatic-events-oracle.ps1` and `run-xlsm-project-oracle.ps1` scripts compare
those tests with real Excel. Build the matching test executable with
`cargo test -p calcvba-host --test <name> --no-run`, then pass its path through
`-RuntimeTestExecutable`. The scripts record owned process identities, start a
PID-scoped UIA dialog handler before compiling/running, and compile through
VBE Debug -> Compile. They do not change Trust Center, macro security, Excel
COM registrations or HKLM. Cleanup must stay restricted to recorded processes.

VBA behaviour is decided by public specifications and reproducible black-box
Office observations. VM3 is the reference interpreter for the JIT; observed
Excel differences become regression tests. Existing engine behaviour is not a
substitute for VBA compatibility.

## Office files and source folders

`calcvba-office` reads VBA directly from `.xlsm`, `.xlsb`, `.xlam`, `.xls`
and standalone `vbaProject.bin` on supported Rust platforms. It preserves source
attributes, module roles, ordered references, project properties and forms.
Hosts can pass `project.to_manifest()` to `calcvba_host::Module::load_project`.
External referenced VBA projects require the host's project graph and host
object models still come from the host.

```powershell
cargo run -p calcvba-cli -- inspect-vba book.xlsm
cargo run -p calcvba-cli -- export-vba book.xlsm --out-dir source
# Edit the UTF-8 .bas/.cls/.frm files, retaining their Attribute lines.
cargo run -p calcvba-cli -- write-vba book.xlsm --project source/VBAProject.basproj --out changed.xlsm
```

The folder uses `CalcVBA.Sdk` and VBE source/form formats. Its Codepage controls
write-back conversion; characters that cannot round-trip are refused before
output is written. Saving replaces source streams and invalidates compiled
caches so Office recompiles. Locked projects, encrypted Office packages and signed
write-back are refused. Existing output paths are refused. See the
[project format](docs/spec/BASPROJ_SPEC_V1.md) for persistence details.

`scripts/run-office-project-oracle.ps1` runs all 33 saved round 14–15 projects
on VM3 and the JIT, then checks edited projects in all four Office formats and
VBE form import. Pass a new `-Directory`, the built `xlsm_projects` test
executable as `-RuntimeTestExecutable`, and the built CLI as `-CliPath`.
The complete batch reuses one recorded, owned Excel process.

## License

[MIT](LICENSE.md).
