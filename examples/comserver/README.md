# COM DLL example

`OutputType=ComServer` builds an in-process COM server. Class source attributes
choose Private, PublicNotCreatable, MultiUse or GlobalMultiUse; the project file
contains no separate Instancing overrides. SingleUse and GlobalSingleUse require
an out-of-process ComExe and are rejected for this DLL target.

For a developer build:

```powershell
cargo build -p calcvba-cli -p calcvba-comhost
./target/debug/calcvba-cli.exe build examples/comserver/ComServer.basproj --target ComServer --out-dir temp/com-example
```

Keep the emitted DLL, type library, `.oxi` and `.comserver.json` together. The DLL
provides the COM class factory, apartment registration, Automation dispatch,
connection points and supported dual interfaces. It executes the project's
class code through a VM3 session on the owning VBA thread. This is a VM3 COM host;
there is no attempt to run JIT code followed by an interpreter fallback.
An explicit RuntimeFlavor=Jit is refused at build time.

Private classes stay inside the project. PublicNotCreatable classes have type
information and can be returned by another class; they have no external class
factory. MultiUse creates independent instances. GlobalMultiUse adds global
members to the referencing client's binding context using the type library's
`appobject` flag. PublicNotCreatable uses `noncreatable`.
See Microsoft's [appobject](https://learn.microsoft.com/en-us/windows/win32/midl/appobject)
and [type library attributes](https://learn.microsoft.com/en-us/windows/win32/midl/type-library-attributes).

Use `scripts/run-com-server-clients.ps1 -OracleSession <existing-session>` for
an Excel VBA comparison. It creates fresh private test identities, checks the
class factories and type library, registers only those identities under HKCU,
compiles an early-bound client with visible VBE Debug -> Compile, then runs it
in the already-owned Excel and compares the original ordinary VBA source.
It checks instance state, ByRef aliases and writes before faults, rich Variant
arrays, class visibility, dual properties, global state, events and error 457.
The script unregisters the test identities in `finally` and verifies their keys
are absent. It does not register Excel add-ins or imported Office interfaces,
change security settings, write HKLM, launch Excel or quit another session.

COM work remains before EXE output: callback re-entry can panic on the borrowed
session (bd-79wxv), Object parameters still reject foreign Automation objects
(bd-styhq), and factories/dual methods need the original VBA error information
through IErrorInfo (bd-1t1qc). These cases are not covered by the passing example.
