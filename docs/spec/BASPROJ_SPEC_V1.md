# .basproj project files

`oxvba-project` reads XML project files. This is a project loader, not an
MSBuild engine. Source paths are relative to the project file. For example:

```xml
<Project Sdk="OxVba.Sdk/0.1.0">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <ProjectName>Hello</ProjectName>
    <EntryPoint>Main.Main</EntryPoint>
  </PropertyGroup>
  <ItemGroup>
    <Module Include="Main.bas" />
  </ItemGroup>
</Project>
```

`PropertyGroup` holds scalar properties; later values replace earlier ones.
`OutputType` accepts HostModule, Library, NativeLibrary, Exe, ComServer or ComExe.
`ProjectName` names the VBA project and `EntryPoint` can select Module.Procedure.
`DefineConstants` supplies conditional compilation values. `RuntimeFlavor`
and the existing host configuration properties are parsed into the manifest.

`ItemGroup` accepts Module, ClassModule and DocumentModule source items,
ProjectReference, COMReference and NativeReference. Class items can carry
VB_PredeclaredId, VB_Exposed, VB_GlobalNamespace and VB_Creatable metadata,
plus Instancing, ProgId and Description for COM packaging. Document items
carry HostDocumentType supplied by the host.

ProjectReference Include points to another project. COMReference Include names
a library, with Guid, VersionMajor, VersionMinor, Lcid and ImportLib metadata.
References retain their declared order. NativeReference describes a native
library. Import Project includes another project fragment. Convention discovery
can supply source files when explicit module items are absent.

BuildTarget describes packaging: the historical Bundle spelling selects the
current .oxi package; WrapperExe, WrapperLibrary and WrappedComServer describe
runtime-backed outputs. CLI `build --target OxImage` explicitly selects .oxi.
CLI `build --target NativeLibrary` requires `OutputType=NativeLibrary` and emits
a Windows x64 DLL using the embedded prebuilt runtime template. It runs no Rust
compiler or linker. Native EXE output remains unfinished.

## 2.5.1 NativeExport — named C exports

```xml
<NativeExport Include="Add">
  <Module>Numbers</Module>
  <Procedure>AddNumbers</Procedure>
  <CallingConvention>Stdcall</CallingConvention>
  <Ordinal>7</Ordinal>
</NativeExport>
```

`Include` is the exported name. Module and Procedure are required, matched with
VBA name rules. CallingConvention accepts Stdcall (default) or Cdecl; Windows
x64 uses the same convention for both. Ordinal is optional, nonzero, and unique.
Export names are case-sensitive and unique. The procedure must be Public in a
standard module. Its actual signature comes from binding, including ByVal/ByRef
and exact types; cached declaration strings do not supply the ABI.

The initial native implementation admits ByVal Long, Double and Boolean Functions
with numeric locals, direct calls to admitted procedures, and Double Abs/Sqr
library calls. Native signatures support up to 32 parameters. Unsupported
reachable signatures, state and instructions fail before the DLL is written.
The DLL contains no interpreter fallback. Other numeric widths and the remaining
native boundary types are unfinished. See [the native template design](NATIVE_EXPORT_PE_TEMPLATE_V1.md).

CalcVBA has no XLL target, registration metadata, conversions or host support.
Addin is rejected as an OutputType. Category, Description, ArgumentDescriptions
and unimplemented StringEncoding/ErrorPolicy metadata on NativeExport are
rejected. A future external VBA library can build ordinary named exports with
the pointer/layout features it needs.

The parser, loader and project tests in `crates/oxvba-project` are the source
for exact accepted attributes and diagnostics.
