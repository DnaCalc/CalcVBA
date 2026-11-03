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
`OutputType` accepts HostModule, Library, Exe, Addin, ComServer or ComExe.
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
NativeExport metadata can be parsed, but genuine native per-program DLL/EXE
emission remains unfinished. Parsed metadata alone does not establish that an
output target or all of its signatures can execute.

The parser, loader and project tests in `crates/oxvba-project` are the source
for exact accepted attributes and diagnostics.
