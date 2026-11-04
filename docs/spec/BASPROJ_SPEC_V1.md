# .basproj project files

`calcvba-project` reads XML project files. This is a project loader, not an
MSBuild engine. Source paths are relative to the project file. For example:

```xml
<Project Sdk="CalcVBA.Sdk/0.1.0">
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

`ItemGroup` accepts Module, ClassModule, DocumentModule and FormModule source
items, ProjectReference, COMReference and NativeReference. VB_* attributes,
including class exposure, creatability, predeclared identity,
name and description, belong in source Attribute lines. Duplicate VB* and
Description module metadata in XML is refused. Instancing and ProgId describe
COM packaging. Document items may supply HostDocumentType for custom hosts;
Office document identities can be recovered from the source VB_Base attribute.

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
  <StringEncoding>Bstr</StringEncoding>
  <ErrorPolicy>Zero</ErrorPolicy>
</NativeExport>
```

`Include` is the exported name. Module and Procedure are required, matched with
VBA name rules. CallingConvention accepts Stdcall (default) or Cdecl; Windows
x64 uses the same convention for both. Ordinal is optional, nonzero, and unique.
Export names are case-sensitive and unique. The procedure must be Public in a
standard module. Its actual signature comes from binding, including ByVal/ByRef
and exact types; cached declaration strings do not supply the ABI.

The native implementation admits ByVal/ByRef Byte, Integer, Long, LongLong/LongPtr,
Single, Double, Currency, Date and Boolean Functions and Subs
with numeric locals, direct calls to admitted procedures, and Double Abs/Sqr
library calls. CLng/CDbl/CBool specialize numeric arguments without Variant boxing.
Long integer division/modulo and Boolean/Long And/Or/Xor/Eqv/Imp/Not are admitted.
Native signatures support up to 32 parameters and calls may cross standard modules.
Native call depth is limited to 128 frames; exhaustion follows the zero-on-error
return policy. Generated code uses baseline x64 instructions. Unsupported
reachable signatures and instructions fail before the DLL is written.
BSTR String, UDT, SAFEARRAY, Variant and Object boundary cases are implemented,
with module/Static state initialized by the owning thread on its first call.
StringEncoding accepts Bstr (default), Utf16 and Ansi for ByVal String inputs;
String results remain caller-owned BSTRs. ErrorPolicy accepts Zero (default) or
HResult (HRESULT return, with a trailing result pointer for Functions).
CalcVbaGetLastError is a reserved export returning the error number and optional
BSTR source/description outputs; callers free those BSTRs with SysFreeString.
CalcVbaShutdown is also reserved. Call it on the owning thread before unloading
the DLL to release module-held strings, arrays, Variants and COM references.
It returns a VBA status number and closes the context permanently; later VBA
export calls return error 5. It does not run cleanup under the loader lock.
The DLL contains no interpreter fallback. Some owning
constructs remain unfinished. See [the native template design](NATIVE_EXPORT_PE_TEMPLATE_V1.md).

CalcVBA has no XLL target, registration metadata, conversions or host support.
Addin is rejected as an OutputType. Category, Description, ArgumentDescriptions
metadata on NativeExport are
rejected, including empty elements and unsupported NativeExport attributes.
A future external VBA library can build ordinary named exports with
the pointer/layout features it needs.

The parser, loader and project tests in `crates/calcvba-project` are the source
for exact accepted attributes and diagnostics.

## Office project folders

The SDK name is `CalcVBA.Sdk`. A `.basproj` is valid MSBuild XML, with
PropertyGroup and ItemGroup elements, but CalcVBA does not execute MSBuild
targets, conditions, property expansion or SDK tasks. A real MSBuild SDK can
be added later without changing the file extension.

Module items list `.bas` standard modules, `.cls` class or document modules,
and `.frm` UserForms with adjacent `.frx` resources. Their item names are
Module, ClassModule, DocumentModule and FormModule. Files are UTF-8 without a
BOM. Source headers retain VBE Attribute lines, including member attributes,
VB_Name, VB_Base and all VB_* flags. VBE class and form preambles are retained.
VBExposed, VBCreatable, VBPredeclaredId and VBGlobalNamespace XML metadata is
refused: edit the source Attribute instead. StreamName, DocumentVersion,
ReadOnly and Private describe persistence and packaging, not duplicate VB_*
attributes. Document identity comes from the source VB_Base, not a filename.

An exported Office folder records Codepage, SysKind, CompatibilityVersion,
Lcid, LcidInvoke, ProjectName, ProjectDescription, HelpFile, HelpContextId,
LibFlags, VersionMajor, VersionMinor and DefineConstants. DefineConstants uses
semicolon separators in XML and colon separators in MS-OVBA. Office project
arguments have unique VBA names and signed 16-bit numeric values; saving
normalizes Boolean/hex numeric spellings and refuses string values or overflow.
Source-level `#Const` expressions are a separate VBA language feature. Files remain UTF-8
when Codepage is a legacy Windows or Macintosh encoding. Workbook write-back
converts all source and metadata to that encoding and refuses any character
that cannot round-trip exactly. A conversion failure writes no output workbook.

COMReference items keep the ordered persisted Libid, plus compiler reference
identity fields. Control references retain OriginalLibid, TwiddledLibid,
ExtendedName, ExtendedLibid, ControlGuid and Cookie. Intrinsic=true marks the
Office host library, which is not an extra dir reference. VBA is a compiler
intrinsic. VbaProjectReference retains AbsoluteLibid, RelativeLibid and version;
compiling that external project still requires a supplied project graph through
`Project::from_folder(...).to_manifest()`. External workbook locations are not
automatically converted into local `.basproj` dependencies.
OfficeProjectProperty items retain the original project ID, unlocked protection
records, host extenders, workspace preferences and other PROJECT properties.
The Include value identifies `section|key`; Value holds the original value.

The round trip is workbook -> new `.basproj` folder -> changed project -> new
workbook. The workbook supplies its existing sheets, cells, host documents and
non-VBA parts; the folder supplies the complete module roster, sources,
references and project properties. Document additions/removals or VB_Base
changes must be accompanied by corresponding host-document changes; this
source operation does not create or rename worksheet objects. UserForms retain
VBE `.frm`/`.frx` persistence; this does not add a forms runtime.

The portable reader accepts `.xlsm`, `.xlsb`, `.xlam`, `.xls` and standalone
vbaProject.bin. Persistence follows [MS-OVBA](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc).
ZIP-based files follow OPC relationships; legacy workbooks use
their compound-file project storage. Locked projects and encrypted Office packages
are refused. Signed-project write-back is currently refused. Saving rebuilds
the VBA storage, sets every source offset to zero, removes module p-code
prefixes and cache streams, and invalidates `_VBA_PROJECT` so Office recompiles.
An observed Excel source-only loading quirk requires UTF-16 help-path copies
in dir when caches are invalidated. Native cached files use the specified MBCS
copies. The reader accepts both and checks their value against PROJECT.
Quoted PROJECT descriptions follow the native Excel save format.
Unrelated ZIP parts and legacy workbook streams are preserved. Save-as refuses
an existing destination. It never changes Office trust or security settings.

Instancing accepts Private (1), PublicNotCreatable (2), SingleUse (3),
GlobalSingleUse (4), MultiUse (5) and GlobalMultiUse (6). SingleUse and
GlobalSingleUse require OutputType=ComExe on Windows: an out-of-process COM
server. ComServer denotes an in-process COM server. A standalone EXE uses
OutputType=Exe with a Public Sub Main in a standard `.bas` module. New
ComServer, Exe and ComExe output work follows the Office-file work; parsing an
output type does not claim its implementation is complete.
