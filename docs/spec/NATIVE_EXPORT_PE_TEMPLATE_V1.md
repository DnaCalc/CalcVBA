# Native exports through a prebuilt PE template, V1

Owner direction (2026-11-03):
- The first implementation is named C exports and `.dll` compilation.
- CalcVBA gets no special `.xll` support (owner, 2026-11-03). An `.xll` is an ordinary CalcVBA
  native-library project, provided as a template, example or wizard; packaging is still open.
- The `XLOPER12` conversions, `xlfRegister` calls and `xlAutoOpen`/`xlAutoFree12` are written as
  regular VBA in a standard VBA library for `.xll` work, made as elegant as possible. That library
  doubles as a test of the compiler's output performance, the language's expressiveness and the
  runtime.
- The user's build needs no Rust toolchain and no linker. CalcVBA's own build infrastructure
  patches export structures and compiled code into a prebuilt template DLL that is designed to be
  patched.
- No single generic entry into the engine and no Excel-DNA-style thunk table. Each export is
  compiled and connected along a direct path: a native-ABI wrapper calls the compiled procedure.

Status: owner-directed design; N1-N5 implementation is tracked by bd-2pntk.
N1 is implemented; N2 currently admits x64 ByVal/ByRef Byte, Integer, Long,
LongLong/LongPtr, Single, Double, Currency, Date and Boolean Functions and Subs,
direct calls across standard modules, numeric CLng/CDbl/CBool conversions and
Double Abs/Sqr helpers. Long integer division/modulo and Boolean/Long bitwise
operators use the shared VM3/JIT numeric lowering. Generated code targets baseline x64, not the build CPU's
optional instructions. Calls have at most 32 parameters and 128 active native
procedure frames; exceeding the depth returns VBA status 28 through the ordinary
zero-on-error wrapper. N3 now admits BSTR strings (and terminated UTF-16/ANSI
ByVal inputs), native UDTs, typed SAFEARRAYs, Automation Variants and COM object
references. Module/Static state initialize once, on the first owning-thread call,
outside DllMain. ErrorPolicy=HResult and CalcVbaGetLastError are implemented.
The interop fixture is called from C, VM3/JIT Declare and one reused Excel;
ValueSuite also executes the same VBA bodies directly through VM3/JIT.

The numeric-width fixture also checks casts, narrow storage, values above Double's
exact integer range, and overflow through all three clients.
Remaining refusals include Optional/ParamArray signatures,
project classes, unsupported native Declare signatures and Variant array receivers.
Diagnostics identify the construct and procedure's source declaration. Native UDTs
include owning Variant fields and fixed strings, with native Automation layout
separate from the runtime's canonical record layout. Record SAFEARRAYs carry
registration-free IRecordInfo owners; ReDim Preserve retains their element owners.
Office VBA's raw UDT array descriptors retain their original descriptor form,
with removed owning fields cleared explicitly during ReDim.
ByRef String() ReDim preserves an incoming array's Unicode or ANSI representation.
For a null incoming array, StringEncoding=Ansi selects VBA Declare's ANSI descriptor;
the default remains Unicode for direct native clients and local arrays.
Call CalcVbaShutdown on the owning thread before FreeLibrary
to release module owners outside loader lock. It closes the context; later calls
return error 5 and cannot initialize it again. Object Is queries canonical IUnknown identity, releasing both
temporary references on every path. Distinct interfaces of one object compare equal.
General scalar/String/Variant/UDT pointers and ordinary native
Declare calls are compiled inside exported bodies. ByRef ANSI String uses an
owning byte BSTR cell, including replacement by the callee. UDT Declare arguments
convert Unicode fields to their ANSI representation; Variant fields retain Unicode.
String, Variant and UDT returns use the Windows x64 return rules. Private/forced
ByRef arguments and ByVal aggregate parameters remain refused.
Each compiled Declare call site retains its library, type conversion plan and
libffi CIF until CalcVbaShutdown; subsequent native calls reuse that preparation.
Unallocated typed UDT arrays retain their record prototype across initialization,
copies and Erase, while projecting a null SAFEARRAY pointer at a native boundary.
Fixed-array UDT fields support checked multidimensional access, including owning
String/Variant, fixed-string and nested record elements. Signed literal bounds
retain their inline layout. The record suite compares C, VM3/JIT Declare and Excel.

N4's ordinary VBA library in `examples/native/xll` converts numeric Variants,
counted Unicode strings, Boolean, errors, missing/nil, arrays and single/multiple
area references. Registration, callbacks and buffer release remain ordinary VBA.
The clients exercise both the native DLL and the same source on VM3/JIT and
Excel VBA. Worksheet cases compare individually in Excel and DNA Calc's XLL
host; the template README identifies host limitations and ownership rules.
The compiler and runtime contain no XLL-specific operation or registration metadata.

## 1. Declaration

- `<NativeExport>` in the `.basproj` is the single declaration. Its selection fields are already parsed by
  `calcvba-project` (model.rs) and specified in `docs/spec/BASPROJ_SPEC_V1.md` §2.5.1.
- Its fields are export name, `Module`, `Procedure`, `CallingConvention` and `Ordinal`.
- Add a `StringEncoding` field per export or per parameter: `Bstr` (default), `Utf16` or `Ansi`.
- Validation:
  - the procedure must be Public, in a standard module, with a fully typed signature;
  - names must be unique;
  - ordinals must not collide.
- The bound signature comes from the binder, never from strings. The project file only selects
  procedures and options.
- New `OutputType=NativeLibrary` produces `<Project>.dll`.
- Update the stale `examples/xll`, `scripts/stage-xll-*` and `BASPROJ_SPEC_V1.md:64` in the
  same change: remove the XLL-specific `OutputType=Addin` claims and the Category/Description
  metadata as CalcVBA features. Registration metadata belongs to the VBA library (§8).

## 2. Native ABI per export

These rules follow the current Declare implementation and scalar ABI tests (the historical
`docs/spec/HAL_DECLARE_ABI_SPEC_V1.md` is retained in git before 78e386cf), so a CalcVBA or
Excel VBA `Declare` of the export round-trips exactly.

**Calling convention**
- x64 has one convention (Microsoft x64).
- Future x86 builds use undecorated `Stdcall` (default) or `Cdecl`. Current Cranelift has no x86-32 backend; x86 needs a separate code generator before a template can be offered.

**ByVal parameters**

| VBA type | Native type |
|---|---|
| `Byte` | `u8` |
| `Integer` | `i16` |
| `Long` | `i32` |
| `LongLong` | `i64` |
| `LongPtr` | `isize` |
| `Single` | `f32` |
| `Double` | `f64` |
| `Currency` | `i64` (scaled by 10^4) |
| `Date` | `f64` |
| `Boolean` | `VARIANT_BOOL` (`i16`, -1/0) |
| `String` | `BSTR`, or `LPWSTR`/`LPSTR` per `StringEncoding` |
| `Object` | `IDispatch*` (typed interfaces as their pointer) |

**ByRef parameters**
- A ByRef scalar becomes a pointer to the scalar, with the Declare alias rules.
- A ByRef UDT becomes a pointer to a struct with VBA layout.
- A ByRef `Variant` becomes `VARIANT*`.
- A ByRef array becomes `SAFEARRAY**`.

**Returns**
- Scalars return in registers.
- `String` returns a `BSTR`; the caller frees it.
- An Automation `VARIANT` result (24 bytes on Windows x64) uses a hidden result pointer as the first argument;
  the callee also returns that pointer in RAX.
- UDT results follow the Windows x64 aggregate rules: eligible plain 1-, 2-, 4-
  and 8-byte aggregates return in RAX; other aggregates use the hidden pointer
  and return it in RAX. Do not apply the hidden-pointer rule to every UDT.
  The fixture covers both a four-byte register result and a larger owning UDT
  with a hidden result pointer. ByVal UDT parameters remain refused.

**Unhandled VBA error in an export**
- The export returns the type's zero value.
- The error (number, source, description) is stored in the owning runtime context. Native exceptions and Rust panics must never cross the foreign boundary; unwind tables alone do not provide exception containment.
- The template exports `CalcVbaGetLastError` to read it, and `Err.LastDllError` semantics stay intact.
- Optional per export: `ErrorPolicy=HResult`, where the export returns an HRESULT and the VBA
  result goes through a trailing out pointer.
- The error accessor has C signature `int32_t CalcVbaGetLastError(BSTR*, BSTR*)`.
  Either output can be null. Non-null outputs receive owning BSTRs; free them
  with SysFreeString. Success clears the number. An explicit Err.Raise preserves
  its source/description; default Source is the project name, as observed in Excel.
  Dynamic source, description, help file and help context arguments are converted
  before retaining independent owners, so activation cleanup cannot invalidate them.
- `int32_t CalcVbaGetLastErrorHelp(BSTR*, int32_t*)` returns the same error number
  and optional help file/context. Free a returned BSTR with SysFreeString. Successful
  exports and CalcVbaShutdown release retained text and reset these fields. Both
  accessors leave output storage untouched on foreign-thread calls.
- `int32_t CalcVbaShutdown(void)` releases owning module globals and returns a VBA
  status number. Repeated shutdown is harmless. Calls from a foreign thread still
  return 1444 and leave all owners intact. No VBA cleanup runs in DllMain.
- VBA Declare's ANSI String() SAFEARRAY form (VT_EMPTY with FADF_BSTR) is converted
  at element access. Ordinary VT_BSTR SAFEARRAYs retain Unicode contents.

## 3. Direct compiled path (AOT)

1. **One lowering, two module backends.**
   - Shared lowering and import declaration are generic over `cranelift_module::Module`.
   - The JIT still owns `JITModule` and retains its existing entry paths.
   - `PeModule` (in `calcvba-aot`) lays functions and data out at chosen section
     offsets and resolves relocations itself. Its x64 backend does not infer
     optional CPU features from the build host.
   - Do not go through `cranelift-object` and a COFF link step.
2. **Whole project compiled native.** For every procedure reachable from the exports:
   - calls between project procedures are direct `rel32` calls;
   - VBA library, runtime, string, Variant and COM operations call runtime helpers in the template
     by RVA, resolved at patch time from the template's helper manifest (§4). That is a direct
     `rel32` in the same image: no import table and no base relocations for code;
   - string literals and constants go to read-only data in the patched section, addressed
     RIP-relative.
3. **Per-export wrapper (Cranelift-generated).**
   - It has the native signature from §2. It enters the runtime: a fast check that the caller is the
     owning thread (§3a); on the first call per process it initialises module state once, outside loader lock.
   - It converts the parameters, calls the procedure body directly with the internal typed
     convention `(ctx, typed args…, out) -> status`, maps the status by error policy, converts the
     result and returns.
   - There is no dispatcher, no token lookup, and no Variant boxing of typed values.
4. **No interpreter fallback in an exported DLL.** If the AOT lowering cannot compile a reachable
   construct, the build fails with a typed refusal naming the construct and source location.
5. **Image metadata.**
   - The `.oxi` (type info, COM binding metadata, source map) goes into the read-only `.cvbimg`
     section.
   - The runtime finds it by section lookup on its own module, not through resource APIs.

## 3a. Threading: none (owner, 2026-11-03)

- CalcVBA does not support multi-threading. Each loaded DLL has one runtime context.
- That context is bound to the thread of the first call. Module state lives there.
- A call from any other thread does nothing. It returns the type's zero value, or `RPC_E_WRONG_THREAD` (0x8001010E)
  under `ErrorPolicy=HResult`, and sets a defined Win32 last-error value on the rejected calling thread. It must not mutate the owning runtime context from that thread.
- Such calls are never marshalled and never run concurrently.
- No locks, no per-thread contexts, and no thread-safety analysis in the compiler. Claiming the owning thread uses an atomic thread ID (or equivalent OS primitive) so simultaneous first calls cannot create a data race. All context reads and writes follow the successful thread check.
- The initial template atomically claims a retained OS thread handle. The entry
  check first verifies that the owner has not terminated, using a zero-timeout
  thread-state query, and then compares the thread ID. It does not rely on an ID
  remaining unique after termination. DllMain closes only that handle on unload.
  A rejected caller gets Win32 ERROR_INVALID_THREAD_ID (1444), without a context write.

## 4. Template DLL (prebuilt per architecture; CalcVBA developer-time build)

**What the template contains**
- A Rust `cdylib` that statically links the runtime crates: runtime, lib, hal, com and rt-abi.
  It does not link the compiler, VM3 or Cranelift.
- The application DllMain performs no VBA work, allocation or image loading. The Rust/CRT startup may initialise its own runtime; do not replace the CRT entry with an ad hoc /ENTRY. DisableThreadLibraryCalls is optional and can fail for a CRT using static TLS.

**Made patchable by a "template finisher" (developer-time, in CalcVBA's own build)**
- Reserve spare section-header slots and `SizeOfHeaders` room in the developer link (large file alignment), then verify them in the finisher. Do not shift linked sections or any existing function/data RVA.
- Require `.reloc` last in the linked template. The patcher preserves its fixups, appends generated sections and moves/rebuilds only the relocation section at the new end. All other template RVAs remain unchanged.
- Developer-only helper exports keep helper code alive through linking. The finisher records their RVAs and removes the public export directory. Only selected wrappers and intentional public runtime APIs are exported in a finished program.
- Section `.cvbhlp` holds the helper manifest: helper name, ABI version and RVA for every runtime
  entry the lowering may call, plus the runtime ABI hash. The patcher refuses a mismatch.
- Build without Control Flow Guard at first. CFG needs the guard function table to be extended
  for generated functions; that is a later milestone.
  MSVC CRT objects can retain the inactive CF_INSTRUMENTED marker with /GUARD:NO.
  It is accepted only with the GuardCF DLL flag clear, no guard function table
  or count, and no other guard flags. Actual CFG-enabled templates are refused.

**Templates ship with CalcVBA**
- `calcvba-native-x64.dll` first. ARM64 and x86 are later work; x86 also needs a code generator that supports it.
- Each is built once by CalcVBA's release build and is never rebuilt on the user's machine.

## 5. Patcher (user build time, pure Rust, new crate `calcvba-pe`)

**Input:** the template, the PeModule layout, the export list and the `.oxi`.

**New sections**
- `.cvbtext` (RX): code.
- `.cvbrdat` (R): constants, export directory, unwind info.
- `.cvbdata` (RW): per-process runtime slots.
- `.cvbimg` (R): the image.

**Steps**
1. Write `IMAGE_EXPORT_DIRECTORY`: names sorted, name-pointer and ordinal tables, address table
   pointing at the wrappers.
2. Set `DataDirectory[EXPORT]`.
3. Preserve template UNWIND_INFO RVAs and merge the template's and generated `RUNTIME_FUNCTION` entries into one sorted
   `.pdata`, so stack unwinding and debuggers can traverse generated frames. VBA errors use explicit status returns, and Rust panics are contained inside runtime helpers. Cranelift emits
   Windows x64 unwind info.
4. Rebuild `.reloc` with any absolute pointers in data (target: none in code). Preserve existing template fixups. Test loading at a nonpreferred base to exercise these fixups.
5. Update `SizeOfImage` and the checksum.
6. Leave the result unsigned; signing is the user's step after the build.
7. Write a map file (export to procedure to RVA) for diagnostics. PDB output is a later milestone.

**Verification in tests**
- The structure is parsed back with an independent reader (the `object` crate or llvm-readobj).
- The DLL loads with `LoadLibrary` and every export resolves with `GetProcAddress`.
- A stack walk through a generated frame works.

## 6. Clients and checks

- Each export is called from three clients and the results compared:
  - a C test harness;
  - CalcVBA itself via `Declare`;
  - Excel's own VBA via `Declare` on this machine (the real-world oracle).
- The build refuses a construct rather than miscompiling it. The AOT path runs the same
  conformance cases as VM3/JIT, and the two must agree.

## 7. Milestones

| Milestone | Scope | Accept when |
|---|---|---|
| N1 | Template DLL plus finisher, and `calcvba-pe` patcher | A patched hand-written stub export (returns 42) with `.pdata`/`.reloc` correct loads, resolves by name and stack-walks |
| N2 | Lowering generic over `Module`; `PeModule`; helper manifest; ByVal numeric/Boolean exports; direct calls between procedures; VBA library helpers | Called from C, CalcVBA `Declare` and Excel VBA `Declare` with identical results |
| N3 | Strings (BSTR/LPWSTR/LPSTR), ByRef scalars, UDTs, SAFEARRAY arrays, Variant, objects; error policy and `CalcVbaGetLastError`; module state; single-thread rule (§3a) | All three clients agree |
| N4 | General pointer/layout/library support exercised by ordinary VBA (§8) | An external VBA library can use N1-N3 exports and pointer/Declare operations; CalcVBA has no XLL target, intrinsic or registration metadata |
| N5 | CFG guard-table patching; PDB; ARM64; x86 only with a supporting code generator; optionally AOT COM server | Separate templates and ABI tests for each admitted architecture |

## 8. `.xll` as an ordinary VBA project (not a CalcVBA feature)

**What the add-in author writes**
- `xlAutoOpen`, `xlAutoClose`, `xlAutoFree12` and `xlAddInManagerInfo12` are plain Public VBA
  functions, exported by name through `<NativeExport>` like any other export.
- Worksheet functions are plain exports. Simple ones need no `XLOPER12` at all: a
  `Function Add(ByVal a As Double, ByVal b As Double) As Double` export registers with type text
  `BBB`. Functions that need Excel's general values take and return `XLOPER12*` as `LongPtr`.

**The standard VBA library for `.xll` work (VBA source; shipped as a template/example, not in
CalcVBA's crates)**
- `XLOPER12` as a VBA UDT with typed accessors.
- Conversions between `XLOPER12` and Variant/arrays: strings, `xltypeMulti`, errors, missing/nil,
  references.
- Callback to Excel: `MdCallBack12` found with `GetProcAddress(GetModuleHandle(0))` and called
  through `DispCallFunc` (oleaut32).
- `xlfRegister`/`xlfUnregister` helpers driven by a registration table written in VBA.
- Return buffers marked `xlbitDLLFree`, released in the author's `xlAutoFree12`.
- `xlfCaller`, `xlCoerce` and `xlGetName` helpers.
- No thread-safe (`$`) registration: CalcVBA is single-threaded (§3a), so functions register for
  Excel's main calculation thread only.

**What CalcVBA must provide for this, and nothing XLL-specific**
- N1-N3 exports, with `LongPtr` and pointer parameters and returns.
- Fast `VarPtr`/`StrPtr`, `Declare` to `RtlMoveMemory` and `DispCallFunc` inside an exported DLL.
- UDT layout matching C, and module state and initialisation under the single-thread rule (§3a).
- Where the library is slow or awkward, fix the compiler or runtime generally (inlining,
  `Declare` call cost, UDT copies), never by adding XLL intrinsics.

## 9. Implementation limits and refusal rules

The owner design describes the intended full runtime. N1 may retain only the
runtime helpers required by its test; N2 grows the manifest as shared lowering
becomes usable. Linking runtime crates does not mean every helper is retained.
The compiler, Cranelift, JIT allocator and VM3 must stay out of the template.
The current JIT includes helper shims and owning-entry machinery coupled to JIT
execution. Extract reusable ABI pieces before admitting those operations in AOT;
never embed the JIT merely to reuse a helper or silently execute an interpreter.

The finisher/patcher accepts only unsigned AMD64 PE32+ DLL or EXE templates with normal
file/section alignment, adequate spare headers, relocation data and no CFG.
Reject malformed bounds, overlapping sections, unsupported relocation kinds,
unresolved helpers, ABI/hash mismatches, duplicate names/ordinals, out-of-range
relative displacements and export targets outside generated executable code.
Patch an in-memory copy and publish the output only after validation succeeds.
The output has no executable overlay and requests no runtime executable allocation.

`OutputType=Exe` selects the EXE template and the project's parameterless Sub Main.
The native builder emits the same procedure bodies and owning wrappers as DLL
output. It fills a unique startup record with checked Main, shutdown and error
entry RVAs. Startup reports unhandled errors and releases state on the owning
thread. EXE builds refuse explicit NativeExport entries, a DLL template, missing
Main or unsupported reachable bodies before publishing an executable.

The helper manifest is versioned and stored in .cvbhlp. It contains the runtime
ABI identifier and sorted (name, signature version, RVA) entries. Changing a
helper signature changes the ABI identifier/version; an import must match the
manifest before its relocation is applied. Relative code references are checked
at build time. Unsupported absolute code relocations are refused in the initial
implementation; absolute data pointers require IMAGE_REL_BASED_DIR64 entries.

N2 admission is explicit: every reachable signature, instruction and call must
be admitted. A numeric boundary does not imply that an arbitrary numeric-looking
body is supported. Refusals identify procedure and instruction/source span when
available. Only procedure-reachable code is emitted. JIT tests remain required.
The scalar Windows Declare ABI in current executable tests supersedes the old
HAL document's obsolete M0/profile descriptions. VBA Boolean uses i16; noncanonical
foreign Boolean payloads remain a separate existing compatibility issue.
A ByVal String in ordinary VBA Declare defaults to ANSI marshaling; callers of a
BSTR export must deliberately supply the BSTR pointer using the appropriate ABI.

The .basproj signature comes from binding, including accessibility, standard
module identity, parameter modes and exact types. Do not trust optional cached
signature fields in NativeExportDescriptor. Apply the bound StringEncoding and
ErrorPolicy to the generated boundary; reject unknown metadata values.

## 10. Primary references and corrected historical claims

- [Microsoft PE format](https://learn.microsoft.com/en-us/windows/win32/debug/pe-format).
- [Microsoft x64 calling convention](https://learn.microsoft.com/en-us/cpp/build/x64-calling-convention).
- [Microsoft x64 unwind tables](https://learn.microsoft.com/en-us/cpp/build/exception-handling-x64).
- [Windows thread handles and IDs](https://learn.microsoft.com/en-us/windows/win32/procthread/thread-handles-and-identifiers)
  and [zero-timeout thread-state queries](https://learn.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-waitforsingleobject).
- Cranelift 0.133.1 Module trait and x64 unwind emitter (the pinned dependency sources).
- .basproj §2.5.1 and HAL Declare ABI spec as retained in git before 78e386cf.
- Owner direction: calcvba-native-export-design.md, 2026-11-03.

The historical AOT note is superseded. Module does not provide a universal
finalize_definitions method (it is backend-specific). GetProcAddress reads the
export directory on demand; the old claim that all exports are resolved before
DllMain is too broad. Our design still writes a complete image before loading.
Appending mapped sections with RIP-relative code does not force a fixed base.
The object crate is used as an independent reader, not assumed to be a complete
PE image linker. Merely writing .pdata does not contain a panic or provide VBA
error handling. No XLL-specific compiler or runtime feature is planned.
