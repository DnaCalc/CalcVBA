# Runtime Variant as the real Automation VARIANT: design and plan (2026-11-13)

Read at CalcVBA `219403627` (master, pulled). Excel facts in section 2 were captured on this
machine (Excel x64, VBA7) with `scripts/run-excel-probe-session.ps1` on a private desktop; the
probe modules are kept under `temp/variant-layout-probe/` and are to become checked-in
oracle cases (section 9). Everything else was read from the code; **V** marks a claim verified
by reading the code, **I** one inferred from it that still needs a test.

## Decisions (owner, 2026-11-13)

1. The runtime cell becomes the platform's real `VARIANT`, and **arrays and records go with
   it**: `BRECORD` with a real `IRecordInfo`, and Automation `SAFEARRAY`s.
2. On Windows, `SAFEARRAY`s are allocated by oleaut32; elsewhere by CalcVBA with the identical
   layout and prefixes.
3. No compatibility with CalcVBA's old internal layout. No benchmarking against it: optimise
   later against the correct types. Borrow-instead-of-copy improvements are made as each
   step lands. A temporary VM3 slowdown is acceptable.
4. `ProcRef` (0xFFF0) stays for now as an internal tag (section 3.4), to be revisited. If
   compatibility problems show up, generate real callback thunks so `AddressOf` yields a code
   address as in VBA.
5. Close the `VarPtr` gap now: `VarPtr`, `StrPtr` and ByRef `Declare` arguments use real
   addresses, and writes through them are seen by later reads, in VM3 and the JIT alike.
6. Prepare the shared value crate as part of this work (section 7). CalcDb does not need it
   yet.
7. Excel/VBA oracle probing on this machine is allowed.

The goal is true VBA compatibility, with VM3 and the JIT behaving identically.

## 1. History: why the cell is 16 bytes

| When | Commit / bead | What happened |
|---|---|---|
| 2026-02-26 | `11ac22592` | `PLAN.md:200` proposed a 24-byte `repr(C)` union "for ABI compatibility with COM VARIANT" (the 24 came from a 16-byte Decimal payload, not `BRECORD`). `BRAINSTORM.md:254-260` argued for a non-COM 16-byte cell, four per cache line. |
| 2026-02-26 | `7e0a794f6` | First `variant.rs`: `repr(C, align(16)) { vtype, payload: [u8; 14] }`, the cache-line design. |
| 2026-02-27 | `f944c62d4` | "Align Variant plan and runtime to COM VARIANT layout": today's `vt + reserved1..3 + data`. It removed the cache-line rationale; the text it removed already claimed the Windows `VARIANT` "is 16 bytes". |
| 2026-04-20 | `61dd72a46`, `bd-t8rr.2.3` | The x64 VARIANT/SAFEARRAY fact pack (later deleted) "confirmed" 16 bytes with a probe "derived from the documented field layout"; judging from its wording it left out the `BRECORD` arm. |
| 2026-06-22 | `eaf574ba8` | `docs/spec/OXVBA_REPRESENTATION_LAYOUT_DOCTRINE_V1.md:17-20` states the 16-byte overlay as fact. |
| 2026-07-11 | `a62f773a1`, `bd-59co.2.2.20/.21` | "Exact 16-byte VBA/COM-shaped Variant core" (`variant.rs:178`); the Miri bead treats 16 as given. |
| 2026-11-04 | `736c79301`, `3f034effe` | First correct 24 in the code (`NATIVE_EXPORT_PE_TEMPLATE_V1.md:124`, `vba_record.rs:1792`), worked around rather than fixed. |

No performance, Miri, 32-bit, record or "simulated x64 carrier" reason was recorded for 16
bytes on x64. The one performance argument belonged to the earlier non-COM cell and was
withdrawn by the commit that adopted the COM layout.

## 2. What Excel shows (x64, captured 2026-11-13)

**Cells and strides.** `LenB(LongPtr)` = 8. `VarPtr(a(1)) - VarPtr(a(0))`: Variant 24, Long 4,
String 8, a UDT `{Long, Variant, Integer}` 40.

**Cell bytes** (offset 0 first):

| Value | Bytes 0-23 |
|---|---|
| `CLng(5)` | `0300 0000 0000 0000` `05000000 00000000` `00000000 00000000` |
| `1.5#` | `0500 …` `000000000000F83F` `0000000000000000` |
| `CDec("12345678901234567890.5")` | `0E00 0100 06000000` `396C362F819F4EB1` `0000000000000000`: `wReserved`=vt, scale 1 at 2, sign at 3, `Hi32` at 4, `Lo64` at 8 |
| `True` | vt `&HB`, payload `FFFF` |
| `CVErr(2042)` | vt `&HA`, `scode` `&H800A07FA` |
| Missing `Optional x As Variant` | vt `&HA`, `scode` `&H80020004`, `IsMissing` True |
| `Nothing` | vt `&H9`, null payload; a `Collection` is also `&H9` |

VBA does not clear unused payload bytes: after `v = CDec(...)` then `v = CLng(1)`, bytes 12-15
still held old data. CalcVBA zeroing them is unobservable except through `CopyMemory`, and need
not copy the junk.

**Arrays.** The `SAFEARRAY*` is at offset 8. `Array(1, 2)` and `Dim a() As Variant` are
`vt &H200C`, `fFeatures &H880` (`FADF_HAVEVARTYPE | FADF_VARIANT`), `cbElements` 24, and the
`vt` prefix at `psa - 4` is `&HC`. A `Long()` is `&H2003`, `&H80`, 4, prefix 3. A `String()` is
`&H2008`, `&H180` (`FADF_BSTR`), 8. An unallocated `Long()` in a Variant is `&H2003` with a null
pointer.

**Copies.** `CopyMemory` of 24 bytes between Variants round-trips Double, Decimal and String
(the copy shares the BSTR); 16 bytes is enough for Decimal. Writes through `VarPtr` reach the
variable: a Long, a Variant's payload, a Variant array element, a UDT field. `StrPtr(s) - 4`
holds `LenB(s)`.

**UDTs.** `{Long n; Double d}`: `Len` 12, `LenB` 16. `{Long a; Variant v; Integer b}`: `Len` 30,
`LenB` 40, `v` at 8, `b` at 32. A nested `{Byte x; TWithVar t; Variant arr(1 To 2); String s;
String * 3 f}`: `Len` 90, `LenB` 112, `t` at 8, `arr(1)` at 48, `arr(2)` at 72, `s` at 96.
`VarPtr(fixedString)` gives the address of a temporary, not of the field.

**Records in Variants.** A `Type` from a standard module cannot be put in a Variant at all:
"Only user-defined types defined in public object modules can be coerced to or from a variant
or passed to late-bound functions" (compile error). In Excel VBA, `VT_RECORD` arises only from
type-library records (and public UDTs of referenced libraries).

**ByRef Variant parameters** (`Sub P(v As Variant)`):

| Argument | Callee's `v` |
|---|---|
| A Variant variable or Variant array element | The caller's own cell (`VarPtr` equal) |
| A typed variable, typed array element or UDT field (Long, String, Double, Boolean) | A `VT_BYREF \| vt` cell whose pointer is the storage itself (`&H4003`, `&H4008`, `&H4005`, `&H400B`) |
| A typed array | `VT_BYREF \| VT_ARRAY \| vt` (`&H6003`) |
| An `As Object` variable | `&H4009` pointing at the variable |
| An `As Collection` variable | `&H4009` pointing at a temporary holding the same pointer; `Set` through it is copied back |
| An expression `n + 0`, or `(n)` | A plain copy (`&H3`) |
| The same cell passed on to a nested call | The same cell |

Assignment through such a reference coerces to the target's type: `"12"` to a Long gives 12,
`2.5` gives 2 and `3.5` gives 4 (banker's rounding), `"abc"` raises 13, `Null` 94 and `1E20` 6,
each leaving the target unchanged. `TypeName`/`VarType` see through the reference (`Long`, 3).

So VBA's own runtime cells hold `VT_BYREF` values. That is the model CalcVBA adopts for ByRef
binding, which is also what COM receives, with no conversion.

## 3. Target layout

| Offset (64-bit) | Offset (32-bit/wasm32) | Field |
|---|---|---|
| 0 | 0 | `vt: u16` (for `VT_DECIMAL`, `DECIMAL.wReserved`) |
| 2, 4, 6 | 2, 4, 6 | `wReserved1..3` (for `VT_DECIMAL`: `scale`, `sign`, `Hi32`) |
| 8 | 8 | union: `llVal`, `dblVal`, `cyVal`, `date`, `bstrVal`, `punkVal`, `pdispVal`, `parray`, `byref`, `scode`, `boolVal`, …; `DECIMAL.Lo64` |
| 8, 16 | 8, 12 | `BRECORD { pvRecord, pRecInfo }` |
| size 24, align 8 | size 16, align 8 | |

### 3.1 The cell definition

| Option | For | Against |
|---|---|---|
| A. Widen the sealed byte carrier to `[u8; 2 * size_of::<usize>()]` | Smallest diff | Pointers stay exposed-provenance `u64`s (`variant.rs:344-357`); 32-bit stays a fiction. |
| B. A real `#[repr(C)]` union with typed fields, and an outer union with `DECIMAL` | Native-width pointers with provenance; 32-bit and wasm32 correct by construction; field-for-field the platform header; Miri strict provenance works | Larger diff in `variant.rs` |
| C. `windows_sys::VARIANT` on Windows, own type elsewhere | No layout of our own on Windows | Two truths; Drop/Clone would differ by platform |

**Choice: B**, one portable definition, with compile-time asserts on Windows that size and every
field offset equal `windows_sys::Win32::System::Variant::VARIANT`.

*As built (step 4a):* the cell has the platform size and layout, but the payload stays the sealed
byte carrier: the existing eight bytes plus a zeroed union tail (`BRECORD.pRecInfo` on 64-bit,
empty on 32-bit), with pointers stored at native width in the low bytes. That gives the exact
layout and correct 32-bit cells with a small diff. The typed union (B's provenance benefit) is
done when the value types move into `calcvba-automation`, after records and arrays (steps 5-6)
have rewritten the pointer payloads. `Variant` stays
`repr(transparent)` over it. Addresses that VBA code sees as integers (`VarPtr`, `StrPtr`,
`ObjPtr`) use `expose_provenance` / `with_exposed_provenance` at that boundary only.

### 3.2 Records

`VT_RECORD` holds `pvRecord` (the record data in Automation layout) and `pRecInfo` (a real
`IRecordInfo`). Today's boxed `RecordPayload::{Com, Vba}` (`variant.rs:372-404`) and the
two-layout split (`vba_record.rs:514-523` vs `:1788-1843`; Variant width is the only
difference) are deleted.

- Typelib records keep the OS `IRecordInfo` (`GetRecordInfoFromGuids`,
  `windows_variant.rs:1687`).
- VBA `Type`s get CalcVBA's `IRecordInfo`, which exists on Windows today
  (`automation_record.rs:111-413`) and becomes portable, built from `VbaRecordLayout`.
- A VBA `Type` value (a local, a field, an array element) is its Automation-layout bytes in
  place, exactly as VBA stores it. Excel never puts a standard-module `Type` in a Variant
  (section 2), so the `IRecordInfo` for VBA `Type`s is needed for typed `SAFEARRAY`s of UDTs, for
  `Declare` and COM parameters typed as the UDT, and for public UDTs of referenced CalcVBA
  projects. Assigning a standard-module `Type` to a Variant must be a compile error, as in
  VBA.

### 3.3 Off Windows: the portable COM stand-in

`IUnknown`, `IDispatch` and `IRecordInfo` are CalcVBA's vtables with the Windows layout. First,
`RawRuntimeIUnknownVtbl` (`object_ref.rs:725-731`) changes from `extern "C"` with
`iid: RuntimeGuid` by value to `extern "system"` with `*const GUID` (V); `HostObject` uses the
same table (`host-api/src/object.rs:532-558`, V). Then a project object can sit in `punkVal`
and be handed to foreign code.

### 3.4 Tags

| Tag | Today (V) | Target |
|---|---|---|
| Automation scalars, `VT_BSTR`, `VT_DISPATCH`, `VT_ERROR`, `VT_DECIMAL` | Real values | Unchanged |
| `VT_UNKNOWN` (0x0D) | Decodes to `Object` (`variant.rs:50`) | Distinct |
| `VT_RECORD` | Boxed CalcVBA payload | 3.2 |
| `ArrayVariant` 0x200C | The only array tag; element type in the `SAFEARRAY` or in `reserved1` (`variant.rs:439-452`) | `VT_ARRAY \| vt`; unallocated typed arrays are `VT_ARRAY \| vt` with null `parray` |
| `VT_BYREF \| vt` | Never in runtime cells | **In runtime cells**, as in VBA (section 2): ByRef binding of a typed variable, element or field to a Variant parameter. Reads and writes go through; writes coerce to the target type. |
| `ProcRef` 0xFFF0 | AddressOf procedure index (`variant.rs:615-627`) | Stays internal for now (decision 4): never in an address-exposed cell, an array or a record; every boundary refuses it with a VBA error. To revisit; real thunks if compatibility needs them. |

## 4. Address-of and storage stability

Today every `VarPtr`/`StrPtr` result and every ByRef `Declare` argument is a copy in a
side buffer (a pin in a process registry, or a boxed cell in the HAL), with write-back only
to simple variables and only when the pointer expression is written directly in the call (V,
`pointer_helpers.rs:1534-1581`, `vm3/src/lib.rs:1818-1834, 4952-5119`, `jit/helpers.rs:4660-4956`).
The obstacle is not moving storage: VM3 and the generic JIT read every operand as a deep clone
(`vm3/src/lib.rs:5354-5374`), so the helpers never see the slot.

**One address-of primitive, shared by both backends.**
`OxInst::AddrOf { dst, place_path, kind }` (base place + field chain + element indices),
lowered from `VarPtr`, `StrPtr`, `ObjPtr` and every ByRef native argument, and backed by one
runtime function `place_address(slot: *mut VARIANT, path) -> *mut u8`:

| Kind | Address |
|---|---|
| Variant | the cell |
| Typed scalar | `&cell.payload` (offset 8) |
| String | `&cell.bstrVal`; `StrPtr` reads `cell.bstrVal` |
| UDT | `pvRecord` + Automation field offset per field |
| Array | `&cell.parray` |
| Element | `pvData + linear index × cbElements`, bounds-checked (error 9) |

VM3 calls it after `resolve`, the JIT after `slot_mut`, so both produce the same address for
globals (one allocation) and stable addresses for locals.

**Stable storage** (V for today's shape):

| Storage | VM3 today | JIT today | Change |
|---|---|---|---|
| Locals | `Frame.locals: Vec<Variant>`, never resized | `JitFrame.locals: Vec<Variant>` | `Box<[UnsafeCell<Variant>]>`; never hold `&`/`&mut` across a native call |
| ByRef params | alias to caller `Loc` | `SlotAlias` | `VT_BYREF` cells (section 2) for typed targets; the caller's cell for Variant targets |
| Temporaries | `HashMap<usize, Variant>`, moves on rehash | `Vec<Variant>` | VM3 to `Vec`, sized from `func.temps`; released at `StmtBoundary` (VBA's statement lifetime) |
| Globals | `programs[p].globals: Vec<Variant>` | same store | Never resized after load |
| Class fields | `RefCell<BTreeMap<i32, Variant>>`; assigning Empty removes the entry | same | `Box<[Variant]>` by field ordinal, Empty kept |
| Array elements | runtime `SafeArray` | same | oleaut-owned `pvData`; addresses change only on `ReDim`, as in VBA |
| UDT fields | canonical layout ≠ Automation | same | Automation layout in place (3.2) |
| Typed-tier slots | `RawActivation.slots: [ScalarBits; 128]` | `PhysicalSlots::Stack` | Functions using `AddrOf` stay on the generic tier at first; then address-taken slots become real memory in the typed tiers |

The pin registry, `lookup_pointer`, `free_pins`, `PtrWritebackKind`, `DeclarePtrWriteback` and
the binder's write-back detection (`calcvba-bind/src/call.rs:1019-1064`) are deleted.
`TrackedBstrAccountingToken` and the live BSTR counters need a redesign, because native code
may now free or replace runtime BSTRs legitimately.

## 5. The interop sweep: optimal approach at each point

General rules, now that the layouts agree:

- **Borrow** a runtime cell for a ByVal `[in]` argument (bit-copy into `rgvarg`/the call frame,
  never cleared by the callee).
- **Point** at the runtime storage for ByRef: `VT_BYREF | VT_VARIANT` to the cell, or
  `VT_BYREF | vt` to the payload of a typed variable, as VBA does.
- **Adopt** results by move into the destination cell: no copy and clear.
- **Copy** only where VBA does: parenthesised or expression arguments, ByVal parameters the
  callee assigns, ANSI conversions for `Declare`, coercion to a declared type, and QI to a
  declared interface.
- One marshalling layer: VARIANT to VARIANT. `ComValue`, `DynamicValue`, `ComInvokeValue` and
  the legacy `i32` token transport are deleted (`calcvba-com/src/model.rs`,
  `dynamic_object.rs`, `windows_invoke.rs:867, 978, 1072`, `compat.rs`). This also fixes lost
  subtypes: today Integer and Byte go out as `VT_I4` and come back as Long, and Single comes
  back as Double (`model.rs:103-216`, `windows_variant.rs:2066-2079, 2333`,
  `windows_vtable.rs:1154, 1456`, V).

### 5.1 COM client (VBA calls out)

CalcVBA has two client stacks (V): stack A (VM3 and the generic JIT, through `ComHal` into
`windows_bridge.rs`, `windows_invoke.rs`, `windows_vtable.rs`, `windows_variant.rs`, with
foreign objects as `i32` handles into a bindings map) and stack B (the typed host stack,
`calcvba_host_api::HostObject` into `typed_slots/*`), which already borrows argument views,
uses real `VT_BYREF` views, calls `pfnDeferredFillIn`, and enumerates lazily. **Stack B's model
becomes the one marshaller.**

| Point | Today (V) | Optimal | Copies that stay | VM3/JIT |
|---|---|---|---|---|
| Late-bound `Invoke` (`vm3/lib.rs:4758`, `jit/helpers.rs:4508-4653`, `windows_bridge.rs:922-1147`) | Each argument read as a copy, then `DynamicValue` → `ComValue` → fresh VARIANT; BSTRs reallocated, arrays rebuilt per element; `GetIDsOfNames` per call; retry with `PROPERTYGET` re-marshals; ByRef supports Long only (`windows_invoke.rs:727-741`); passing `Nothing` reaches an `.expect()` (`dynamic_object.rs:93`, `model.rs:286`) | Borrowed `rgvarg` for ByVal; `VT_BYREF \| VT_VARIANT` to the cell or `VT_BYREF \| vt` to typed storage for ByRef; result adopted into the destination; DISPID cache per object and name; one `Invoke` with `METHOD \| PROPERTYGET` | Expression/parenthesised arguments; fixed-length strings | Shared HAL today; the JIT write-back skips VM3's For-Each array lock check (`helpers.rs:4644` vs `lib.rs:5377-5388`), and the JIT refuses constant arguments (`lowering.rs:2645`). Both go away with pointer-into-slot. |
| Early-bound vtable (`windows_vtable.rs:316-535`, `windows_invoke.rs:3222`) | Strings go through `to_string` → `SysAllocString`, truncating at NUL and losing lone surrogates (`:689-708`, `:885-907`); a boxed cell per ByRef; ByRef writes discarded on failure HRESULT (`:502-504`); `[lcid]` synthesised as 0 | `[in] BSTR` = the runtime BSTR; ByRef = `&cell.payload`; retval written into the destination cell's payload, then `vt` set; ByRef record = `pvRecord` of the variable | QI to the declared IID; Variant-content coercion to a declared type | Shared helper; stack B's typed path is a third marshaller, folded in |
| Property Let/Set, default members, CallByName (`windows_invoke.rs:1252-1571`, `vm3/lib.rs:2578-2616`, `jit/helpers.rs:4999-5100`) | Arguments silently dropped when `requires_argument` is false (`:1294-1329`) | As above; `Set` passes the IDispatch; excess arguments raise 450 | QI for `Set` when the variable holds another interface | Error construction differs only (5 in both) |
| For Each over COM (`windows_variant.rs:349-431`) | Enumerator drained eagerly into a `SafeArray`; any `VT_UNKNOWN` result that is an `IEnumVARIANT` becomes an array | Lazy: `Next(1, &loopvar_cell, &fetched)`, as stack B | QI for a typed loop variable | VM3 uses `Fault::from_hal`, the JIT error 13 by default (`helpers.rs:7096-7105`): unify |
| Events/WithEvents (`windows_connection_point.rs`, `vm3/lib.rs:2375, 3949-4064`) | One Advise per handled DISPID; handlers queued and run after the source's `Invoke` returns; ByRef ignored except one Boolean; **the JIT never subscribes** (`helpers.rs:3291-3350`) | One sink per WithEvents variable covering all DISPIDs; handler runs synchronously inside `Invoke`; ByRef parameters are views on `rgvarg` | Coercion to the handler's declared type | **Divergent: fix in the JIT** |
| CreateObject/GetObject/New/`Is`/`ObjPtr` (`com.rs:247-395`, `windows_client.rs:245-343`) | Objects are `i32` handles; bindings never released before state reset (I); `GetObject` ignores `class` | Interface pointer in `punkVal`/`pdispVal`; `Is` compares QI(IUnknown); `Set x = Nothing` releases; `New` uses the coclass CLSID | None | Shared |
| Typelib records (`windows_variant.rs:145-317, 1346-1684`) | `RecordCreateCopy` in, `deep_clone` out, `PutField` per field; 1-D record arrays lose their LBound | `pvRecord` borrowed; record `SAFEARRAY` passed as `&cell.parray` | `RecordCopy` for ByVal; `IsMatchingType` failure | Shared |
| Arrays (`windows_variant.rs:1218-1300, 1949-1994`) | Variant arrays rebuilt with `SafeArrayCreateVector(..., 0, n)`, losing LBound; a legacy `VT_I4` token fallback | `cell.parray` (ByVal view) or `&cell.parray` (ByRef); the callee's `SafeArrayRedim`/`Destroy` is legitimate | Locked arrays raise 10 (via `fLocks`) | Shared |
| Project objects passed to COM (`windows_runtime_object.rs:33`) | A wrapper whose `Invoke` returns `DISP_E_MEMBERNOTFOUND`, cached until reset, delaying `Class_Terminate` | The object's own COM-correct IDispatch (3.3) | None | Shared |

### 5.2 Declare and pointer helpers

| Point | Today (V) | Optimal | Copies that stay |
|---|---|---|---|
| ByVal scalar (`dynlink.rs:1217-1245`) | `variant_i32` truncates Double with `as i32` | The coerced payload read from the cell | None |
| ByRef scalar (`dynlink.rs:824-956`) | `Box<T>` copy, written back afterwards | `&cell.payload` | None; a typed mismatch is a compile error in VBA |
| ByVal/ByRef String (`dynlink.rs:878-899, 1246-1318`) | ByVal: a plain byte buffer with no BSTR prefix; `vbNullString` becomes `""`; conversion via lossy `as_str` | An ANSI BSTR temporary (NULL for a null BSTR); after the call, convert back and replace `cell.bstrVal` in place; ByRef passes `&temp` and adopts what the callee left | The ANSI conversion and copy-back (VBA semantics) |
| ByRef Variant (`dynlink.rs:1116-1137`) | `ComValue` → boxed VARIANT; Objects refused | The cell itself | None |
| Arrays (`dynlink.rs:934-999`) | `SafeArrayCreate` + `SafeArrayPutElement` per element, read back per element | `&cell.parray` | String arrays and UDT arrays with strings: ANSI copy (to confirm with Excel) |
| Records / `As Any` (`dynlink.rs:1073-1147`) | ANSI projection for every record | `pvRecord` when the record (recursively) has no strings; real address for scalars and elements | ANSI projection with copy-back when it has strings |
| Array element ByRef | A temporary plus write-back; `VarPtr(a(i))` ignores the index (`calcvba-bind/src/call.rs:1025-1033`) | `pvData + offset` | None |
| ByVal Object / ByVal Variant | Fall back to an `i32` (`dynlink.rs:1266`) | The interface pointer; on x64 a pointer to a VARIANT copy | The VARIANT copy for ByVal Variant |
| Unix FFI (`dynlink.rs:707-790, 1149-1192`) | ByRef refused | The same direct-address scheme | ANSI/UTF-16 conversion as the platform requires |
| `VarPtr`/`StrPtr`/`ObjPtr`/`VarPtrArray` (`pointer_helpers.rs`) | Pins of copies; `ObjPtr` already real | Section 4 | Statement-lifetime temporaries for r-values (`StrPtr("lit")`) |
| `AddressOf` (`callback_thunks.rs`) | Windows x64 only; 32 fixed thunks; only for `CallWindowProc` (`vm3/lib.rs:5167-5177`); **not wired in the JIT** (`helpers.rs:4660-4904`) | JIT parity with VM3 now; real per-procedure thunks when decision 4 is revisited | None |
| Call preparation (`windows_ffi_bridge.rs:149-181`) | A libffi CIF built on every call | One prepared CIF per call site, as AOT does (`native-template/src/foreign.rs:55-140`) | None |

`interop_plan.rs` (identity and verification only, x64 hard-coded at `:277`) gains a per-parameter
marshal class (Direct, AnsiStr, AnsiRecord, ArrayPtr, Callback) and an architecture, and is
built once for VM3, the JIT and AOT; today VM3 and the JIT each build it.

The AOT native lane already passes real slot addresses and real BSTRs (`jit/native_values.rs:1196-1366`,
`native-template/src/foreign.rs:200-264`); it is the model for VM3 and the generic JIT.

### 5.3 Inbound: native exports, COM servers, hosting

| Point | Today (V) | Optimal | Copies that stay | VM3/JIT |
|---|---|---|---|---|
| Native export ABI (`jit/native.rs:690-830, 1718-2130`) | Already real VARIANT cells; ByRef binds the caller's pointer; result moved out | Unchanged; literal 24s become shared constants; assigning to a ByRef Variant holding `VT_BYREF` writes through (today it replaces the reference, `values.rs:109-127`, I) | ByVal Variant/String parameters the body assigns; `VariantCopyInd` of `VT_BYREF` inputs; ANSI/UTF-8 import | JIT only by design |
| Native template runtime services (`native-template/src/services.rs:233-505`, `members.rs:186-330`) | Native cell → runtime `Variant` → `ComValue` → native cell: about three copies per string; arrays via `SafeArrayGetElement` per element | `read` borrows the cell (after `VT_BYREF` deref); `write` moves; arrays are the same oleaut `SAFEARRAY` | ANSI string arrays; coercion | AOT only |
| Project objects returned from exports (`services.rs:444-451`, `classes.rs:12-19`) | A compat pointer with three vtable slots in a `VT_DISPATCH` cell | The object's COM-correct IDispatch (3.3) | None | — |
| COM EXE server (`native-template/src/server.rs:93-256`) | `DispInvoke` over the compiled vtable and type info | Keep | None | AOT only |
| COM DLL server (`calcvba-comhost/src/lib.rs:1060-1300, 3720-3800, 4109-4271`) | `VARIANT` → `ComValue` → `Variant` (subtypes lost); **named arguments dropped** (`:1110-1118`); write-back index wrong with named arguments (I); a new wrapper per returned object, so `Is` fails (I); **VM3 only** (`calcvba-engine/src/engine.rs:766-767`) | `DispInvoke` over a generated vtable and typelib, as the EXE server; ByRef Variant formals bound to `pvarVal`; ByVal borrowed; results moved into `pVarResult`; one wrapper per object; route through `calcvba-host`'s `Session` so the JIT runs it too | Typed byref of a different `vt` (oleaut reports a type mismatch) | **Divergent: admit the JIT** |
| Host/embedding API (`host-api/src/value.rs:97-168, 306+`, `rt-abi/src/host_entry/activation.rs:186-347`, `calcvba-host/src/lib.rs:828-1003`) | `Value` is a declared-type carrier, not a second Variant; ByRef copied in and back; results cloned via `retain_bits` although `take_result` exists | `Value` reuses the shared `BStr`/`SafeArray`/`Variant`; `ValueSlot::Record` becomes `pvRecord` + `IRecordInfo`; ByRef `Variant`/`String` arguments alias the host's cell; results by `take_result` | ByVal arguments the body assigns; typed coercion | Shared frames (`RawActivation`) |
| Wire encoding (`variant.rs:307-341, 482-563`) | `[u8; 16]`; used only by `pointer_helpers.rs` and tests | A validated decoder for tests and debugging at `[u8; size_of::<VARIANT>()]`; `pointer_helpers` copies cells directly | None | — |
| Built-ins (Collection `collection.rs:43-52`, Dictionary as VBA source, `Err`) | Store runtime Variants directly | Follow the type | None | Shared |
| wasm32 | VM3 only; no Variant crosses to JS | 16-byte cell with native-width pointers | None | VM3 only on wasm |

## 6. Defects found by the sweep

These are bugs against VBA independent of the layout, found while reading. Most are fixed by
the steps below; each is filed as a bead so none is lost:

1. A stored `VarPtr` result used twice writes into freed memory: pins are freed at the first
   `Declare` that receives the address (`vm3/lib.rs:5013, 5117`, `jit/helpers.rs:4806-4811`, I).
2. Boxed ByRef/`As Any` cells and pins are as small as the declared type, so
   `CopyMemory arr(0), src, 16` writes past them (`dynlink.rs:825-833, 1073-1107`,
   `pointer_helpers.rs:1274-1285`, I).
3. `VarPtr(a(i))` ignores `i` (`calcvba-bind/src/call.rs:1025-1033`, V).
4. `vbNullString` to `ByVal As String` passes `""`, not NULL (`dynlink.rs:1252-1257`, V).
5. `ByVal` Object and `ByVal` Variant to `Declare` pass an `i32` (`dynlink.rs:1266`, V).
6. `ComValue` loses Integer/Byte/SignedByte/UnsignedInteger and Single subtypes in both
   directions (V).
7. Early-bound `[in] BSTR` is truncated at an embedded NUL (`windows_vtable.rs:689-708`, V).
8. Passing `Nothing` to a late-bound call reaches an `.expect()` (V for the chain).
9. COM objects are never released before state reset (I).
10. Late-bound ByRef supports only `Long` (V).
11. 1-D Variant and record arrays passed to COM lose their LBound (V).
12. The JIT never subscribes WithEvents to COM sources (V).
13. The JIT has no `AddressOf` registration (V).
14. The COM DLL server drops named arguments (V) and runs VM3 only (V).
15. Each returned project object gets a fresh COM wrapper, breaking `Is` in clients (I).
16. Off Windows, `StrPtr` pins a buffer without a length prefix (`pointer_helpers.rs:1522-1526`, V).
17. `BStr::as_str` (lossy UTF-16) on the Declare ANSI path and about 20 library sites (V).
18. The JIT takes 15.2 s on the `udt_nested_arrays` bench fixture, VM3 39 ms (measured).

## 7. The shared value crate: `calcvba-automation`

Moves out of `calcvba-runtime` (dependency evidence from each file's `use` lines, V):

- `VarType` (with `VT_UNKNOWN`, `VT_ARRAY | vt`, `VT_BYREF | vt`), the union cell, the owning
  `Variant` (Drop, Clone, move), and the validated decoder.
- `bstr.rs` in full.
- The `SAFEARRAY` descriptor, bounds, element access, and oleaut or portable allocation,
  including record arrays through `IRecordInfo*`.
- `abi`: `VARIANT_SIZE`, `VARIANT_ALIGN`, `VT_OFFSET`, `PAYLOAD_OFFSET`, `RECINFO_OFFSET`.
- A COM core: `GUID`; `IUnknownVtbl`, `IDispatchVtbl`, `IRecordInfoVtbl` (`extern "system"`,
  `*const GUID`); a ref-counted `UnknownRef`.
- The portable `IRecordInfo`, from `automation_record.rs`, driven by a field-descriptor
  schema.
- `Decimal96`, `CurrencyValue`, `F64Value`.
- The `live_counters` and `owning_boundary` test hooks, behind features.

Stays in `calcvba-runtime`: `CompatObjectBase`, project instances and termination,
`object_owner`, `collection`; `VbaRecordLayout`/`VbaRecord` semantics, which build the shared
`IRecordInfo`; the `ProcRef` extension; VBA array helpers; `coerce`, arithmetic, string
compare, dates; `pointer_helpers`, `call_frame`, `interop_plan`, `builtins`, `callback_thunks`.
`com_record.rs` is deleted. `calcvba-runtime` re-exports the shared types so existing paths
keep working for DnaCalc. Today `Variant`, `SafeArray`, `VbaRecord`, `ObjectRef` and `BStr`
form one dependency cycle; it breaks once records are `BRECORD` and arrays carry `IRecordInfo*`,
because the shared crate then needs only the `IRecordInfo` interface. `Variant` stays `!Send`.

DnaCalc (`calc-vba`, `calc-host-vba`, pinned at `219403627`) needs, when it moves its pin:
`VarType::ArrayVariant` checks (`calc-vba/src/runtime/mod.rs:644, 806`) replaced by an array
predicate, and the new crate at the same revision. Its matches on `VarType` already have
catch-all arms. CalcDb's stage-0 plan (`CalcDb/docs/book/00-stage0-plan.md:13, 159`) already
defines the canonical VARIANT as the platform ABI's.

## 8. Migration order

Tracked in DnaCalc as epic `dc-yqq0`; step *n* is `dc-yqq0.n`. Defect 17's library part is
`dc-yqq0.11`, defect 18 is `dc-cvvw`; every other defect in section 6 is named in the step
bead that fixes it.

Each step builds and passes `./scripts/check.ps1`, and keeps VM3 and the JIT in agreement:
both use one `Variant` type, and the generic JIT stores no Variant offsets of its own
(`lowering.rs` passes descriptors to Rust helpers taking `*mut Variant`, V).

1. **Oracle cases.** Check in the section 2 probes as Excel oracle cases that also run under
   VM3 and the JIT, with today's CalcVBA failures recorded as known differences.
2. **The shared crate and constants.** Create `calcvba-automation` with the `abi` constants
   and the COM core types; the native/AOT JIT (`native.rs:98, 593, 624, 1447, 1487-1517, 1568,
   1908`, `native_values.rs:806, 1316`) and `native-template/src/foreign.rs:127` use them.
3. **COM-correct vtables.** `extern "system"`, `REFIID` by pointer, for every runtime and
   `HostObject` vtable.
4. **The cell.** Platform size (4a), `VT_ARRAY | vt` tags (4b), `VT_UNKNOWN` (4c), the wire
   decoder, size and offset tests, specs corrected. `VT_BYREF` cells moved to step 7, where
   ByRef binding and address-of change together
   (`OXVBA_REPRESENTATION_LAYOUT_DOCTRINE_V1.md:17-20`, `VBA_TYPE_SYSTEM_V1.md`,
   `OXVBA_POINTER_HELPERS_CONTRACT_V1.md`). VM3 hot paths move from clone to borrow as they
   are touched. Miri on `calcvba-runtime`, `calcvba-automation`, `calcvba-jit` and the
   `calcvba-com` mock (`scripts/run-miri.sh`).
5. **Records.** `BRECORD`, Automation layout in place, portable `IRecordInfo`;
   `automation_shape` and `com_record.rs` deleted; the Variant-assignment compile error for
   standard-module `Type`s.
6. **SAFEARRAYs.** oleaut32 allocation on Windows, the standard prefix elsewhere, `IRecordInfo`
   at the `FADF_RECORD` slot; VM3 and the JIT re-read `pvData` after every foreign call.
7. **Address-of and storage** (section 4), with `VT_BYREF` binding for ByRef Variant
   parameters; the pin registry deleted.
8. **One COM marshaller** (5.1), replacing stack A with stack B's model; `ComValue` and the
   legacy transports deleted; JIT WithEvents.
9. **Declare** (5.2): direct addresses, ANSI BSTR temporaries, prepared CIFs, the shared
   interop plan, `AddressOf` in the JIT.
10. **Inbound** (5.3): native template services by borrow and move; comhost on `DispInvoke`
    and on the shared `Session`, so the JIT runs it; host API aliases and `take_result`.

## 9. Tests

**Excel oracle cases** (from section 2, run under Excel, VM3 and the JIT): strides; cell bytes
per subtype; `SAFEARRAY` headers; `CopyMemory` round-trips; writes through `VarPtr`; UDT `Len`,
`LenB` and field offsets; the `StrPtr` prefix; the ByRef Variant parameter table and coercion
errors; the standard-module `Type`-to-Variant compile error.

**To add from the sweep:**
- COM: `VarType` of Integer/Byte/Single/LongLong arguments and results; late-bound ByRef
  String/Double/Variant/array/object; passing `Nothing`; a callee changing a ByRef Variant's
  type; embedded NUL in BSTRs; ByRef writes after a failing HRESULT; the `[lcid]` value; LBound 1
  arrays; callee `ReDim` of a ByRef array; release counts after `Set x = Nothing`; `Is` across
  interfaces; lazy For Each `Next` counts; JIT WithEvents; multi-event sinks.
- Declare: `p = VarPtr(x): CopyMemory ByVal p, …` twice; `CopyMemory a(0), src, 16`;
  `VarPtr(a(3)) = VarPtr(a(0)) + 12`; Variant and String swaps; `vbNullString` arriving as NULL;
  ByVal Object; `ByRef arr()` to `SafeArrayRedim`; String arrays to an ANSI DLL; `AddressOf`
  under the JIT, EnumWindows.
- Inbound: exports returning class instances, checked by an oleaut client; comhost `VarType`
  of `VT_I2`/`VT_UI1` arguments, named arguments, the same object returned twice; host ByRef
  aliasing through callbacks.
- Layout unit tests: on Windows, size and every field offset equal `windows_sys::VARIANT`; on
  32-bit and wasm32, size 16 and `BRECORD` at 8 and 12.

The existing COM, Declare, native-export and comhost suites stay green at every step:
`calcvba-engine/tests/com_matrix_*.rs`, `native_declare_*`, `calcvba-build/tests/native_exports.rs`,
`wrapped_com_server_smoke.rs`, `calcvba-host/tests/typed_*`, `scripts/run-com-*.ps1`.

## 10. Risks

- **R1. Size of the change.** Steps 7-10 rewrite every interop layer. Each step lands with its
  oracle and suite green; no step leaves a boundary converting through a layout that no longer
  exists.
- **R2. Provenance.** Typed pointer fields replace exposed-provenance bytes; addresses visible
  to VBA code use exposed provenance only at `VarPtr`/`StrPtr`/`ObjPtr`. Miri is the check.
- **R3. Address stability.** Slots must not move while foreign code holds an address, and no
  Rust reference may be held across a native call (section 4).
- **R4. Foreign ownership.** Once COM callees and DLLs may free or replace BSTRs and
  `SAFEARRAY`s in runtime cells, live-handle counting and any cached pointer (`pvData`, BSTR)
  must be re-read after every foreign call.
- **R5. Records off Windows.** The portable `IRecordInfo` is new code; a wrong
  `RecordCopy`/`RecordClear` leaks or double-frees.
- **R6. `VT_BYREF` cells.** Every Variant read and write must handle the reference, and
  coercion on write must match VBA (section 2). A missed site reads a pointer as a value.
- **R7. Downstream.** DnaCalc moves its pin together with the new crate (section 7).
