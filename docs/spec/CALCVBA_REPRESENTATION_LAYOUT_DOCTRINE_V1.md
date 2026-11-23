# Runtime value layout

The runtime preserves VBA/Automation carrier layout and ownership where
implemented. Layout tests live in `calcvba-runtime` and native interop tests in
`calcvba-com`; broader carrier families remain incomplete.

## Carrier Requirements

### BSTR

`BStr` is an owned BSTR-shaped UTF-16 allocation. `StrPtr` exposes the BSTR data
pointer for addressable string storage or for a bounded temporary when the input
is not addressable.

### VARIANT

`Variant` is the canonical execution cell and must remain layout-compatible with
Windows `VARIANT`: 24 bytes on 64-bit targets and 16 where pointers are 32-bit, with the
`BRECORD` union arm and the `DECIMAL` overlay of the header,
VARENUM tags, payload ownership, `VARIANT_BOOL`, `CY`, `DATE`, object, BSTR, and
SAFEARRAY pointer payloads.

APIs that expose a variant cell pointer must point at this real cell. Any helper
that serializes or clones bytes must not imply a second non-native wire format.

### SAFEARRAY

`SafeArray` is the runtime array carrier and must expose a COM-compatible
descriptor pointer. Declared scalar arrays must allocate typed SAFEARRAY payloads
such as `VT_I2`, `VT_I4`, `VT_R4`, `VT_R8`, `VT_CY`, `VT_DATE`, `VT_BSTR`, and
`VT_BOOL`; only declared `Variant` arrays should use `VT_VARIANT` by default.

The current implementation now preserves typed storage for VM `ReDim` scalar
arrays. Scalar UDT values and UDT SAFEARRAY elements use descriptor-backed
`VbaRecord` payloads with native field offsets for scalar fields, nested records,
and fixed arrays whose bounds are available as literal metadata. Record export to
foreign COM record descriptors and broader record addressability remain
in-progress exact-layout lanes.

### IUnknown

`ObjectRef` must expose a COM-compatible identity pointer whose first field is a
vtable pointer with `QueryInterface`, `AddRef`, and `Release` shape compatible
with the supported runtime interfaces. COM-backed and runtime-backed objects may
have different dispatch implementations, but their object identity carrier must
not be an arbitrary integer token.

### Numeric Primitives

Numeric primitive payloads must match VBA/COM storage:

- `Integer`/`Long`/`LongLong` as signed 16/32/64-bit integers.
- `Byte` as unsigned 8-bit.
- `Single`/`Double` as IEEE 32/64-bit floats.
- `Currency` as scaled signed 64-bit `CY`.
- `Date` as OLE Automation `f64`.
- `Boolean` as `VARIANT_BOOL` (`-1` true, `0` false) where stored in VARIANT or
  SAFEARRAY carriers.

Semantic coercion and overflow behavior still belongs to arithmetic/binder/VM
logic; it must not change the physical layout of an already stored carrier.
