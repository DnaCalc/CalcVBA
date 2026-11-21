# DISPIDs and vtable slots Excel's VBA gives a class (dc-yqq0.21, Windows half)

Measured 2026-11-21 on this machine's Excel (64-bit) on a private desktop. A VBA class instance was
passed to `OxVba.TestEventServer.DescribeDispatch`, which reads the object's own
`IDispatch::GetTypeInfo(0)`, then the interface half of the dual through `GetRefTypeOfImplType(-1)`.
Probe sources: `temp/objid2` (the `Widget` below) and `temp/objid3` (`W2`-`W8`). The raw output is
in the appendix.

## Rules

1. **The type.** `GetTypeInfo` returns `TKIND_DISPATCH` with flags `0x1C2` (`FDUAL | FNONEXTENSIBLE |
   FOLEAUTOMATION | FCANCREATE`), named after the class, with a fresh GUID per class. Its interface
   half (`TKIND_INTERFACE`) has the same name, GUID and member list. `GetRefTypeOfImplType(0)` and
   `GetImplTypeFlags(0)` fail with `TYPE_E_ELEMENTNOTFOUND`, although `cImplTypes` is 1.
2. **Public fields** get `0x40030000 + n`, with n counting every field in source order. A
   `Long` or `String` field has get and let accessors (invkind 2, 4); an `Object` field has get, let
   and set, and its let is `FUNCFLAG_FRESTRICTED` (0x1); a `Variant` field has get, let and set.
3. **Properties** get `0x68030000 + n`, one number per property name. Properties are numbered first,
   in reverse source order: in `W4`, `P1`, `P2` and `P3` are 2, 1 and 0.
4. **Other procedures** (Function and Sub) get `0x60030000 + n`. n continues after the properties, in
   source order. `Private` and `Friend` procedures take a number too, and don't appear in the type
   info (`W5`: `S` Private is 0, `F` Friend is 1, `A` is 2).
5. **Special members.** `Class_Initialize`, `Class_Terminate` and events leave the procedure DISPID
   counter unchanged. `VB_UserMemId = 0` gives the procedure or all accessors of the property
   DISPID 0 and leaves the ordinary counter unchanged. The property keeps its declaration
   position in the vtable.
6. **Vtable order** (`oVft`, 8 bytes per slot, the first at 56 after IDispatch's seven): fields first
   (each accessor its own slot), then members in source order, a property's accessors in a row.
   `Private` and `Friend` procedures take hidden slots after the public ones (`cbSizeVft` 80 for `W5`,
   one public member). `Widget` has two such procedures and a third hidden slot for its nonempty
   `Class_Terminate`. The empty `Class_Terminate` in `W3` adds no slot.
7. **Signatures.** The dispatch half reports VBA's return type (`ret` = VT code); the interface half
   returns `HRESULT` (25) with an added retval parameter for functions and property gets. `Optional`
   parameters leave `cParamsOpt` at 0 (Optional is marked per parameter); a `ParamArray` sets
   `cParamsOpt` to -1.

## Follow-up probes

The field sequence `Private one`, `Private two`, `Public Value`, `Private three`, `Public Entry`
gave `Value` DISPID `0x40030002` and `Entry` DISPID `0x40030004`. Their first vtable entries were at
bytes 56 and 72, respectively.

Removing `Class_Terminate` from `Widget` reduced `cbSizeVft` from 200 to 192 on both halves with
all 15 public entries unchanged. Removing its event or private fields left the size at 200.
Removing both the `Friend` and `Private` procedures while keeping `Class_Terminate` gave 184.
With one public method, an empty `Class_Initialize` kept the vtable at 64 bytes, a nonempty
`Class_Initialize` or a `Class_Terminate` containing only `Exit Sub` gave 72, and two nonempty
lifecycle procedures gave 80.

With a default property `Item` between `First` and `Last`, Excel reported `First = 0x68030001`,
`Item Get/Let = 0`, `Last = 0x68030000` and `Other = 0x60030002`. For a default method `Item`,
Excel reported `First = 0x60030000`, `Item = 0` and `Last = 0x60030001`. Late-bound default
invocation returned 42 for both fixtures.

## Appendix: raw output

`Widget` (fields `Value As Long`, `Name As String`, `Child As Object`; event `Changed`; private
`mX`, `mObj`; `Twice`, `Reset`, `X` get/let, `Obj` get/set, `WithArgs(a, Optional b, Optional c = 3)`,
`Rest(ParamArray items())`, Friend `Hidden`, Private `Secret`, `Class_Terminate`):

```text
describe:
type Widget kind=TKIND_DISPATCH flags=0x1C2 guid=18c68922-d160-4880-944c-4564bf267d76 funcs=15 vars=0 impl=1 vtbl=200
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x40030000 invkind=2 oVft=56 params=0 opt=0 flags=0x0 ret=3 Value
  func memid=0x40030000 invkind=4 oVft=64 params=1 opt=0 flags=0x0 ret=24 Value
  func memid=0x40030001 invkind=2 oVft=72 params=0 opt=0 flags=0x0 ret=8 Name
  func memid=0x40030001 invkind=4 oVft=80 params=1 opt=0 flags=0x0 ret=24 Name
  func memid=0x40030002 invkind=2 oVft=88 params=0 opt=0 flags=0x0 ret=9 Child
  func memid=0x40030002 invkind=4 oVft=96 params=1 opt=0 flags=0x1 ret=24 Child
  func memid=0x40030002 invkind=8 oVft=104 params=1 opt=0 flags=0x0 ret=24 Child
  func memid=0x60030002 invkind=1 oVft=112 params=0 opt=0 flags=0x0 ret=3 Twice
  func memid=0x60030003 invkind=1 oVft=120 params=0 opt=0 flags=0x0 ret=24 Reset
  func memid=0x68030001 invkind=2 oVft=128 params=0 opt=0 flags=0x0 ret=5 X
  func memid=0x68030001 invkind=4 oVft=136 params=1 opt=0 flags=0x0 ret=24 X
  func memid=0x68030000 invkind=2 oVft=144 params=0 opt=0 flags=0x0 ret=9 Obj
  func memid=0x68030000 invkind=8 oVft=152 params=1 opt=0 flags=0x0 ret=24 Obj
  func memid=0x60030004 invkind=1 oVft=160 params=3 opt=0 flags=0x0 ret=3 WithArgs,a,b,c
  func memid=0x60030005 invkind=1 oVft=168 params=1 opt=-1 flags=0x0 ret=3 Rest,items
--dual
type Widget kind=TKIND_INTERFACE flags=0x1C2 guid=18c68922-d160-4880-944c-4564bf267d76 funcs=15 vars=0 impl=1 vtbl=200
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x40030000 invkind=2 oVft=56 params=1 opt=0 flags=0x0 ret=25 Value
  func memid=0x40030000 invkind=4 oVft=64 params=1 opt=0 flags=0x0 ret=25 Value
  func memid=0x40030001 invkind=2 oVft=72 params=1 opt=0 flags=0x0 ret=25 Name
  func memid=0x40030001 invkind=4 oVft=80 params=1 opt=0 flags=0x0 ret=25 Name
  func memid=0x40030002 invkind=2 oVft=88 params=1 opt=0 flags=0x0 ret=25 Child
  func memid=0x40030002 invkind=4 oVft=96 params=1 opt=0 flags=0x1 ret=25 Child
  func memid=0x40030002 invkind=8 oVft=104 params=1 opt=0 flags=0x0 ret=25 Child
  func memid=0x60030002 invkind=1 oVft=112 params=1 opt=0 flags=0x0 ret=25 Twice
  func memid=0x60030003 invkind=1 oVft=120 params=0 opt=0 flags=0x0 ret=25 Reset
  func memid=0x68030001 invkind=2 oVft=128 params=1 opt=0 flags=0x0 ret=25 X
  func memid=0x68030001 invkind=4 oVft=136 params=1 opt=0 flags=0x0 ret=25 X
  func memid=0x68030000 invkind=2 oVft=144 params=1 opt=0 flags=0x0 ret=25 Obj
  func memid=0x68030000 invkind=8 oVft=152 params=1 opt=0 flags=0x0 ret=25 Obj
  func memid=0x60030004 invkind=1 oVft=160 params=4 opt=0 flags=0x0 ret=25 WithArgs,a,b,c
  func memid=0x60030005 invkind=1 oVft=168 params=2 opt=-1 flags=0x0 ret=25 Rest,items

```

`W2`-`W8`:

```text
== W2
type W2 kind=TKIND_DISPATCH flags=0x1C2 guid=be64e44d-5173-45ca-a55f-fbc9c70f19b5 funcs=3 vars=0 impl=1 vtbl=80
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x60030000 invkind=1 oVft=56 params=0 opt=0 flags=0x0 ret=3 A
  func memid=0x60030001 invkind=1 oVft=64 params=0 opt=0 flags=0x0 ret=3 B
  func memid=0x60030002 invkind=1 oVft=72 params=0 opt=0 flags=0x0 ret=3 C
--dual
type W2 kind=TKIND_INTERFACE flags=0x1C2 guid=be64e44d-5173-45ca-a55f-fbc9c70f19b5 funcs=3 vars=0 impl=1 vtbl=80
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x60030000 invkind=1 oVft=56 params=1 opt=0 flags=0x0 ret=25 A
  func memid=0x60030001 invkind=1 oVft=64 params=1 opt=0 flags=0x0 ret=25 B
  func memid=0x60030002 invkind=1 oVft=72 params=1 opt=0 flags=0x0 ret=25 C
== W3
type W3 kind=TKIND_DISPATCH flags=0x1C2 guid=515250b8-7732-4a38-8d79-37510133f122 funcs=2 vars=0 impl=1 vtbl=72
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x60030000 invkind=1 oVft=56 params=0 opt=0 flags=0x0 ret=3 A
  func memid=0x60030001 invkind=1 oVft=64 params=0 opt=0 flags=0x0 ret=3 B
--dual
type W3 kind=TKIND_INTERFACE flags=0x1C2 guid=515250b8-7732-4a38-8d79-37510133f122 funcs=2 vars=0 impl=1 vtbl=72
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x60030000 invkind=1 oVft=56 params=1 opt=0 flags=0x0 ret=25 A
  func memid=0x60030001 invkind=1 oVft=64 params=1 opt=0 flags=0x0 ret=25 B
== W4
type W4 kind=TKIND_DISPATCH flags=0x1C2 guid=dee4d075-5744-4ce2-94f7-f8b0a14b290e funcs=3 vars=0 impl=1 vtbl=80
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x68030002 invkind=2 oVft=56 params=0 opt=0 flags=0x0 ret=3 P1
  func memid=0x68030001 invkind=2 oVft=64 params=0 opt=0 flags=0x0 ret=3 P2
  func memid=0x68030000 invkind=2 oVft=72 params=0 opt=0 flags=0x0 ret=3 P3
--dual
type W4 kind=TKIND_INTERFACE flags=0x1C2 guid=dee4d075-5744-4ce2-94f7-f8b0a14b290e funcs=3 vars=0 impl=1 vtbl=80
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x68030002 invkind=2 oVft=56 params=1 opt=0 flags=0x0 ret=25 P1
  func memid=0x68030001 invkind=2 oVft=64 params=1 opt=0 flags=0x0 ret=25 P2
  func memid=0x68030000 invkind=2 oVft=72 params=1 opt=0 flags=0x0 ret=25 P3
== W5
type W5 kind=TKIND_DISPATCH flags=0x1C2 guid=9370ba69-fff8-40b6-9584-d27270561e5b funcs=1 vars=0 impl=1 vtbl=80
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x60030002 invkind=1 oVft=56 params=0 opt=0 flags=0x0 ret=3 A
--dual
type W5 kind=TKIND_INTERFACE flags=0x1C2 guid=9370ba69-fff8-40b6-9584-d27270561e5b funcs=1 vars=0 impl=1 vtbl=80
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x60030002 invkind=1 oVft=56 params=1 opt=0 flags=0x0 ret=25 A
== W6
type W6 kind=TKIND_DISPATCH flags=0x1C2 guid=d074f780-2b32-4581-8c43-d380a660b02d funcs=3 vars=0 impl=1 vtbl=80
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x60030000 invkind=1 oVft=56 params=0 opt=0 flags=0x0 ret=3 Zed
  func memid=0x60030001 invkind=1 oVft=64 params=0 opt=0 flags=0x0 ret=3 Alpha
  func memid=0x60030002 invkind=1 oVft=72 params=0 opt=0 flags=0x0 ret=3 Mid
--dual
type W6 kind=TKIND_INTERFACE flags=0x1C2 guid=d074f780-2b32-4581-8c43-d380a660b02d funcs=3 vars=0 impl=1 vtbl=80
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x60030000 invkind=1 oVft=56 params=1 opt=0 flags=0x0 ret=25 Zed
  func memid=0x60030001 invkind=1 oVft=64 params=1 opt=0 flags=0x0 ret=25 Alpha
  func memid=0x60030002 invkind=1 oVft=72 params=1 opt=0 flags=0x0 ret=25 Mid
== W7
type W7 kind=TKIND_DISPATCH flags=0x1C2 guid=c1d8041c-3fe3-4edc-b3ed-499f076a536d funcs=8 vars=0 impl=1 vtbl=120
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x40030000 invkind=2 oVft=56 params=0 opt=0 flags=0x0 ret=3 F1
  func memid=0x40030000 invkind=4 oVft=64 params=1 opt=0 flags=0x0 ret=24 F1
  func memid=0x40030001 invkind=2 oVft=72 params=0 opt=0 flags=0x0 ret=12 F2
  func memid=0x40030001 invkind=4 oVft=80 params=1 opt=0 flags=0x0 ret=24 F2
  func memid=0x40030001 invkind=8 oVft=88 params=1 opt=0 flags=0x0 ret=24 F2
  func memid=0x68030001 invkind=2 oVft=96 params=0 opt=0 flags=0x0 ret=3 Zed
  func memid=0x68030000 invkind=2 oVft=104 params=0 opt=0 flags=0x0 ret=3 Alpha
  func memid=0x60030002 invkind=1 oVft=112 params=0 opt=0 flags=0x0 ret=24 M
--dual
type W7 kind=TKIND_INTERFACE flags=0x1C2 guid=c1d8041c-3fe3-4edc-b3ed-499f076a536d funcs=8 vars=0 impl=1 vtbl=120
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x40030000 invkind=2 oVft=56 params=1 opt=0 flags=0x0 ret=25 F1
  func memid=0x40030000 invkind=4 oVft=64 params=1 opt=0 flags=0x0 ret=25 F1
  func memid=0x40030001 invkind=2 oVft=72 params=1 opt=0 flags=0x0 ret=25 F2
  func memid=0x40030001 invkind=4 oVft=80 params=1 opt=0 flags=0x0 ret=25 F2
  func memid=0x40030001 invkind=8 oVft=88 params=1 opt=0 flags=0x0 ret=25 F2
  func memid=0x68030001 invkind=2 oVft=96 params=1 opt=0 flags=0x0 ret=25 Zed
  func memid=0x68030000 invkind=2 oVft=104 params=1 opt=0 flags=0x0 ret=25 Alpha
  func memid=0x60030002 invkind=1 oVft=112 params=0 opt=0 flags=0x0 ret=25 M
== W8
type W8 kind=TKIND_DISPATCH flags=0x1C2 guid=9e93f6c9-8c86-4712-a305-67d28be1e3b2 funcs=1 vars=0 impl=1 vtbl=64
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x60030000 invkind=1 oVft=56 params=0 opt=0 flags=0x0 ret=3 A
--dual
type W8 kind=TKIND_INTERFACE flags=0x1C2 guid=9e93f6c9-8c86-4712-a305-67d28be1e3b2 funcs=1 vars=0 impl=1 vtbl=64
  !impl 0 Element not found. (Exception from HRESULT: 0x8002802B (TYPE_E_ELEMENTNOTFOUND))
  func memid=0x60030000 invkind=1 oVft=56 params=1 opt=0 flags=0x0 ret=25 A

```
