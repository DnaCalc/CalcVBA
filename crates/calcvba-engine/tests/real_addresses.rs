//! `VarPtr`, `StrPtr` and `ObjPtr` give the real addresses of VBA storage on VM3 and the JIT,
//! and writes through them reach the variables. Every expected line is Excel x64's
//! (`temp/variant-layout-probe/addr7/AddrProbe.bas`, captured 2026-11-14 in one owned Excel),
//! except the UDT offsets, which are section 2 of the variant-layout design. The ByRef Variant
//! lines (`VT_BYREF` cells) belong to step 7b.
#![cfg(target_os = "windows")]

use calcvba_engine::{Engine, HostConfig};
use calcvba_hal::model::HostPolicy;
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};

const PROBE: &str = r#"
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal n As LongPtr)

Private Type TMix
    a As Long
    v As Variant
    b As Integer
End Type

Public out As String
Private mLong As Long
Private mDepth(1 To 3) As LongPtr

Private Function Ptr(ByVal p As LongPtr) As LongPtr
    Dim v As LongPtr
    CopyMemory VarPtr(v), p, LenB(v)
    Ptr = v
End Function

Private Function I16(ByVal p As LongPtr) As Long
    Dim v As Integer
    CopyMemory VarPtr(v), p, 2
    I16 = v And &HFFFF&
End Function

Private Function CalleeLong(n As Long) As LongPtr
    CalleeLong = VarPtr(n)
End Function

Private Function CalleeLongByVal(ByVal n As Long) As LongPtr
    CalleeLongByVal = VarPtr(n)
End Function

Private Function CalleeVariant(v As Variant, ByRef vt As Long) As LongPtr
    CalleeVariant = VarPtr(v)
    vt = I16(VarPtr(v))
End Function

Private Function Nested(v As Variant) As LongPtr
    Dim vt As Long
    Nested = CalleeVariant(v, vt)
End Function

Private Sub Recurse(ByVal depth As Long)
    Dim local1 As Long
    mDepth(depth) = VarPtr(local1)
    If depth < 3 Then Recurse depth + 1
End Sub

Private Function StaticAddr() As LongPtr
    Static s As Long
    StaticAddr = VarPtr(s)
End Function

Sub Main()
    Dim x As Long, p1 As LongPtr, p2 As LongPtr, vt As Long
    Dim s As String, v As Variant, w As Variant, d() As Long, va As Variant, o As Collection
    Dim c As AddrClass, t As TMix, ts(0 To 1) As TMix

    p1 = VarPtr(x): x = 5: p2 = VarPtr(x)
    out = out & "LOCAL_STABLE " & (p1 = p2) & vbLf
    p2 = CalleeLong(x)
    out = out & "BYREF_LONG_SAME " & (p1 = p2) & vbLf
    p2 = CalleeLongByVal(x)
    out = out & "BYVAL_LONG_SAME " & (p1 = p2) & vbLf
    p2 = CalleeLong((x))
    out = out & "PAREN_LONG_SAME " & (p1 = p2) & vbLf
    p1 = VarPtr(v)
    p2 = CalleeVariant(v, vt)
    out = out & "BYREF_VARIANT_OF_VARIANT same=" & (p1 = p2) & " vt=&H" & Hex$(vt) & vbLf
    p2 = Nested(v)
    out = out & "NESTED_VARIANT same=" & (p1 = p2) & vbLf

    Recurse 1
    out = out & "RECURSE distinct=" & (mDepth(1) <> mDepth(2) And mDepth(2) <> mDepth(3)) & vbLf
    out = out & "STATIC_STABLE " & (StaticAddr() = StaticAddr()) & vbLf
    p1 = VarPtr(mLong): mLong = 3
    out = out & "MODULE_STABLE " & (p1 = VarPtr(mLong)) & vbLf

    s = "abc"
    out = out & "STR varptr_holds_strptr=" & (Ptr(VarPtr(s)) = StrPtr(s)) & vbLf
    p1 = StrPtr(s): s = "abcd"
    out = out & "STR_REASSIGN strptr_changed=" & (p1 <> StrPtr(s)) & " varptr_same=" & (VarPtr(s) = VarPtr(s)) & vbLf
    out = out & "STRPTR_NULL=" & StrPtr(vbNullString) & " STRPTR_EMPTY_NONZERO=" & (StrPtr("") <> 0) & vbLf
    out = out & "OBJPTR_NOTHING=" & ObjPtr(o) & vbLf
    Set o = New Collection
    out = out & "OBJPTR_VARPTR " & (Ptr(VarPtr(o)) = ObjPtr(o)) & vbLf

    v = "xy"
    out = out & "VARIANT_STR payload_is_strptr=" & (Ptr(VarPtr(v) + 8) = StrPtr(v)) & vbLf

    ReDim d(0 To 1)
    out = out & "ELEM d(1)-d(0)=" & (VarPtr(d(1)) - VarPtr(d(0))) & vbLf
    va = Array(1, 2, 3)
    out = out & "VARIANT_ARRAY_ELEM v(1)-v(0)=" & (VarPtr(va(1)) - VarPtr(va(0))) & " v0_is_pvdata=" & (VarPtr(va(0)) = Ptr(Ptr(VarPtr(va) + 8) + 16)) & vbLf

    Set c = New AddrClass
    out = out & "CLASS_FIELD_STABLE " & (c.FieldAddr() = c.FieldAddr()) & vbLf
    c.SetField 9
    x = 5
    CopyMemory c.FieldAddr(), VarPtr(x), 4
    out = out & "CLASS_FIELD_WRITE " & c.GetField() & vbLf

    x = 1
    CopyMemory VarPtr(x), VarPtr(mLong), 4
    out = out & "WRITE_LOCAL " & x & vbLf
    w = 7&
    CopyMemory VarPtr(w) + 8, VarPtr(mLong), 4
    out = out & "WRITE_VARIANT_PAYLOAD " & w & vbLf

    out = out & "UDT v=" & (VarPtr(t.v) - VarPtr(t)) & " b=" & (VarPtr(t.b) - VarPtr(t)) & " stride=" & (VarPtr(ts(1)) - VarPtr(ts(0))) & vbLf
    t.a = 1
    CopyMemory VarPtr(t.a), VarPtr(mLong), 4
    out = out & "WRITE_UDT_FIELD " & t.a & vbLf
    ts(1).b = 4
    x = 6
    CopyMemory VarPtr(ts(1).b), VarPtr(x), 2
    out = out & "WRITE_UDT_ELEMENT_FIELD " & ts(1).b & vbLf
    d(1) = 0
    CopyMemory VarPtr(d(1)), VarPtr(mLong), 4
    out = out & "WRITE_ELEMENT " & d(1) & vbLf
End Sub
"#;

const CLASS: &str = r#"
Option Explicit
Private mField As Long

Public Function FieldAddr() As LongPtr
    FieldAddr = VarPtr(mField)
End Function

Public Sub SetField(ByVal n As Long)
    mField = n
End Sub

Public Function GetField() As Long
    GetField = mField
End Function
"#;

const EXPECTED: &str = "LOCAL_STABLE True
BYREF_LONG_SAME True
BYVAL_LONG_SAME False
PAREN_LONG_SAME False
BYREF_VARIANT_OF_VARIANT same=True vt=&H0
NESTED_VARIANT same=True
RECURSE distinct=True
STATIC_STABLE True
MODULE_STABLE True
STR varptr_holds_strptr=True
STR_REASSIGN strptr_changed=True varptr_same=True
STRPTR_NULL=0 STRPTR_EMPTY_NONZERO=True
OBJPTR_NOTHING=0
OBJPTR_VARPTR True
VARIANT_STR payload_is_strptr=True
ELEM d(1)-d(0)=4
VARIANT_ARRAY_ELEM v(1)-v(0)=24 v0_is_pvdata=True
CLASS_FIELD_STABLE True
CLASS_FIELD_WRITE 5
WRITE_LOCAL 3
WRITE_VARIANT_PAYLOAD 3
UDT v=8 b=32 stride=40
WRITE_UDT_FIELD 3
WRITE_UDT_ELEMENT_FIELD 6
WRITE_ELEMENT 3
";

fn module(name: &str, kind: ModuleKind, source: &str) -> ModuleUnit {
    ModuleUnit {
        module_name: name.into(),
        module_kind: kind,
        attributes: ModuleAttributes::named(name),
        source: source.into(),
    }
}

#[test]
fn pointer_helpers_give_real_addresses_on_vm3_and_jit() {
    let manifest = SymbolProjectManifest {
        project_name: "Proj".into(),
        project_kind: ProjectKind::Source,
        modules: vec![
            module("AddrProbe", ModuleKind::Procedural, PROBE),
            module("AddrClass", ModuleKind::Class, CLASS),
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    for (name, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(HostPolicy::native());
        let values = engine
            .execute_manifest_with_variant_snapshot(&manifest)
            .unwrap_or_else(|e| panic!("{name}: {:?}: {}", e.phase(), e.message()));
        let out = values
            .iter()
            .find_map(|v| {
                v.as_bstr()
                    .map(|s| s.to_string())
                    .filter(|s| s.contains("LOCAL_STABLE"))
            })
            .unwrap_or_else(|| panic!("{name}: no probe output in {values:?}"));
        assert_eq!(out, EXPECTED, "{name}");
    }
}
