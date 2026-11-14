//! An array element or UDT field passed ByRef is the storage itself on VM3 and the JIT, as in
//! VBA: the callee's `VarPtr` is the element's, its writes reach the element at once (seen
//! through the array before the call returns), and nested calls pass the same storage on.
//! Source and expected output are Excel x64's (`temp/variant-layout-probe/byref7b2/ElemRefs.bas`,
//! captured 2026-11-14 in one owned Excel).
use calcvba_engine::{Engine, HostConfig};

const MODULE: &str = r#"
Option Explicit

Private Type TRec
    a As Long
    s As String
    v As Variant
    d() As Long
End Type

Private Function AddrLong(n As Long) As LongPtr
    AddrLong = VarPtr(n)
End Function

Private Sub BumpLong(n As Long)
    n = n + 1
End Sub

Private Sub AppendStr(s As String)
    s = s & "!"
End Sub

Private Sub SetVar(v As Variant)
    v = "set"
End Sub

Private Sub Outer(n As Long)
    BumpLong n
End Sub

Private Sub GrowArr(d() As Long)
    ReDim Preserve d(0 To 2)
    d(2) = 9
End Sub

Private Sub WriteThenRead(n As Long, ByRef seen As Long, a() As Long)
    n = 42
    seen = a(1)
End Sub

Public Function RunAll() As String
    Dim out As String, a(0 To 2) As Long, t As TRec, ts(0 To 1) As TRec, s(0 To 1) As String, va(0 To 1) As Variant, seen As Long
    out = out & "ELEM_ADDR " & (AddrLong(a(1)) = VarPtr(a(1))) & vbLf
    out = out & "FIELD_ADDR " & (AddrLong(t.a) = VarPtr(t.a)) & vbLf
    out = out & "ELEM_FIELD_ADDR " & (AddrLong(ts(1).a) = VarPtr(ts(1).a)) & vbLf
    a(1) = 5: BumpLong a(1)
    out = out & "BUMP_ELEM " & a(1) & vbLf
    t.a = 7: BumpLong t.a
    out = out & "BUMP_FIELD " & t.a & vbLf
    ts(1).a = 1: Outer ts(1).a
    out = out & "NESTED_ELEM_FIELD " & ts(1).a & vbLf
    s(0) = "x": AppendStr s(0)
    out = out & "STR_ELEM " & s(0) & vbLf
    t.s = "y": AppendStr t.s
    out = out & "STR_FIELD " & t.s & vbLf
    SetVar va(1)
    out = out & "VAR_ELEM " & va(1) & vbLf
    SetVar t.v
    out = out & "VAR_FIELD " & t.v & vbLf
    ReDim t.d(0 To 0)
    GrowArr t.d
    out = out & "ARR_FIELD ub=" & UBound(t.d) & " d2=" & t.d(2) & vbLf
    a(1) = 0
    WriteThenRead a(1), seen, a
    out = out & "ALIAS_SEEN " & seen & " a1=" & a(1) & vbLf
    RunAll = out
End Function

"#;

const EXPECTED: &str = "ELEM_ADDR True
FIELD_ADDR True
ELEM_FIELD_ADDR True
BUMP_ELEM 6
BUMP_FIELD 8
NESTED_ELEM_FIELD 2
STR_ELEM x!
STR_FIELD y!
VAR_ELEM set
VAR_FIELD set
ARR_FIELD ub=2 d2=9
ALIAS_SEEN 42 a1=42
";

#[test]
fn byref_elements_and_fields_are_the_storage_itself_on_vm3_and_jit() {
    let source = format!(
        "Public result As String{nl}{MODULE}{nl}Sub Main(){nl}    result = RunAll(){nl}End Sub{nl}",
        nl = "
"
    );
    for (name, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let values = Engine::new(config)
            .execute_source_with_variant_snapshot_clean(&source)
            .unwrap_or_else(|e| panic!("{name}: {e:?}"));
        let out = values
            .iter()
            .find_map(|v| {
                v.as_bstr()
                    .map(|s| s.to_string())
                    .filter(|s| s.contains("ELEM_ADDR"))
            })
            .unwrap_or_else(|| panic!("{name}: no output in {values:?}"));
        assert_eq!(out, EXPECTED, "{name}");
    }
}
