//! A typed variable passed to a ByRef `Variant` parameter is a `VT_BYREF` reference to it, on
//! VM3 and the JIT: reads see the variable, writes reach it coerced to its type (with VBA's
//! errors leaving it unchanged), and `ReDim`, element writes, `For Each`, `Erase` and `Set` act
//! on the caller's variable. Source and expected output are Excel x64's
//! (`temp/variant-layout-probe/byref7b/ByRefOps.bas`, captured 2026-11-14 in one owned Excel).
use calcvba_engine::{Engine, HostConfig};

const MODULE: &str = r#"
Option Explicit

Private Sub Grow(v As Variant)
    ReDim Preserve v(0 To 3)
    v(3) = 7
End Sub

Private Sub SetElem(v As Variant)
    v(1) = 5
End Sub

Private Function SumAll(v As Variant) As Long
    Dim x As Variant, s As Long
    For Each x In v
        s = s + x
    Next x
    SumAll = s
End Function

Private Sub Wipe(v As Variant)
    Erase v
End Sub

Private Sub Inner(v As Variant)
    v = v + 1
End Sub

Private Sub Outer(v As Variant)
    Inner v
End Sub

Private Sub Append(v As Variant)
    v = v & "b"
End Sub

Private Sub AssignTo(v As Variant, ByVal value As Variant)
    v = value
End Sub

Private Function TryAssign(v As Variant, ByVal value As Variant) As String
    On Error Resume Next
    v = value
    TryAssign = "err=" & Err.Number
End Function

Private Sub SetObj(v As Variant)
    Set v = New Collection
End Sub

Private Sub SetNothing(v As Variant)
    Set v = Nothing
End Sub

Private Function Kind(v As Variant) As String
    Kind = TypeName(v) & "/" & VarType(v) & "/" & IsArray(v)
End Function

Private Function Bounds(v As Variant) As String
    Bounds = LBound(v) & ":" & UBound(v)
End Function

Public Function RunAll() As String
    Dim out As String, a() As Long, n As Long, s As String, d As Double, b As Boolean
    Dim bt As Byte, o As Object, i As Integer, c As Currency, dt As Date, fa(1 To 3) As Long
    ReDim a(0 To 1): a(0) = 1
    Grow a
    out = out & "GROW ub=" & UBound(a) & " a0=" & a(0) & " a3=" & a(3) & vbLf
    SetElem a
    out = out & "SETELEM a1=" & a(1) & vbLf
    out = out & "SUM " & SumAll(a) & vbLf
    out = out & "KIND_ARRAY " & Kind(a) & " BOUNDS " & Bounds(a) & vbLf
    Wipe a
    On Error Resume Next
    n = UBound(a)
    out = out & "WIPE err=" & Err.Number & vbLf
    Err.Clear
    On Error GoTo 0
    fa(1) = 4
    out = out & "FIXED_SUM " & SumAll(fa) & " KIND " & Kind(fa) & vbLf
    out = out & "FIXED_GROW " & TryGrow(fa) & vbLf
    n = 1: Outer n
    out = out & "NESTED_LONG " & n & vbLf
    d = 1.5: Outer d
    out = out & "NESTED_DOUBLE " & d & vbLf
    i = 32767
    out = out & "INTEGER_OVERFLOW " & TryAssign(i, 40000) & " i=" & i & vbLf
    s = "a": Append s
    out = out & "APPEND " & s & vbLf
    AssignTo b, 5
    out = out & "BOOL_FROM_5 " & b & vbLf
    out = out & "BYTE_300 " & TryAssign(bt, 300) & " bt=" & bt & vbLf
    AssignTo c, "1.23456"
    out = out & "CURRENCY " & c & vbLf
    AssignTo dt, "2026-11-14"
    out = out & "DATE " & Format$(dt, "yyyy-mm-dd") & vbLf
    SetObj o
    out = out & "SETOBJ " & TypeName(o) & vbLf
    SetNothing o
    out = out & "SETNOTHING " & (o Is Nothing) & vbLf
    out = out & "STRING_FROM_NULL " & TryAssign(s, Null) & " s=" & s & vbLf
    out = out & "LONG_FROM_EMPTY " & TryAssign(n, Empty) & " n=" & n & vbLf
    RunAll = out
End Function

Private Function TryGrow(v As Variant) As String
    On Error Resume Next
    ReDim v(0 To 5)
    TryGrow = "err=" & Err.Number
End Function

"#;

const EXPECTED: &str = "GROW ub=3 a0=1 a3=7
SETELEM a1=5
SUM 13
KIND_ARRAY Long()/8195/True BOUNDS 0:3
WIPE err=9
FIXED_SUM 4 KIND Long()/8195/True
FIXED_GROW err=10
NESTED_LONG 2
NESTED_DOUBLE 2.5
INTEGER_OVERFLOW err=6 i=32767
APPEND ab
BOOL_FROM_5 True
BYTE_300 err=6 bt=0
CURRENCY 1.2346
DATE 2026-11-14
SETOBJ Collection
SETNOTHING True
STRING_FROM_NULL err=94 s=ab
LONG_FROM_EMPTY err=0 n=0
";

#[test]
fn typed_variables_bound_to_byref_variants_are_references_on_vm3_and_jit() {
    let source = format!(
        "Public result As String
{MODULE}
Sub Main()
    result = RunAll()
End Sub
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
                    .filter(|s| s.contains("GROW"))
            })
            .unwrap_or_else(|| panic!("{name}: no output in {values:?}"));
        assert_eq!(out.replace("\r\n", "\n"), EXPECTED, "{name}");
    }
}
