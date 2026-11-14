//! While an array element (or a field of an element) is out as a ByRef argument, VBA locks
//! every array on its path: `ReDim`, `ReDim Preserve`, `Erase` and replacing the array raise
//! error 10 until the call returns, even when the call raised. Replacing the record that holds a
//! locked field array is allowed; the reference then writes to the detached array. Source and
//! expected output are Excel x64's (`temp/variant-layout-probe/redim7c`, LockProbe, EdgeProbe and
//! RecProbe without the `cLocks` readings, captured 2026-11-14 in one owned Excel).
use calcvba_engine::{Engine, HostConfig};

const MODULE: &str = r#"
Option Explicit

Private Type TIn
    a As Long
    d() As Long
End Type

Private g() As Long
Private h() As Long
Private gv As Variant
Private gt As TIn
Private ts() As TIn

Private Function Bounds(ByRef d() As Long) As String
    On Error GoTo fail
    Bounds = LBound(d) & ".." & UBound(d)
    Exit Function
fail:
    Bounds = "err=" & Err.Number
End Function

Private Function TryRedim(n As Long) As String
    On Error GoTo fail
    ReDim g(0 To 5)
    TryRedim = "ok"
    Exit Function
fail:
    TryRedim = "err=" & Err.Number
End Function

Private Function TryPreserve(n As Long) As String
    On Error GoTo fail
    ReDim Preserve g(0 To 5)
    TryPreserve = "ok"
    Exit Function
fail:
    TryPreserve = "err=" & Err.Number
End Function

Private Function TryErase(n As Long) As String
    On Error GoTo fail
    Erase g
    TryErase = "ok"
    Exit Function
fail:
    TryErase = "err=" & Err.Number
End Function

Private Function TryAssign(n As Long) As String
    On Error GoTo fail
    Dim other() As Long
    ReDim other(0 To 1)
    g = other
    TryAssign = "ok"
    Exit Function
fail:
    TryAssign = "err=" & Err.Number
End Function

Private Function Nested(n As Long) As String
    Nested = TryRedim(n)
End Function

Private Function TryRedimH(n As Long) As String
    On Error GoTo fail
    ReDim h(0 To 5)
    TryRedimH = "ok"
    Exit Function
fail:
    TryRedimH = "err=" & Err.Number
End Function

Private Function Inner(n As Long) As Long
    Inner = n
End Function

Private Function PairRedimH(m As Long, k As Long) As String
    PairRedimH = TryRedimH(m)
End Function

Private Function RedimOuter(n As Long) As String
    On Error GoTo fail
    ReDim ts(0 To 3)
    RedimOuter = "ok"
    Exit Function
fail:
    RedimOuter = "err=" & Err.Number
End Function

Private Function RedimInner(n As Long) As String
    On Error GoTo fail
    ReDim ts(1).d(0 To 7)
    RedimInner = "ok"
    Exit Function
fail:
    RedimInner = "err=" & Err.Number
End Function

Private Function OverwriteVariant(x As Variant) As String
    On Error GoTo fail
    gv = 5
    OverwriteVariant = "ok"
    Exit Function
fail:
    OverwriteVariant = "err=" & Err.Number
End Function

Private Function EraseRecords(n As Long) As String
    On Error GoTo fail
    Erase ts
    EraseRecords = "ok"
    Exit Function
fail:
    EraseRecords = "err=" & Err.Number
End Function

Private Function WriteAfter(n As Long) As String
    n = 77
    WriteAfter = "g1=" & g(1)
End Function

Private Function AssignSized(n As Long) As String
    Dim other As TIn
    ReDim other.d(0 To 4)
    other.d(1) = 9
    gt = other
    n = 66
    AssignSized = "bounds=" & Bounds(gt.d) & " d1=" & gt.d(1)
End Function

Private Sub Boom(n As Long)
    Err.Raise 5
End Sub

Public Function RunAll() As String
    Dim out As String, e As Long
    ReDim g(0 To 2)
    out = out & "ELEM_REDIM " & TryRedim(g(1)) & " ub=" & UBound(g) & vbLf
    out = out & "ELEM_PRESERVE " & TryPreserve(g(1)) & " ub=" & UBound(g) & vbLf
    out = out & "ELEM_ERASE " & TryErase(g(1)) & vbLf
    out = out & "ELEM_ASSIGN " & TryAssign(g(1)) & " ub=" & UBound(g) & vbLf
    out = out & "NESTED_REDIM " & Nested(g(1)) & vbLf
    out = out & "SCALAR_ARG " & TryRedim(e) & " ub=" & UBound(g) & vbLf
    ReDim g(0 To 2)
    out = out & "BYVAL_ELEM " & TryRedim((g(1))) & " ub=" & UBound(g) & vbLf
    ReDim h(0 To 2)
    out = out & "ARG_IN_ARG " & PairRedimH(g(0), Inner(h(1))) & " ubh=" & UBound(h) & vbLf

    ReDim ts(0 To 2)
    ReDim ts(1).d(0 To 2)
    out = out & "REDIM_OUTER_VIA_FIELD " & RedimOuter(ts(1).a) & vbLf
    out = out & "REDIM_OUTER_VIA_INNER " & RedimOuter(ts(1).d(1)) & vbLf
    out = out & "REDIM_INNER_VIA_INNER " & RedimInner(ts(1).d(1)) & vbLf
    out = out & "REDIM_INNER_VIA_FIELD " & RedimInner(ts(1).a) & " ub=" & UBound(ts(1).d) & vbLf
    out = out & "ERASE_RECORDS " & EraseRecords(ts(1).a) & vbLf

    gv = Array(1, 2, 3)
    out = out & "VARIANT_OVERWRITE " & OverwriteVariant(gv(1)) & vbLf

    ReDim g(0 To 2)
    out = out & "WRITE_AFTER " & WriteAfter(g(1)) & vbLf
    ReDim gt.d(0 To 2)
    out = out & "ASSIGN_SIZED " & AssignSized(gt.d(1)) & vbLf

    On Error Resume Next
    Boom g(1)
    out = out & "AFTER_ERROR err=" & Err.Number
    Err.Clear
    On Error GoTo 0
    out = out & " " & TryRedim(e) & vbLf
    RunAll = out
End Function

"#;

const EXPECTED: &str = "ELEM_REDIM err=10 ub=2
ELEM_PRESERVE err=10 ub=2
ELEM_ERASE err=10
ELEM_ASSIGN err=10 ub=2
NESTED_REDIM err=10
SCALAR_ARG ok ub=5
BYVAL_ELEM ok ub=5
ARG_IN_ARG ok ubh=5
REDIM_OUTER_VIA_FIELD err=10
REDIM_OUTER_VIA_INNER err=10
REDIM_INNER_VIA_INNER err=10
REDIM_INNER_VIA_FIELD ok ub=7
ERASE_RECORDS err=10
VARIANT_OVERWRITE err=10
WRITE_AFTER g1=77
ASSIGN_SIZED bounds=0..4 d1=9
AFTER_ERROR err=5 ok
";

#[test]
fn byref_elements_lock_their_arrays_on_vm3_and_jit() {
    let source = format!(
        "Public result As String{nl}{MODULE}{nl}Sub Main(){nl}    result = RunAll(){nl}End Sub{nl}",
        nl = "\n"
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
                    .filter(|s| s.contains("ELEM_REDIM"))
            })
            .unwrap_or_else(|| panic!("{name}: no output in {values:?}"));
        assert_eq!(out, EXPECTED, "{name}");
    }
}
