//! `For Each` over an array locks it (`cLocks` + 1) as VBA does: `ReDim`, `Erase` and replacing
//! the array raise error 10 inside the loop; `Exit For` and the normal end release the lock, a
//! `GoTo` out of the loop keeps it until the procedure returns; the loop walks the live array.
//! Source and expected output are Excel x64's (`temp/variant-layout-probe/redim7c/ForEachProbe.bas`,
//! captured 2026-11-14 in one owned Excel). The probe reads `cLocks` through `CopyMemory`, so it
//! runs with the native Declare lane.
#![cfg(any(windows, all(target_os = "linux", target_pointer_width = "64")))]

use calcvba_engine::{Engine, HostConfig};
use calcvba_hal::HostPolicy;

const MODULE: &str = r#"
Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal n As LongPtr)

Private Type TRec
    a As Long
    d() As Long
End Type

Private g() As Long
Private gv As Variant
Private gt As TRec

Private Function Ptr(ByVal addr As LongPtr) As LongPtr
    Dim p As LongPtr
    CopyMemory VarPtr(p), addr, 8
    Ptr = p
End Function

Private Function SlotOf(v As Variant) As LongPtr
    SlotOf = Ptr(VarPtr(v) + 8)
End Function

Private Function Locks(ByVal psa As LongPtr) As Long
    Dim n As Long
    If psa = 0 Then Locks = -1: Exit Function
    CopyMemory VarPtr(n), psa + 8, 4
    Locks = n
End Function

Private Function GLocks() As Long
    GLocks = Locks(Ptr(SlotOf(g)))
End Function

Private Sub LeakForEach()
    Dim x As Variant
    For Each x In g
        GoTo out
    Next
out:
End Sub

Private Function LocalLeak() As String
    Dim loc() As Long, x As Variant
    ReDim loc(0 To 2)
    For Each x In loc
        GoTo out
    Next
out:
    On Error Resume Next
    ReDim loc(0 To 5)
    LocalLeak = "redim_err=" & Err.Number & " locks=" & Locks(Ptr(SlotOf(loc)))
    Err.Clear
    For Each x In loc
        Exit For
    Next
    LocalLeak = LocalLeak & " after_second_loop=" & Locks(Ptr(SlotOf(loc)))
End Function

Public Function RunAll() As String
    Dim out As String, x As Variant, i As Long, total As Long
    ReDim g(0 To 2)
    out = out & "IDLE locks=" & GLocks() & vbLf
    On Error Resume Next
    For Each x In g
        ReDim g(0 To 5)
        out = out & "FOREACH_REDIM err=" & Err.Number & " locks=" & GLocks() & vbLf
        Err.Clear
        Erase g
        out = out & "FOREACH_ERASE err=" & Err.Number & vbLf
        Err.Clear
        Exit For
    Next
    On Error GoTo 0
    out = out & "FOREACH_EXIT ub=" & UBound(g) & " locks=" & GLocks() & vbLf
    For Each x In g
    Next
    out = out & "FOREACH_DONE locks=" & GLocks() & vbLf
    g(0) = 1: g(1) = 2: g(2) = 3
    For Each x In g
        If x = 1 Then g(2) = 30
        total = total + x
    Next
    out = out & "FOREACH_LIVE total=" & total & vbLf
    For i = 0 To 0
        For Each x In g
            GoTo jumped
        Next
    Next
jumped:
    out = out & "FOREACH_GOTO locks=" & GLocks() & vbLf
    On Error Resume Next
    ReDim g(0 To 1)
    out = out & "AFTER_GOTO_REDIM err=" & Err.Number & vbLf
    Err.Clear
    On Error GoTo 0
    LeakForEach
    out = out & "LEAK_AFTER_CALL locks=" & GLocks() & vbLf
    out = out & "LOCAL_LEAK " & LocalLeak() & vbLf
    ReDim gt.d(0 To 2)
    On Error Resume Next
    For Each x In gt.d
        ReDim gt.d(0 To 5)
        out = out & "FIELD_FOREACH_REDIM err=" & Err.Number & vbLf
        Err.Clear
        Exit For
    Next
    On Error GoTo 0
    out = out & "FIELD_FOREACH_AFTER ub=" & UBound(gt.d) & vbLf
    gv = Array(1, 2, 3)
    On Error Resume Next
    For Each x In gv
        gv = 5
        out = out & "VARIANT_FOREACH_OVERWRITE err=" & Err.Number & vbLf
        Err.Clear
        Exit For
    Next
    On Error GoTo 0
    out = out & "VARIANT_FOREACH_AFTER vt=" & VarType(gv) & vbLf
    RunAll = out
End Function
"#;

const EXPECTED: &str = "IDLE locks=0
FOREACH_REDIM err=10 locks=1
FOREACH_ERASE err=10
FOREACH_EXIT ub=2 locks=0
FOREACH_DONE locks=0
FOREACH_LIVE total=33
FOREACH_GOTO locks=1
AFTER_GOTO_REDIM err=10
LEAK_AFTER_CALL locks=1
LOCAL_LEAK redim_err=10 locks=1 after_second_loop=1
FIELD_FOREACH_REDIM err=10
FIELD_FOREACH_AFTER ub=2
VARIANT_FOREACH_OVERWRITE err=10
VARIANT_FOREACH_AFTER vt=8204
";

/// The probe's `CopyMemory` as the platform's C library has it (`RtlMoveMemory` on Windows,
/// `memmove` in glibc), so the probe runs off Windows too.
fn platform(source: &str) -> String {
    if cfg!(windows) {
        source.to_string()
    } else {
        source.replace(
            r#"Lib "kernel32" Alias "RtlMoveMemory""#,
            r#"Lib "libc.so.6" Alias "memmove""#,
        )
    }
}

#[test]
fn for_each_locks_the_array_on_vm3_and_jit() {
    let module = platform(MODULE);
    let source = format!(
        "Public result As String{nl}{module}{nl}Sub Main(){nl}    result = RunAll(){nl}End Sub{nl}",
        nl = "
"
    );
    for (name, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(HostPolicy::native());
        let values = engine
            .execute_source_with_variant_snapshot_clean(&source)
            .unwrap_or_else(|e| panic!("{name}: {e:?}"));
        let out = values
            .iter()
            .find_map(|v| {
                v.as_bstr()
                    .map(|s| s.to_string())
                    .filter(|s| s.contains("FOREACH_REDIM"))
            })
            .unwrap_or_else(|| panic!("{name}: no output in {values:?}"));
        assert_eq!(out, EXPECTED, "{name}");
    }
}
