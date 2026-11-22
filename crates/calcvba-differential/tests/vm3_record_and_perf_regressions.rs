//! Regression tests for VM3 record-copy elimination, in-place field updates,
//! and VM3/JIT parity across optimized dispatch, value writes, arrays, strings,
//! and class calls (dc-17aqo.9).

use calcvba_differential::{Executor, run, run_modules};
use calcvba_symbol::manifest::ModuleKind;

fn assert_vm3_jit_match(label: &str, source: &str) {
    let vm3 = run(Executor::Vm3, source);
    let jit = run(Executor::Jit, source);
    assert!(
        vm3.unsupported.is_none(),
        "{label}: vm3 unsupported: {:?}",
        vm3.unsupported
    );
    assert!(
        jit.unsupported.is_none(),
        "{label}: jit unsupported: {:?}",
        jit.unsupported
    );
    assert_eq!(vm3.result, jit.result, "{label}: result mismatch");
    assert_eq!(vm3.raised, jit.raised, "{label}: raised mismatch");
    assert_eq!(vm3.err, jit.err, "{label}: err mismatch");
    if let Some(balance) = vm3.handle_balance {
        assert!(balance.is_zero(), "{label}: vm3 handle leak: {balance:?}");
    }
    if let Some(balance) = jit.handle_balance {
        assert!(balance.is_zero(), "{label}: jit handle leak: {balance:?}");
    }
}

fn assert_vm3_jit_match_modules(label: &str, modules: &[(&str, ModuleKind, &str)]) {
    let vm3 = run_modules(Executor::Vm3, modules, "Vm3PerfTest");
    let jit = run_modules(Executor::Jit, modules, "Vm3PerfTest");
    assert!(
        vm3.unsupported.is_none(),
        "{label}: vm3 unsupported: {:?}",
        vm3.unsupported
    );
    assert!(
        jit.unsupported.is_none(),
        "{label}: jit unsupported: {:?}",
        jit.unsupported
    );
    assert_eq!(vm3.result, jit.result, "{label}: result mismatch");
    assert_eq!(vm3.raised, jit.raised, "{label}: raised mismatch");
    assert_eq!(vm3.err, jit.err, "{label}: err mismatch");
    if let Some(balance) = vm3.handle_balance {
        assert!(balance.is_zero(), "{label}: vm3 handle leak: {balance:?}");
    }
}

#[test]
fn append_preserves_utf16_units_and_handles_self_concatenation() {
    assert_vm3_jit_match(
        "append UTF-16",
        r#"
Public result As String
Sub Main()
    Dim text As String, suffix As String, i As Long, v As Variant, odd As String
    text = String$(260, "x")
    suffix = ChrW(0) & ChrW(&HD800)
    For i = 1 To 64
        text = text & suffix
        text = text & "y"
    Next i
    text = text & text
    odd = ChrB(65) & ChrB(66)
    v = Null
    On Error Resume Next
    text = v & v
    result = CStr(Err.Number) & ":" & CStr(AscW(odd)) & ":" & CStr(AscW(Mid$(text, 262, 1))) & ":" & text
End Sub
"#,
    );
}

/// Updating a scalar field of a UDT or a UDT-array element must mutate the
/// record in place without deep-cloning the record or reallocating sibling
/// BSTR string fields (matches Excel VBA storage stability and eliminates
/// per-update record/BSTR copies in VM3).
#[test]
fn record_set_and_udt_array_field_update_preserve_sibling_string_bstr_address() {
    let source = r#"
Private Type Rec
    Id As Long
    Amount As Double
    Name As String
End Type

Public localStrStable As Boolean
Public localVarStable As Boolean
Public arrayStrStable As Boolean
Public arrayVarStable As Boolean
Public byrefStrStable As Boolean
Public finalAmount As Double

Private Sub BumpByRef(r As Rec)
    r.Amount = r.Amount + 2.5
End Sub

Sub Main()
    Dim localRec As Rec
    localRec.Id = 7
    localRec.Amount = 10#
    localRec.Name = "stable-local"

    Dim pStrBefore As LongPtr, pVarBefore As LongPtr
    pStrBefore = StrPtr(localRec.Name)
    pVarBefore = VarPtr(localRec.Amount)

    Dim i As Long, pad As String
    localStrStable = True
    localVarStable = True
    For i = 1 To 51
        pad = "pad-" & i
        localRec.Amount = localRec.Amount + 0.5
        If StrPtr(localRec.Name) <> pStrBefore Then localStrStable = False
        If VarPtr(localRec.Amount) <> pVarBefore Then localVarStable = False
    Next i
    localStrStable = localStrStable And (pStrBefore <> 0)
    localVarStable = localVarStable And (pVarBefore <> 0)

    BumpByRef localRec
    byrefStrStable = (StrPtr(localRec.Name) = pStrBefore)

    Dim arr() As Rec
    ReDim arr(1 To 4)
    arr(2).Id = 2
    arr(2).Amount = 1#
    arr(2).Name = "stable-array"

    Dim aStrBefore As LongPtr, aVarBefore As LongPtr
    aStrBefore = StrPtr(arr(2).Name)
    aVarBefore = VarPtr(arr(2).Amount)

    Dim k As Long
    k = 2
    arrayStrStable = True
    arrayVarStable = True
    For i = 1 To 51
        pad = "arr-pad-" & i
        arr(k).Amount = arr(k).Amount + i * 0.25
        If StrPtr(arr(2).Name) <> aStrBefore Then arrayStrStable = False
        If VarPtr(arr(2).Amount) <> aVarBefore Then arrayVarStable = False
    Next i
    arrayStrStable = arrayStrStable And (aStrBefore <> 0)
    arrayVarStable = arrayVarStable And (aVarBefore <> 0)
    finalAmount = localRec.Amount + arr(2).Amount
End Sub
"#;

    let vm3 = run(Executor::Vm3, source);
    assert!(
        vm3.unsupported.is_none(),
        "unsupported: {:?}",
        vm3.unsupported
    );
    assert!(!vm3.raised, "unexpected fault: {:?}", vm3.err);
    let snap = vm3.result.expect("vm3 run succeeded");
    assert_eq!(
        snap[0],
        calcvba_differential::Canon::Bool(true),
        "local UDT RecordSet must not reallocate sibling String BSTR: {snap:?}"
    );
    assert_eq!(
        snap[1],
        calcvba_differential::Canon::Bool(true),
        "local UDT RecordSet must keep field storage in place: {snap:?}"
    );
    assert_eq!(
        snap[2],
        calcvba_differential::Canon::Bool(true),
        "UDT array element field update must not reallocate sibling String BSTR: {snap:?}"
    );
    assert_eq!(
        snap[3],
        calcvba_differential::Canon::Bool(true),
        "UDT array element field update must mutate element in place: {snap:?}"
    );
    assert_eq!(
        snap[4],
        calcvba_differential::Canon::Bool(true),
        "ByRef UDT RecordSet must not reallocate sibling String BSTR: {snap:?}"
    );
    assert_eq!(
        snap[5],
        calcvba_differential::canon(&calcvba_runtime::Variant::from_f64(
            10.0 + 25.5 + 2.5 + 1.0 + (51.0 * 52.0 / 2.0) * 0.25
        ))
    );
    if let Some(balance) = vm3.handle_balance {
        assert!(balance.is_zero(), "handle leak: {balance:?}");
    }
}

#[test]
fn vm3_jit_parity_loops_arith_edge_cases_and_fused_branches() {
    assert_vm3_jit_match(
        "long_and_double_loops_and_div_edges",
        r#"
Public sumLong As Long
Public sumDbl As Double
Public errZeroZero As Long
Public errDblEmpty As Long
Public errIntDivMin As Long
Public modMin As Long
Public nullBranchTaken As Long
Public strCmpRes As Boolean

Sub Main()
    Dim i As Long, acc As Long
    For i = 1 To 500
        acc = (acc + (i Mod 17) * 3) Mod 1000003
    Next i
    sumLong = acc

    Dim x As Double, dacc As Double
    For i = 1 To 200
        x = i * 0.5
        dacc = dacc + Sqr(x) / (1# + x)
    Next i
    sumDbl = dacc

    On Error Resume Next
    Dim z As Variant
    Err.Clear
    z = 0 / 0
    errZeroZero = Err.Number

    Dim e As Variant, d0 As Double
    Err.Clear
    z = d0 / e
    errDblEmpty = Err.Number

    Dim minL As Long, neg1 As Long
    minL = -2147483647 - 1
    neg1 = -1
    Err.Clear
    z = minL \ neg1
    errIntDivMin = Err.Number

    Err.Clear
    modMin = minL Mod neg1
    On Error GoTo 0

    Dim nv As Variant
    nv = Null
    If nv = 1 Then
        nullBranchTaken = 1
    Else
        nullBranchTaken = 2
    End If

    Dim s1 As String, s2 As String
    s1 = "10"
    s2 = "2"
    strCmpRes = (s1 < s2)
End Sub
"#,
    );
}

#[test]
fn vm3_jit_parity_udt_arrays_strings_and_polymorphic_classes() {
    assert_vm3_jit_match(
        "udt_array_errors_and_byref_qsort",
        r#"
Private Type Rec
    Id As Long
    Amount As Double
    Name As String
End Type

Public sumAmt As Double
Public oobErr As Long
Public qsortSum As Double
Public strResult As String
Public oddLenB As Long

Private Sub QSort(a() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Double, t As Double
    i = lo: j = hi: p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p: i = i + 1: Loop
        Do While a(j) > p: j = j - 1: Loop
        If i <= j Then
            t = a(i): a(i) = a(j): a(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QSort a, lo, j
    If i < hi Then QSort a, i, hi
End Sub

Sub Main()
    Dim r() As Rec, i As Long, k As Long
    ReDim r(1 To 10)
    For i = 1 To 10
        r(i).Id = i
        r(i).Name = "R" & i
    Next i
    For i = 1 To 100
        k = 1 + i Mod 10
        r(k).Amount = r(k).Amount + i * 0.25
    Next i
    For i = 1 To 10
        sumAmt = sumAmt + r(i).Amount + Len(r(i).Name)
    Next i

    On Error Resume Next
    Err.Clear
    r(11).Amount = 1#
    oobErr = Err.Number
    On Error GoTo 0

    Dim a() As Double, seed As Long
    ReDim a(1 To 64)
    seed = 12345
    For i = 1 To 64
        seed = (seed * 1103 + 12345) Mod 65521
        a(i) = seed / 7#
    Next i
    QSort a, 1, 64
    For i = 1 To 64 Step 7
        qsortSum = qsortSum + a(i) * i
    Next i

    Dim s As String
    For i = 1 To 40
        s = s & Chr$(65 + i Mod 26)
    Next i
    strResult = UCase$(Mid$(s, 5, 10)) & "|" & Replace(s, "A", "00") & "|" & Trim$(" " & Left$(s, 4) & " ")
    oddLenB = LenB(ChrB(65) & ChrB(66) & ChrB(67))
End Sub
"#,
    );

    let cls_a = r#"
Option Explicit
Private mVal As Double
Public Sub Add(ByVal x As Double)
    mVal = mVal + x
End Sub
Public Property Get Value() As Double
    Value = mVal
End Property
"#;
    let cls_b = r#"
Option Explicit
Private mVal As Double
Public Sub Add(ByVal x As Double)
    mVal = mVal + x * 10#
End Sub
Public Property Get Value() As Double
    Value = mVal
End Property
"#;
    let main_mod = r#"
Option Explicit
Public earlyTotal As Double
Public lateTotal As Double
Public fibRes As Long

Private Function Fib(ByVal n As Long) As Long
    If n < 2 Then
        Fib = n
    Else
        Fib = Fib(n - 1) + Fib(n - 2)
    End If
End Function

Sub Main()
    Dim a As ClsA, b As ClsB, o As Object, i As Long
    Set a = New ClsA
    Set b = New ClsB
    For i = 1 To 50
        a.Add i Mod 7
    Next i
    earlyTotal = a.Value

    For i = 1 To 50
        If i Mod 2 = 0 Then
            Set o = a
        Else
            Set o = b
        End If
        o.Add CDbl(i Mod 5)
    Next i
    lateTotal = a.Value + b.Value
    fibRes = Fib(12)
End Sub
"#;
    assert_vm3_jit_match_modules(
        "polymorphic_late_and_early_class_calls",
        &[
            ("Main", ModuleKind::Procedural, main_mod),
            ("ClsA", ModuleKind::Class, cls_a),
            ("ClsB", ModuleKind::Class, cls_b),
        ],
    );
}

#[test]
fn as_new_local_is_pruned_on_callee_return_before_reusing_frame_depth() {
    assert_vm3_jit_match(
        "as_new_pruned_on_return",
        r#"
Option Explicit
Public countRes As Long
Public modeRes As Long

Private Function MakeAndCount() As Long
    Dim c As New Collection
    c.Add 10
    c.Add 20
    MakeAndCount = c.Count
End Function

Private Function CheckMode(ByVal mode As Long) As Long
    If mode = 1 Then
        CheckMode = 100
    Else
        CheckMode = mode + 42
    End If
End Function

Sub Main()
    countRes = MakeAndCount()
    modeRes = CheckMode(0)
End Sub
"#,
    );
}
