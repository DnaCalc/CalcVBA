//! Defects found by DNA Calc's VBA validation corpus (dc-2z5n), round 4: the built-in
//! Scripting.Dictionary (dc-6eey), a UDT scalar field passed ByRef (dc-9z6t), `TypeOf` on a
//! constant (dc-pye0), library-qualified calls (dc-8f3x). Both executors must agree.

use calcvba_differential::{Canon, Executor, RunOutcome, canon, run};
use calcvba_runtime::Variant;

fn outcome(executor: Executor, source: &str) -> RunOutcome {
    run(executor, source)
}

fn first(executor: Executor, source: &str) -> Canon {
    let outcome = outcome(executor, source);
    assert!(
        outcome.unsupported.is_none(),
        "{executor:?} unsupported: {:?}",
        outcome.unsupported
    );
    let snap = outcome
        .result
        .unwrap_or_else(|e| panic!("{executor:?} run failed: {e}"));
    snap.first()
        .cloned()
        .unwrap_or_else(|| panic!("{executor:?}: empty snapshot"))
}

fn both(source: &str, expected: &Canon) {
    for executor in [Executor::Vm3, Executor::Jit] {
        assert_eq!(&first(executor, source), expected, "{executor:?}");
    }
}

fn both_error(source: &str, number: i32) {
    for executor in [Executor::Vm3, Executor::Jit] {
        let outcome = outcome(executor, source);
        assert!(outcome.raised, "{executor:?}: expected error {number}");
        assert_eq!(outcome.err.number, number, "{executor:?}");
    }
}

#[test]
fn dictionary_late_bound_basic_members() {
    both(
        "Public r As Variant
Sub Main()
    Dim d As Object, k, s As String
    Set d = CreateObject(\"Scripting.Dictionary\")
    d.Add \"b\", 2
    d.Add \"a\", 1
    d(\"c\") = 3
    d.Item(\"a\") = d.Item(\"a\") + 10
    For Each k In d
        s = s & k & d(k) & \";\"
    Next k
    s = s & d.Count & d.Exists(\"A\") & d.Exists(\"a\")
    d.Remove \"b\"
    s = s & d.Count & UBound(d.Keys) & UBound(d.Items)
    d.RemoveAll
    s = s & d.Count & UBound(d.Keys)
    r = s
End Sub
",
        &canon(&Variant::from_string("b2;a11;c3;3FalseTrue2110-1")),
    );
}

#[test]
fn dictionary_early_bound_text_compare_and_missing_item() {
    both(
        "Public r As Variant
Sub Main()
    Dim d As New Scripting.Dictionary
    d.CompareMode = 1
    d.Add \"Key\", 5
    Dim v
    v = d(\"nothing\")
    r = d.Exists(\"KEY\") & d.Count & IsEmpty(v) & d(\"key\")
End Sub
",
        &canon(&Variant::from_string("True2True5")),
    );
}

#[test]
fn dictionary_duplicate_key_raises_457_and_missing_remove_32811() {
    both_error(
        "Sub Main()
    Dim d As Object
    Set d = CreateObject(\"Scripting.Dictionary\")
    d.Add 1, 1
    d.Add 1, 2
End Sub
",
        457,
    );
    both_error(
        "Sub Main()
    Dim d As Object
    Set d = CreateObject(\"Scripting.Dictionary\")
    d.Remove \"x\"
End Sub
",
        32811,
    );
}

#[test]
fn dictionary_holds_objects_and_numeric_keys() {
    both(
        "Public r As Variant
Sub Main()
    Dim d As Object, c As New Collection
    Set d = CreateObject(\"Scripting.Dictionary\")
    c.Add 7
    d.Add 1, \"one\"
    Set d(\"obj\") = c
    r = d.Exists(1&) & d.Exists(\"1\") & d(\"obj\").Count & d(1)
End Sub
",
        &canon(&Variant::from_string("TrueFalse1one")),
    );
}

#[test]
fn fixed_array_assigned_to_dynamic_array_function_result() {
    both(
        "Public r As Variant
Function G() As Double()
    Dim a(1 To 2, 1 To 3) As Double
    a(1, 1) = 1
    a(2, 3) = 6
    G = a
End Function
Sub Main()
    Dim v
    Dim d() As Double
    Dim a(1 To 2) As Double
    a(1) = 5
    d = a
    a(1) = 9
    v = G()
    r = v(2, 3) + 10 * UBound(v, 2) + d(1) * 1000
End Sub
",
        &canon(&Variant::from_f64(5036.0)),
    );
}

#[test]
fn udt_scalar_field_passed_byref_to_a_typed_scalar_parameter() {
    both(
        "Public r As Variant
Private Type VS
    Z As Single
    N As Long
End Type
Private Function Bump(ByVal x As Single, ByRef z As Single) As Single
    z = z + x
    Bump = z
End Function
Private Function Inc(ByRef n As Long) As Long
    n = n + 1
    Inc = n
End Function
Sub Main()
    Dim vs(1 To 2) As VS
    Dim t As VS
    vs(2).Z = 1
    t.N = 5
    r = Bump(1.5!, vs(2).Z) * 100 + Inc(t.N) * 10 + vs(2).Z + t.N
End Sub
",
        &canon(&Variant::from_f64(250.0 + 60.0 + 2.5 + 6.0)),
    );
}

#[test]
fn typeof_on_a_constant_is_false() {
    both(
        "Public r As Variant
Sub Main()
    If TypeOf 5 Is 5 Then
        r = 1
    Else
        r = 0
    End If
End Sub
",
        &canon(&Variant::from_i16(0)),
    );
}

#[test]
fn redim_of_an_indexed_and_a_with_field() {
    both(
        "Public r As Variant
Private Type Conn
    Buf() As Byte
End Type
Dim c(1 To 2) As Conn
Sub Main()
    Dim i As Long
    i = 2
    ReDim c(i).Buf(0 To 4)
    With c(1)
        ReDim .Buf(0 To 2)
    End With
    r = UBound(c(i).Buf) * 10 + UBound(c(1).Buf)
End Sub
",
        &canon(&Variant::from_i32(42)),
    );
}

#[test]
fn dictionary_compare_method_constants_resolve_unqualified() {
    both(
        "Public r As Variant
Sub Main()
    Dim d As Object
    Set d = CreateObject(\"Scripting.Dictionary\")
    d.CompareMode = TextCompare
    d.Add \"A\", 1
    r = d.Exists(\"a\") & d.CompareMode & BinaryCompare
End Sub
",
        &canon(&Variant::from_string("True10")),
    );
}
