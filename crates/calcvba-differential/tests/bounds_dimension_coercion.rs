use calcvba_differential::{Canon, Executor, run};

#[test]
fn array_bound_dimensions_coerce_literals_and_typed_locals_before_writing_results() {
    let source = r#"
Option Explicit
Public result As String
Sub Main()
    Dim m(1 To 3, 0 To 4, 2 To 6) As Long
    Dim d As Double, s As String, v As Variant, x As Long
    result = LBound(m, "3") & "," & UBound(m, 1.5) & "," & UBound(m, 2.5)
    d = 1.5
    s = "3"
    v = 2.5
    result = result & ";" & UBound(m, d) & "," & LBound(m, s) & "," & UBound(m, v)
    x = 41
    On Error Resume Next
    x = LBound(m, Null)
    result = result & ";" & x & "," & Err.Number
    Err.Clear
    x = UBound(m, Empty)
    result = result & ";" & x & "," & Err.Number
    Err.Clear
    x = LBound(m, "abc")
    result = result & ";" & x & "," & Err.Number
End Sub
"#;
    for executor in [Executor::Vm3, Executor::Jit] {
        let outcome = run(executor, source);
        assert!(
            outcome.unsupported.is_none(),
            "{executor:?}: {:?}",
            outcome.unsupported
        );
        let values = outcome
            .result
            .unwrap_or_else(|e| panic!("{executor:?}: {e}"));
        assert!(
            values.contains(&Canon::Str("2,4,4;4,2,4;41,94;41,9;41,13".into())),
            "{executor:?}: {values:?}"
        );
    }
}
