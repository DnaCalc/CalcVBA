use calcvba_differential::{Canon, Executor, run};

#[test]
fn format_masks_and_null_alias_match_excel_on_both_engines() {
    let source = r#"
Option Explicit
Public result As String, errors As String
Sub Main()
    Dim v As Variant
    result = "[" & Format("hello", "!@@@@@@@@") & "]|" & Format("abc", "\[@\]")
    result = result & "|" & Format("hello", "@ ""suffix""") & "|[" & Format(123, "@@@@@") & "]"
    result = result & "|" & Format(True, "@") & "|" & Format(-5, "@;neg")
    result = result & "|[" & Format(Empty, "0.00") & "]|[" & Format(Null) & "]|[" & Format$(Null, "0") & "]"
    v = 41
    On Error Resume Next
    v = Format(Array(1, 2), "0")
    errors = v & "," & Err.Number
    Err.Clear
    v = Format(CVErr(2007), "0")
    errors = errors & ";" & v & "," & Err.Number
    Err.Clear
    v = Format(5, Null)
    errors = errors & ";" & v & "," & Err.Number
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
        for expected in [
            "[hello   ]|[a]bc|h suffixello|[  123]|True|-5|[]|[]|[]",
            "41,13;41,13;41,94",
        ] {
            assert!(
                values.contains(&Canon::Str(expected.into())),
                "{executor:?}: {values:?}"
            );
        }
    }
}
