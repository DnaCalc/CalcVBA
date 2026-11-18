use calcvba_differential::{Canon, Executor, run};

#[test]
fn measured_missing_constants_and_library_state_are_bound_on_both_backends() {
    let source = r#"
Option Explicit
Public result As String
Sub Main()
    Dim forms As Object, again As Object, x As Variant
    Set forms = UserForms
    Set again = UserForms
    result = vbIMENoOp & "," & VBA.VbIMEStatus.vbIMEOn & "," & vbIMEOff
    result = result & ";" & SystemColorConstants.vbMsgBox & "," & VBA.SystemColorConstants.vbMsgBoxText
    result = result & ";" & (forms Is again) & "," & CallByName(forms, "Count", vbGet)
    VBA.Calendar = 1
    result = result & ";" & Calendar
    VBA.DateTime.Calendar = 0
    result = result & "," & VBA.Calendar
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
            values.contains(&Canon::Str(
                "0,1,2;-2147483625,-2147483624;True,0;1,0".into()
            )),
            "{executor:?}: {values:?}"
        );
    }
}
