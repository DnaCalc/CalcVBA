use calcvba_differential::{Canon, Executor, run};

#[test]
fn err_object_aliases_keep_identity_and_use_live_runtime_state() {
    let source = r#"
Option Explicit
Public result As String
Sub Main()
    Dim e As ErrObject, other As Object, boxed As Variant, n As Variant
    Set e = Err
    Set other = Err
    Set boxed = Err
    On Error Resume Next
    Err.Number = 23
    n = Err
    result = TypeName(Err) & "," & IsObject(Err) & "," & VarType(Err)
    result = result & ";" & n & "," & e.Number & "," & other.Number & "," & boxed.Number
    result = result & ";" & (e Is other) & "," & (ObjPtr(e) = ObjPtr(Err))
    CallByName e, "Description", vbLet, "alias"
    result = result & ";" & Err.Description
    CallByName e, "Raise", vbMethod, 5, "AliasSource", "Raised"
    result = result & ";" & e.Number & "," & Err.Source & "," & other.Description
    e.Clear
    result = result & ";" & Err.Number & "," & e.Description
    e.Raise Description:="named", Number:=11
    result = result & ";" & Err.Number & "," & Err.Description
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
                "ErrObject,True,3;23,23,23,23;True,True;alias;5,AliasSource,Raised;0,;11,named"
                    .into()
            )),
            "{executor:?}: {values:?}"
        );
    }
}
