//! VBA 7.1 regular-expression classes on both execution backends.
use calcvba_differential::{Canon, Executor, run};

#[test]
fn regexp_creation_typed_results_defaults_named_calls_and_identity() {
    let source = r#"
Public result As String
Sub Main()
    Dim r As RegExp, r2 As VBA.RegExp, late As Object
    Dim mc As VBA.MatchCollection, m As VBA.Match, sm As VBA.SubMatches
    Set r = New RegExp
    Set r2 = New VBA.RegExp
    Set late = CreateObject("VBScript.RegExp")
    r.Pattern = "(a)(b)?"
    r.Global = True
    Set mc = r.Execute(sourceString:="aa")
    Set m = mc.Item(Index:=1)
    Set sm = m.SubMatches
    result = TypeName(r) & "," & TypeName(r2) & "," & TypeName(late)
    result = result & ";" & mc.Count & "," & m.FirstIndex & "," & m.Length
    result = result & ";" & sm.Count & "," & sm(0) & "," & TypeName(sm(1))
    result = result & ";" & m & "," & (mc(1) Is m)
    result = result & ";" & r.Replace(replaceVar:="[$1]", sourceString:="aa")
    result = result & ";" & (TypeOf r Is RegExp) & "," & (TypeOf mc Is VBA.MatchCollection)
    result = result & ";" & CallByName(r, "Test", vbMethod, "a")
    result = result & ";" & CallByName(mc, "Count", vbGet)
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
                "RegExp,RegExp,RegExp;2,1,1;2,a,Empty;a,True;[a][a];True,True;True;2".into()
            )),
            "{executor:?}: {values:?}"
        );
    }
}
