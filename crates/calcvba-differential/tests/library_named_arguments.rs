use calcvba_differential::{Canon, Executor, run};

#[test]
fn reversed_library_names_and_optional_gaps_use_the_declared_parameters() {
    let source = r#"
Option Explicit
Public result As String
Sub Main()
    Dim a As Variant, o As Object
    result = VBA.Left(Length:=2, String:="abcd")
    result = result & ";" & Mid(Start:=2, String:="abcd")
    result = result & ";" & Replace("AaAa", "a", "x", Compare:=vbTextCompare)
    result = result & ";" & InStrRev(Compare:=vbTextCompare, StringMatch:="B", StringCheck:="abcb")
    result = result & ";" & IIf(FalsePart:="wrong", TruePart:="right", Expression:=True)
    result = result & ";" & Year(Date:=DateSerial(Day:=15, Month:=1, Year:=2020))
    result = result & ";" & CLng(Expression:="7") & "," & TypeName(VarName:=1%)
    a = Filter(Compare:=vbTextCompare, SourceArray:=Array("x", "B"), Match:="b")
    result = result & ";" & Join(Delimiter:="|", SourceArray:=a)
    result = result & ";" & SLN(Life:=3, Salvage:=10, Cost:=100)
    Set o = GetObject(Class:="Scripting.Dictionary", PathName:="")
    result = result & ";" & TypeName(o)
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
                "ab;bcd;xxxx;4;right;2020;7,Integer;B;30;Dictionary".into()
            )),
            "{executor:?}: {values:?}"
        );
    }
}
