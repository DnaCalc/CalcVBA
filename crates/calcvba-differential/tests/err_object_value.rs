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

#[test]
fn err_raise_source_and_description_preserve_lone_surrogates_on_vm3_and_jit() {
    let source = r#"
Option Explicit
Public directCodes As String
Public inheritedCodes As String
Public aliasCodes As String
Public propCodes As String

Private Function UnitsOf(ByVal s As String) As String
    Dim i As Long, out As String
    For i = 1 To Len(s)
        If i > 1 Then out = out & ","
        out = out & CStr(CLng(AscW(Mid$(s, i, 1))) And &HFFFF&)
    Next i
    UnitsOf = out
End Function

Private Sub RaiseFromCallee(ByVal s As String, ByVal d As String, ByVal h As String)
    Err.Raise 1002, s, d, h
End Sub

Sub Main()
    Dim hi As String, lo As String, endLo As String
    Dim src As String, desc As String, hlp As String
    hi = ChrW(&HD800)
    lo = ChrW(&HDC00)
    endLo = ChrW(&HDFFF)
    src = "S" & hi & "r"
    desc = "D" & lo & endLo
    hlp = "H" & hi & ".chm"

    ' 1. Direct Err.Raise caught by On Error Resume Next, plus Error()
    On Error Resume Next
    Err.Raise 1000, src, desc, hlp
    directCodes = CStr(Err.Number) & "|" & UnitsOf(Err.Source) & "|" & UnitsOf(Err.Description) & "|" & UnitsOf(Err.HelpFile) & "|" & UnitsOf(Error())

    ' 2. Direct Err property writes + Err.Raise inheritance
    Err.Clear
    Err.Source = hi & "A"
    Err.Description = lo & "B"
    Err.HelpFile = endLo & "C"
    Err.Raise 1001
    inheritedCodes = CStr(Err.Number) & "|" & UnitsOf(Err.Source) & "|" & UnitsOf(Err.Description) & "|" & UnitsOf(Err.HelpFile)

    ' 3. Late-bound ErrObject / CallByName path (rt-abi library_object)
    Dim e As ErrObject
    Set e = Err
    e.Clear
    CallByName e, "Raise", vbMethod, 1003, src, desc, hlp
    aliasCodes = CStr(e.Number) & "|" & UnitsOf(e.Source) & "|" & UnitsOf(e.Description) & "|" & UnitsOf(e.HelpFile)

    ' 4. Cross-procedure propagation caught by On Error GoTo
    On Error GoTo Handler
    RaiseFromCallee src, desc, hlp
    Exit Sub
Handler:
    propCodes = CStr(Err.Number) & "|" & UnitsOf(Err.Source) & "|" & UnitsOf(Err.Description) & "|" & UnitsOf(Err.HelpFile)
    Resume Next
End Sub
"#;
    // "S" & &HD800 & "r" -> 83,55296,114
    // "D" & &HDC00 & &HDFFF -> 68,56320,57343
    // "H" & &HD800 & ".chm" -> 72,55296,46,99,104,109
    // &HD800 & "A" -> 55296,65
    // &HDC00 & "B" -> 56320,66
    // &HDFFF & "C" -> 57343,67
    let expected_direct = "1000|83,55296,114|68,56320,57343|72,55296,46,99,104,109|68,56320,57343";
    let expected_inherited = "1001|55296,65|56320,66|57343,67";
    let expected_alias = "1003|83,55296,114|68,56320,57343|72,55296,46,99,104,109";
    let expected_prop = "1002|83,55296,114|68,56320,57343|72,55296,46,99,104,109";

    for executor in [Executor::Vm3, Executor::Jit] {
        let outcome = run(executor, source);
        assert!(
            outcome.unsupported.is_none(),
            "{executor:?}: {:?}",
            outcome.unsupported
        );
        if let Some(balance) = outcome.handle_balance {
            assert!(
                balance.is_zero(),
                "{executor:?}: leaked live handles: {balance:?}"
            );
        }
        let values = outcome
            .result
            .unwrap_or_else(|e| panic!("{executor:?}: {e}"));
        assert!(
            values.contains(&Canon::Str(expected_direct.into())),
            "{executor:?}: missing directCodes `{expected_direct}` in {values:?}"
        );
        assert!(
            values.contains(&Canon::Str(expected_inherited.into())),
            "{executor:?}: missing inheritedCodes `{expected_inherited}` in {values:?}"
        );
        assert!(
            values.contains(&Canon::Str(expected_alias.into())),
            "{executor:?}: missing aliasCodes `{expected_alias}` in {values:?}"
        );
        assert!(
            values.contains(&Canon::Str(expected_prop.into())),
            "{executor:?}: missing propCodes `{expected_prop}` in {values:?}"
        );
    }
}
