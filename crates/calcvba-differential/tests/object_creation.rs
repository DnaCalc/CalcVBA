use calcvba_differential::{Canon, Executor, run};

#[test]
fn computed_dictionary_creation_uses_fresh_initialized_objects_on_both_backends() {
    let source = r#"
Option Explicit
Public result As String
Sub Main()
    Dim class As String, server As String, path As String
    Dim o As Object, other As Object, v As Variant, number As Long
    class = "Scripting." & "Dictionary"
    Set o = CreateObject(class, server)
    o.Add "a", 7
    Set other = GetObject(path, class)
    result = TypeName(o) & ";" & VarType(o) & ";" & o("a") & ";" & other.Count & ";" & (o Is other)
    v = class
    Set other = CreateObject(Class:=v)
    result = result & ";" & other.Count & ";" & (o Is other)
    On Error Resume Next
    number = 123
    Set other = CreateObject(number)
    result = result & ";" & Err.Number
    Err.Clear
    Set other = GetObject()
    result = result & ";" & Err.Number
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
            values.contains(&Canon::Str("Dictionary;9;7;0;False;0;False;429;13".into())),
            "{executor:?}: {values:?}"
        );
    }
}

#[test]
fn project_creation_functions_shadow_library_names_even_for_dictionary_literals() {
    let source = r#"
Option Explicit
Public result As String
Private Function CreateObject(ByVal Class As Variant) As Long
    CreateObject = 41
End Function
Private Function GetObject(ByVal PathName As Variant, ByVal Class As Variant) As Long
    GetObject = 42
End Function
Sub Main()
    result = CreateObject("Scripting.Dictionary") & "," & GetObject("", "Scripting.Dictionary")
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
            values.contains(&Canon::Str("41,42".into())),
            "{executor:?}: {values:?}"
        );
    }
}

#[test]
fn a_public_project_factory_shadows_only_unqualified_library_calls() {
    use calcvba_differential::run_modules;
    use calcvba_symbol::manifest::ModuleKind::Procedural;
    let main = r#"
Public result As String
Sub Main()
    result = CreateObject("Scripting.Dictionary") & "," & TypeName(VBA.CreateObject("Scripting.Dictionary"))
End Sub
"#;
    let factory = r#"
Public Function CreateObject(ByVal Class As Variant) As Long
    CreateObject = 41
End Function
"#;
    for executor in [Executor::Vm3, Executor::Jit] {
        let outcome = run_modules(
            executor,
            &[("Main", Procedural, main), ("Factory", Procedural, factory)],
            "Main",
        );
        assert!(
            outcome.unsupported.is_none(),
            "{executor:?}: {:?}",
            outcome.unsupported
        );
        let values = outcome
            .result
            .unwrap_or_else(|e| panic!("{executor:?}: {e}"));
        assert!(
            values.contains(&Canon::Str("41,Dictionary".into())),
            "{executor:?}: {values:?}"
        );
    }
}
