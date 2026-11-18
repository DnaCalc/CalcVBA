use calcvba_differential::{Canon, Executor, run};

#[test]
fn shell_keeps_its_double_return_type_with_a_virtual_process_host() {
    // run uses the deterministic host: no OS process is launched. Its task token
    // still crosses the public VBA Shell boundary as Double.
    let source = r#"
Public result As String
Sub Main()
    Dim task As Variant
    task = Shell(PathName:="virtual-process-probe", WindowStyle:=vbHide)
    result = TypeName(task) & "," & (task > 0)
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
            values.contains(&Canon::Str("Double,True".into())),
            "{executor:?}: {values:?}"
        );
    }
}
