//! JIT lowers `Debug.Print` (CallNative DebugPrint) through the same runtime
//! helper as vm3. Before this lowering, any program using `Debug.Print` was
//! declined with `RUN-E-JIT-UNSUPPORTED` (M4-4). The web-sourced repro that
//! surfaced the gap was the Rosetta 100-doors sample (bd-ivaha.37).

use calcvba_differential::{canon, parity::compare_entry};
use calcvba_hal::HostOutputChannel;
use calcvba_runtime::Variant;
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};

fn compile(source: &str) -> calcvba_oxir::OxProgram {
    let manifest = SymbolProjectManifest {
        project_name: "VBAProject".into(),
        project_kind: ProjectKind::Source,
        modules: vec![ModuleUnit {
            module_name: "Probe".into(),
            module_kind: ModuleKind::Procedural,
            attributes: ModuleAttributes::named("Probe"),
            source: source.into(),
        }],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let core = calcvba_bind::bind_program(&manifest, &calcvba_symbol::CatalogTypeLibResolver)
        .expect("bind the fixture");
    calcvba_oxir::elaborate::elaborate(&core).expect("elaborate the fixture")
}

fn assert_jit_matches_vm3(source: &str) {
    let comparison = compare_entry(&[compile(source)]);
    for run in [&comparison.vm3, &comparison.jit] {
        assert!(run.outcome.result.is_ok(), "{run:?}\n{source}");
        assert!(run.outcome.unsupported.is_none(), "{run:?}\n{source}");
    }
    assert!(
        comparison.differences().is_empty(),
        "{comparison:?}\n{source}"
    );
}

#[test]
fn debug_print_concatenated_variant_local() {
    assert_jit_matches_vm3(
        r#"
Dim OpenCount As Long
Sub Main()
    Dim i As Long
    Dim Label As String
    OpenCount = 0
    For i = 1 To 10
        Label = "Closed"
        If i = 4 Then
            Label = "Open"
            OpenCount = OpenCount + 1
        End If
        Debug.Print "Door " & i & " is " & Label
    Next i
End Sub
"#,
    );
}

#[test]
fn debug_print_scalar_constants() {
    assert_jit_matches_vm3(
        r#"
Dim Done As Long
Sub Main()
    Debug.Print "plain string"
    Debug.Print 42
    Debug.Print 3.5
    Debug.Print True
    Done = 1
End Sub
"#,
    );
}

#[test]
fn debug_print_blank_line_and_semicolon_args() {
    assert_jit_matches_vm3(
        r#"
Dim Count As Long
Sub Main()
    Debug.Print
    Debug.Print "a"; "b"
    Debug.Print "x"; 7; "y"
    Count = 2
End Sub
"#,
    );
}

#[test]
fn debug_print_does_not_write_a_destination() {
    assert_jit_matches_vm3(
        r#"
Dim Value As Long
Sub Main()
    Value = 41
    Debug.Print "value is " & Value
    Value = Value + 1
End Sub
"#,
    );
}

#[test]
fn compares_globals_locals_err_and_host_output() {
    let programs = [compile(
        r#"
Public Count As Long
Sub Main()
    Dim current As Long
    current = 42
    Count = current + 1
    Debug.Print "before fault"
    On Error Resume Next
    Err.Raise 9
    Debug.Print "after fault"
End Sub
"#,
    )];
    let comparison = compare_entry(&programs);
    for run in [&comparison.vm3, &comparison.jit] {
        let values = run.outcome.result.as_ref().expect("successful entry");
        assert_eq!(
            values,
            &[canon(&Variant::from_i32(43)), canon(&Variant::from_i32(42))]
        );
        assert_eq!(run.outcome.err.number, 9);
        assert_eq!(run.outcome.err.description, "Subscript out of range");
        assert_eq!(
            run.output,
            [
                (HostOutputChannel::Host, "before fault".into()),
                (HostOutputChannel::Host, "after fault".into()),
            ]
        );
    }
    assert!(comparison.differences().is_empty(), "{comparison:?}");
    assert!(!comparison.has_opaque_values());
}

#[test]
fn compares_uncaught_vba_errors_and_preserves_prior_output() {
    let programs = [compile(
        r#"
Sub Main()
    Debug.Print "about to fail"
    Err.Raise 5, "fixture", "controlled error"
End Sub
"#,
    )];
    let comparison = compare_entry(&programs);
    for run in [&comparison.vm3, &comparison.jit] {
        assert!(run.outcome.raised, "{run:?}");
        assert_eq!(run.outcome.err.number, 5);
        assert_eq!(run.outcome.err.source, "fixture");
        assert_eq!(run.outcome.err.description, "controlled error");
        assert_eq!(
            run.output,
            [(HostOutputChannel::Host, "about to fail".into())]
        );
    }
    assert!(comparison.differences().is_empty(), "{comparison:?}");
}
