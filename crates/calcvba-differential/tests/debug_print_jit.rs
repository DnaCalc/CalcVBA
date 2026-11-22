//! JIT lowers `Debug.Print` (CallNative DebugPrint) through the same runtime
//! helper as vm3. Before this lowering, any program using `Debug.Print` was
//! declined with `RUN-E-JIT-UNSUPPORTED` (M4-4). The web-sourced repro that
//! surfaced the gap was the Rosetta 100-doors sample (bd-ivaha.37).

use calcvba_differential::{Canon, canon, parity::compare_entry};
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

#[test]
fn compares_nested_array_and_udt_values() {
    let programs = [compile(
        r#"
Private Type Sample
    values(1 To 2) As Long
    label As String
End Type
Sub Main()
    Dim item As Sample, nested As Variant
    item.values(2) = 42
    item.label = "sample"
    nested = Array(Array(CLng(42)), CDbl(1.25))
End Sub
"#,
    )];
    let comparison = compare_entry(&programs);
    let values = [&comparison.vm3, &comparison.jit]
        .map(|run| run.outcome.result.as_ref().expect("aggregate snapshot"));
    let [vm3, jit] = values.map(|values| {
        let [Canon::Record { fields, .. }, nested] = values.as_slice() else {
            panic!("expected record and array: {values:?}");
        };
        (fields, nested)
    });
    assert_eq!(vm3, jit);
    assert!(
        matches!(vm3.0.as_slice(), [Canon::Array { bounds: Some(bounds), elements, .. }, Canon::StringBytes(_)]
        if bounds == &[(1, 2)] && elements[1] == canon(&Variant::from_i32(42)))
    );
    assert!(matches!(vm3.1, Canon::Array { elements, .. }
        if matches!(elements.first(), Some(Canon::Array { elements, .. })
            if elements == &[canon(&Variant::from_i32(42))])));
    assert!(!comparison.has_opaque_values());
    assert!(matches!(
        &values[0][0],
        Canon::Record {
            identity: Some(_),
            ..
        }
    ));
    assert!(comparison.differences().is_empty(), "{comparison:?}");
}

#[cfg(target_os = "linux")]
#[test]
fn compares_exact_scratch_file_contents() {
    let programs = [compile(
        r#"
Sub Main()
    MkDir "C:\parity"
    Open "C:\parity\result.txt" For Output As #1
    Print #1, "sample";
    Close #1
End Sub
"#,
    )];
    let comparison = calcvba_differential::parity::compare_entry_with_files(&programs).unwrap();
    for run in [&comparison.vm3, &comparison.jit] {
        assert!(run.outcome.result.is_ok(), "{run:?}");
        assert_eq!(run.outcome.err.number, 0);
        assert_eq!(run.files[std::path::Path::new("parity")], None);
        assert_eq!(
            run.files[std::path::Path::new("parity/result.txt")].as_deref(),
            Some(b"sample".as_slice())
        );
    }
    assert!(comparison.differences().is_empty(), "{comparison:?}");
}

#[test]
fn compares_typed_host_cell_writes() {
    let programs = [compile(
        r#"
Sub Main()
    Dim sheet As Object
    Set sheet = CreateObject("CalcVBA.ParitySheet")
    sheet.WriteCell CLng(2), CLng(3), CDbl(1.25)
    Set sheet = Nothing
End Sub
"#,
    )];
    let comparison = calcvba_differential::parity::compare_entry_with_cells(&programs);
    for run in [&comparison.vm3, &comparison.jit] {
        assert!(run.outcome.result.is_ok(), "{run:?}");
        let [write] = run.cells.as_slice() else {
            panic!("expected one cell write: {run:?}");
        };
        assert_eq!((write.row, write.column), (2, 3));
        assert!(matches!(&write.value, Canon::Raw { tag: 5, bytes, .. }
            if f64::from_ne_bytes(*bytes) == 1.25));
    }
    assert!(comparison.differences().is_empty(), "{comparison:?}");
}
