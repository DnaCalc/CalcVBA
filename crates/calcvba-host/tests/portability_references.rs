use calcvba_host::portability::{Host, Support, TypeLibraryState, Verdict};
use calcvba_host::provider::Provider;
use calcvba_host_api as api;

#[test]
fn dormant_unknown_references_are_visible_in_the_project_report() {
    let mut project = api::Project::module("Main", "Public Sub Answer()\nEnd Sub\n");
    for (name, libid) in [
        ("MissingControls", "{1967BEAC-A012-4000-8000-000000000001}"),
        ("MissingTools", "{1967BEAC-A012-4000-8000-000000000002}"),
    ] {
        let mut reference = api::TypeLibraryReference::named(name);
        reference.libid = Some(libid.into());
        project.references.push(reference);
    }

    let provider = Provider::default();
    for host in [Host::Linux, Host::Windows] {
        let report = provider.check_portability(&project, host).unwrap();
        assert_ne!(report.verdict, Verdict::Portable, "{report:#?}");
        let unavailable: Vec<_> = report
            .findings
            .iter()
            .filter(|finding| finding.feature == "Early-bound reference")
            .map(|finding| (finding.verdict, finding.detail.as_deref().unwrap()))
            .collect();
        assert_eq!(unavailable.len(), 2, "{report:#?}");
        assert_eq!(unavailable[0].0, Support::Reject);
        assert_eq!(unavailable[1].0, Support::Reject);
        assert!(unavailable[0].1.contains("MissingControls"));
        assert!(unavailable[1].1.contains("MissingTools"));
        assert_eq!(report.references.len(), 2);
        for (finding, reference) in report.references.iter().zip(&project.references) {
            assert_eq!(finding.name, reference.name);
            assert_eq!(finding.guid, reference.libid);
            assert_eq!(finding.state, TypeLibraryState::Unknown);
        }
    }
}

#[test]
#[cfg(target_os = "linux")]
fn faithful_session_refuses_only_procedures_that_name_an_unknown_library() {
    use calcvba_host::portability::Mode;
    use calcvba_host_api::RuntimeProvider;

    let mut project = api::Project::module(
        "Main",
        r#"Option Explicit
Private entered As Long
Public Function Answer() As Long
    Answer = 42 + entered
End Function
Public Sub NeedLibrary()
    Dim widget As MissingControls.Widget
    entered = 1
    Set widget = New MissingControls.Widget
End Sub
"#,
    );
    let mut reference = api::TypeLibraryReference::named("MissingControls");
    reference.libid = Some("{1967BEAC-A012-4000-8000-000000000001}".into());
    project.references.push(reference);

    let provider = Provider::default().with_portability_mode(Mode::WindowsFaithful);
    for backend in [api::Backend::Interpreter, api::Backend::Jit] {
        let session = provider.load(&project, backend, api::HOST_API_ABI).unwrap();
        let answer = session
            .bind("Main", "Answer", api::ProcedureKind::Function)
            .unwrap();
        let needs_library = session
            .bind("Main", "NeedLibrary", api::ProcedureKind::Sub)
            .unwrap();
        let error = session.call_bound(&needs_library, &[], 10_000).unwrap_err();
        assert!(error.to_string().contains("MissingControls"), "{error}");
        assert_eq!(
            session.call_bound(&answer, &[], 10_000).unwrap(),
            api::Variant::from_i32(42)
        );
    }
}
