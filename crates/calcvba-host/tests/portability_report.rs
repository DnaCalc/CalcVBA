use calcvba_host::portability::{FeatureClass, Host, Support, Verdict};
use calcvba_host::provider::Provider;
use calcvba_host_api as api;

#[test]
fn provider_checks_single_project_and_treats_host_libraries_as_host_services() {
    let mut project = api::Project::module(
        "Main",
        "Sub Run()\nDim s As String\ns = Left$(\"abc\", 1)\nEnd Sub\n",
    );
    project.libraries.push(api::ObjectLibrary {
        name: "HostModel".into(),
        interfaces: vec![api::ObjectInterface {
            name: "Sheet".into(),
            iid: [1; 16],
            members: vec![api::ObjectMember {
                name: "Value".into(),
                slot: 0,
                kind: api::MemberKind::PropertyGet,
                parameters: vec![],
                result: api::ScalarType::Double,
                is_default: true,
            }],
        }],
    });

    let provider = Provider::default();
    let report = provider.check_portability(&project, Host::Linux).unwrap();
    assert_eq!(report.verdict, Verdict::HostDependentValues, "{report:#?}");
    assert!(
        report.findings.iter().any(|f| {
            f.feature == "Early-bound reference"
                && f.class == FeatureClass::Host
                && f.verdict == Support::Host
                && f.detail.as_deref() == Some("HostModel")
        }),
        "{report:#?}"
    );
}

#[test]
fn provider_checks_project_graph_with_cross_project_shadowing() {
    let helper = api::ProjectNode {
        project: api::Project {
            name: "Helper".into(),
            modules: vec![api::SourceModule {
                name: "Utils".into(),
                source: "Public Function Dir() As String\nDir = \"ok\"\nEnd Function\n".into(),
            }],
            libraries: vec![],
            references: vec![],
        },
        references: vec![],
    };
    let entry = api::ProjectNode {
        project: api::Project {
            name: "Entry".into(),
            modules: vec![api::SourceModule {
                name: "Main".into(),
                source: "Sub Run()\nDim x As String\nx = Dir()\nEnd Sub\n".into(),
            }],
            libraries: vec![],
            references: vec![],
        },
        references: vec![api::ProjectReference::Source("Helper".into())],
    };
    let graph = api::ProjectGraph {
        root: "Entry".into(),
        projects: vec![helper, entry],
    };

    let provider = Provider::default();
    let report = provider
        .check_graph_portability(&graph, Host::Linux)
        .unwrap();
    assert_eq!(report.verdict, Verdict::Portable, "{report:#?}");
    assert!(
        !report.findings.iter().any(|f| f.feature == "Dir"),
        "Dir should be shadowed by Helper.Utils.Dir: {report:#?}"
    );
    assert!(
        report.findings.iter().any(|f| {
            f.feature == "Early-bound reference"
                && f.class == FeatureClass::Pure
                && f.verdict == Support::Native
                && f.detail.as_deref() == Some("Helper")
        }),
        "{report:#?}"
    );
}

#[test]
#[cfg(target_os = "linux")]
fn provider_enforces_portability_mode_across_load_graph_and_reload() {
    use calcvba_host::portability::Mode;
    use calcvba_host_api::RuntimeProvider;

    let provider = Provider::default().with_portability_mode(Mode::WindowsFaithful);
    let unsupported = api::Project::module(
        "Main",
        "Private Declare PtrSafe Sub Absent Lib \"kernel32\" ()\n\
         Public Function Answer() As Long\n\
             Answer = 42\n\
         End Function\n",
    );
    for backend in [api::Backend::Interpreter, api::Backend::Jit] {
        let session = provider
            .load(&unsupported, backend, api::HOST_API_ABI)
            .unwrap();
        let handle = session
            .bind("Main", "Answer", api::ProcedureKind::Function)
            .unwrap();
        assert_eq!(
            session.call_bound(&handle, &[], 10_000).unwrap(),
            api::Variant::from_i32(42)
        );
    }

    // Faithful reload publishes the newly compiled code and expires old handles.
    let mut valid = api::Project::module(
        "Main",
        "Public Function Answer() As Long\nAnswer = 7\nEnd Function\n",
    );
    valid.libraries.push(api::ObjectLibrary {
        name: "HostModel".into(),
        interfaces: vec![api::ObjectInterface {
            name: "Sheet".into(),
            iid: [2; 16],
            members: vec![api::ObjectMember {
                name: "Value".into(),
                slot: 0,
                kind: api::MemberKind::PropertyGet,
                parameters: vec![],
                result: api::ScalarType::Double,
                is_default: true,
            }],
        }],
    });
    let mut session = provider
        .load(&valid, api::Backend::Interpreter, api::HOST_API_ABI)
        .expect("host library project should pass windows-faithful");
    let handle = session
        .bind("Main", "Answer", api::ProcedureKind::Function)
        .unwrap();
    assert_eq!(
        session.call_bound(&handle, &[], 10_000).unwrap(),
        api::Variant::from_i32(7)
    );
    session.reload(&unsupported).unwrap();
    assert!(matches!(
        session.call_bound(&handle, &[], 10_000),
        Err(api::Error::StaleHandle)
    ));
    let handle = session
        .bind("Main", "Answer", api::ProcedureKind::Function)
        .unwrap();
    assert_eq!(
        session.call_bound(&handle, &[], 10_000).unwrap(),
        api::Variant::from_i32(42)
    );

    // Faithful mode can load a graph with a dormant unsupported procedure.
    let bad_helper = api::ProjectNode {
        project: api::Project {
            name: "Helper".into(),
            modules: vec![api::SourceModule {
                name: "Utils".into(),
                source: "Public Sub Bad()\nSendKeys \"x\"\nEnd Sub\n".into(),
            }],
            libraries: vec![],
            references: vec![],
        },
        references: vec![],
    };
    let good_helper = api::ProjectNode {
        project: api::Project {
            name: "Helper".into(),
            modules: vec![api::SourceModule {
                name: "Utils".into(),
                source: "Public Function Dir() As String\nDir = \"ok\"\nEnd Function\n".into(),
            }],
            libraries: vec![],
            references: vec![],
        },
        references: vec![],
    };
    let entry = api::ProjectNode {
        project: api::Project {
            name: "Entry".into(),
            modules: vec![api::SourceModule {
                name: "Main".into(),
                source: "Public Function Run() As String\nRun = Dir()\nEnd Function\n".into(),
            }],
            libraries: vec![],
            references: vec![],
        },
        references: vec![api::ProjectReference::Source("Helper".into())],
    };
    let bad_graph = api::ProjectGraph {
        root: "Entry".into(),
        projects: vec![bad_helper, entry.clone()],
    };
    provider
        .load_graph(
            &bad_graph,
            &[],
            api::Backend::Interpreter,
            api::HOST_API_ABI,
        )
        .unwrap();

    // Strict mode allows shadowed Dir() across a project graph, then refuses bad_graph on reload_graph.
    let strict_provider = Provider::default().with_portability_mode(Mode::Strict);
    let good_graph = api::ProjectGraph {
        root: "Entry".into(),
        projects: vec![good_helper, entry],
    };
    let mut graph_session = strict_provider
        .load_graph(
            &good_graph,
            &[],
            api::Backend::Interpreter,
            api::HOST_API_ABI,
        )
        .expect("shadowed Dir should pass strict mode");
    let reload_graph_err = graph_session
        .reload_graph(&bad_graph)
        .unwrap_err()
        .to_string();
    assert!(
        reload_graph_err.contains("PORT-E-UNSUPPORTED") && reload_graph_err.contains("SendKeys"),
        "{reload_graph_err}"
    );
}

#[test]
#[cfg(target_os = "linux")]
fn module_load_with_strict_portability_rejects_unsupported_source() {
    use calcvba_host::Module;
    use calcvba_host::portability::Mode;

    let err = match Module::load_with_portability(
        "Main",
        "Sub Main()\nSendKeys \"a\"\nEnd Sub\n",
        Mode::Strict,
    ) {
        Ok(_) => panic!("expected rejection"),
        Err(err) => err.to_string(),
    };
    assert!(
        err.contains("PORT-E-UNSUPPORTED") && err.contains("SendKeys"),
        "{err}"
    );
}
