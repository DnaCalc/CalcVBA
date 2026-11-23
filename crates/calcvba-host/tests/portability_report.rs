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
