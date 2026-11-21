use calcvba_symbol::{
    manifest::{ProjectKind, ProjectReference, SymbolProjectManifest},
    portability::{Host, Support, Verdict, check_project},
};

#[test]
fn office_default_references_use_host_services() {
    let project = SymbolProjectManifest {
        project_name: "Book".into(),
        project_kind: ProjectKind::Host,
        modules: Vec::new(),
        references: [
            ("stdole", "00020430-0000-0000-C000-000000000046"),
            ("Office", "2DF8D04C-5BFA-101B-BDE5-00AA0044DE52"),
        ]
        .into_iter()
        .map(|(name, guid)| ProjectReference::TypeLibrary {
            name: name.into(),
            guid: Some(guid.into()),
            version_major: None,
            version_minor: None,
            lcid: None,
            import_lib: None,
        })
        .collect(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let report = check_project(&[project], Host::Linux).unwrap();
    assert_eq!(report.verdict, Verdict::HostDependentValues);
    assert_eq!(report.findings.len(), 2);
    assert!(
        report
            .findings
            .iter()
            .all(|finding| finding.verdict == Support::Host)
    );
}
