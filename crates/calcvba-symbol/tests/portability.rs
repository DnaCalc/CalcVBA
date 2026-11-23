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
        codepage: None,
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

#[test]
fn browser_reports_missing_modal_and_file_services() {
    use calcvba_symbol::portability::check_source;
    let report = check_source(
        "M",
        r#"
Sub Main()
    MsgBox "hello"
    Dim answer As String
    answer = InputBox("value")
    Open "C:\probe.txt" For Input As #1
    SaveSetting "Probe", "State", "Value", answer
End Sub
"#,
        Host::Browser,
    )
    .unwrap();
    for feature in ["MsgBox", "InputBox", "Open", "SaveSetting"] {
        assert!(
            report
                .findings
                .iter()
                .any(|use_| use_.feature == feature && use_.verdict == Support::Reject),
            "{feature}: {report:?}"
        );
    }
}

#[test]
fn shadowing_across_modules_enums_projects_and_qualifiers_suppresses_false_library_hits() {
    use calcvba_symbol::manifest::{
        ModuleAttributes, ModuleKind, ModuleUnit, ReferencedProjectManifest,
    };

    let helper_project = SymbolProjectManifest {
        project_name: "HelperLib".into(),
        project_kind: ProjectKind::Library,
        modules: vec![
            ModuleUnit {
                module_name: "PublicHelpers".into(),
                module_kind: ModuleKind::Procedural,
                attributes: ModuleAttributes::named("PublicHelpers"),
                source: "Public Const Dir = \"local\"\n\
                         Public Enum CustomMode\n\
                             Shell = 1\n\
                         End Enum\n"
                    .into(),
            },
            ModuleUnit {
                module_name: "HiddenHelpers".into(),
                module_kind: ModuleKind::Procedural,
                attributes: ModuleAttributes::named("HiddenHelpers"),
                source: "Option Private Module\n\
                         Public Function Environ(ByVal x As String) As String\n\
                             Environ = x\n\
                         End Function\n"
                    .into(),
            },
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
        codepage: None,
    };

    let main_project = SymbolProjectManifest {
        project_name: "MainProj".into(),
        project_kind: ProjectKind::Source,
        modules: vec![
            ModuleUnit {
                module_name: "SharedVars".into(),
                module_kind: ModuleKind::Procedural,
                attributes: ModuleAttributes::named("SharedVars"),
                source: "Public Now As Long\n\
                         Public Enum LocalFlags\n\
                             SendKeys = 2\n\
                         End Enum\n"
                    .into(),
            },
            ModuleUnit {
                module_name: "Main".into(),
                module_kind: ModuleKind::Procedural,
                attributes: ModuleAttributes::named("Main"),
                source: "Sub Test()\n\
                             Dim FileSystem As Object\n\
                             Dim x As Variant\n\
                             x = Dir\n\
                             x = Shell\n\
                             x = Now\n\
                             x = SendKeys\n\
                             x = FileSystem.Kill\n\
                             x = Environ(\"PATH\")\n\
                             x = VBA.FileSystem.Dir(\"a\")\n\
                         End Sub\n"
                    .into(),
            },
        ],
        references: vec![ProjectReference::Project {
            referenced_project_name: "HelperLib".into(),
        }],
        reference_projects: vec![ReferencedProjectManifest {
            project_name: helper_project.project_name.clone(),
            project_kind: helper_project.project_kind,
            modules: helper_project.modules.clone(),
            conditional_compilation_target: None,
        }],
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
        codepage: None,
    };

    let report = check_project(&[helper_project, main_project], Host::Linux).unwrap();
    let main_findings: Vec<_> = report
        .findings
        .iter()
        .filter(|f| f.project == "MainProj" && f.module.as_deref() == Some("Main"))
        .collect();

    // Shadowed symbols (Dir, Shell, Now, SendKeys, local FileSystem.Kill) must NOT be flagged.
    for shadowed in ["Shell", "Now", "SendKeys", "Kill"] {
        assert!(
            !main_findings.iter().any(|f| f.feature == shadowed),
            "shadowed {shadowed} should not be reported: {main_findings:#?}"
        );
    }
    // Only the explicitly VBA-qualified Dir on line 10 should be reported, not line 4.
    let dir_lines: Vec<_> = main_findings
        .iter()
        .filter(|f| f.feature == "Dir")
        .map(|f| f.line)
        .collect();
    assert_eq!(dir_lines, vec![Some(10)], "{main_findings:#?}");
    // Option Private Module in HelperLib does NOT shadow Environ in MainProj.
    assert!(
        main_findings
            .iter()
            .any(|f| f.feature == "Environ" && f.line == Some(9)),
        "expected Environ on line 9: {main_findings:#?}"
    );
}

#[test]
fn error_lock_unlock_and_createobject_const_and_servername_are_classified() {
    use calcvba_symbol::portability::check_source;

    let report = check_source(
        "M",
        "Private Const DICT_PROG = \"Scripting.Dictionary\"\n\
         Sub Main()\n\
             Error 5\n\
             Lock #1\n\
             Unlock #1\n\
             Dim a As Object, b As Object, c As Object\n\
             Set a = CreateObject(DICT_PROG)\n\
             Set b = CreateObject((\"Scripting.Dictionary\"))\n\
             Set c = CreateObject(\"Scripting.Dictionary\", \"remote-server\")\n\
             Call CreateObject(DICT_PROG)\n\
         End Sub\n",
        Host::Linux,
    )
    .unwrap();

    for (feature, line, verdict) in [
        ("Error", 3, Support::Native),
        ("Lock", 4, Support::Shim),
        ("Unlock", 5, Support::Shim),
        ("CreateObject", 7, Support::Replace),
        ("CreateObject", 8, Support::Replace),
        ("CreateObject", 9, Support::Reject),
        ("CreateObject", 10, Support::Replace),
    ] {
        assert!(
            report
                .findings
                .iter()
                .any(|f| { f.feature == feature && f.line == Some(line) && f.verdict == verdict }),
            "missing ({feature}, line {line}, {verdict:?}) in {:#?}",
            report.findings
        );
    }
}

#[test]
fn early_bound_scripting_reference_distinguishes_portable_dictionary_from_other_classes() {
    use calcvba_symbol::manifest::{ModuleAttributes, ModuleKind, ModuleUnit};

    let make_project = |source: &str| SymbolProjectManifest {
        project_name: "P".into(),
        project_kind: ProjectKind::Source,
        modules: vec![ModuleUnit {
            module_name: "M".into(),
            module_kind: ModuleKind::Procedural,
            attributes: ModuleAttributes::named("M"),
            source: source.into(),
        }],
        references: vec![ProjectReference::TypeLibrary {
            name: "Scripting".into(),
            guid: Some("{420B2830-E718-11CF-893D-00A0C9054228}".into()),
            version_major: Some(1),
            version_minor: Some(0),
            lcid: Some(0),
            import_lib: None,
        }],
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
        codepage: None,
    };

    let dict_only = check_project(
        &[make_project(
            "Sub Main()\nDim d As New Scripting.Dictionary\nd.Add \"a\", 1\nEnd Sub\n",
        )],
        Host::Linux,
    )
    .unwrap();
    assert_eq!(
        dict_only.verdict,
        Verdict::PortableWithShims,
        "{dict_only:#?}"
    );
    assert!(
        dict_only
            .findings
            .iter()
            .any(|f| f.feature == "Early-bound reference" && f.verdict == Support::Replace),
        "{dict_only:#?}"
    );

    let text_stream = check_project(
        &[make_project(
            "Sub Main()\nDim ts As Scripting.TextStream\nEnd Sub\n",
        )],
        Host::Linux,
    )
    .unwrap();
    assert_eq!(
        text_stream.verdict,
        Verdict::NeedsWindows,
        "{text_stream:#?}"
    );
    assert!(
        text_stream
            .findings
            .iter()
            .any(|f| f.feature == "Early-bound reference" && f.verdict == Support::Reject),
        "{text_stream:#?}"
    );
}
