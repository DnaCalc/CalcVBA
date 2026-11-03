use calcvba_host_api::{self as api, RuntimeProvider, SessionExt};
use oxvba_embed::{CodeCache, provider::Provider};
use std::rc::Rc;

fn backends() -> Vec<api::Backend> {
    let mut values = vec![api::Backend::Interpreter];
    if cfg!(feature = "jit") {
        values.push(api::Backend::Jit);
    }
    values
}
#[test]
fn source_project_global_byref_reads_the_owning_unit_store() {
    let graph = api::ProjectGraph {
        root: "App".into(),
        projects: vec![
            node(
                source(
                    "Lib",
                    "Helpers",
                    "Private bias As Long\nPublic Sub Bump(ByRef n As Long)\nbias=1000\nn=n+2\nEnd Sub\nPublic Function State() As Long\nState=bias\nEnd Function\n",
                ),
                &[],
            ),
            node(
                source(
                    "App",
                    "Entry",
                    "Private count As Long\nPublic Function Run() As Long\ncount=40\nHelpers.Bump count\nRun=count\nEnd Function\nPublic Function LibraryState() As Long\nLibraryState=Helpers.State()\nEnd Function\n",
                ),
                &["Lib"],
            ),
        ],
    };
    for backend in backends() {
        let session = Provider::default()
            .load_graph(&graph, &[], backend, api::HOST_API_ABI)
            .unwrap();
        assert_eq!(
            long(run(&*session, "Entry", "Run", &[], api::ValueType::Long)),
            42
        );
        assert_eq!(
            long(run(
                &*session,
                "Entry",
                "LibraryState",
                &[],
                api::ValueType::Long
            )),
            1000
        );
    }
}
fn source(name: &str, module: &str, text: &str) -> api::Project {
    let mut project = api::Project::module(module, text);
    project.name = name.into();
    project
}
fn node(project: api::Project, references: &[&str]) -> api::ProjectNode {
    api::ProjectNode {
        project,
        references: references
            .iter()
            .map(|s| api::ProjectReference::Source((*s).into()))
            .collect(),
    }
}
fn run(
    session: &dyn api::Session,
    module: &str,
    name: &str,
    args: &[api::ValueSlot<'_>],
    result: api::ValueType,
) -> api::Value {
    let types = args.iter().map(|value| value.ty()).collect::<Vec<_>>();
    let bound = session
        .bind_values(module, name, &types, result)
        .unwrap_or_else(|e| panic!("bind {module}.{name}: {e}"));
    let mut invocation = session
        .begin_values(
            &bound,
            args,
            api::InvocationFrame::new(10, None, Rc::new(())),
            10000,
        )
        .unwrap();
    for _ in 0..10000 {
        match invocation
            .resume(1)
            .unwrap_or_else(|e| panic!("resume {module}.{name}: {e}"))
        {
            api::ValueProgress::Suspended => {}
            api::ValueProgress::Returned(value) => return value,
        }
    }
    panic!("call never completed")
}
fn long(value: api::Value) -> i32 {
    match value {
        api::Value::Long(v) => v,
        other => panic!("expected Long: {other:?}"),
    }
}
fn text(value: api::Value) -> String {
    match value {
        api::Value::String(v) => v.as_str().to_string(),
        other => panic!("expected String: {other:?}"),
    }
}

#[test]
fn optional_defaults_missing_aliases_and_named_calls_cross_modules() {
    let mut project = source(
        "OptionalApp",
        "Entry",
        r#"
Public Function CallDefault() As Long
    CallDefault = Helpers.Advance()
End Function
Public Function Aliases() As String
    Dim n As Long
    Dim first As Long
    Dim second As Long
    n = 40
    first = Helpers.Advance(n)
    second = Helpers.Advance((n))
    Aliases = first & ":" & second & ":" & n
End Function
Public Function Omitted() As Boolean
    Omitted = Helpers.MissingFlag()
End Function
Public Function PresentEmpty() As Boolean
    PresentEmpty = Helpers.MissingFlag(Empty)
End Function
Public Function NamedCall() As Double
    NamedCall = Helpers.Mix(right:=2.5)
End Function
Public Function HostDefault(Optional ByRef n As Long = 40, Optional ByVal label As String = "x") As String
    n = n + 2
    HostDefault = label & n
End Function
Public Function HostMissing(Optional ByVal value As Variant) As Boolean
    HostMissing = IsMissing(value)
End Function
Public Function ZeroDefault(Optional ByVal value As Long) As Long
    ZeroDefault = value
End Function
"#,
    );
    project.modules.push(api::SourceModule {
        name: "Helpers".into(),
        source: r#"
Public Function Advance(Optional ByRef n As Long = 40) As Long
    n = n + 2
    Advance = n
End Function
Public Function MissingFlag(Optional ByRef value As Variant) As Boolean
    MissingFlag = IsMissing(value)
    value = "written"
End Function
Public Function Mix(Optional ByVal left As Long = 4, Optional ByVal right As Double = 0.5) As Double
    Mix = left + right
End Function
"#
        .into(),
    });
    for backend in backends() {
        let session = Provider::default()
            .load(&project, backend, api::HOST_API_ABI)
            .unwrap();
        for (name, expected) in [("Omitted", true), ("PresentEmpty", false)] {
            let bound = session
                .bind("Entry", name, api::ProcedureKind::Function)
                .unwrap();
            assert_eq!(
                session.call_bound(&bound, &[], 1000).unwrap().as_bool(),
                Some(expected)
            );
        }
        assert_eq!(
            long(run(
                &*session,
                "Entry",
                "CallDefault",
                &[],
                api::ValueType::Long
            )),
            42
        );
        assert_eq!(
            text(run(
                &*session,
                "Entry",
                "Aliases",
                &[],
                api::ValueType::String
            )),
            "42:44:42"
        );
        assert!(matches!(
            run(&*session, "Entry", "Omitted", &[], api::ValueType::Boolean),
            api::Value::Boolean(true)
        ));
        assert!(matches!(
            run(
                &*session,
                "Entry",
                "PresentEmpty",
                &[],
                api::ValueType::Boolean
            ),
            api::Value::Boolean(false)
        ));
        assert!(matches!(
            run(&*session, "Entry", "NamedCall", &[], api::ValueType::Double),
            api::Value::Double(6.5)
        ));
        assert_eq!(
            text(run(
                &*session,
                "Entry",
                "HostDefault",
                &[],
                api::ValueType::String
            )),
            "x42"
        );
        assert_eq!(
            text(run(
                &*session,
                "Entry",
                "HostDefault",
                &[api::ValueSlot::Long(3)],
                api::ValueType::String
            )),
            "x5"
        );
        assert!(matches!(
            run(
                &*session,
                "Entry",
                "HostMissing",
                &[],
                api::ValueType::Boolean
            ),
            api::Value::Boolean(true)
        ));
        let empty = api::Variant::empty();
        assert!(matches!(
            run(
                &*session,
                "Entry",
                "HostMissing",
                &[api::ValueSlot::Variant(&empty)],
                api::ValueType::Boolean
            ),
            api::Value::Boolean(false)
        ));
        assert_eq!(
            long(run(
                &*session,
                "Entry",
                "ZeroDefault",
                &[],
                api::ValueType::Long
            )),
            0
        );
        let typed = session
            .bind_typed::<(), i32>("Entry", "ZeroDefault")
            .unwrap();
        assert_eq!(session.call_typed(&typed, (), 1000).unwrap(), 0);
        assert!(
            session
                .bind_values(
                    "Entry",
                    "HostDefault",
                    &[api::ValueType::Double],
                    api::ValueType::String
                )
                .is_err()
        );
    }
}

fn graph() -> api::ProjectGraph {
    api::ProjectGraph {
        root: "App".into(),
        projects: vec![
            node(
                source(
                    "App",
                    "Entry",
                    r#"
Private counter As Long
Public Function Run(Optional ByVal n As Long = 40) As Long
    counter = counter + 10
    Run = Helpers.Bump(n) + counter
End Function
Public Function Count() As Long
    Count = counter
End Function
Public Function LibraryCount() As Long
    LibraryCount = Helpers.Count()
End Function
Public Function ForeignMissing() As Boolean
    ForeignMissing = Helpers.MissingFlag()
End Function
Public Function ForeignNamed() As Double
    ForeignNamed = Helpers.Mix(right:=2.5)
End Function
Public Function Handled() As String
    Dim n As Long
    On Error GoTo caught
    Helpers.Fail n
    Exit Function
caught:
    Handled = n & ":" & Err.Number & ":" & Err.Source & ":" & Err.Description
End Function
"#,
                ),
                &["Library"],
            ),
            node(
                source(
                    "Library",
                    "Helpers",
                    r#"
Private counter As Long
Public Function Bump(ByRef n As Long, Optional ByVal delta As Long = 2) As Long
    counter = counter + 1
    n = n + delta
    Bump = n + counter
End Function
Public Function Count() As Long
    Count = counter
End Function
Public Function MissingFlag(Optional ByRef value As Variant) As Boolean
    MissingFlag = IsMissing(value)
    value = "written"
End Function
Public Function Mix(Optional ByVal left As Long = 4, Optional ByVal right As Double = 0.5) As Double
    Mix = left + right
End Function
Public Sub Fail(ByRef n As Long)
    n = 41
    Err.Raise 600, "Library", "from library"
End Sub
"#,
                ),
                &[],
            ),
        ],
    }
}

#[test]
fn source_project_calls_keep_stores_aliases_errors_and_old_leases() {
    for backend in backends() {
        let mut graph = graph();
        let mut session = Provider::default()
            .load_graph(&graph, &[], backend, api::HOST_API_ABI)
            .unwrap();
        assert_eq!(
            long(run(&*session, "Entry", "Run", &[], api::ValueType::Long)),
            53
        );
        assert!(matches!(
            run(
                &*session,
                "Entry",
                "ForeignMissing",
                &[],
                api::ValueType::Boolean
            ),
            api::Value::Boolean(true)
        ));
        assert!(matches!(
            run(
                &*session,
                "Entry",
                "ForeignNamed",
                &[],
                api::ValueType::Double
            ),
            api::Value::Double(6.5)
        ));
        let foreign = session
            .bind("Entry", "ForeignMissing", api::ProcedureKind::Function)
            .unwrap();
        assert_eq!(
            session.call_bound(&foreign, &[], 1000).unwrap().as_bool(),
            Some(true)
        );
        assert_eq!(
            long(run(&*session, "Entry", "Run", &[], api::ValueType::Long)),
            64
        );
        assert_eq!(
            long(run(&*session, "Entry", "Count", &[], api::ValueType::Long)),
            20
        );
        assert_eq!(
            long(run(
                &*session,
                "Entry",
                "LibraryCount",
                &[],
                api::ValueType::Long
            )),
            2
        );
        assert_eq!(
            text(run(
                &*session,
                "Entry",
                "Handled",
                &[],
                api::ValueType::String
            )),
            "41:600:Library:from library"
        );
        let bound = session
            .bind_values("Entry", "Run", &[], api::ValueType::Long)
            .unwrap();
        let mut old = session
            .begin_values(
                &bound,
                &[],
                api::InvocationFrame::new(22, None, Rc::new(())),
                10000,
            )
            .unwrap();
        assert!(matches!(
            old.resume(1).unwrap(),
            api::ValueProgress::Suspended
        ));
        assert!(session.reset().is_err());
        graph.projects[1].project.modules[0].source = graph.projects[1].project.modules[0]
            .source
            .replace("n = n + delta", "n = n + delta + 100");
        session.reload_graph(&graph).unwrap();
        assert!(matches!(
            session.call_values(
                &bound,
                &[],
                &mut api::TypedContext {
                    remaining_blocks: 100
                }
            ),
            Err(api::Error::StaleHandle)
        ));
        let old_result = loop {
            if let api::ValueProgress::Returned(value) = old.resume(1).unwrap() {
                break value;
            }
        };
        assert_eq!(long(old_result), 75);
        assert_eq!(
            long(run(&*session, "Entry", "Run", &[], api::ValueType::Long)),
            153
        );
        session.reset().unwrap();
        assert_eq!(
            long(run(&*session, "Entry", "Run", &[], api::ValueType::Long)),
            153
        );
    }
}

#[test]
fn graph_validation_preflight_visibility_and_cache_identity() {
    let provider = Provider::default();
    let mut missing = graph();
    missing.projects.pop();
    assert!(provider.compile_graph(&missing).is_err());
    let mut cyclic = graph();
    cyclic.projects[1]
        .references
        .push(api::ProjectReference::Source("App".into()));
    assert!(provider.compile_graph(&cyclic).is_err());
    let mut duplicate = graph();
    duplicate.projects[0]
        .references
        .push(api::ProjectReference::Source("library".into()));
    assert!(provider.compile_graph(&duplicate).is_err());
    let mut private = graph();
    private.projects[1].project.modules[0].source =
        "Option Private Module\n".to_string() + &private.projects[1].project.modules[0].source;
    assert!(provider.compile_graph(&private).is_err());
    for backend in backends() {
        // Len is now admitted by the owning runtime. Keep its cross-project
        // success separate from the preflight refusal checked below.
        let mut supported = graph();
        supported.projects[1].project.modules[0].source = supported.projects[1].project.modules[0]
            .source
            .replace("n = n + delta", "n = Len(\"supported\")");
        let session = provider
            .load_graph(&supported, &[], backend, api::HOST_API_ABI)
            .unwrap();
        assert_eq!(
            long(run(&*session, "Entry", "Run", &[], api::ValueType::Long)),
            20
        );

        let mut unsupported = graph();
        unsupported.projects[1].project.modules[0].source = unsupported.projects[1].project.modules
            [0]
        .source
        .replace("n = n + delta", "n = Abs(-1)");
        let session = provider
            .load_graph(&unsupported, &[], backend, api::HOST_API_ABI)
            .unwrap();
        let bound = session
            .bind_values("Entry", "Run", &[], api::ValueType::Long)
            .unwrap();
        assert!(matches!(
            session.begin_values(
                &bound,
                &[],
                api::InvocationFrame::new(33, None, Rc::new(())),
                10000
            ),
            Err(api::Error::UnsupportedSignature(_))
        ));
        assert_eq!(
            long(run(&*session, "Entry", "Count", &[], api::ValueType::Long)),
            0
        );
        let backend = match backend {
            api::Backend::Interpreter => oxvba_embed::Backend::Interpreter,
            api::Backend::Jit => oxvba_embed::Backend::Jit,
        };
        let mut cache = CodeCache::new(4, 1_000_000);
        let g = graph();
        let one = cache
            .prepare(provider.compile_graph(&g).unwrap(), backend)
            .unwrap();
        let same = cache
            .prepare(provider.compile_graph(&g).unwrap(), backend)
            .unwrap();
        assert_eq!(one.code_identity(), same.code_identity());
        let mut changed = g;
        changed.projects[1].project.modules[0]
            .source
            .push_str("\nPublic Function Extra() As Long\nExtra=17\nEnd Function\n");
        let other = cache
            .prepare(provider.compile_graph(&changed).unwrap(), backend)
            .unwrap();
        assert_ne!(one.code_identity(), other.code_identity());
    }
}

#[test]
fn diamond_dependencies_share_one_store_and_ordered_references_choose_once() {
    let diamond = api::ProjectGraph {
        root: "app".into(),
        projects: vec![
            node(
                source(
                    "App",
                    "Entry",
                    "Public Function Run() As Long\nRun=LeftHelpers.Take()+RightHelpers.Take()\nEnd Function\n",
                ),
                &["LeftProject", "RightProject"],
            ),
            node(
                source(
                    "LeftProject",
                    "LeftHelpers",
                    "Public Function Take() As Long\nTake=State.NextValue()\nEnd Function\n",
                ),
                &["Leaf"],
            ),
            node(
                source(
                    "RightProject",
                    "RightHelpers",
                    "Public Function Take() As Long\nTake=State.NextValue()\nEnd Function\n",
                ),
                &["Leaf"],
            ),
            node(
                source(
                    "Leaf",
                    "State",
                    "Private counter As Long\nPublic Function NextValue() As Long\ncounter=counter+1\nNextValue=counter\nEnd Function\n",
                ),
                &[],
            ),
        ],
    };
    let mut ordered = api::ProjectGraph {
        root: "App".into(),
        projects: vec![
            node(
                source(
                    "App",
                    "Entry",
                    "Public Function Run() As Long\nRun=Pick()\nEnd Function\n",
                ),
                &["First", "Second"],
            ),
            node(
                source(
                    "First",
                    "FirstHelpers",
                    "Public Function Pick() As Long\nPick=11\nEnd Function\n",
                ),
                &[],
            ),
            node(
                source(
                    "Second",
                    "SecondHelpers",
                    "Public Function Pick() As Long\nPick=22\nEnd Function\n",
                ),
                &[],
            ),
        ],
    };
    for backend in backends() {
        let session = Provider::default()
            .load_graph(&diamond, &[], backend, api::HOST_API_ABI)
            .unwrap();
        assert_eq!(
            long(run(&*session, "Entry", "Run", &[], api::ValueType::Long)),
            3
        );
        assert_eq!(
            long(run(&*session, "Entry", "Run", &[], api::ValueType::Long)),
            7
        );
        for (references, expected) in [(["First", "Second"], 11), (["Second", "First"], 22)] {
            ordered.projects[0].references = references
                .into_iter()
                .map(|name| api::ProjectReference::Source(name.into()))
                .collect();
            let session = Provider::default()
                .load_graph(&ordered, &[], backend, api::HOST_API_ABI)
                .unwrap();
            assert_eq!(
                long(run(&*session, "Entry", "Run", &[], api::ValueType::Long)),
                expected
            );
        }
    }
}

#[test]
fn source_project_calls_retain_nominal_host_interfaces_and_runtime_defaults() {
    let iid = [85; 16];
    let library = api::TypedObjectLibrary {
        name: "HostValues".into(),
        interfaces: vec![api::TypedObjectInterface {
            name: "Cell".into(),
            iid,
            members: vec![api::TypedObjectMember {
                name: "Value".into(),
                slot: 0,
                kind: api::TypedMemberKind::PropertyGet,
                parameters: vec![],
                result: api::ValueType::Double,
                is_default: true,
            }],
        }],
    };
    let graph = api::ProjectGraph {
        root: "App".into(),
        projects: vec![
            node(
                source(
                    "App",
                    "Entry",
                    "Public Function Read(ByVal cell As HostValues.Cell) As Double\nRead=Helpers.Read(cell)\nEnd Function\nPublic Function Echo(ByVal cell As HostValues.Cell) As HostValues.Cell\nSet Echo=Helpers.Echo(cell)\nEnd Function\n",
                ),
                &["Library"],
            ),
            node(
                source(
                    "Library",
                    "Helpers",
                    "Public Function Read(ByVal cell As HostValues.Cell) As Double\nDim value As Object\nSet value=cell\nRead=value+1#\nEnd Function\nPublic Function Echo(ByVal cell As HostValues.Cell) As HostValues.Cell\nSet Echo=cell\nEnd Function\n",
                ),
                &[],
            ),
        ],
    };
    for backend in backends() {
        let count = Rc::new(std::cell::Cell::new(0));
        let calls = count.clone();
        let object = api::HostObject::new(
            iid,
            vec![
                api::HostSlot::typed(vec![], api::ValueType::Double, move |_, _| {
                    calls.set(calls.get() + 1);
                    Ok(api::Value::Double(41.0))
                })
                .unwrap(),
            ],
        );
        let session = Provider::default()
            .load_graph(
                &graph,
                std::slice::from_ref(&library),
                backend,
                api::HOST_API_ABI,
            )
            .unwrap();
        assert!(matches!(
            run(
                &*session,
                "Entry",
                "Read",
                &[api::ValueSlot::Object(object.argument())],
                api::ValueType::Double
            ),
            api::Value::Double(42.0)
        ));
        assert_eq!(count.get(), 1);
        let api::Value::Object(Some(echo)) = run(
            &*session,
            "Entry",
            "Echo",
            &[api::ValueSlot::Object(object.argument())],
            api::ValueType::Interface(iid),
        ) else {
            panic!("nominal object result expected");
        };
        assert!(std::ptr::eq(echo.header(), object.header()));
        assert_eq!(count.get(), 1);
    }
}
