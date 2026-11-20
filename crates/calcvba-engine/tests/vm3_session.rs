//! W8-c: the host session (`ProjectRuntimeSession`) built from an `OxImage` via
//! `Engine::prepare_image_session` — the path the in-process COM server uses, on vm3 and on
//! the JIT (dc-yqq0.10: the server is no longer vm3 only). Bind a project, then create class
//! instances and invoke their members, with each backend.

use std::cell::RefCell;
use std::rc::Rc;

use calcvba_bundle::ProjectMemberKind;
use calcvba_engine::{Engine, ExecBackend, HostConfig};
use calcvba_hal::model::HostPolicy;
use calcvba_oxir::OxImage;
use calcvba_runtime::Variant;
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};

const BACKENDS: [ExecBackend; 2] = [ExecBackend::Vm3, ExecBackend::Jit];

fn class_module(name: &str, source: &str) -> ModuleUnit {
    let mut attributes = ModuleAttributes::named(name);
    attributes.vb_exposed = true;
    attributes.vb_creatable = true;
    ModuleUnit {
        module_name: name.to_string(),
        module_kind: ModuleKind::Class,
        attributes,
        source: source.to_string(),
    }
}

fn session_for(
    backend: ExecBackend,
    modules: Vec<ModuleUnit>,
) -> calcvba_engine::ProjectRuntimeSession {
    let manifest = SymbolProjectManifest {
        project_name: "DemoServer".to_string(),
        project_kind: ProjectKind::Library,
        modules,
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let typelibs = calcvba_symbol::CatalogTypeLibResolver;
    let program = calcvba_bind::bind_program(&manifest, &typelibs).expect("bind project");
    let oxp = calcvba_oxir::elaborate::elaborate(&program).expect("elaborate to OxIR");
    let mut engine = Engine::new(HostConfig { backend });
    engine.set_host_policy(HostPolicy::native());
    engine
        .prepare_image_session(OxImage::new(vec![oxp]))
        .unwrap_or_else(|err| panic!("prepare {backend:?} image session: {err}"))
}

#[test]
fn image_session_creates_and_invokes_a_class_member() {
    // The build emits an OxImage (.oxi); the COM server loads it into a ProjectRuntimeSession and
    // drives it: create a class instance, invoke a member with by-value args.
    let source = r#"
Option Explicit

Public Function Twice(ByVal n As Long) As Long
    Twice = n + n
End Function

Public Property Get Name() As String
    Name = "Widget"
End Property
"#;
    for backend in BACKENDS {
        let mut session = session_for(backend, vec![class_module("Widget", source)]);
        let widget = session
            .create_class_instance("Widget")
            .unwrap_or_else(|err| panic!("create Widget on {backend:?}: {err}"));
        let twice = session
            .invoke_member_values(
                widget.clone(),
                "Twice",
                Some(ProjectMemberKind::Method),
                vec![Variant::from_i32(21)],
            )
            .unwrap_or_else(|err| panic!("invoke Twice on {backend:?}: {err}"));
        assert_eq!(twice.as_i32(), Some(42), "{backend:?} create + invoke");

        let name = session
            .invoke_member_values(
                widget,
                "Name",
                Some(ProjectMemberKind::PropertyGet),
                Vec::new(),
            )
            .unwrap_or_else(|err| panic!("invoke Name on {backend:?}: {err}"));
        assert_eq!(
            name.as_bstr().map(|s| s.as_str()),
            Some("Widget".to_string()),
            "{backend:?}"
        );
    }
}

#[test]
fn image_session_runs_initialize_keeps_state_and_writes_back_byref() {
    let source = r#"
Option Explicit
Private mCount As Long

Private Sub Class_Initialize()
    mCount = 100
End Sub

Public Function Bump(ByRef target As Long, ByVal by As Long) As Long
    target = target + by
    mCount = mCount + 1
    Bump = mCount
End Function

Public Sub Fail()
    Err.Raise 457, "Widget", "duplicate"
End Sub
"#;
    for backend in BACKENDS {
        let mut session = session_for(backend, vec![class_module("Widget", source)]);
        let widget = session.create_class_instance("Widget").unwrap();
        let mut args = vec![Variant::from_i32(2), Variant::from_i32(5)];
        let count = session
            .invoke_member_values_with_writebacks(
                widget.clone(),
                "Bump",
                Some(ProjectMemberKind::Method),
                &mut args,
                &[0, 1],
            )
            .unwrap_or_else(|err| panic!("Bump on {backend:?}: {err}"));
        assert_eq!(
            count.as_i32(),
            Some(101),
            "{backend:?}: Initialize ran, state kept"
        );
        assert_eq!(args[0].as_i32(), Some(7), "{backend:?}: ByRef written back");
        let count = session
            .invoke_member_values_with_writebacks(
                widget.clone(),
                "Bump",
                Some(ProjectMemberKind::Method),
                &mut args,
                &[0, 1],
            )
            .unwrap();
        assert_eq!(count.as_i32(), Some(102), "{backend:?}: second call");
        assert_eq!(args[0].as_i32(), Some(12), "{backend:?}");

        let error = session
            .invoke_member_values(widget, "Fail", Some(ProjectMemberKind::Method), Vec::new())
            .expect_err("Err.Raise surfaces");
        let fault = error.vba_fault().expect("a VBA fault");
        assert_eq!(fault.code, 457, "{backend:?}");
        assert_eq!(fault.message, "duplicate", "{backend:?}");
    }
}

#[test]
fn image_session_delivers_raise_event_to_the_host_sink() {
    let source = r#"
Option Explicit
Public Event Changed(ByVal value As Long)

Public Sub Fire(ByVal value As Long)
    RaiseEvent Changed(value * 2)
End Sub
"#;
    for backend in BACKENDS {
        let mut session = session_for(backend, vec![class_module("Widget", source)]);
        let seen = Rc::new(RefCell::new(Vec::new()));
        let record = seen.clone();
        session.set_project_event_sink(move |_, _, args| {
            record.borrow_mut().push(args[0].as_i32());
            Ok(())
        });
        let widget = session.create_class_instance("Widget").unwrap();
        session
            .invoke_member_values(
                widget,
                "Fire",
                Some(ProjectMemberKind::Method),
                vec![Variant::from_i32(21)],
            )
            .unwrap_or_else(|err| panic!("Fire on {backend:?}: {err}"));
        assert_eq!(*seen.borrow(), vec![Some(42)], "{backend:?}");
    }
}

#[test]
fn image_session_rejects_an_unknown_class() {
    for backend in BACKENDS {
        let mut session = session_for(
            backend,
            vec![class_module(
                "Widget",
                "Public Function Id() As Long\n    Id = 1\nEnd Function\n",
            )],
        );
        assert!(
            session.create_class_instance("Nope").is_err(),
            "{backend:?}: an undeclared class is can't-create-object"
        );
    }
}
