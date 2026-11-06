use calcvba_bundle::ProjectMemberKind;
use calcvba_engine::{Engine, HostConfig, ProjectRuntimeSession};
use calcvba_runtime::{ObjectRef, Variant};
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};
use std::{cell::RefCell, rc::Rc};

fn session() -> ProjectRuntimeSession {
    let mut panel = ModuleAttributes::named("Panel");
    panel.vb_predeclared_id = true;
    panel.vb_exposed = true;
    let mut button = ModuleAttributes::named("Button");
    button.vb_exposed = true;
    let manifest = SymbolProjectManifest {
        project_name: "Forms".into(),
        project_kind: ProjectKind::Library,
        modules: vec![
            ModuleUnit {
                module_name: "Metrics".into(),
                module_kind: ModuleKind::Procedural,
                attributes: ModuleAttributes::named("Metrics"),
                source: "Public initialized As Long\nPublic terminated As Long\n".into(),
            },
            ModuleUnit {
                module_name: "Panel".into(),
                module_kind: ModuleKind::Class,
                attributes: panel,
                source: r#"
Public Event BeforeUnload()
Private number As Long
Private WithEvents child As Button
Private Sub Class_Initialize()
    initialized=initialized+1
    number=initialized
    Set child=New Button
End Sub
Private Sub Class_Terminate()
    terminated=terminated+1
End Sub
Public Function NumberValue() As Long
    NumberValue=number
End Function
Public Sub CloseForm()
    RaiseEvent BeforeUnload
    Set child=Nothing
End Sub
Public Function HeldControl() As Button
    child.BindParent Me
    Set HeldControl=child
End Function
Private Sub child_Click()
    number=number+10
End Sub
"#
                .into(),
            },
            ModuleUnit {
                module_name: "Button".into(),
                module_kind: ModuleKind::Class,
                attributes: button,
                source: r#"Public Event Click()
Private parent As Object
Public Sub BindParent(ByVal value As Object)
    Set parent=value
End Sub
Public Sub DetachParent()
    Set parent=Nothing
End Sub
Public Function Alive() As Long
    Alive=42
End Function
"#
                .into(),
            },
            ModuleUnit {
                module_name: "Driver".into(),
                module_kind: ModuleKind::Class,
                attributes: ModuleAttributes::named("Driver"),
                source: r#"
Public Function DefaultPanel() As Panel
    Set DefaultPanel=Panel
End Function
Public Function Counts() As String
    Counts=CStr(initialized) & ":" & CStr(terminated)
End Function
"#
                .into(),
            },
        ],
        references: vec![],
        reference_projects: vec![],
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let core =
        calcvba_bind::bind_program(&manifest, &calcvba_symbol::CatalogTypeLibResolver).unwrap();
    let program = calcvba_oxir::elaborate::elaborate(&core).unwrap();
    Engine::new(HostConfig::vm3())
        .prepare_image_session(calcvba_oxir::OxImage::new(vec![program]))
        .unwrap()
}
fn invoke(session: &mut ProjectRuntimeSession, object: &ObjectRef, name: &str) -> Variant {
    session
        .invoke_member_values(
            object.clone(),
            name,
            Some(ProjectMemberKind::Method),
            vec![],
        )
        .unwrap()
}
fn counts(session: &mut ProjectRuntimeSession, driver: &ObjectRef, expected: &str) {
    assert_eq!(
        invoke(session, driver, "Counts")
            .as_bstr()
            .unwrap()
            .as_str(),
        expected
    );
}

#[test]
fn noncreating_lookup_reset_and_final_host_release_observe_the_same_default_slot() {
    let mut session = session();
    let access = session.default_instances();
    let callback_access = access.clone();
    let resets = Rc::new(RefCell::new(vec![]));
    let observed = resets.clone();
    session.set_project_event_sink(move |owner, _, _| {
        observed.borrow_mut().push(
            callback_access
                .reset_if_same("Panel", &owner)
                .map_err(|e| e.to_string())?,
        );
        Ok(())
    });
    let driver = session.create_class_instance("Driver").unwrap();
    assert!(access.peek("Panel").unwrap().is_none());
    counts(&mut session, &driver, "0:0");
    let explicit = session.create_class_instance("Panel").unwrap();
    invoke(&mut session, &explicit, "CloseForm");
    assert_eq!(*resets.borrow(), [false]);
    assert!(access.peek("Panel").unwrap().is_none());
    counts(&mut session, &driver, "1:0");
    let held = explicit.clone();
    session.release_class_instance(explicit);
    counts(&mut session, &driver, "1:0");
    session.release_class_instance(held);
    counts(&mut session, &driver, "1:1");

    let default = invoke(&mut session, &driver, "DefaultPanel")
        .as_object_ref()
        .unwrap();
    let peek = access.peek("Forms.Panel").unwrap().unwrap();
    assert_eq!(peek.raw_iunknown(), default.raw_iunknown());
    drop(peek);
    let explicit = session.create_class_instance("Panel").unwrap();
    assert!(!access.reset_if_same("Panel", &explicit).unwrap());
    session.release_class_instance(explicit);
    counts(&mut session, &driver, "3:2");
    invoke(&mut session, &default, "CloseForm");
    assert_eq!(*resets.borrow(), [false, true]);
    assert!(access.peek("Panel").unwrap().is_none());
    counts(&mut session, &driver, "3:2");
    assert_eq!(
        invoke(&mut session, &default, "NumberValue").as_i32(),
        Some(2)
    );
    let replacement = invoke(&mut session, &driver, "DefaultPanel")
        .as_object_ref()
        .unwrap();
    assert_ne!(replacement.raw_iunknown(), default.raw_iunknown());
    assert!(!access.reset_if_same("Panel", &default).unwrap());
    counts(&mut session, &driver, "4:2");
    session.release_class_instance(default);
    counts(&mut session, &driver, "4:3");
    assert!(access.reset_if_same("Panel", &replacement).unwrap());
    session.release_class_instance(replacement);
    counts(&mut session, &driver, "4:4");
    session.release_class_instance(driver);
    drop(session);
    assert!(access.peek("Panel").is_err());
}

#[test]
fn unload_preserves_a_control_that_holds_its_parent_until_that_reference_is_released() {
    let mut session = session();
    let access = session.default_instances();
    let reset = access.clone();
    session.set_project_event_sink(move |owner, _, _| {
        reset
            .reset_if_same("Panel", &owner)
            .map(|_| ())
            .map_err(|e| e.to_string())
    });
    let driver = session.create_class_instance("Driver").unwrap();
    let form = invoke(&mut session, &driver, "DefaultPanel")
        .as_object_ref()
        .unwrap();
    let control = invoke(&mut session, &form, "HeldControl")
        .as_object_ref()
        .unwrap();
    invoke(&mut session, &form, "CloseForm");
    assert!(access.peek("Panel").unwrap().is_none());
    session.release_class_instance(form);
    counts(&mut session, &driver, "1:0");
    assert_eq!(invoke(&mut session, &control, "Alive").as_i32(), Some(42));
    invoke(&mut session, &control, "DetachParent");
    counts(&mut session, &driver, "1:1");
    session.release_class_instance(control);
    session.release_class_instance(driver);
}
