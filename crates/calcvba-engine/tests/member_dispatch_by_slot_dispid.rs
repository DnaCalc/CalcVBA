//! End-to-end VM3 and JIT verification of descriptor-backed member dispatch by DISPID, vtable
//! slot, and `RuntimeDispatchPlanCache` (dc-yqq0.22): late-bound methods/properties/public fields,
//! default members (DISPID 0), `For Each` enumerators (DISPID -4), `CallByName`, `Implements`,
//! and session `invoke_member_values`.

use calcvba_bundle::ProjectMemberKind;
use calcvba_engine::{Engine, ExecBackend, HostConfig};
use calcvba_hal::model::HostPolicy;
use calcvba_oxir::OxImage;
use calcvba_runtime::Variant;
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};

const WIDGET_CLASS: &str = r#"
Public Value As Long
Public Tag As Variant
Public Child As Object
Private mSecret As Long

Private Sub Class_Initialize()
    mSecret = 7
End Sub

Public Function Twice() As Long
    Twice = Value * 2
End Function

Public Property Get Secret() As Long
    Secret = mSecret
End Property

Public Property Let Secret(ByVal v As Long)
    mSecret = v
End Property

Public Property Get DefaultVal() As Long
Attribute DefaultVal.VB_UserMemId = 0
    DefaultVal = Value
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Dim c As New Collection
    c.Add Value
    c.Add Value + 1
    Set NewEnum = c.[_NewEnum]
End Property

Friend Function FriendAdd(ByVal n As Long) As Long
    FriendAdd = Value + n
End Function
"#;

const ISHAPE_CLASS: &str = r#"
Public Function Area() As Long
End Function
"#;

const BOX_CLASS: &str = r#"
Implements IShape
Public Side As Long

Private Function IShape_Area() As Long
    IShape_Area = Side * Side
End Function
"#;

const MAIN_MODULE: &str = r#"
Public OutTwice As Long
Public OutDefault As Long
Public OutEnumSum As Long
Public OutTagText As String
Public OutTagChildTwice As Long
Public OutChildTwice As Long
Public OutCallByName As Long
Public OutSecret As Long
Public OutFriend As Long
Public OutIfaceArea As Long
Public OutLoopSum As Long

Public Sub Main()
    Dim w As New Widget
    Dim o As Object
    Set o = w

    ' Late-bound public field Let + Get + method call (cached through RuntimeDispatchPlanCache).
    o.Value = 21
    OutTwice = o.Twice()

    ' Default member (DISPID 0) in value context.
    OutDefault = o + 100

    ' Custom For Each enumerator (DISPID -4).
    Dim item As Variant
    Dim total As Long
    For Each item In o
        total = total + CLng(item)
    Next item
    OutEnumSum = total

    ' Public Variant field accepts both Let and Set late-bound.
    o.Tag = "ok"
    OutTagText = o.Tag
    Set o.Tag = w
    OutTagChildTwice = o.Tag.Twice()

    ' Public Object field accepts Set and returns object late-bound.
    Set o.Child = w
    OutChildTwice = o.Child.Twice()

    ' CallByName on both public field and method.
    CallByName o, "Value", VbLet, 15
    OutCallByName = CallByName(o, "Value", VbGet) + CallByName(o, "Twice", VbMethod)

    ' Property Let / Get through late-bound dispatch.
    o.Secret = 99
    OutSecret = o.Secret

    ' Friend method callable on typed receiver.
    OutFriend = w.FriendAdd(5)

    ' Implements interface dispatch.
    Dim b As New CBox
    b.Side = 6
    Dim s As IShape
    Set s = b
    OutIfaceArea = s.Area()

    ' Repeated late-bound loop hitting RuntimeDispatchPlanCache.
    Dim i As Long
    Dim acc As Long
    For i = 1 To 10
        acc = acc + o.Twice()
    Next i
    OutLoopSum = acc
End Sub
"#;

fn manifest() -> SymbolProjectManifest {
    SymbolProjectManifest {
        project_name: "DispatchTest".to_string(),
        project_kind: ProjectKind::Source,
        modules: vec![
            ModuleUnit {
                module_name: "MainMod".to_string(),
                module_kind: ModuleKind::Procedural,
                attributes: ModuleAttributes::named("MainMod"),
                source: MAIN_MODULE.to_string(),
            },
            ModuleUnit {
                module_name: "Widget".to_string(),
                module_kind: ModuleKind::Class,
                attributes: ModuleAttributes::named("Widget"),
                source: WIDGET_CLASS.to_string(),
            },
            ModuleUnit {
                module_name: "IShape".to_string(),
                module_kind: ModuleKind::Class,
                attributes: ModuleAttributes::named("IShape"),
                source: ISHAPE_CLASS.to_string(),
            },
            ModuleUnit {
                module_name: "CBox".to_string(),
                module_kind: ModuleKind::Class,
                attributes: ModuleAttributes::named("CBox"),
                source: BOX_CLASS.to_string(),
            },
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
        codepage: None,
    }
}

#[test]
fn late_bound_fields_dispids_enumerator_implements_and_loop_match_on_vm3_and_jit() {
    let project = manifest();
    for (label, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let engine = Engine::new(config);
        let values = engine
            .execute_project_closure_with_variant_snapshot(std::slice::from_ref(&project))
            .unwrap_or_else(|err| panic!("{label}: {err:?}"));
        assert_eq!(values[0].as_i32(), Some(42), "{label} OutTwice");
        assert_eq!(values[1].as_i32(), Some(121), "{label} OutDefault");
        assert_eq!(values[2].as_i32(), Some(43), "{label} OutEnumSum");
        assert_eq!(
            values[3].as_bstr().map(|s| s.as_str()).as_deref(),
            Some("ok"),
            "{label} OutTagText"
        );
        assert_eq!(values[4].as_i32(), Some(42), "{label} OutTagChildTwice");
        assert_eq!(values[5].as_i32(), Some(42), "{label} OutChildTwice");
        assert_eq!(values[6].as_i32(), Some(45), "{label} OutCallByName");
        assert_eq!(values[7].as_i32(), Some(99), "{label} OutSecret");
        assert_eq!(values[8].as_i32(), Some(20), "{label} OutFriend");
        assert_eq!(values[9].as_i32(), Some(36), "{label} OutIfaceArea");
        assert_eq!(values[10].as_i32(), Some(300), "{label} OutLoopSum");
    }
}

#[test]
fn session_invoke_member_dispatches_methods_and_public_fields_on_vm3_and_jit() {
    let mut attributes = ModuleAttributes::named("Widget");
    attributes.vb_exposed = true;
    attributes.vb_creatable = true;
    let project = SymbolProjectManifest {
        project_name: "DispatchSession".to_string(),
        project_kind: ProjectKind::Library,
        modules: vec![ModuleUnit {
            module_name: "Widget".to_string(),
            module_kind: ModuleKind::Class,
            attributes,
            source: WIDGET_CLASS.to_string(),
        }],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
        codepage: None,
    };
    let typelibs = calcvba_symbol::CatalogTypeLibResolver;
    let program = calcvba_bind::bind_program(&project, &typelibs).expect("bind project");
    let oxp = calcvba_oxir::elaborate::elaborate(&program).expect("elaborate to OxIR");

    for backend in [ExecBackend::Vm3, ExecBackend::Jit] {
        let mut engine = Engine::new(HostConfig { backend });
        engine.set_host_policy(HostPolicy::native());
        let mut session = engine
            .prepare_image_session(OxImage::new(vec![oxp.clone()]))
            .unwrap_or_else(|err| panic!("{backend:?} prepare_image_session: {err}"));
        let widget = session
            .create_class_instance("Widget")
            .unwrap_or_else(|err| panic!("{backend:?} create Widget: {err}"));

        // Public field Let + Get through session invoke (plan-cached, no name scan).
        session
            .invoke_member_values(
                widget.clone(),
                "Value",
                Some(ProjectMemberKind::PropertyLet),
                vec![Variant::from_i32(21)],
            )
            .unwrap_or_else(|err| panic!("{backend:?} Value Let: {err}"));
        let value = session
            .invoke_member_values(
                widget.clone(),
                "Value",
                Some(ProjectMemberKind::PropertyGet),
                Vec::new(),
            )
            .unwrap_or_else(|err| panic!("{backend:?} Value Get: {err}"));
        assert_eq!(value.as_i32(), Some(21), "{backend:?} Value Get");

        let twice = session
            .invoke_member_values(widget, "Twice", Some(ProjectMemberKind::Method), Vec::new())
            .unwrap_or_else(|err| panic!("{backend:?} Twice: {err}"));
        assert_eq!(twice.as_i32(), Some(42), "{backend:?} Twice");
    }
}
