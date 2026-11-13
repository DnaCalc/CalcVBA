//! dc-tey4.5 (Wasabi): a module-level variable declared with a module Enum is a Long, so the
//! typed source-project route verifies (`m_LastError As WasabiError` was an untyped Object).
use calcvba_host::provider::Provider;
use calcvba_host_api::{self as api, RuntimeProvider};
#[path = "support/typed_host.rs"]
#[allow(dead_code)]
mod support;

const SOURCE: &str = "Option Explicit
Public Enum Fail
    F_NONE = 0
    F_BAD = 7
End Enum
Private m_Last As Fail
Private Sub SetIt(ByVal e As Fail)
    Static calls As Long
    calls = calls + 1
    m_Last = e
End Sub
Public Function Go() As Double
    m_Last = F_NONE
    SetIt F_BAD
    Go = m_Last
End Function
";

#[test]
fn module_level_enum_variable_loads_and_holds_a_long() {
    for backend in support::backends() {
        let session = Provider::default()
            .load(&api::Project::module("Api", SOURCE), backend, api::HOST_API_ABI)
            .unwrap_or_else(|e| panic!("{backend:?}: {e:?}"));
        let bound = session.bind_slots("Api", "Go", &[], api::ScalarType::Double).unwrap();
        let frame = api::InvocationFrame::new(1, None, std::rc::Rc::new(0u64));
        let mut invocation = session.begin_invocation(&bound, &[], frame, 10_000).unwrap();
        let mut result = None;
        for _ in 0..100_000 {
            match invocation.resume(1).unwrap() {
                api::InvocationProgress::Suspended => {}
                api::InvocationProgress::Returned(api::OwnedValue::Double(v)) => {
                    result = Some(v);
                    break;
                }
                api::InvocationProgress::Returned(other) => panic!("{backend:?}: {other:?}"),
            }
        }
        assert_eq!(result, Some(7.0), "{backend:?}");
    }
}

const ENUMERATOR_CLASS: &str = "Private m_n As Long
Public Property Set EnumVariant(newEnum_ As Object)
    m_n = 1
End Property
Public Property Get EnumVariant() As Object
Attribute EnumVariant.VB_UserMemId = -4
    Set EnumVariant = Nothing
End Property
Public Function Self() As Object
    Set Self = Me
End Function
";

#[test]
fn enumerator_attribute_names_the_property_get_only() {
    // VBA-KeyedCollection's EnumHelper: a Set and a Get of one name; the attribute marks the Get.
    let manifest = calcvba_symbol::manifest::SymbolProjectManifest {
        project_name: "HostProject".into(),
        project_kind: calcvba_symbol::manifest::ProjectKind::Library,
        modules: vec![calcvba_symbol::manifest::ModuleUnit {
            module_name: "EnumHelper".into(),
            module_kind: calcvba_symbol::manifest::ModuleKind::Class,
            attributes: calcvba_symbol::manifest::ModuleAttributes::named("EnumHelper"),
            source: ENUMERATOR_CLASS.into(),
        }],
        references: vec![],
        reference_projects: vec![],
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    calcvba_host::Module::load_project(&manifest, &calcvba_symbol::CatalogTypeLibResolver)
        .expect("a class with an enumerator property pair loads");
}
