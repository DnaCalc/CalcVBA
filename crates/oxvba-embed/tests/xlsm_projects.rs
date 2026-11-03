#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_host_api as api;
use oxvba_embed::{Backend, ProcedureKind, provider::Provider};
use oxvba_hal::{
    adapters::builder::HostBuilder,
    model::{HalProfileId, HostPolicy},
};
#[path = "support/owned_excel.rs"]
mod owned_excel;
const SETUP: &str = r#"
Public Function WindowApplication(ByVal win As Excel.Window) As Excel.Application
 Set WindowApplication = win.Application
End Function
Public Function ActiveBook(ByVal app As Excel.Application) As Excel.Workbook
 Set ActiveBook = app.ActiveWorkbook
End Function
Public Function BookPath(ByVal book As Excel.Workbook) As String
 BookPath = book.FullName
End Function
Public Function SheetForCodeName(ByVal book As Excel.Workbook, ByVal code As String) As Excel.Worksheet
 Dim ws As Excel.Worksheet
 For Each ws In book.Worksheets
  If ws.CodeName = code Then
   Set SheetForCodeName = ws
   Exit Function
  End If
 Next ws
End Function
"#;
fn call(
    session: &oxvba_embed::OwnedSession,
    module: &str,
    name: &str,
    args: &[api::ValueSlot<'_>],
) -> api::Value {
    let bound = session.bind(module, name, ProcedureKind::Function).unwrap();
    session
        .call_values(
            &bound,
            args,
            &mut api::TypedContext {
                remaining_blocks: 100_000,
            },
        )
        .unwrap_or_else(|e| panic!("{module}.{name}: {e:?}"))
}
fn object(value: api::Value) -> api::HostObject {
    let api::Value::Object(Some(object)) = value else {
        panic!("object expected")
    };
    object
}
#[test]
fn internal_byval_variant_calls_copy_typed_values_on_both_runtimes() {
    let project = api::Project::module(
        "Api",
        r#"
Public Function Echo(ByVal boxed As Variant) As String
 Echo = CStr(boxed)
End Function
Public Function RunProbe() As String
 RunProbe = Echo("abc") & ":" & Echo(12&) & ":" & Echo(True)
End Function
"#,
    );
    let module = Provider::default().compile(&project).unwrap();
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let session = module.clone().into_session(backend).unwrap();
        let api::Value::String(result) = call(&session, "Api", "RunProbe", &[]) else {
            panic!("string")
        };
        assert_eq!(result.as_str(), "abc:12:True");
    }
}
#[test]
fn internal_byval_object_arguments_unbox_private_variant_values() {
    let module = Provider::default()
        .compile(&api::Project::module(
            "Api",
            r#"
Public Function Accept(ByVal obj As Object) As Boolean
 Accept = obj Is Nothing
End Function
Public Function RunProbe() As Boolean
 Dim boxed As Variant
 Set boxed = Nothing
 RunProbe = Accept(boxed)
End Function
"#,
        ))
        .unwrap();
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let session = module.clone().into_session(backend).unwrap();
        assert!(matches!(
            call(&session, "Api", "RunProbe", &[]),
            api::Value::Boolean(true)
        ));
    }
}
#[test]
fn fixed_variant_arrays_keep_nonzero_bounds_and_coordinate_order() {
    let project = api::Project::module(
        "Api",
        r#"
Public Function RunArray() As String
 Dim data(2& To 3&, -1& To 0&) As Variant
 data(2&, -1&) = 11&
 data(2&, 0&) = 22&
 data(3&, -1&) = 33&
 data(3&, 0&) = 44&
 RunArray = data(2&, -1&) & ":" & data(2&, 0&) & ":" & data(3&, -1&) & ":" & data(3&, 0&)
 RunArray = RunArray & ":" & LBound(data, 1&) & ":" & UBound(data, 1&)
 RunArray = RunArray & ":" & LBound(data, 2&) & ":" & UBound(data, 2&)
End Function
"#,
    );
    let module = Provider::default().compile(&project).unwrap();
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let session = module.clone().into_session(backend).unwrap();
        let api::Value::String(value) = call(&session, "Api", "RunArray", &[]) else {
            panic!("string")
        };
        assert_eq!(value.as_str(), "11:22:33:44:2:3:-1:0");
    }
}
#[test]
#[ignore = "requires the owned PID/HWND saved-workbook oracle supervisor"]
fn owned_xlsm_project() {
    let snapshot: serde_json::Value = serde_json::from_slice(
        &std::fs::read(std::env::var("OXVBA_XLSM_PROJECT").unwrap()).unwrap(),
    )
    .unwrap();
    let mut project = api::Project::module("CalcVbaHostSetup", SETUP);
    project.name = snapshot["name"].as_str().unwrap().into();
    let mut classes = Vec::new();
    let mut documents = Vec::new();
    for module in snapshot["modules"].as_array().unwrap() {
        let name = module["name"].as_str().unwrap();
        assert_ne!(name, "CalcVbaHostSetup");
        project.modules.push(api::SourceModule {
            name: name.into(),
            source: module["source"].as_str().unwrap().into(),
        });
        match module["kind"].as_i64().unwrap() {
            1 => {}
            2 => classes.push(name.into()),
            100 => documents.push((name.into(), module["host_type"].as_str().unwrap().into())),
            kind => panic!("unsupported actual project module kind {kind}"),
        }
    }
    for reference in snapshot["references"].as_array().unwrap() {
        assert!(!reference["broken"].as_bool().unwrap());
        project.references.push(api::TypeLibraryReference {
            name: reference["name"].as_str().unwrap().into(),
            libid: Some(reference["guid"].as_str().unwrap().into()),
            major: Some(reference["major"].as_u64().unwrap() as u16),
            minor: Some(reference["minor"].as_u64().unwrap() as u16),
            lcid: Some(0),
            import_library: None,
        });
    }
    let (_com, pointer) = owned_excel::attach();
    let host = HostBuilder::new()
        .profile(HalProfileId::Windows)
        .policy(HostPolicy::interactive_dev())
        .build();
    // SAFETY: exact owned NativeOM attachment transfers this same-apartment reference.
    let raw = unsafe {
        host.com()
            .bind_native_dispatch_object_variant("Owned.ExcelWindow", pointer)
    }
    .unwrap();
    let module = Provider::with_host(host.clone())
        .with_class_modules(classes)
        .with_document_modules(documents.clone())
        .compile(&project)
        .unwrap();
    let interfaces = module
        .com_interfaces()
        .iter()
        .filter_map(|i| match i {
            oxvba_oxir::ComInterface::Com(m) => Some(m.clone()),
            _ => None,
        })
        .collect::<Vec<_>>();
    let window = interfaces
        .iter()
        .find(|i| i.name.eq_ignore_ascii_case("Excel.Window"))
        .unwrap();
    let members = window
        .members
        .iter()
        .filter(|m| m.name == "Application")
        .cloned()
        .collect::<Vec<_>>();
    // SAFETY: compiler-owned metadata and retained same-apartment native root.
    let window = unsafe {
        host.com().prepare_typed_object_with_locale(
            raw.as_object_ref().unwrap(),
            &members,
            &interfaces,
            1033,
        )
    }
    .unwrap();
    let backend = match std::env::var("OXVBA_XLSM_BACKEND").unwrap().as_str() {
        "vm3" => Backend::Interpreter,
        "jit" => Backend::Jit,
        _ => panic!("backend"),
    };
    let session = module.into_session(backend).unwrap();
    let app = object(call(
        &session,
        "CalcVbaHostSetup",
        "WindowApplication",
        &[api::ValueSlot::Object(window.argument())],
    ));
    let book = object(call(
        &session,
        "CalcVbaHostSetup",
        "ActiveBook",
        &[api::ValueSlot::Object(app.argument())],
    ));
    let api::Value::String(path) = call(
        &session,
        "CalcVbaHostSetup",
        "BookPath",
        &[api::ValueSlot::Object(book.argument())],
    ) else {
        panic!("path")
    };
    assert_eq!(
        path.as_str(),
        std::env::var("OXVBA_OWNED_EXCEL_WORKBOOK").unwrap()
    );
    for (name, host_type) in documents {
        if host_type == "Excel.Workbook" {
            session.attach_document(&name, &book).unwrap();
        } else {
            let code = api::BStr::from(name.as_str());
            let sheet = object(call(
                &session,
                "CalcVbaHostSetup",
                "SheetForCodeName",
                &[
                    api::ValueSlot::Object(book.argument()),
                    api::ValueSlot::String(&code),
                ],
            ));
            session.attach_document(&name, &sheet).unwrap();
        }
    }
    let api::Value::String(result) = call(
        &session,
        snapshot["entry_module"].as_str().unwrap(),
        "RunProbe",
        &[],
    ) else {
        panic!("string return")
    };
    let result = serde_json::json!({"kind":"String","units":result.to_utf16_units()});
    std::fs::write(
        std::env::var("OXVBA_XLSM_RESULT").unwrap(),
        serde_json::to_vec(&result).unwrap(),
    )
    .unwrap();
}
