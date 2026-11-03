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
#[test]
fn saved_macro_fixtures_admit_unchanged_bodies_on_both_backends() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/support/xlsm");
    let cases: serde_json::Value =
        serde_json::from_str(include_str!("support/xlsm/cases.json")).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let id = case["id"].as_str().unwrap();
        let mut project = api::Project::module(
            "Api",
            std::fs::read_to_string(root.join(format!("{id}-entry.bas"))).unwrap(),
        );
        let mut documents: Vec<(String, String)> = vec![
            ("ThisWorkbook".into(), "Excel.Workbook".into()),
            ("Sheet1".into(), "Excel.Worksheet".into()),
        ];
        for (name, _) in &documents {
            let source = case["documents"][name]
                .as_str()
                .map(|f| std::fs::read_to_string(root.join(f)).unwrap())
                .unwrap_or_default();
            project.modules.push(api::SourceModule {
                name: name.clone(),
                source,
            });
        }
        let mut classes = Vec::new();
        for extra in case["extras"].as_array().unwrap() {
            let name = extra["name"].as_str().unwrap().to_owned();
            if extra["kind"].as_i64() == Some(2) {
                classes.push(name.clone());
            }
            project.modules.push(api::SourceModule {
                name,
                source: std::fs::read_to_string(root.join(extra["file"].as_str().unwrap()))
                    .unwrap(),
            });
        }
        project.references.push(api::TypeLibraryReference {
            name: "Excel".into(),
            libid: Some("00020813-0000-0000-C000-000000000046".into()),
            major: Some(1),
            minor: Some(9),
            lcid: Some(0),
            import_library: None,
        });
        let result = Provider::default()
            .with_document_modules(std::mem::take(&mut documents))
            .with_class_modules(classes)
            .compile(&project)
            .map_err(|e| format!("{e:?}"))
            .and_then(|module| {
                for backend in [Backend::Interpreter, Backend::Jit]
                    .into_iter()
                    .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
                {
                    let session = module
                        .clone()
                        .into_session(backend)
                        .map_err(|e| format!("{e:?}"))?;
                    let bound = session
                        .bind("Api", "RunProbe", ProcedureKind::Function)
                        .map_err(|e| format!("{e:?}"))?;
                    session
                        .bind_values(&bound, &[], api::ValueType::String)
                        .map_err(|e| format!("{e:?}"))?;
                }
                Ok(())
            });
        if let Err(error) = result {
            failures.push(format!("{id}: {error:?}"));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn redim_preserve_keeps_coordinates_and_failed_resize_keeps_allocation() {
    let source = r#"
Public Function RunProbe() As String
 Dim a() As Long, v As Variant
 ReDim a(2 To 3, -1 To 0)
 a(2, -1) = 11: a(2, 0) = 12: a(3, -1) = 21: a(3, 0) = 22
 ReDim Preserve a(2 To 3, -1 To 1)
 v = a
 RunProbe = a(2, -1) & ":" & a(2, 0) & ":" & a(3, -1) & ":" & a(3, 0) & ":" & a(3, 1)
 On Error Resume Next
 ReDim Preserve a(2 To 4, -1 To 1)
 RunProbe = RunProbe & ":" & Err.Number & ":" & UBound(a, 1) & ":" & VarType(v)
End Function
"#;
    let module = Provider::default()
        .compile(&api::Project::module("Api", source))
        .unwrap();
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let session = module.clone().into_session(backend).unwrap();
        let api::Value::String(value) = call(&session, "Api", "RunProbe", &[]) else {
            panic!("string")
        };
        assert_eq!(value.as_str(), "11:12:21:22:0:9:3:8195");
    }
}
#[test]
fn boxed_redim_inherits_type_and_preserve_rejects_type_change() {
    let module = Provider::default()
        .compile(&api::Project::module(
            "Api",
            r#"
Public Function RunProbe() As String
 Dim a() As Long, v As Variant
 ReDim a(1 To 2): a(1) = 6: v = a
 ReDim Preserve v(1 To 3) As Long
 RunProbe = VarType(v) & ":" & v(1) & ":" & UBound(v)
 ReDim v(0 To 1)
 RunProbe = RunProbe & ":" & VarType(v) & ":" & v(0)
 ReDim v(1 To 2) As Double
 On Error Resume Next
 ReDim Preserve v(1 To 3) As Variant
 RunProbe = RunProbe & ":" & Err.Number & ":" & VarType(v) & ":" & UBound(v)
End Function
"#,
        ))
        .unwrap();
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let session = module.clone().into_session(backend).unwrap();
        let api::Value::String(value) = call(&session, "Api", "RunProbe", &[]) else {
            panic!("string");
        };
        assert_eq!(value.as_str(), "8195:6:3:8195:0:13:8197:2");
        let bound = session
            .bind("Api", "RunProbe", ProcedureKind::Function)
            .unwrap();
        let general = session.call_bound(&bound, &[], 100_000).unwrap();
        assert_eq!(
            general.as_bstr().unwrap().as_str(),
            "8195:6:3:8195:0:13:8197:2"
        );
    }
}

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
fn declared_comparison_rules_keep_string_numeric_variant_and_null_distinct() {
    let module = Provider::default()
        .compile(&api::Project::module(
            "Api",
            r#"
Public Function RunProbe() As String
 Dim a As Variant, b As Variant, s As String, n As Long, result As Variant
 a = 2#: s = ""
 RunProbe = CStr(a <> s) & ":" & CStr(s < a)
 a = 5&: b = "4"
 RunProbe = RunProbe & ":" & CStr(a < b) & ":" & CStr(a = "5")
 n = 5
 RunProbe = RunProbe & ":" & CStr(n > b)
 a = Empty: b = "0"
 RunProbe = RunProbe & ":" & CStr(a = b)
 a = Null: result = (a = s)
 RunProbe = RunProbe & ":" & VarType(result)
End Function
"#,
        ))
        .unwrap();
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let session = module.clone().into_session(backend).unwrap();
        let api::Value::String(value) = call(&session, "Api", "RunProbe", &[]) else {
            panic!("string")
        };
        assert_eq!(value.as_str(), "True:True:True:True:True:False:1");
        let bound = session
            .bind("Api", "RunProbe", ProcedureKind::Function)
            .unwrap();
        let general = session.call_bound(&bound, &[], 100_000).unwrap();
        assert_eq!(
            general.as_bstr().unwrap().as_str(),
            "True:True:True:True:True:False:1"
        );
    }
}
#[test]
fn class_array_members_redim_preserve_and_copy_on_both_runtimes() {
    let mut project = api::Project::module(
        "Api",
        r#"
Public Function RunProbe() As String
 Dim data As Payload
 Set data = New Payload
 data.Fill
 data.PutItem 2, 3, 29
 data.Values = data.Matrix
 data.PutItem 1, 1, 99
 RunProbe = data.Values(1, 1) & ":" & data.GetItem(1, 1) & ":" & data.GetItem(2, 3)
End Function
"#,
    );
    project.modules.push(api::SourceModule {
        name: "Payload".into(),
        source: include_str!("support/xlsm/Payload.cls").into(),
    });
    let module = Provider::default()
        .with_class_modules(vec!["Payload".into()])
        .compile(&project)
        .unwrap();
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let session = module.clone().into_session(backend).unwrap();
        let api::Value::String(value) = call(&session, "Api", "RunProbe", &[]) else {
            panic!("string")
        };
        assert_eq!(value.as_str(), "11:99:29");
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
        .policy(HostPolicy::native())
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
    let startup = snapshot
        .get("startup")
        .filter(|v| !v.is_null())
        .map(|startup| {
            let route = module
                .event_routes()
                .iter()
                .find(|route| {
                    route.native.as_ref().is_some_and(|native| {
                        native
                            .name
                            .eq_ignore_ascii_case(startup["event"].as_str().unwrap())
                    })
                })
                .unwrap();
            let native = route.native.as_ref().unwrap();
            let iid =
                oxvba_runtime::RuntimeGuid::parse(native.connection_point_iid.as_deref().unwrap())
                    .unwrap();
            let mut bytes = [0; 16];
            bytes[..4].copy_from_slice(&iid.data1.to_le_bytes());
            bytes[4..6].copy_from_slice(&iid.data2.to_le_bytes());
            bytes[6..8].copy_from_slice(&iid.data3.to_le_bytes());
            bytes[8..].copy_from_slice(&iid.data4);
            (
                startup["module"].as_str().unwrap().to_owned(),
                bytes,
                native.dispatch_member_id.unwrap(),
            )
        });
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
    if let Some((module, iid, dispid)) = startup {
        // Host loading completed before subscriptions could observe the native
        // Open. Deliver its lifecycle notification through compiler routes.
        session
            .deliver_document_event(
                &module,
                iid,
                dispid,
                &mut [],
                &mut api::TypedContext {
                    remaining_blocks: 100_000,
                },
            )
            .unwrap();
    }
    let bound = session
        .bind(
            snapshot["entry_module"].as_str().unwrap(),
            "RunProbe",
            ProcedureKind::Function,
        )
        .unwrap();
    let returned = session.call_values(
        &bound,
        &[],
        &mut api::TypedContext {
            remaining_blocks: 100_000,
        },
    );
    let result = if let Some(expected) = snapshot["expected_error"].as_i64() {
        let error =
            returned.expect_err("unhandled event must abort the caller despite Resume Next");
        let oxvba_embed::Error::Runtime { number, .. } = error else {
            panic!("{error:?}")
        };
        assert_eq!(i64::from(number), expected);
        assert!(matches!(
            session.take_event_error(),
            Some(oxvba_embed::Error::Runtime { number: 11, .. })
        ));
        serde_json::json!({"kind":"RuntimeError","number":number})
    } else {
        let api::Value::String(result) = returned.unwrap() else {
            panic!("string return")
        };
        serde_json::json!({"kind":"String","units":result.to_utf16_units()})
    };
    std::fs::write(
        std::env::var("OXVBA_XLSM_RESULT").unwrap(),
        serde_json::to_vec(&result).unwrap(),
    )
    .unwrap();
}
