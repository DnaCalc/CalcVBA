#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_host_api as api;
use oxvba_embed::{Backend, ProcedureKind, provider::Provider};
use oxvba_hal::{
    adapters::builder::HostBuilder,
    model::{HalProfileId, HostPolicy},
};
use std::{
    cell::{Cell, RefCell},
    rc::Rc,
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
Public Function FirstSheet(ByVal book As Excel.Workbook) As Excel.Worksheet
    Set FirstSheet = book.Worksheets.Item(1)
End Function
Public Function BookPath(ByVal book As Excel.Workbook) As String
    BookPath = book.FullName
End Function
"#;
fn project() -> api::Project {
    let mut p = api::Project::module("Api", include_str!("support/automatic/Api.bas"));
    p.modules.extend([
        api::SourceModule {
            name: "HostSetup".into(),
            source: SETUP.into(),
        },
        api::SourceModule {
            name: "Watcher".into(),
            source: include_str!("support/automatic/Watcher.cls").into(),
        },
        api::SourceModule {
            name: "Sheet1".into(),
            source: include_str!("support/automatic/Sheet1.cls").into(),
        },
        api::SourceModule {
            name: "ThisWorkbook".into(),
            source: include_str!("support/automatic/ThisWorkbook.cls").into(),
        },
    ]);
    p.references.push(api::TypeLibraryReference {
        name: "Excel".into(),
        libid: Some("00020813-0000-0000-C000-000000000046".into()),
        major: Some(1),
        minor: Some(9),
        lcid: Some(0),
        import_library: None,
    });
    p
}
fn configured(provider: Provider) -> Provider {
    provider
        .with_class_modules(vec!["Watcher".into()])
        .with_document_modules(vec![
            ("Sheet1".into(), "Excel.Worksheet".into()),
            ("ThisWorkbook".into(), "Excel.Workbook".into()),
        ])
}
fn provider() -> Provider {
    configured(Provider::default())
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
                remaining_blocks: 20000,
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
const EXPECTED: &str =
    "1:1:WA1;DA1;|nested=3:3:WA1;DA1;WA2;DA2;|rebind=4:WA3;DA3;|release=4:1:True:DA5;|close=1";

#[derive(Default)]
struct IdleSource {
    callback: RefCell<Option<api::object::EventCallback>>,
    subscriptions: Cell<u32>,
    removals: Cell<u32>,
}
struct IdleServices(Rc<IdleSource>);
struct IdleSubscription(Rc<IdleSource>);
impl Drop for IdleSubscription {
    fn drop(&mut self) {
        let callback = self.0.callback.borrow_mut().take();
        drop(callback);
        self.0.removals.set(self.0.removals.get() + 1);
    }
}
impl api::object::HostObjectServices for IdleServices {
    fn subscribe_event(
        &self,
        descriptor: &api::object::EventDescriptor,
        callback: api::object::EventCallback,
    ) -> Result<api::object::EventSubscription, i32> {
        assert_eq!(descriptor.name, "Change");
        assert!(matches!(
            descriptor.parameters.as_slice(),
            [api::object::EventParameter::Object(Some(_))]
        ));
        assert!(self.0.callback.borrow_mut().replace(callback).is_none());
        self.0.subscriptions.set(self.0.subscriptions.get() + 1);
        Ok(api::object::EventSubscription::new(IdleSubscription(
            self.0.clone(),
        )))
    }
}
#[test]
fn idle_document_callbacks_enter_private_handlers_reset_and_drop_cleanly() {
    let mut project = api::Project::module(
        "Api",
        r#"
Public Counter As Long
Public Function ReadCounter() As Long
 ReadCounter = Counter
End Function
"#,
    );
    project.modules.push(api::SourceModule {
        name: "Sheet1".into(),
        source: r#"
Private Sub Worksheet_Change(ByVal Target As Excel.Range)
 Counter = Counter + 1
End Sub
"#
        .into(),
    });
    project.references = excel_reference();
    let module = Provider::default()
        .with_document_modules(vec![("Sheet1".into(), "Excel.Worksheet".into())])
        .compile(&project)
        .unwrap();
    let iid = module
        .com_interfaces()
        .iter()
        .find_map(|interface| match interface {
            oxvba_oxir::ComInterface::Com(metadata)
                if metadata.name.eq_ignore_ascii_case("Excel.Worksheet") =>
            {
                metadata.iid
            }
            _ => None,
        })
        .unwrap_or_else(|| {
            panic!(
                "retained interfaces: {:?}",
                module
                    .com_interfaces()
                    .iter()
                    .map(|i| match i {
                        oxvba_oxir::ComInterface::Com(m) => m.name.as_str(),
                        _ => "project",
                    })
                    .collect::<Vec<_>>()
            )
        });
    let mut bytes = [0; 16];
    bytes[..4].copy_from_slice(&iid.data1.to_le_bytes());
    bytes[4..6].copy_from_slice(&iid.data2.to_le_bytes());
    bytes[6..8].copy_from_slice(&iid.data3.to_le_bytes());
    bytes[8..].copy_from_slice(&iid.data4);
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let state = Rc::new(IdleSource::default());
        // SAFETY: both payload and services retain this same canonical Rc identity;
        // no native pointer or member table is fabricated for the model source.
        let source = unsafe {
            api::HostObject::new_with_services(
                bytes,
                vec![],
                state.clone(),
                Rc::as_ptr(&state).cast(),
                Rc::new(IdleServices(state.clone())),
            )
        };
        let session = module.clone().into_session(backend).unwrap();
        session.attach_document("Sheet1", &source).unwrap();
        for reset in [false, true] {
            if reset {
                session.reset().unwrap();
            }
            for expected in 1..=2 {
                let callback = state.callback.borrow().clone().unwrap();
                assert!(api::with_native_callback_context(|_| ()).is_err());
                callback(&mut [api::Value::Object(None)]).unwrap();
                assert!(
                    matches!(call(&session,"Api","ReadCounter",&[]),api::Value::Long(value) if value==expected)
                );
            }
        }
        assert_eq!(state.subscriptions.get(), 2);
        assert_eq!(state.removals.get(), 1);
        drop(session);
        assert!(state.callback.borrow().is_none());
        assert_eq!(state.removals.get(), 2);
    }
}
fn excel_reference() -> Vec<api::TypeLibraryReference> {
    vec![api::TypeLibraryReference {
        name: "Excel".into(),
        libid: Some("00020813-0000-0000-C000-000000000046".into()),
        major: Some(1),
        minor: Some(9),
        lcid: Some(0),
        import_library: None,
    }]
}
#[test]
fn compiler_retains_document_and_withevents_routes_and_rejects_wrong_handlers() {
    let module = provider().compile(&project()).unwrap();
    assert_eq!(module.event_routes().len(), 3);
    assert_eq!(
        module
            .event_routes()
            .iter()
            .filter(|r| r.document_class.is_some())
            .count(),
        2
    );
    assert!(module.event_routes().iter().all(|r| {
        r.native
            .as_ref()
            .is_some_and(|n| n.connection_point_iid.is_some())
    }));
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let session = module.clone().into_session(backend).unwrap();
        let bound = session
            .bind("Api", "RunProbe", ProcedureKind::Function)
            .unwrap();
        session
            .bind_values(&bound, &[], api::ValueType::String)
            .unwrap();
    }
    for (old, new) in [
        ("ByVal Target As Excel.Range", "ByRef Target As Excel.Range"),
        (
            "ByVal Target As Excel.Range",
            "ByVal Target As Excel.Worksheet",
        ),
    ] {
        let mut p = project();
        p.modules
            .iter_mut()
            .find(|m| m.name == "Sheet1")
            .unwrap()
            .source = p
            .modules
            .iter()
            .find(|m| m.name == "Sheet1")
            .unwrap()
            .source
            .replace(old, new);
        let error = provider()
            .compile(&p)
            .err()
            .expect("incorrect native handler must fail compile");
        assert!(format!("{error:?}").contains("event handler"), "{error:?}");
    }
}
#[test]
#[ignore = "requires the owned PID/HWND oracle supervisor"]
fn owned_excel_automatic_events() {
    let (_com, pointer) = owned_excel::attach();
    let host = HostBuilder::new()
        .profile(HalProfileId::Windows)
        .policy(HostPolicy::interactive_dev())
        .build();
    // SAFETY: the checked owned NativeOM attachment transferred this reference.
    let raw = unsafe {
        host.com()
            .bind_native_dispatch_object_variant("Owned.ExcelWindow", pointer)
    }
    .unwrap();
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let module = configured(Provider::with_host(host.clone()))
            .compile(&project())
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
        // SAFETY: unchanged compiler metadata and this retained same-apartment root.
        let window = unsafe {
            host.com().prepare_typed_object_with_locale(
                raw.as_object_ref().unwrap(),
                &members,
                &interfaces,
                1033,
            )
        }
        .unwrap();
        let session = module.into_session(backend).unwrap();
        let app = object(call(
            &session,
            "HostSetup",
            "WindowApplication",
            &[api::ValueSlot::Object(window.argument())],
        ));
        let book = object(call(
            &session,
            "HostSetup",
            "ActiveBook",
            &[api::ValueSlot::Object(app.argument())],
        ));
        let api::Value::String(path) = call(
            &session,
            "HostSetup",
            "BookPath",
            &[api::ValueSlot::Object(book.argument())],
        ) else {
            panic!("path")
        };
        assert_eq!(
            path.as_str(),
            std::env::var("OXVBA_OWNED_EXCEL_WORKBOOK").unwrap()
        );
        let sheet = object(call(
            &session,
            "HostSetup",
            "FirstSheet",
            &[api::ValueSlot::Object(book.argument())],
        ));
        session.attach_document("ThisWorkbook", &book).unwrap();
        session.attach_document("Sheet1", &sheet).unwrap();
        for run in 0..2 {
            if run == 1 {
                session.reset().unwrap();
            }
            let api::Value::String(result) = call(&session, "Api", "RunProbe", &[]) else {
                panic!("string")
            };
            println!("{backend:?} automatic run={run}: {}", result.as_str());
            assert_eq!(result.as_str(), EXPECTED);
        }
    }
}
