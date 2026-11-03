#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_hal::{
    adapters::builder::HostBuilder,
    model::{HalProfileId, HostPolicy},
};
use calcvba_host::{Backend, ProcedureKind, provider::Provider};
use calcvba_host_api as api;
#[path = "support/owned_excel.rs"]
mod owned_excel;
fn project() -> api::Project {
    let mut p = api::Project::module(
        "Api",
        format!(
            "{}\n{}",
            include_str!("support/excel_surface.bas"),
            include_str!("support/excel_chains.bas")
        ),
    );
    p.references.push(api::TypeLibraryReference {
        name: "Excel".into(),
        libid: Some("{00020813-0000-0000-C000-000000000046}".into()),
        major: Some(1),
        minor: Some(9),
        lcid: Some(0),
        import_library: None,
    });
    p
}
fn call(
    session: &calcvba_host::OwnedSession,
    name: &str,
    args: &[api::ValueSlot<'_>],
) -> api::Value {
    let bound = session.bind("Api", name, ProcedureKind::Function).unwrap();
    session
        .call_values(
            &bound,
            args,
            &mut api::TypedContext {
                remaining_blocks: 10000,
            },
        )
        .unwrap_or_else(|e| panic!("{name}: {e:?}"))
}
fn object(value: api::Value) -> api::HostObject {
    let api::Value::Object(Some(value)) = value else {
        panic!("object expected")
    };
    value
}
#[test]
fn object_chains_are_admitted_without_backend_fallback() {
    let module = Provider::default().compile(&project()).unwrap();
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let session = module.clone().into_session(backend).unwrap();
        for (name, ty) in [
            ("ObjectChains", api::ValueType::HostObject),
            ("VariantChains", api::ValueType::Variant),
        ] {
            let bound = session.bind("Api", name, ProcedureKind::Function).unwrap();
            session
                .bind_values(&bound, &[ty], api::ValueType::String)
                .unwrap();
        }
    }
}
#[test]
#[ignore = "requires the owned PID/HWND oracle supervisor"]
fn owned_excel_object_variant_chains() {
    let (_com, pointer) = owned_excel::attach();
    let host = HostBuilder::new()
        .profile(HalProfileId::Windows)
        .policy(HostPolicy::native())
        .build();
    // SAFETY: exact owned NativeOM reference transferred by the checked attach.
    let raw = unsafe {
        host.com()
            .bind_native_dispatch_object_variant("Owned.ExcelWindow", pointer)
    }
    .unwrap();
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let module = Provider::with_host(host.clone())
            .compile(&project())
            .unwrap();
        let interfaces = module
            .com_interfaces()
            .iter()
            .filter_map(|i| match i {
                calcvba_oxir::ComInterface::Com(m) => Some(m.clone()),
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
        // SAFETY: unchanged compiler-owned metadata, retained same-apartment root.
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
            "WindowApplication",
            &[api::ValueSlot::Object(window.argument())],
        ));
        let book = object(call(
            &session,
            "ActiveBook",
            &[api::ValueSlot::Object(app.argument())],
        ));
        let api::Value::String(path) = call(
            &session,
            "BookPath",
            &[api::ValueSlot::Object(book.argument())],
        ) else {
            panic!("path expected")
        };
        assert_eq!(
            path.as_str(),
            std::env::var("OXVBA_OWNED_EXCEL_WORKBOOK").unwrap()
        );
        let sheet = object(call(
            &session,
            "FirstSheet",
            &[api::ValueSlot::Object(book.argument())],
        ));
        let variant = api::Variant::from_object_ref(sheet.object_ref());
        for (name, arg, expected) in [
            (
                "ObjectChains",
                api::ValueSlot::Object(sheet.argument()),
                "A2:12:19:38:9:A2",
            ),
            ("VariantChains", api::ValueSlot::Variant(&variant), "E2:14"),
            (
                "ChainErrors",
                api::ValueSlot::Object(sheet.argument()),
                "91:424:438",
            ),
            (
                "ChainSnapshot",
                api::ValueSlot::Object(sheet.argument()),
                "7:12:19:38:3:12:14",
            ),
        ] {
            let api::Value::String(value) = call(&session, name, &[arg]) else {
                panic!("string expected")
            };
            println!("{backend:?} {name}={}", value.as_str());
            assert_eq!(value.as_str(), expected);
        }
    }
}
