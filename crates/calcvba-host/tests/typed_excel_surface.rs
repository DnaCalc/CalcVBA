#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_hal::{
    adapters::builder::HostBuilder,
    model::{HalProfileId, HostPolicy},
};
use calcvba_host::{Backend, ProcedureKind, provider::Provider};
use calcvba_host_api as api;
use std::rc::Rc;
#[path = "support/owned_excel.rs"]
mod owned_excel;

fn project() -> api::Project {
    let mut project = api::Project::module("Api", include_str!("support/excel_surface.bas"));
    project.references.push(api::TypeLibraryReference {
        name: "Excel".into(),
        libid: Some("{00020813-0000-0000-C000-000000000046}".into()),
        major: Some(1),
        minor: Some(9),
        lcid: Some(0),
        import_library: None,
    });
    project
}
fn call(
    session: &calcvba_host::OwnedSession,
    name: &str,
    args: &[api::ValueSlot<'_>],
    result: api::ValueType,
) -> api::Value {
    let bound = session.bind("Api", name, ProcedureKind::Function).unwrap();
    session
        .bind_values(
            &bound,
            &args.iter().map(|a| a.ty()).collect::<Vec<_>>(),
            result,
        )
        .unwrap_or_else(|e| panic!("{name}: {e:?}"));
    let mut invocation = session
        .begin_values(
            &bound,
            args,
            api::InvocationFrame::new(13, None, Rc::new(())),
            10000,
        )
        .unwrap();
    loop {
        if let api::ValueProgress::Returned(value) = invocation
            .resume(1)
            .unwrap_or_else(|e| panic!("{name}: {e:?}"))
        {
            return value;
        }
    }
}
fn object(value: api::Value) -> api::HostObject {
    let api::Value::Object(Some(object)) = value else {
        panic!("native object expected")
    };
    object
}
#[test]
fn installed_surface_retains_compiler_metadata() {
    let module = Provider::default().compile(&project()).unwrap();
    let range = module
        .com_interfaces()
        .iter()
        .find(|i| i.name().eq_ignore_ascii_case("Excel.Range"))
        .unwrap();
    let cells = range
        .com_members()
        .unwrap()
        .iter()
        .find(|m| m.name == "Cells")
        .unwrap();
    assert_eq!(
        cells.return_wire_type,
        Some(calcvba_com::TypeLibWireType::InterfacePointer {
            name: "Excel.Range".into()
        })
    );
    assert!(calcvba_com::typed_slots::plan::is_dispatch(cells));
    assert!(module.procedures().any(|p| p.name == "RangeShape"));
    for name in ["Excel.Range", "Excel.Sheets", "Excel.Workbooks"] {
        let interface = module
            .com_interfaces()
            .iter()
            .find(|i| i.name().eq_ignore_ascii_case(name))
            .unwrap();
        assert!(
            interface
                .com_members()
                .unwrap()
                .iter()
                .any(|m| m.token == -4),
            "{name} retains the complete enumerator descriptor"
        );
    }
}
#[test]
#[ignore = "requires the owned PID/HWND oracle supervisor"]
fn owned_excel_surface_from_native_window_uses_prebound_calls() {
    let (_com, pointer) = owned_excel::attach();
    let host = HostBuilder::new()
        .profile(HalProfileId::Windows)
        .policy(HostPolicy::native())
        .build();
    // SAFETY: attach transfers one owned dispatch for the checked owned window.
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
        // SAFETY: unchanged compiler metadata for the live owned window, same STA.
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
            api::ValueType::HostObject,
        ));
        let book = object(call(
            &session,
            "ActiveBook",
            &[api::ValueSlot::Object(app.argument())],
            api::ValueType::HostObject,
        ));
        let path = call(
            &session,
            "BookPath",
            &[api::ValueSlot::Object(book.argument())],
            api::ValueType::String,
        );
        let api::Value::String(path) = path else {
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
            api::ValueType::HostObject,
        ));
        let range = object(call(
            &session,
            "SheetRange",
            &[api::ValueSlot::Object(sheet.argument())],
            api::ValueType::HostObject,
        ));
        for (name, arg, expected, transports) in [
            (
                "FormatKinds",
                range.argument(),
                "7:5|6:5|$0.00|1.2346",
                (0, 13),
            ),
            (
                "RangeShape",
                range.argument(),
                "$F$2|$H$2:$I$3|$F$1:$F$3|2:2:4|F1:G2",
                (0, 13),
            ),
            (
                "RangeFormula",
                range.argument(),
                "=2+3:5:5.00|=G1*2:=RC[1]*2:8",
                (0, 14),
            ),
            (
                "ApplicationOptions",
                app.argument(),
                "False:False|5",
                (10, 0),
            ),
        ] {
            let before = host.com().com_dispatch_transport_counts();
            let value = call(
                &session,
                name,
                &[api::ValueSlot::Object(arg)],
                api::ValueType::String,
            );
            let api::Value::String(value) = value else {
                panic!("String expected")
            };
            println!("{backend:?} {name}={}", value.as_str());
            assert_eq!(value.as_str(), expected, "{backend:?} {name}");
            assert_eq!(
                host.com().com_dispatch_transport_counts(),
                (before.0 + transports.0, before.1 + transports.1),
                "exact native/prebound calls for {name}"
            );
        }
        assert!(matches!(
            call(
                &session,
                "EnsureSheets",
                &[api::ValueSlot::Object(book.argument())],
                api::ValueType::Long
            ),
            api::Value::Long(2)
        ));
        for (name, arg, expected, transports) in [
            (
                "RangeIteration",
                range.argument(),
                "F1:1;G1:11;F2:1;G2:11;nothing=True:9",
                (5, 22),
            ),
            ("RangeFirst", range.argument(), "F1:F1", (1, 3)),
            (
                "SheetIteration",
                book.argument(),
                "2:Run13First;Run13Second;nothing=True",
                (7, 0),
            ),
            ("BookIteration", app.argument(), "1:True", (4, 0)),
            ("SheetItem", book.argument(), "Run13Second", (3, 0)),
        ] {
            let before = host.com().com_dispatch_transport_counts();
            let api::Value::String(value) = call(
                &session,
                name,
                &[api::ValueSlot::Object(arg)],
                api::ValueType::String,
            ) else {
                panic!("String expected")
            };
            println!("{backend:?} {name}={}", value.as_str());
            assert_eq!(value.as_str(), expected);
            assert_eq!(
                host.com().com_dispatch_transport_counts(),
                (before.0 + transports.0, before.1 + transports.1),
                "live enumeration transport {name}"
            );
        }
        let api::Value::String(value) = call(
            &session,
            "BookItem",
            &[
                api::ValueSlot::Object(app.argument()),
                api::ValueSlot::Object(book.argument()),
            ],
            api::ValueType::String,
        ) else {
            panic!("String expected")
        };
        assert_eq!(value.as_str(), "oracle-bootstrap.xlsx");
        assert!(matches!(
            call(&session, "NothingIteration", &[], api::ValueType::Long),
            api::Value::Long(424)
        ));
        // Two independently prepared locale views reach the same owned cells.
        // Cross-read them to distinguish setter provenance from read locale.
        // SAFETY: same exact retained descriptors and owned NativeOM receiver.
        let default_window = unsafe {
            host.com().prepare_typed_object_with_locale(
                raw.as_object_ref().unwrap(),
                &members,
                &interfaces,
                0x400,
            )
        }
        .unwrap();
        let default_app = object(call(
            &session,
            "WindowApplication",
            &[api::ValueSlot::Object(default_window.argument())],
            api::ValueType::HostObject,
        ));
        let default_book = object(call(
            &session,
            "ActiveBook",
            &[api::ValueSlot::Object(default_app.argument())],
            api::ValueType::HostObject,
        ));
        let default_sheet = object(call(
            &session,
            "FirstSheet",
            &[api::ValueSlot::Object(default_book.argument())],
            api::ValueType::HostObject,
        ));
        let default_range = object(call(
            &session,
            "SheetRange",
            &[api::ValueSlot::Object(default_sheet.argument())],
            api::ValueType::HostObject,
        ));
        let address = call(
            &session,
            "RangeShape",
            &[api::ValueSlot::Object(default_range.argument())],
            api::ValueType::String,
        );
        assert!(
            matches!(address, api::Value::String(value) if value.as_str() == "$F$2|$H$2:$I$3|$F$1:$F$3|2:2:4|F1:G2")
        );
        for (writer, expected) in [
            (
                &default_range,
                ["5:5|$0.00|1.234567", "5:5|\\$0.00|1.234567"],
            ),
            (&range, ["6:5|R0.00|1.2346", "6:5|$0.00|1.2346"]),
        ] {
            call(
                &session,
                "SetCurrencyFormat",
                &[api::ValueSlot::Object(writer.argument())],
                api::ValueType::Long,
            );
            for (reader, expected) in [&default_range, &range].into_iter().zip(expected) {
                let api::Value::String(value) = call(
                    &session,
                    "ReadCurrencyFormat",
                    &[api::ValueSlot::Object(reader.argument())],
                    api::ValueType::String,
                ) else {
                    panic!("classification expected")
                };
                println!("{backend:?} format provenance: {}", value.as_str());
                assert_eq!(value.as_str(), expected);
            }
        }
    }
    host.com()
        .release_object_variant(raw.as_object_ref().unwrap())
        .unwrap();
}
