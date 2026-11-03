#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_host_api as api;
use oxvba_com::windows_connection_point::SynchronousEventArgument as Incoming;
use oxvba_embed::{Backend, OwnedSession, ProcedureKind, provider::Provider};
use oxvba_hal::{
    adapters::builder::HostBuilder,
    model::{HalProfileId, HostPolicy},
};
use std::{
    cell::{Cell, RefCell},
    rc::{Rc, Weak},
};
#[path = "support/owned_excel.rs"]
mod owned_excel;
struct Target {
    session: Weak<OwnedSession>,
    sheet: api::HostObject,
    range_iid: api::InterfaceId,
    change: oxvba_embed::BoundProcedure,
    close: oxvba_embed::BoundProcedure,
    changes: Cell<usize>,
    closes: Cell<usize>,
    errors: Cell<usize>,
}
thread_local! {static TARGET: RefCell<Option<Rc<Target>>>=const {RefCell::new(None)};}
struct Scope;
impl Drop for Scope {
    fn drop(&mut self) {
        TARGET.with(|s| s.borrow_mut().take());
    }
}
fn receive_change(args: &[Incoming]) -> Result<(), i32> {
    let target = TARGET.with(|s| s.borrow().clone()).ok_or(5)?;
    target.changes.set(target.changes.get() + 1);
    let result = api::with_native_callback_context(|context| {
        let session = target.session.upgrade().ok_or(api::Error::StaleHandle)?;
        let object = oxvba_com::typed_slots::prepare_typed_event_object(
            &target.sheet,
            &args[0],
            target.range_iid,
        )
        .map_err(api::Error::UnsupportedSignature)?
        .ok_or(api::Error::ArgumentMismatch)?;
        session
            .call_values(
                &target.change,
                &[api::ValueSlot::Object(object.argument())],
                context,
            )
            .map_err(|_| api::Error::ArgumentMismatch)?;
        Ok::<(), api::Error>(())
    })
    .and_then(|r| r);
    if result.is_err() {
        target.errors.set(target.errors.get() + 1);
        return Err(5);
    }
    Ok(())
}
fn receive_close(args: &[Incoming]) -> Result<(), i32> {
    let target = TARGET.with(|s| s.borrow().clone()).ok_or(5)?;
    target.closes.set(target.closes.get() + 1);
    let [Incoming::BooleanByRef(cancel)] = args else {
        return Err(13);
    };
    let mut values = [api::Value::Boolean(cancel.get())];
    let result = api::with_native_callback_context(|context| {
        let session = target.session.upgrade().ok_or(api::Error::StaleHandle)?;
        session
            .call_values_with_writebacks(&target.close, &mut values, context)
            .map_err(|_| api::Error::ArgumentMismatch)
    })
    .and_then(|r| r);
    if let api::Value::Boolean(value) = values[0] {
        cancel.set(value)
    }
    if result.is_err() {
        target.errors.set(target.errors.get() + 1);
        return Err(5);
    }
    Ok(())
}
fn project() -> api::Project {
    let mut p = api::Project::module(
        "Api",
        format!(
            "{}\n{}",
            include_str!("support/excel_events.bas"),
            include_str!("support/excel_surface.bas")
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
fn invoke(
    session: &OwnedSession,
    name: &str,
    args: &[api::ValueSlot<'_>],
    result: api::ValueType,
    context: &mut api::TypedContext,
) -> api::Value {
    let bound = session.bind("Api", name, ProcedureKind::Function).unwrap();
    session
        .bind_values(
            &bound,
            &args.iter().map(|a| a.ty()).collect::<Vec<_>>(),
            result,
        )
        .unwrap_or_else(|e| panic!("{name}: {e:?}"));
    session
        .call_values(&bound, args, context)
        .unwrap_or_else(|e| panic!("{name}: {e:?}"))
}
fn object(v: api::Value) -> api::HostObject {
    let api::Value::Object(Some(v)) = v else {
        panic!("object expected")
    };
    v
}
fn metadata() -> (
    oxvba_com::TypeLibEventMetadata,
    oxvba_com::TypeLibEventMetadata,
) {
    let request = oxvba_com::TypeLibResolveRequest {
        reference_name: "Excel".into(),
        requested_coclass: None,
        importlib_hint: None,
        libid_hint: Some("00020813-0000-0000-C000-000000000046".into()),
        major_version_hint: Some(1),
        minor_version_hint: Some(9),
        lcid_hint: Some(0),
    };
    let get = |name: &str, class: &str| {
        let mut request = request.clone();
        request.requested_coclass = Some(class.into());
        let identity = oxvba_com::resolve_known_typelib_identity(&request).unwrap();
        let blob = oxvba_com::build_typelib_metadata(&identity);
        blob.events
            .iter()
            .find(|e| e.name.eq_ignore_ascii_case(name))
            .unwrap_or_else(|| panic!("missing {class}.{name}; events={:?}", blob.events))
            .clone()
    };
    (get("Change", "Worksheet"), get("BeforeClose", "Workbook"))
}
#[test]
fn installed_event_metadata_is_exact() {
    let (change, close) = metadata();
    assert_eq!(
        change.parameter_types,
        [oxvba_com::TypeLibParamType::Object]
    );
    assert_eq!(
        close.parameter_types,
        [oxvba_com::TypeLibParamType::ByRefBoolean]
    );
    println!("change={change:?}; close={close:?}");
}
#[test]
#[ignore = "requires the owned PID/HWND oracle supervisor"]
fn owned_excel_events_reenter_and_cancel_before_native_return() {
    let (_com, pointer) = owned_excel::attach();
    let (change_event, close_event) = metadata();
    let host = HostBuilder::new()
        .profile(HalProfileId::Windows)
        .policy(HostPolicy::native())
        .build();
    // SAFETY: exact owned NativeOM window reference transferred by attach.
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
        // SAFETY: unchanged exact compiler plans for the owned same-apartment window.
        let window = unsafe {
            host.com().prepare_typed_object_with_locale(
                raw.as_object_ref().unwrap(),
                &members,
                &interfaces,
                1033,
            )
        }
        .unwrap();
        let session = Rc::new(module.into_session(backend).unwrap());
        let mut context = api::TypedContext {
            remaining_blocks: 10000,
        };
        let app = object(invoke(
            &session,
            "WindowApplication",
            &[api::ValueSlot::Object(window.argument())],
            api::ValueType::HostObject,
            &mut context,
        ));
        let book = object(invoke(
            &session,
            "ActiveBook",
            &[api::ValueSlot::Object(app.argument())],
            api::ValueType::HostObject,
            &mut context,
        ));
        let sheet = object(invoke(
            &session,
            "FirstSheet",
            &[api::ValueSlot::Object(book.argument())],
            api::ValueType::HostObject,
            &mut context,
        ));
        let range = object(invoke(
            &session,
            "EventRange",
            &[api::ValueSlot::Object(sheet.argument())],
            api::ValueType::HostObject,
            &mut context,
        ));
        let target = Rc::new(Target {
            session: Rc::downgrade(&session),
            sheet: sheet.clone(),
            range_iid: range.header().iid,
            change: session
                .bind("Api", "Worksheet_Change", ProcedureKind::Sub)
                .unwrap(),
            close: session
                .bind("Api", "Workbook_BeforeClose", ProcedureKind::Sub)
                .unwrap(),
            changes: Cell::new(0),
            closes: Cell::new(0),
            errors: Cell::new(0),
        });
        session
            .bind_values(
                &target.change,
                &[api::ValueType::HostObject],
                api::ValueType::Void,
            )
            .unwrap();
        session
            .bind_values(
                &target.close,
                &[api::ValueType::Boolean],
                api::ValueType::Void,
            )
            .unwrap();
        TARGET.with(|s| s.replace(Some(target.clone())));
        let _scope = Scope;
        // SAFETY: exact trusted imported Worksheet/Workbook source event metadata.
        let change = unsafe {
            oxvba_com::typed_slots::advise_typed_event(
                &sheet,
                change_event.clone(),
                Rc::new(receive_change),
            )
        }
        .unwrap();
        // SAFETY: same exact source-interface and apartment contract.
        let close = unsafe {
            oxvba_com::typed_slots::advise_typed_event(
                &book,
                close_event.clone(),
                Rc::new(receive_close),
            )
        }
        .unwrap();
        for (suppress, expected, deliveries, transports) in [
            (true, "1:True:20", 1, (3, 8)),
            (false, "2:True:20", 3, (1, 10)),
        ] {
            let before = host.com().com_dispatch_transport_counts();
            let v = invoke(
                &session,
                "StartChange",
                &[
                    api::ValueSlot::Object(range.argument()),
                    api::ValueSlot::Boolean(suppress),
                ],
                api::ValueType::String,
                &mut context,
            );
            assert!(
                matches!(&v,api::Value::String(s)if s.as_str()==expected),
                "{backend:?}: {v:?}"
            );
            assert_eq!(target.changes.get(), deliveries);
            assert_eq!(
                host.com().com_dispatch_transport_counts(),
                (before.0 + transports.0, before.1 + transports.1)
            );
            println!("{backend:?} suppress={suppress}: {v:?}");
        }
        let before = host.com().com_dispatch_transport_counts();
        let v = invoke(
            &session,
            "TryClose",
            &[api::ValueSlot::Object(book.argument())],
            api::ValueType::String,
            &mut context,
        );
        assert!(
            matches!(&v,api::Value::String(s)if s.as_str()=="1:oracle-bootstrap.xlsx"),
            "{backend:?}: {v:?}"
        );
        assert_eq!(target.closes.get(), 1);
        assert_eq!(
            host.com().com_dispatch_transport_counts(),
            (before.0 + 2, before.1)
        );
        assert_eq!(target.errors.get(), 0);
        println!("{backend:?} close={v:?}");
        close.close().unwrap();
        change.close().unwrap();
        let v = invoke(
            &session,
            "StartChange",
            &[
                api::ValueSlot::Object(range.argument()),
                api::ValueSlot::Boolean(false),
            ],
            api::ValueType::String,
            &mut context,
        );
        assert!(matches!(&v,api::Value::String(s)if s.as_str()=="0:True:20"));
        assert_eq!(target.changes.get(), 3);
    }
}
