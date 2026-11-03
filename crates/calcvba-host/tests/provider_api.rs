//! The functions below use only the host API. Runtime choice is supplied at the
//! application composition edge, so the library never links a concrete backend.
use calcvba_host_api::{self as api, SessionExt};

fn client(provider: &dyn api::RuntimeProvider, backend: api::Backend) {
    let info = provider.info();
    assert_eq!(info.host_api_abi, api::HOST_API_ABI);
    assert_eq!(info.typed_entry_abi, api::TYPED_ENTRY_ABI);
    assert!(matches!(
        provider.load(&api::Project::module("Api", "invalid"), backend, 0),
        Err(api::Error::AbiMismatch { .. })
    ));
    let source = "Option Private Module\nPublic Function Add2(ByVal x As Double, ByVal y As Double) As Double\nAdd2 = x+y\nEnd Function\nPublic Function OptionalValue(Optional ByVal x As Long = 42) As Long\nOptionalValue=x\nEnd Function\nPrivate Sub Hidden()\nEnd Sub\nPublic Sub Touch()\nEnd Sub\nPublic Property Get Size() As Long\nSize=9\nEnd Property\nPublic Property Let Size(ByVal value As Long)\nEnd Property\n";
    let project = api::Project::module("Api", source);
    let mut session = provider.load(&project, backend, api::HOST_API_ABI).unwrap();
    assert_eq!(session.procedures().len(), 6);
    assert!(session.procedures().iter().all(|p| p.option_private_module));
    let optional = session
        .procedures()
        .iter()
        .find(|p| p.name == "OptionalValue")
        .unwrap();
    assert!(optional.parameters[0].optional && !optional.parameters[0].by_ref);
    assert_eq!(optional.parameters[0].ty, api::Type::Long);
    assert_eq!(
        optional.parameters[0].default.as_ref().unwrap().as_i32(),
        Some(42)
    );
    assert!(matches!(
        session.bind("Api", "Hidden", api::ProcedureKind::Sub),
        Err(api::Error::UnknownProcedure(_))
    ));
    let touch = session
        .bind("Api", "Touch", api::ProcedureKind::Sub)
        .unwrap();
    assert_eq!(
        session.call_bound(&touch, &[], 100).unwrap(),
        api::Variant::empty()
    );
    let getter = session
        .bind("Api", "Size", api::ProcedureKind::PropertyGet)
        .unwrap();
    let getter = session.bind_typed_handle::<(), i32>(&getter).unwrap();
    assert_eq!(session.call_typed(&getter, (), 100).unwrap(), 9);
    let dynamic = session
        .bind("Api", "Add2", api::ProcedureKind::Function)
        .unwrap();
    let typed = session
        .bind_typed::<(f64, f64), f64>("Api", "Add2")
        .unwrap();
    assert_eq!(typed.generation(), dynamic.generation());
    assert_eq!(session.call_typed(&typed, (19.0, 23.0), 100).unwrap(), 42.0);
    assert_eq!(
        session
            .call_bound(
                &dynamic,
                &[api::Variant::from_f64(19.0), api::Variant::from_f64(23.0)],
                100
            )
            .unwrap()
            .as_f64(),
        Some(42.0)
    );
    let foreign = provider.load(&project, backend, api::HOST_API_ABI).unwrap();
    assert!(matches!(
        foreign.call_typed(&typed, (19.0, 23.0), 100),
        Err(api::Error::StaleHandle)
    ));
    let forged = api::BoundProcedure::from_provider((), dynamic.generation());
    assert!(matches!(
        session.call_bound(&forged, &[], 100),
        Err(api::Error::StaleHandle)
    ));
    assert!(
        session
            .reload(&api::Project::module("Api", "not VBA"))
            .is_err()
    );
    assert_eq!(session.call_typed(&typed, (19.0, 23.0), 100).unwrap(), 42.0);
    session.reset().unwrap();
    assert_eq!(session.call_typed(&typed, (19.0, 23.0), 100).unwrap(), 42.0);
    session.reload(&project).unwrap();
    assert!(matches!(
        session.call_typed(&typed, (19.0, 23.0), 100),
        Err(api::Error::StaleHandle)
    ));
    assert!(matches!(
        session.call_bound(&dynamic, &[], 100),
        Err(api::Error::StaleHandle)
    ));
    let next = session
        .bind_typed::<(f64, f64), f64>("Api", "Add2")
        .unwrap();
    assert_eq!(next.generation(), typed.generation() + 1);
    drop(foreign);
    assert_eq!(session.call_typed(&next, (20.0, 22.0), 100).unwrap(), 42.0);
}
#[test]
fn runtime_composition_supplies_the_existing_host_contract_to_an_api_only_client() {
    let provider = calcvba_host::provider::Provider::default();
    client(&provider, api::Backend::Interpreter);
    if cfg!(feature = "jit") {
        client(&provider, api::Backend::Jit);
    }
}
