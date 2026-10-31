use calcvba_host_api::{self as api, RuntimeProvider, SessionExt};
use oxvba_embed::provider::Provider;
use std::{
    cell::{Cell, RefCell},
    rc::{Rc, Weak},
};
#[path = "support/typed_host.rs"]
mod support;

#[test]
fn range_like_host_slots_execute_with_typed_arguments_and_ownership() {
    for backend in support::backends() {
        let session = Provider::default()
            .load(&support::project(), backend, api::HOST_API_ABI)
            .unwrap();
        let calls = Rc::new(Cell::new(0));
        let owned = Rc::new(());
        let weak = Rc::downgrade(&owned);
        let counter = calls.clone();
        let object = api::HostObject::new(
            support::IID,
            vec![
                api::HostSlot::double_long(move |_, offset| {
                    let _keep = &owned;
                    counter.set(counter.get() + 1);
                    Ok(40.0 + f64::from(offset))
                }),
                api::HostSlot::double(|_| Ok(41.0)),
            ],
        );
        let retained = object.clone();
        let read = session
            .bind_typed::<(api::ObjectArg<'_>, i32), f64>("Api", "ReadRange")
            .unwrap();
        let value = session
            .bind_typed::<(api::ObjectArg<'_>,), f64>("Api", "ReadValue")
            .unwrap();
        assert_eq!(
            session
                .call_typed(&read, (object.argument(), 1), 100)
                .unwrap(),
            42.0
        );
        assert_eq!(
            session
                .call_typed(&value, (object.argument(),), 100)
                .unwrap(),
            82.0
        );
        assert_eq!(calls.get(), 1);
        drop(object);
        assert!(weak.upgrade().is_some());
        assert_eq!(
            session
                .call_typed(&read, (retained.argument(), 2), 100)
                .unwrap(),
            43.0
        );
        drop(retained);
        assert!(
            weak.upgrade().is_none(),
            "last object release must drop callback captures"
        );
    }
}

#[test]
fn direct_native_primary_entry_and_api_slots_match_for_mixed_object_argument_positions() {
    for backend in support::backends() {
        let provider = Provider::default();
        let module = provider.compile(&support::project()).unwrap();
        let engine_backend = match backend {
            api::Backend::Interpreter => oxvba_embed::Backend::Interpreter,
            api::Backend::Jit => oxvba_embed::Backend::Jit,
        };
        let engine = module.prepare(engine_backend).unwrap();
        let direct = engine
            .bind_typed::<(f64, api::ObjectArg<'_>, i32, f64), f64>("Api", "Mixed")
            .unwrap();
        let session = provider
            .load(&support::project(), backend, api::HOST_API_ABI)
            .unwrap();
        let slots = session
            .bind_typed::<(f64, api::ObjectArg<'_>, i32, f64), f64>("Api", "Mixed")
            .unwrap();
        let object = api::HostObject::new(
            support::IID,
            vec![api::HostSlot::double_long(|_, offset| {
                Ok(f64::from(offset) + 10.0)
            })],
        );
        let args = (19.0, object.argument(), 3, 10.0);
        assert_eq!(engine.call_typed(&direct, args, 100).unwrap(), 42.0);
        assert_eq!(session.call_typed(&slots, args, 100).unwrap(), 42.0);
    }
}

#[test]
fn host_slots_refuse_wrong_interface_bounds_and_native_signature() {
    for backend in support::backends() {
        let session = Provider::default()
            .load(&support::project(), backend, api::HOST_API_ABI)
            .unwrap();
        let read = session
            .bind_slots(
                "Api",
                "ReadRange",
                &[api::ScalarType::HostObject, api::ScalarType::Long],
                api::ScalarType::Double,
            )
            .unwrap();
        for (object, error) in [
            (
                api::HostObject::new(
                    [9; 16],
                    vec![api::HostSlot::double_long(|_, _| {
                        panic!("wrong IID must not enter")
                    })],
                ),
                438,
            ),
            (api::HostObject::new(support::IID, vec![]), 438),
            (
                api::HostObject::new(
                    support::IID,
                    vec![api::HostSlot::double(|_| {
                        panic!("wrong signature must not enter")
                    })],
                ),
                13,
            ),
            (
                api::HostObject::new(
                    support::IID,
                    vec![api::HostSlot::double_long(|_, _| Err(11))],
                ),
                11,
            ),
            (
                api::HostObject::new(
                    support::IID,
                    vec![api::HostSlot::double_long(|_, _| {
                        panic!("caught at host ABI")
                    })],
                ),
                5,
            ),
        ] {
            let result = session.call_slots(
                &read,
                &[object.borrow_slot(), api::Slot::Long(1)],
                &mut api::TypedContext {
                    remaining_blocks: 100,
                },
            );
            assert!(
                matches!(result, Err(api::Error::Runtime { number, .. }) if number == error),
                "{backend:?}: expected {error}, got {result:?}"
            );
        }
    }
}

#[test]
fn same_session_typed_callback_reentry_shares_budget_and_keeps_outer_activation() {
    for backend in support::backends() {
        let session: Rc<Box<dyn api::Session>> = Rc::new(
            Provider::default()
                .load(&support::project(), backend, api::HOST_API_ABI)
                .unwrap(),
        );
        let weak = Rc::downgrade(&session);
        let twice = session.bind_typed::<(f64,), f64>("Api", "Twice").unwrap();
        let dynamic = session
            .bind("Api", "Twice", api::ProcedureKind::Function)
            .unwrap();
        let calls = Rc::new(Cell::new(0));
        let counter = calls.clone();
        let object = api::HostObject::new(
            support::IID,
            vec![api::HostSlot::double_long(move |context, offset| {
                let session = weak.upgrade().unwrap();
                assert!(matches!(
                    session.call_typed(&twice, (20.0,), 100),
                    Err(api::Error::Reentry)
                ));
                assert!(matches!(
                    session.call_bound(&dynamic, &[api::Variant::from_f64(20.0)], 100),
                    Err(api::Error::Reentry)
                ));
                assert!(matches!(session.reset(), Err(api::Error::Reentry)));
                counter.set(counter.get() + 1);
                session
                    .call_typed_in_context(&twice, (f64::from(offset),), context)
                    .map_err(status)
            })],
        );
        let read = session
            .bind_typed::<(api::ObjectArg<'_>, i32), f64>("Api", "ReadRange")
            .unwrap();
        let mut context = api::TypedContext {
            remaining_blocks: 100,
        };
        assert_eq!(
            session
                .call_typed_in_context(&read, (object.argument(), 20), &mut context)
                .unwrap(),
            41.0
        );
        assert!(
            context.remaining_blocks < 98,
            "outer and nested blocks must both consume budget"
        );
        assert_eq!(calls.get(), 1);
        assert!(matches!(
            session.call_typed(&read, (object.argument(), 20), 1),
            Err(api::Error::BudgetExceeded)
        ));
        assert_eq!(
            session
                .call_typed(&read, (object.argument(), 21), 100)
                .unwrap(),
            43.0
        );
    }
}

#[test]
fn recursive_host_reentry_is_bounded_and_recovers_without_leaked_owners() {
    for backend in support::backends() {
        let session: Rc<Box<dyn api::Session>> = Rc::new(
            Provider::default()
                .load(&support::project(), backend, api::HOST_API_ABI)
                .unwrap(),
        );
        let weak_session = Rc::downgrade(&session);
        let read = session
            .bind_slots(
                "Api",
                "ReadRange",
                &[api::ScalarType::HostObject, api::ScalarType::Long],
                api::ScalarType::Double,
            )
            .unwrap();
        let recursive = read.clone();
        let weak_object: Rc<RefCell<Weak<api::HostObject>>> = Rc::default();
        let callback_object = weak_object.clone();
        let object = Rc::new(api::HostObject::new(
            support::IID,
            vec![api::HostSlot::double_long(move |context, offset| {
                if offset == 0 {
                    return Ok(0.0);
                }
                let session = weak_session.upgrade().unwrap();
                let object = callback_object.borrow().upgrade().unwrap();
                session
                    .call_slots(
                        &recursive,
                        &[object.borrow_slot(), api::Slot::Long(offset - 1)],
                        context,
                    )
                    .map(|r| r.as_double())
                    .map_err(status)
            })],
        ));
        *weak_object.borrow_mut() = Rc::downgrade(&object);
        let call = |offset, blocks| {
            session.call_slots(
                &read,
                &[object.borrow_slot(), api::Slot::Long(offset)],
                &mut api::TypedContext {
                    remaining_blocks: blocks,
                },
            )
        };
        assert_eq!(call(4, 1000).unwrap().as_double(), 5.0);
        assert!(matches!(
            call(1000, 10000),
            Err(api::Error::Runtime { number: 28, .. })
        ));
        assert!(matches!(call(1000, 10), Err(api::Error::BudgetExceeded)));
        assert_eq!(call(1, 1000).unwrap().as_double(), 2.0);
        drop(object);
        assert!(weak_object.borrow().upgrade().is_none());
    }
}
fn status(error: api::Error) -> i32 {
    match error {
        api::Error::Runtime { number, .. } => number,
        api::Error::BudgetExceeded => -1,
        _ => 5,
    }
}

#[test]
fn callback_nesting_limit_covers_calls_across_independent_sessions() {
    for backend in support::backends() {
        let provider = Provider::default();
        let sessions = Rc::new([
            provider
                .load(&support::project(), backend, api::HOST_API_ABI)
                .unwrap(),
            provider
                .load(&support::project(), backend, api::HOST_API_ABI)
                .unwrap(),
        ]);
        let handles = sessions
            .iter()
            .map(|s| {
                s.bind_slots(
                    "Api",
                    "ReadRange",
                    &[api::ScalarType::HostObject, api::ScalarType::Long],
                    api::ScalarType::Double,
                )
                .unwrap()
            })
            .collect::<Vec<_>>();
        let outer = handles[0].clone();
        let calls = Rc::new(Cell::new(0));
        let counter = calls.clone();
        let weak_object: Rc<RefCell<Weak<api::HostObject>>> = Rc::default();
        let callback_object = weak_object.clone();
        let nested = sessions.clone();
        let object = Rc::new(api::HostObject::new(
            support::IID,
            vec![api::HostSlot::double_long(move |context, offset| {
                counter.set(counter.get() + 1);
                if offset == 0 {
                    return Ok(0.0);
                }
                let object = callback_object.borrow().upgrade().unwrap();
                let next = ((offset - 1) % 2) as usize;
                nested[next]
                    .call_slots(
                        &handles[next],
                        &[object.borrow_slot(), api::Slot::Long(offset - 1)],
                        context,
                    )
                    .map(|v| v.as_double())
                    .map_err(status)
            })],
        ));
        *weak_object.borrow_mut() = Rc::downgrade(&object);
        let result = sessions[0].call_slots(
            &outer,
            &[object.borrow_slot(), api::Slot::Long(1000)],
            &mut api::TypedContext {
                remaining_blocks: 10000,
            },
        );
        assert!(matches!(
            result,
            Err(api::Error::Runtime { number: 28, .. })
        ));
        assert_eq!(
            calls.get(),
            64,
            "{backend:?} must bound the whole synchronous stack"
        );
        assert_eq!(
            sessions[0]
                .call_slots(
                    &outer,
                    &[object.borrow_slot(), api::Slot::Long(1)],
                    &mut api::TypedContext {
                        remaining_blocks: 100
                    }
                )
                .unwrap()
                .as_double(),
            2.0
        );
    }
}

#[test]
fn ambiguous_or_unsupported_host_library_descriptors_refuse_before_execution() {
    let provider = Provider::default();
    let mut duplicate_slot = support::project();
    duplicate_slot.libraries[0].interfaces[0].members[1].slot = 0;
    let mut duplicate_iid = support::project();
    let mut alias = duplicate_iid.libraries[0].interfaces[0].clone();
    alias.name = "Alias".into();
    duplicate_iid.libraries[0].interfaces.push(alias);
    let mut unsupported_return = support::project();
    unsupported_return.libraries[0].interfaces[0].members[0].result = api::ScalarType::Boolean;
    let mut duplicate_reference = support::project();
    duplicate_reference
        .references
        .push(api::TypeLibraryReference::named("HostObjects"));
    for project in [
        duplicate_slot,
        duplicate_iid,
        unsupported_return,
        duplicate_reference,
    ] {
        assert!(matches!(
            provider.load(&project, api::Backend::Interpreter, api::HOST_API_ABI),
            Err(api::Error::UnsupportedSignature(_))
        ));
    }
}

#[test]
fn retained_range_descriptor_resolves_each_active_host_invocation() {
    for backend in support::backends() {
        let session = Provider::default()
            .load(&support::project(), backend, api::HOST_API_ABI)
            .unwrap();
        // The descriptor outlives any one invocation. Its host service resolves
        // the current sheet/value, and refuses a call with no active view.
        let active: Rc<Cell<Option<(u32, f64)>>> = Rc::new(Cell::new(None));
        let view = active.clone();
        let sheet_id = 7;
        let object = api::HostObject::new(
            support::IID,
            vec![api::HostSlot::double_long(move |_, offset| {
                let (sheet, value) = view.get().ok_or(91)?;
                if sheet != sheet_id {
                    return Err(9);
                }
                Ok(value + f64::from(offset))
            })],
        );
        let retained = object.clone();
        drop(object);
        let read = session
            .bind_typed::<(api::ObjectArg<'_>, i32), f64>("Api", "ReadRange")
            .unwrap();
        assert!(matches!(
            session.call_typed(&read, (retained.argument(), 0), 100),
            Err(api::Error::Runtime { number: 91, .. })
        ));
        active.set(Some((sheet_id, 41.0)));
        assert_eq!(
            session
                .call_typed(&read, (retained.argument(), 0), 100)
                .unwrap(),
            42.0
        );
        active.set(None);
        assert!(matches!(
            session.call_typed(&read, (retained.argument(), 0), 100),
            Err(api::Error::Runtime { number: 91, .. })
        ));
        active.set(Some((sheet_id, 82.0)));
        assert_eq!(
            session
                .call_typed(&read, (retained.argument(), 0), 100)
                .unwrap(),
            83.0
        );
        active.set(Some((99, 1000.0)));
        assert!(matches!(
            session.call_typed(&read, (retained.argument(), 0), 100),
            Err(api::Error::Runtime { number: 9, .. })
        ));
    }
}
