use calcvba_host_api::{self as api, RuntimeProvider, SessionExt};
use oxvba_embed::provider::Provider;
use std::{cell::Cell, rc::Rc};
#[path = "support/typed_host.rs"]
mod support;

fn project() -> api::Project {
    let mut project = support::project();
    project.modules[0].source.push_str(r#"
Private kept As HostObjects.Range
Private entries As Long
Public Function Hold(ByVal root As Object) As Double
    Set kept = root
    Hold = 1#
End Function
Public Function Fetch() As HostObjects.Range
    Set Fetch = kept
End Function
Public Function Echo(ByVal root As HostObjects.Range) As HostObjects.Range
    Set Echo = root
End Function
Public Function ReadHeld() As Double
    ReadHeld = kept.NumericValue
End Function
Public Function ClearHeld() As Double
    Set kept = Nothing
    ClearHeld = 1#
End Function
Public Function FromStatic(ByVal root As HostObjects.Range, ByVal store As Long) As HostObjects.Range
    Static saved As HostObjects.Range
    If store <> 0 Then
        Set saved = root
    End If
    Set FromStatic = saved
End Function
Public Function Counted(ByVal root As HostObjects.Range, ByVal count As Long) As Double
    Dim total As Double
    entries = entries + 1
    total = root.Read(0)
    Do While count > 0
        total = total + 1#
        count = count - 1
    Loop
    Counted = total + entries
End Function
Public Function EntryCount() As Long
    EntryCount = entries
End Function
Public Function ReadTwice() As Double
    ReadTwice = kept.NumericValue + kept.NumericValue
End Function
Public Property Get HeldProperty() As HostObjects.Range
    Set HeldProperty = kept
End Property
"#);
    project
}
fn frame(id: u64, view: u64) -> api::InvocationFrame {
    api::InvocationFrame::new(id, None, Rc::new(view))
}
fn finish(invocation: &mut dyn api::Invocation, slice: u64) -> api::OwnedValue {
    for _ in 0..100_000 {
        match invocation.resume(slice).unwrap() {
            api::InvocationProgress::Suspended => {}
            api::InvocationProgress::Returned(value) => return value,
        }
    }
    panic!("bounded invocation did not return")
}
fn owned(
    session: &dyn api::Session,
    name: &str,
    args: &[api::Slot<'_>],
    result: api::ScalarType,
) -> api::OwnedValue {
    let types: Vec<_> = args.iter().map(|arg| arg.ty()).collect();
    let bound = session.bind_slots("Api", name, &types, result).unwrap();
    let mut invocation = session
        .begin_invocation(&bound, args, frame(1, 41), 10_000)
        .unwrap();
    finish(&mut *invocation, 1)
}
fn object(value: api::OwnedValue) -> Option<api::HostObject> {
    if let api::OwnedValue::Object(object) = value {
        object
    } else {
        panic!("object result expected: {value:?}")
    }
}

#[test]
fn module_static_and_results_retain_exact_owners_and_reset_releases_them() {
    for backend in support::backends() {
        let session = Provider::default()
            .load(&project(), backend, api::HOST_API_ABI)
            .unwrap();
        let owner = Rc::new(());
        let weak = Rc::downgrade(&owner);
        let table = api::HostObject::new(
            support::IID,
            vec![
                api::HostSlot::double_long(move |_, _| {
                    let _keep = &owner;
                    Ok(40.0)
                }),
                api::HostSlot::double(|_| Ok(41.0)),
            ],
        );
        let pointer = table.argument().as_ptr();
        let echoed = object(owned(
            &*session,
            "Echo",
            &[table.borrow_slot()],
            api::ScalarType::HostObject,
        ))
        .unwrap();
        assert_eq!(echoed.argument().as_ptr(), pointer);
        assert!(matches!(
            owned(
                &*session,
                "Hold",
                &[table.borrow_slot()],
                api::ScalarType::Double
            ),
            api::OwnedValue::Double(1.0)
        ));
        let stored = object(owned(
            &*session,
            "FromStatic",
            &[table.borrow_slot(), api::Slot::Long(1)],
            api::ScalarType::HostObject,
        ))
        .unwrap();
        drop(table);
        drop(echoed);
        drop(stored);
        assert!(weak.upgrade().is_some());
        let fetched = object(owned(&*session, "Fetch", &[], api::ScalarType::HostObject)).unwrap();
        assert_eq!(fetched.argument().as_ptr(), pointer);
        let raw_fetch = session
            .bind_slots("Api", "Fetch", &[], api::ScalarType::HostObject)
            .unwrap();
        assert!(
            matches!(
                session.call_slots(
                    &raw_fetch,
                    &[],
                    &mut api::TypedContext {
                        remaining_blocks: 100
                    }
                ),
                Err(api::Error::UnsupportedSignature(_))
            ),
            "raw object pointers must not escape a completed frame"
        );
        let property = session
            .bind("Api", "HeldProperty", api::ProcedureKind::PropertyGet)
            .unwrap();
        let property = session
            .type_bound_slots(&property, &[], api::ScalarType::HostObject)
            .unwrap();
        let mut property_call = session
            .begin_invocation(&property, &[], frame(7, 41), 100)
            .unwrap();
        let property_value = object(finish(&mut *property_call, 1)).unwrap();
        assert_eq!(property_value.argument().as_ptr(), pointer);
        drop(property_value);
        let clear = session
            .bind("Api", "ClearHeld", api::ProcedureKind::Function)
            .unwrap();
        assert_eq!(
            session.call_bound(&clear, &[], 100).unwrap().as_f64(),
            Some(1.0),
            "dynamic and typed calls share one store"
        );
        assert!(object(owned(&*session, "Fetch", &[], api::ScalarType::HostObject)).is_none());
        assert!(
            weak.upgrade().is_some(),
            "Static and result still retain the object"
        );
        session.reset().unwrap();
        assert!(
            weak.upgrade().is_some(),
            "returned object survives project reset"
        );
        drop(fetched);
        assert!(
            weak.upgrade().is_none(),
            "all session and result owners must release"
        );
    }
}

#[test]
fn saved_object_resolves_successive_views_and_stays_an_object_after_host_invalidation() {
    for backend in support::backends() {
        let session = Provider::default()
            .load(&project(), backend, api::HOST_API_ABI)
            .unwrap();
        let alive = Rc::new(Cell::new(true));
        let valid = alive.clone();
        let table = api::HostObject::new(
            support::IID,
            vec![
                api::HostSlot::double_long(|_, _| Ok(0.0)),
                api::HostSlot::double(move |_| {
                    if !valid.get() {
                        return Err(424);
                    }
                    let frame = api::active_invocation().ok_or(91)?;
                    let (sheet, value) = frame.read_view.downcast_ref::<(u32, f64)>().ok_or(9)?;
                    if *sheet != 7 {
                        return Err(9);
                    }
                    Ok(*value)
                }),
            ],
        );
        let pointer = table.argument().as_ptr();
        owned(
            &*session,
            "Hold",
            &[table.borrow_slot()],
            api::ScalarType::Double,
        );
        drop(table);
        let read = session
            .bind_slots("Api", "ReadHeld", &[], api::ScalarType::Double)
            .unwrap();
        assert!(matches!(
            session.call_slots(
                &read,
                &[],
                &mut api::TypedContext {
                    remaining_blocks: 100
                }
            ),
            Err(api::Error::Runtime { number: 91, .. })
        ));
        for value in [41.0, 82.0] {
            let metadata = api::InvocationFrame::new(10, None, Rc::new((7u32, value)));
            let mut call = session.begin_invocation(&read, &[], metadata, 100).unwrap();
            assert!(matches!(finish(&mut *call,1),api::OwnedValue::Double(v) if v==value));
        }
        let mut wrong = session
            .begin_invocation(
                &read,
                &[],
                api::InvocationFrame::new(11, None, Rc::new((8u32, 999.0))),
                100,
            )
            .unwrap();
        assert!(matches!(
            wrong.resume(1000),
            Err(api::Error::Runtime { number: 9, .. })
        ));
        alive.set(false);
        let mut stale = session
            .begin_invocation(
                &read,
                &[],
                api::InvocationFrame::new(12, None, Rc::new((7u32, 41.0))),
                100,
            )
            .unwrap();
        assert!(matches!(
            stale.resume(1000),
            Err(api::Error::Runtime { number: 424, .. })
        ));
        let still = object(owned(&*session, "Fetch", &[], api::ScalarType::HostObject)).unwrap();
        assert_eq!(
            still.argument().as_ptr(),
            pointer,
            "invalid host view must not erase object identity"
        );
    }
}

#[test]
fn global_receiver_survives_reentrant_replacement_without_borrowing_the_global_cell() {
    for backend in support::backends() {
        let session: Rc<Box<dyn api::Session>> = Rc::new(
            Provider::default()
                .load(&project(), backend, api::HOST_API_ABI)
                .unwrap(),
        );
        let weak_session = Rc::downgrade(&session);
        let owner = Rc::new(());
        let weak = Rc::downgrade(&owner);
        let inside = weak.clone();
        let next = api::HostObject::new(
            support::IID,
            vec![
                api::HostSlot::double_long(|_, _| Ok(0.0)),
                api::HostSlot::double(|_| Ok(82.0)),
            ],
        );
        let hold = session
            .bind_typed::<(api::ObjectArg<'_>,), f64>("Api", "Hold")
            .unwrap();
        let initial = api::HostObject::new(
            support::IID,
            vec![
                api::HostSlot::double_long(|_, _| Ok(0.0)),
                api::HostSlot::double(move |context| {
                    let _keep = &owner;
                    weak_session
                        .upgrade()
                        .unwrap()
                        .call_typed_in_context(&hold, (next.argument(),), context)
                        .unwrap();
                    assert!(
                        inside.upgrade().is_some(),
                        "receiver must survive replacement of its last global reference"
                    );
                    Ok(40.0)
                }),
            ],
        );
        owned(
            &**session,
            "Hold",
            &[initial.borrow_slot()],
            api::ScalarType::Double,
        );
        drop(initial);
        assert!(matches!(
            owned(&**session, "ReadTwice", &[], api::ScalarType::Double),
            api::OwnedValue::Double(122.0)
        ));
        assert!(
            weak.upgrade().is_none(),
            "transient receiver owner must release after its callback"
        );
    }
}

#[test]
fn cancellation_budget_fault_and_panic_release_frames_and_leave_the_session_usable() {
    for backend in support::backends() {
        for failure in [0, 11, 5, -1] {
            let session = Provider::default()
                .load(&project(), backend, api::HOST_API_ABI)
                .unwrap();
            let owner = Rc::new(());
            let weak = Rc::downgrade(&owner);
            let table = api::HostObject::new(
                support::IID,
                vec![api::HostSlot::double_long(move |_, _| {
                    let _keep = &owner;
                    if failure == 5 {
                        panic!("contained host failure");
                    }
                    if failure > 0 {
                        return Err(failure);
                    }
                    Ok(40.0)
                })],
            );
            let bound = session
                .bind_slots(
                    "Api",
                    "Counted",
                    &[api::ScalarType::HostObject, api::ScalarType::Long],
                    api::ScalarType::Double,
                )
                .unwrap();
            let mut invocation = session
                .begin_invocation(
                    &bound,
                    &[table.borrow_slot(), api::Slot::Long(20)],
                    frame(4, 40),
                    if failure == -1 { 0 } else { 1000 },
                )
                .unwrap();
            drop(table);
            if failure == 0 {
                assert!(matches!(
                    invocation.resume(1).unwrap(),
                    api::InvocationProgress::Suspended
                ));
                let count = session
                    .bind("Api", "EntryCount", api::ProcedureKind::Function)
                    .unwrap();
                assert!(matches!(
                    session.call_bound(&count, &[], 100),
                    Err(api::Error::Reentry)
                ));
                drop(invocation);
            } else {
                let result = invocation.resume(100_000);
                if failure == -1 {
                    assert!(matches!(result, Err(api::Error::BudgetExceeded)));
                } else {
                    assert!(
                        matches!(result,Err(api::Error::Runtime {number,..}) if number==failure)
                    );
                }
                assert!(matches!(
                    invocation.resume(1),
                    Err(api::Error::InvocationFinished)
                ));
            }
            assert!(
                weak.upgrade().is_none(),
                "fault/cancel must release argument and local owners"
            );
            session.reset().unwrap();
            let twice = session.bind_typed::<(f64,), f64>("Api", "Twice").unwrap();
            assert_eq!(session.call_typed(&twice, (21.0,), 100).unwrap(), 42.0);
            assert!(api::active_invocation().is_none());
        }
    }
}

#[test]
fn wrong_interface_assignment_preserves_the_previous_owner() {
    for backend in support::backends() {
        let session = Provider::default()
            .load(&project(), backend, api::HOST_API_ABI)
            .unwrap();
        let original = api::HostObject::new(support::IID, vec![]);
        owned(
            &*session,
            "Hold",
            &[original.borrow_slot()],
            api::ScalarType::Double,
        );
        let wrong = api::HostObject::new([9; 16], vec![]);
        let hold = session
            .bind_slots(
                "Api",
                "Hold",
                &[api::ScalarType::HostObject],
                api::ScalarType::Double,
            )
            .unwrap();
        let mut invocation = session
            .begin_invocation(&hold, &[wrong.borrow_slot()], frame(5, 0), 100)
            .unwrap();
        assert!(matches!(
            invocation.resume(1000),
            Err(api::Error::Runtime { number: 13, .. })
        ));
        let fetched = object(owned(&*session, "Fetch", &[], api::ScalarType::HostObject)).unwrap();
        assert_eq!(fetched.argument().as_ptr(), original.argument().as_ptr());
    }
}

#[test]
fn one_operation_slices_resume_effects_once_and_restore_the_pinned_view() {
    for backend in support::backends() {
        let session = Provider::default()
            .load(&project(), backend, api::HOST_API_ABI)
            .unwrap();
        let calls = Rc::new(Cell::new(0));
        let counter = calls.clone();
        let table = api::HostObject::new(
            support::IID,
            vec![api::HostSlot::double_long(move |_, _| {
                counter.set(counter.get() + 1);
                let frame = api::active_invocation().expect("invocation view");
                assert_eq!(frame.id, 17);
                assert_eq!(frame.parent, Some(9));
                Ok(*frame.read_view.downcast_ref::<u64>().unwrap() as f64)
            })],
        );
        let bound = session
            .bind_slots(
                "Api",
                "Counted",
                &[api::ScalarType::HostObject, api::ScalarType::Long],
                api::ScalarType::Double,
            )
            .unwrap();
        let view = Rc::new(40u64);
        let weak = Rc::downgrade(&view);
        let mut invocation = session
            .begin_invocation(
                &bound,
                &[table.borrow_slot(), api::Slot::Long(20)],
                api::InvocationFrame::new(17, Some(9), view),
                10_000,
            )
            .unwrap();
        drop(table);
        assert!(matches!(
            invocation.resume(0).unwrap(),
            api::InvocationProgress::Suspended
        ));
        assert_eq!(calls.get(), 0);
        assert!(matches!(session.reset(), Err(api::Error::Reentry)));
        let result = finish(&mut *invocation, 1);
        assert!(
            matches!(result, api::OwnedValue::Double(61.0)),
            "{backend:?}: {result:?}"
        );
        assert_eq!(calls.get(), 1, "completed host effect must not replay");
        assert!(api::active_invocation().is_none());
        assert!(matches!(
            invocation.resume(1),
            Err(api::Error::InvocationFinished)
        ));
        let count = session.bind_typed::<(), i32>("Api", "EntryCount").unwrap();
        assert_eq!(session.call_typed(&count, (), 100).unwrap(), 1);
        drop(invocation);
        assert!(weak.upgrade().is_none());
    }
}

#[test]
fn reload_publishes_new_signature_while_the_old_activation_keeps_code_and_state() {
    for backend in support::backends() {
        let provider = Provider::default();
        let mut session = provider
            .load(&project(), backend, api::HOST_API_ABI)
            .unwrap();
        let calls = Rc::new(Cell::new(0));
        let counter = calls.clone();
        let table = api::HostObject::new(
            support::IID,
            vec![api::HostSlot::double_long(move |_, _| {
                counter.set(counter.get() + 1);
                Ok(40.0)
            })],
        );
        let old = session
            .bind_slots(
                "Api",
                "Counted",
                &[api::ScalarType::HostObject, api::ScalarType::Long],
                api::ScalarType::Double,
            )
            .unwrap();
        let mut invocation = session
            .begin_invocation(
                &old,
                &[table.borrow_slot(), api::Slot::Long(3)],
                frame(2, 40),
                10_000,
            )
            .unwrap();
        while calls.get() == 0 {
            assert!(matches!(
                invocation.resume(1).unwrap(),
                api::InvocationProgress::Suspended
            ));
        }
        let replacement = api::Project::module(
            "Api",
            "Public Function Counted(ByVal text As String) As String\nCounted = text\nEnd Function\n",
        );
        session.reload(&replacement).unwrap();
        assert!(matches!(
            session.begin_invocation(&old, &[], frame(3, 0), 100),
            Err(api::Error::StaleHandle)
        ));
        let dynamic = session
            .bind("Api", "Counted", api::ProcedureKind::Function)
            .unwrap();
        let input = api::Variant::from_string("new generation");
        assert_eq!(
            session
                .call_bound(&dynamic, std::slice::from_ref(&input), 100)
                .unwrap(),
            input
        );
        session.reset().unwrap();
        drop(session);
        assert!(matches!(
            finish(&mut *invocation, 1),
            api::OwnedValue::Double(44.0)
        ));
        assert_eq!(calls.get(), 1);
    }
}

#[test]
fn typed_ownership_does_not_require_a_member_call_in_the_module() {
    for backend in support::backends() {
        for empty in [false, true] {
            let mut project = support::project();
            if empty {
                project.libraries[0].interfaces[0].members.clear();
            }
            project.modules[0].source="Public Function Echo(ByVal root As HostObjects.Range) As HostObjects.Range\nSet Echo = root\nEnd Function\n".into();
            let session = Provider::default()
                .load(&project, backend, api::HOST_API_ABI)
                .unwrap();
            assert!(
                matches!(&session.procedures()[0].parameters[0].ty,api::Type::Object {iid:Some(iid),..} if *iid==support::IID)
            );
            let table = api::HostObject::new(support::IID, vec![]);
            let result = object(owned(
                &*session,
                "Echo",
                &[table.borrow_slot()],
                api::ScalarType::HostObject,
            ))
            .unwrap();
            assert_eq!(result.argument().as_ptr(), table.argument().as_ptr());
        }
    }
}

#[test]
fn qualified_memberless_interfaces_keep_distinct_nominal_identities() {
    for backend in support::backends() {
        let mut project = api::Project::module(
            "Api",
            r#"
Public Function First(ByVal root As One.Range) As One.Range
    Set First = root
End Function
Public Function Second(ByVal root As Two.Range) As Two.Range
    Set Second = root
End Function
"#,
        );
        for (name, iid) in [("One", [1; 16]), ("Two", [2; 16])] {
            project.libraries.push(api::ObjectLibrary {
                name: name.into(),
                interfaces: vec![api::ObjectInterface {
                    name: "Range".into(),
                    iid,
                    members: vec![],
                }],
            });
        }
        let session = Provider::default()
            .load(&project, backend, api::HOST_API_ABI)
            .unwrap();
        let first = api::HostObject::new([1; 16], vec![]);
        let second = api::HostObject::new([2; 16], vec![]);
        let returned = object(owned(
            &*session,
            "First",
            &[first.borrow_slot()],
            api::ScalarType::HostObject,
        ))
        .unwrap();
        assert_eq!(returned.argument().as_ptr(), first.argument().as_ptr());
        let bound = session
            .bind_slots(
                "Api",
                "First",
                &[api::ScalarType::HostObject],
                api::ScalarType::HostObject,
            )
            .unwrap();
        assert!(matches!(
            session.begin_invocation(&bound, &[second.borrow_slot()], frame(2, 0), 100),
            Err(api::Error::Runtime { number: 13, .. })
        ));
        let returned = object(owned(
            &*session,
            "Second",
            &[second.borrow_slot()],
            api::ScalarType::HostObject,
        ))
        .unwrap();
        assert_eq!(returned.argument().as_ptr(), second.argument().as_ptr());
    }
}

#[test]
fn unused_typed_object_parameters_check_identity_before_entry_and_accept_nothing() {
    for backend in support::backends() {
        let mut project = support::project();
        project.modules[0].source = r#"
Public Function Unused(ByVal root As HostObjects.Range) As Double
    Unused = 7#
End Function
Public Function Untyped(ByVal root As Object) As Double
    Untyped = 7#
End Function
Public Function Use(ByVal root As HostObjects.Range) As Double
    Use = root.NumericValue
End Function
"#
        .into();
        let session = Provider::default()
            .load(&project, backend, api::HOST_API_ABI)
            .unwrap();
        let typed = session
            .bind_typed::<(api::ObjectArg<'_>,), f64>("Api", "Unused")
            .unwrap();
        let bound = session
            .bind_slots(
                "Api",
                "Unused",
                &[api::ScalarType::HostObject],
                api::ScalarType::Double,
            )
            .unwrap();
        let correct = api::HostObject::new(support::IID, vec![]);
        let wrong = api::HostObject::new([9; 16], vec![]);
        assert_eq!(
            session
                .call_typed(&typed, (correct.argument(),), 100)
                .unwrap(),
            7.0
        );
        assert!(matches!(
            session.call_typed(&typed, (wrong.argument(),), 100),
            Err(api::Error::Runtime { number: 13, .. })
        ));
        assert!(matches!(
            session.call_slots(
                &bound,
                &[wrong.borrow_slot()],
                &mut api::TypedContext {
                    remaining_blocks: 100
                }
            ),
            Err(api::Error::Runtime { number: 13, .. })
        ));
        assert!(matches!(
            session.begin_invocation(&bound, &[wrong.borrow_slot()], frame(1, 0), 100),
            Err(api::Error::Runtime { number: 13, .. })
        ));
        assert_eq!(
            session
                .call_slots(
                    &bound,
                    &[api::Slot::Nothing],
                    &mut api::TypedContext {
                        remaining_blocks: 100
                    }
                )
                .unwrap()
                .as_double(),
            7.0
        );
        assert!(matches!(
            owned(
                &*session,
                "Unused",
                &[api::Slot::Nothing],
                api::ScalarType::Double
            ),
            api::OwnedValue::Double(7.0)
        ));
        assert!(matches!(
            owned(
                &*session,
                "Untyped",
                &[wrong.borrow_slot()],
                api::ScalarType::Double
            ),
            api::OwnedValue::Double(7.0)
        ));
        let use_object = session
            .bind_slots(
                "Api",
                "Use",
                &[api::ScalarType::HostObject],
                api::ScalarType::Double,
            )
            .unwrap();
        assert!(matches!(
            session.call_slots(
                &use_object,
                &[api::Slot::Nothing],
                &mut api::TypedContext {
                    remaining_blocks: 100
                }
            ),
            Err(api::Error::Runtime { number: 91, .. })
        ));
    }
}

#[test]
fn a_new_semantic_invocation_repeats_effects_but_an_exhausted_one_cannot_resume() {
    for backend in support::backends() {
        let session = Provider::default()
            .load(&project(), backend, api::HOST_API_ABI)
            .unwrap();
        let calls = Rc::new(Cell::new(0));
        let counter = calls.clone();
        let root = api::HostObject::new(
            support::IID,
            vec![api::HostSlot::double_long(move |_, _| {
                counter.set(counter.get() + 1);
                Ok(40.0)
            })],
        );
        let bound = session
            .bind_slots(
                "Api",
                "Counted",
                &[api::ScalarType::HostObject, api::ScalarType::Long],
                api::ScalarType::Double,
            )
            .unwrap();
        let mut limited = session
            .begin_invocation(
                &bound,
                &[root.borrow_slot(), api::Slot::Long(20)],
                frame(1, 0),
                6,
            )
            .unwrap();
        assert!(matches!(
            limited.resume(100_000),
            Err(api::Error::BudgetExceeded)
        ));
        assert_eq!(calls.get(), 1);
        assert_eq!(limited.remaining_blocks(), 0);
        assert!(matches!(
            limited.resume(1),
            Err(api::Error::InvocationFinished)
        ));
        let mut next = session
            .begin_invocation(
                &bound,
                &[root.borrow_slot(), api::Slot::Long(20)],
                frame(2, 0),
                1000,
            )
            .unwrap();
        assert!(matches!(
            finish(&mut *next, 3),
            api::OwnedValue::Double(62.0)
        ));
        assert_eq!(calls.get(), 2);
    }
}
