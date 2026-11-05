use calcvba_host::provider::Provider;
use calcvba_host_api::{self as api, RuntimeProvider};
use std::{cell::RefCell, collections::HashMap, rc::Rc};
fn backends() -> Vec<api::Backend> {
    let mut values = vec![api::Backend::Interpreter];
    if cfg!(feature = "jit") {
        values.push(api::Backend::Jit);
    }
    values
}
const IID: api::InterfaceId = [71; 16];
type Grid = Rc<RefCell<HashMap<(i32, i32), api::Variant>>>;
type Journal = Rc<RefCell<Vec<String>>>;

fn library() -> api::TypedObjectLibrary {
    use api::{TypedMemberKind as K, ValueType as T};
    let member = |name: &str, slot, kind, parameters, result| api::TypedObjectMember {
        name: name.into(),
        slot,
        kind,
        parameters,
        result,
        is_default: false,
    };
    api::TypedObjectLibrary {
        name: "HostValues".into(),
        interfaces: vec![api::TypedObjectInterface {
            name: "Range".into(),
            iid: IID,
            members: vec![
                member("Value", 0, K::PropertyGet, vec![], T::Variant),
                member("Value2", 1, K::PropertyGet, vec![], T::Variant),
                member(
                    "Cells",
                    2,
                    K::PropertyGet,
                    vec![T::Long, T::Long],
                    T::Interface(IID),
                ),
                member("Address", 3, K::PropertyGet, vec![], T::String),
                member("Rows", 4, K::PropertyGet, vec![], T::Interface(IID)),
                member("Columns", 5, K::PropertyGet, vec![], T::Interface(IID)),
                member("Count", 6, K::PropertyGet, vec![], T::Long),
                member("Value2", 7, K::PropertyLet, vec![T::Variant], T::Void),
            ],
        }],
    }
}
fn project() -> api::Project {
    api::Project::module(
        "Api",
        r#"
Public Function ReadMembers(ByVal root As HostValues.Range) As Double
    Dim first As Double
    first = root.Value2
    ReadMembers = first + root.Cells(2, 3).Value + root.Rows.Count + root.Columns.Count
End Function
Public Function ReadAddress(ByVal root As HostValues.Range) As String
    ReadAddress = root.Cells(2, 3).Address
End Function
Private Sub ReplaceText(ByRef text As String, ByVal root As HostValues.Range)
    text = root.Address
End Sub
Public Function AddressViaCall(ByVal root As HostValues.Range) As String
    Dim text As String
    ReplaceText text, root
    AddressViaCall = text
End Function
Public Function EchoText(ByVal text As String) As String
    EchoText = text
End Function
Public Function EchoValue(ByVal value As Variant) As Variant
    EchoValue = value
End Function
Public Sub SetCell(ByVal root As HostValues.Range, ByVal value As Variant)
    root.Value2 = value
End Sub
Public Function ReadNumeric(ByVal root As HostValues.Range) As Double
    ReadNumeric = root.Value2
End Function
Public Function AddValue(ByVal root As HostValues.Range) As Variant
    AddValue = root.Value2 + 2#
End Function
Private Sub ChangeAliases(ByRef first As Variant, ByRef second As Variant, ByVal root As HostValues.Range)
    first = root.Value2
    second = second + 2#
    Err.Raise 13
End Sub
Public Function VariantAliases(ByVal root As HostValues.Range) As Variant
    Dim value As Variant
    value = 10#
    On Error GoTo AliasFault
    ChangeAliases value, value, root
AliasFault:
    VariantAliases = value
End Function
Public Function ExplainNull(ByVal root As HostValues.Range) As String
    Dim value As Double
    On Error GoTo Explain
    value = root.Value2
    Exit Function
Explain:
    ExplainNull = Err.Description
End Function
"#,
    )
}
fn frame() -> api::InvocationFrame {
    api::InvocationFrame::new(42, Some(9), Rc::new(41u64))
}
fn observe(log: &Journal, text: String) {
    let view = api::active_invocation().expect("retained invocation view");
    assert_eq!((view.id, view.parent), (42, Some(9)));
    assert_eq!(view.read_view.downcast_ref::<u64>(), Some(&41));
    log.borrow_mut().push(text);
}
fn range(
    grid: Grid,
    log: Journal,
    row: i32,
    col: i32,
    rows: i32,
    cols: i32,
    count: i32,
) -> api::HostObject {
    use api::ValueType as T;
    let mut slots = Vec::new();
    let (g, l) = (grid.clone(), log.clone());
    slots.push(
        api::HostSlot::typed(vec![], T::Variant, move |_, _| {
            observe(&l, "Value".into());
            Ok(api::Value::Variant(
                g.borrow()
                    .get(&(row, col))
                    .cloned()
                    .unwrap_or_else(api::Variant::empty),
            ))
        })
        .unwrap(),
    );
    let (g, l) = (grid.clone(), log.clone());
    slots.push(
        api::HostSlot::typed(vec![], T::Variant, move |_, _| {
            observe(&l, "Value2".into());
            Ok(api::Value::Variant(
                g.borrow()
                    .get(&(row, col))
                    .cloned()
                    .unwrap_or_else(api::Variant::empty),
            ))
        })
        .unwrap(),
    );
    let (g, l) = (grid.clone(), log.clone());
    slots.push(
        api::HostSlot::typed(vec![T::Long, T::Long], T::Interface(IID), move |_, args| {
            let [api::ValueSlot::Long(r), api::ValueSlot::Long(c)] = args else {
                return Err(13);
            };
            observe(&l, format!("Cells({r},{c})"));
            Ok(api::Value::Object(Some(range(
                g.clone(),
                l.clone(),
                row + r - 1,
                col + c - 1,
                1,
                1,
                1,
            ))))
        })
        .unwrap(),
    );
    let l = log.clone();
    slots.push(
        api::HostSlot::typed(vec![], T::String, move |_, _| {
            observe(&l, "Address".into());
            let start = format!("${}${row}", char::from(b'A' + (col - 1) as u8));
            let text = if rows == 1 && cols == 1 {
                start
            } else {
                format!(
                    "{start}:${}${}",
                    char::from(b'A' + (col + cols - 2) as u8),
                    row + rows - 1
                )
            };
            Ok(api::Value::String(text.into()))
        })
        .unwrap(),
    );
    let (g, l) = (grid.clone(), log.clone());
    slots.push(
        api::HostSlot::typed(vec![], T::Interface(IID), move |_, _| {
            observe(&l, "Rows".into());
            Ok(api::Value::Object(Some(range(
                g.clone(),
                l.clone(),
                row,
                col,
                rows,
                cols,
                rows,
            ))))
        })
        .unwrap(),
    );
    let (g, l) = (grid.clone(), log.clone());
    slots.push(
        api::HostSlot::typed(vec![], T::Interface(IID), move |_, _| {
            observe(&l, "Columns".into());
            Ok(api::Value::Object(Some(range(
                g.clone(),
                l.clone(),
                row,
                col,
                rows,
                cols,
                cols,
            ))))
        })
        .unwrap(),
    );
    let l = log.clone();
    slots.push(
        api::HostSlot::typed(vec![], T::Long, move |_, _| {
            observe(&l, "Count".into());
            Ok(api::Value::Long(count))
        })
        .unwrap(),
    );
    slots.push(
        api::HostSlot::typed(vec![T::Variant], T::Void, move |_, args| {
            let [api::ValueSlot::Variant(value)] = args else {
                return Err(13);
            };
            observe(&log, "Value2=".into());
            grid.borrow_mut().insert((row, col), (*value).clone());
            Ok(api::Value::Empty)
        })
        .unwrap(),
    );
    api::HostObject::new(IID, slots)
}
fn finish(call: &mut dyn api::ValueInvocation, slice: u64) -> api::Value {
    for _ in 0..10000 {
        match call.resume(slice).unwrap() {
            api::ValueProgress::Suspended => {}
            api::ValueProgress::Returned(value) => return value,
        }
    }
    panic!("typed value invocation did not finish")
}
fn load(provider: &dyn RuntimeProvider, backend: api::Backend) -> Box<dyn api::Session> {
    assert_eq!(provider.value_abi(), Some(api::TYPED_VALUE_ABI));
    provider
        .load_with_object_libraries(&project(), &[library()], backend, api::HOST_API_ABI)
        .unwrap()
}

#[test]
fn common_range_members_use_typed_slots_on_each_backend_without_repeating_callbacks() {
    for backend in backends() {
        for slice in [1, 3, u64::MAX] {
            let session = load(&Provider::default(), backend);
            assert!(
                session
                    .bind_values(
                        "Api",
                        "ReadMembers",
                        &[api::ValueType::Interface(IID)],
                        api::ValueType::Double
                    )
                    .is_ok()
            );
            assert!(matches!(
                session.bind_values(
                    "Api",
                    "ReadMembers",
                    &[api::ValueType::Interface([99; 16])],
                    api::ValueType::Double
                ),
                Err(api::Error::ArgumentMismatch)
            ));
            let grid = Rc::new(RefCell::new(HashMap::from([
                ((1, 1), api::Variant::from_f64(40.0)),
                ((2, 3), api::Variant::from_f64(42.0)),
            ])));
            let log = Rc::new(RefCell::new(Vec::new()));
            let root = range(grid, log.clone(), 1, 1, 3, 4, 12);
            for (name, result) in [
                ("ReadMembers", api::ValueType::Double),
                ("ReadAddress", api::ValueType::String),
                ("AddressViaCall", api::ValueType::String),
            ] {
                log.borrow_mut().clear();
                let bound = session
                    .bind_values("Api", name, &[api::ValueType::HostObject], result)
                    .unwrap();
                let mut call = session
                    .begin_values(
                        &bound,
                        &[api::ValueSlot::Object(root.argument())],
                        frame(),
                        10000,
                    )
                    .unwrap();
                assert!(matches!(
                    call.resume(0).unwrap(),
                    api::ValueProgress::Suspended
                ));
                assert!(log.borrow().is_empty());
                match finish(&mut *call, slice) {
                    api::Value::Double(value) => assert_eq!(value, 89.0),
                    api::Value::String(value) => assert_eq!(
                        value.as_str(),
                        if name == "ReadAddress" {
                            "$C$2"
                        } else {
                            "$A$1:$D$3"
                        }
                    ),
                    value => panic!("unexpected result {value:?}"),
                }
                let expected: Vec<String> = match name {
                    "ReadMembers" => vec![
                        "Value2",
                        "Cells(2,3)",
                        "Value",
                        "Rows",
                        "Count",
                        "Columns",
                        "Count",
                    ],
                    "ReadAddress" => vec!["Cells(2,3)", "Address"],
                    _ => vec!["Address"],
                }
                .into_iter()
                .map(str::to_string)
                .collect();
                assert_eq!(*log.borrow(), expected);
            }
        }
    }
}

#[test]
fn extended_owning_arguments_results_and_subs_preserve_exact_carriers() {
    for backend in backends() {
        let session = load(&Provider::default(), backend);
        let text = api::BStr::from_utf16_units(&[0xd800, 0, 0xdc00]).unwrap();
        let echo = session
            .bind_values(
                "Api",
                "EchoText",
                &[api::ValueType::String],
                api::ValueType::String,
            )
            .unwrap();
        let mut call = session
            .begin_values(&echo, &[api::ValueSlot::String(&text)], frame(), 1000)
            .unwrap();
        drop(text);
        let api::Value::String(value) = finish(&mut *call, 1) else {
            panic!("BSTR expected");
        };
        assert_eq!(value.payload_units(), &[0xd800, 0, 0xdc00]);
        let echo = session
            .bind_values(
                "Api",
                "EchoValue",
                &[api::ValueType::Variant],
                api::ValueType::Variant,
            )
            .unwrap();
        for input in [
            api::Variant::null(),
            api::Variant::from_error_code(2042),
            api::Variant::from_utf16_units(&[0xd800, 0, 0xdc00]),
        ] {
            let mut call = session
                .begin_values(&echo, &[api::ValueSlot::Variant(&input)], frame(), 1000)
                .unwrap();
            let api::Value::Variant(value) = finish(&mut *call, 1) else {
                panic!("Variant expected");
            };
            assert_eq!(value, input);
        }
        let grid = Rc::new(RefCell::new(HashMap::new()));
        let root = range(
            grid.clone(),
            Rc::new(RefCell::new(Vec::new())),
            1,
            1,
            1,
            1,
            1,
        );
        let setter = session
            .bind_values(
                "Api",
                "SetCell",
                &[api::ValueType::HostObject, api::ValueType::Variant],
                api::ValueType::Void,
            )
            .unwrap();
        let input = api::Variant::from_f64(43.0);
        let mut call = session
            .begin_values(
                &setter,
                &[
                    api::ValueSlot::Object(root.argument()),
                    api::ValueSlot::Variant(&input),
                ],
                frame(),
                1000,
            )
            .unwrap();
        assert!(matches!(finish(&mut *call, 1), api::Value::Empty));
        assert_eq!(grid.borrow().get(&(1, 1)), Some(&input));
    }
}

#[test]
fn literal_variant_member_arguments_and_cvar_keep_their_value_types() {
    for backend in backends() {
        let project = api::Project::module(
            "Api",
            r#"
Public Sub LongLiteral(ByVal root As HostValues.Range)
    root.Value2 = 9&
End Sub
Public Sub DoubleLiteral(ByVal root As HostValues.Range)
    root.Value2 = 3.5#
End Sub
Public Sub StringLiteral(ByVal root As HostValues.Range)
    root.Value2 = "text"
End Sub
Public Sub BooleanLiteral(ByVal root As HostValues.Range)
    root.Value2 = True
End Sub
Public Sub ObjectArgument(ByVal root As HostValues.Range)
    root.Value2 = root
End Sub
Public Function Boxed(ByVal value As Variant) As Variant
    Boxed = CVar(value)
End Function
Public Function LiteralBoxed() As Variant
    LiteralBoxed = CVar("text")
End Function
Public Function LocalBoxed() As Variant
    Dim value As String
    value = "local"
    LocalBoxed = CVar(value)
End Function
"#,
        );
        let session = Provider::default()
            .load_with_object_libraries(&project, &[library()], backend, api::HOST_API_ABI)
            .unwrap();
        let grid = Rc::new(RefCell::new(HashMap::new()));
        let log = Rc::new(RefCell::new(Vec::new()));
        let root = range(grid.clone(), log.clone(), 1, 1, 1, 1, 1);
        for (name, expected) in [
            ("LongLiteral", api::Variant::from_i32(9)),
            ("DoubleLiteral", api::Variant::from_f64(3.5)),
            ("StringLiteral", api::Variant::from_string("text")),
            ("BooleanLiteral", api::Variant::from_bool(true)),
        ] {
            let bound = session
                .bind_values(
                    "Api",
                    name,
                    &[api::ValueType::HostObject],
                    api::ValueType::Void,
                )
                .unwrap();
            let mut call = session
                .begin_values(
                    &bound,
                    &[api::ValueSlot::Object(root.argument())],
                    frame(),
                    1000,
                )
                .unwrap();
            assert!(matches!(finish(&mut *call, 1), api::Value::Empty));
            assert_eq!(
                grid.borrow().get(&(1, 1)),
                Some(&expected),
                "{backend:?}: {name}"
            );
        }
        let bound = session
            .bind_values(
                "Api",
                "ObjectArgument",
                &[api::ValueType::HostObject],
                api::ValueType::Void,
            )
            .unwrap();
        let mut call = session
            .begin_values(
                &bound,
                &[api::ValueSlot::Object(root.argument())],
                frame(),
                1000,
            )
            .unwrap();
        assert!(matches!(finish(&mut *call, 1), api::Value::Empty));
        assert_eq!(
            grid.borrow().get(&(1, 1)).unwrap().vtype(),
            calcvba_runtime::VarType::Object
        );
        assert_eq!(
            &*log.borrow(),
            &["Value2=", "Value2=", "Value2=", "Value2=", "Value2="]
        );
        let boxed = session
            .bind_values(
                "Api",
                "Boxed",
                &[api::ValueType::Variant],
                api::ValueType::Variant,
            )
            .unwrap();
        for input in [
            api::Variant::from_i32(17),
            api::Variant::from_string("text"),
            api::Variant::null(),
            api::Variant::empty(),
            api::Variant::from_error_code(2042),
        ] {
            let mut call = session
                .begin_values(&boxed, &[api::ValueSlot::Variant(&input)], frame(), 1000)
                .unwrap();
            let api::Value::Variant(value) = finish(&mut *call, 1) else {
                panic!("Variant result");
            };
            assert_eq!(value, input, "{backend:?}");
        }
        for (name, expected) in [("LiteralBoxed", "text"), ("LocalBoxed", "local")] {
            let bound = session
                .bind_values("Api", name, &[], api::ValueType::Variant)
                .unwrap();
            let mut call = session.begin_values(&bound, &[], frame(), 1000).unwrap();
            let api::Value::Variant(value) = finish(&mut *call, 1) else {
                panic!("Variant result");
            };
            assert_eq!(value.as_bstr().unwrap().as_str(), expected, "{backend:?}");
        }
    }
}

#[test]
fn byval_variant_members_capture_each_argument_before_later_effects() {
    for backend in backends() {
        let mut types = library();
        types.interfaces[0].members.push(api::TypedObjectMember {
            name: "Pair".into(),
            slot: 8,
            kind: api::TypedMemberKind::Method,
            parameters: vec![api::ValueType::Variant, api::ValueType::Variant],
            result: api::ValueType::String,
            is_default: false,
        });
        let project = api::Project::module(
            "Api",
            r#"
Public CurrentValue As Variant
Private Function Mutate() As Long
    CurrentValue = 19&
    Mutate = 19&
End Function
Public Function Run(ByVal root As HostValues.Range) As String
    CurrentValue = 7&
    Run = root.Pair(CurrentValue, Mutate()) & ":" & CurrentValue
End Function
"#,
        );
        let session = Provider::default()
            .load_with_object_libraries(&project, &[types], backend, api::HOST_API_ABI)
            .unwrap();
        let calls = Rc::new(std::cell::Cell::new(0));
        let seen = calls.clone();
        let mut slots = (0..8)
            .map(|_| api::HostSlot::double(|_| panic!("no default or other member read")))
            .collect::<Vec<_>>();
        slots.push(
            api::HostSlot::typed(
                vec![api::ValueType::Variant, api::ValueType::Variant],
                api::ValueType::String,
                move |_, args| {
                    let [
                        api::ValueSlot::Variant(first),
                        api::ValueSlot::Variant(second),
                    ] = args
                    else {
                        return Err(13);
                    };
                    seen.set(seen.get() + 1);
                    Ok(api::Value::String(
                        format!("{}:{}", first.as_i32().unwrap(), second.as_i32().unwrap()).into(),
                    ))
                },
            )
            .unwrap(),
        );
        let root = api::HostObject::new(IID, slots);
        let bound = session
            .bind_values(
                "Api",
                "Run",
                &[api::ValueType::HostObject],
                api::ValueType::String,
            )
            .unwrap();
        let mut call = session
            .begin_values(
                &bound,
                &[api::ValueSlot::Object(root.argument())],
                frame(),
                1000,
            )
            .unwrap();
        let api::Value::String(value) = finish(&mut *call, 1) else {
            panic!("String result");
        };
        assert_eq!(value.as_str(), "7:19:19", "{backend:?}");
        assert_eq!(calls.get(), 1);
    }
}

#[test]
fn scalar_api_refuses_extended_results_before_entering_the_body() {
    for backend in backends() {
        let session = load(&Provider::default(), backend);
        let log = Rc::new(RefCell::new(Vec::new()));
        let root = range(
            Rc::new(RefCell::new(HashMap::new())),
            log.clone(),
            1,
            1,
            1,
            1,
            1,
        );
        let bound = session
            .bind_values(
                "Api",
                "ReadAddress",
                &[api::ValueType::HostObject],
                api::ValueType::String,
            )
            .unwrap();
        assert!(matches!(
            session.begin_invocation(&bound, &[root.borrow_slot()], frame(), 100),
            Err(api::Error::UnsupportedSignature(_))
        ));
        assert!(matches!(
            session.call_slots(
                &bound,
                &[root.borrow_slot()],
                &mut api::TypedContext {
                    remaining_blocks: 100
                }
            ),
            Err(api::Error::UnsupportedSignature(_))
        ));
        assert!(log.borrow().is_empty());
    }
}

#[test]
fn null_errors_and_wrong_interface_results_preserve_their_distinct_outcomes() {
    for backend in backends() {
        let session = load(&Provider::default(), backend);
        let log = Rc::new(RefCell::new(Vec::new()));
        let grid = Rc::new(RefCell::new(HashMap::from([(
            (1, 1),
            api::Variant::null(),
        )])));
        let root = range(grid, log.clone(), 1, 1, 1, 1, 1);
        let add = session
            .bind_values(
                "Api",
                "AddValue",
                &[api::ValueType::HostObject],
                api::ValueType::Variant,
            )
            .unwrap();
        let mut call = session
            .begin_values(
                &add,
                &[api::ValueSlot::Object(root.argument())],
                frame(),
                1000,
            )
            .unwrap();
        let api::Value::Variant(value) = finish(&mut *call, 1) else {
            panic!("Variant expected");
        };
        assert_eq!(value, api::Variant::null());
        assert_eq!(&*log.borrow(), &["Value2"]);
        let numeric = session
            .bind_values(
                "Api",
                "ReadNumeric",
                &[api::ValueType::HostObject],
                api::ValueType::Double,
            )
            .unwrap();
        let mut call = session
            .begin_values(
                &numeric,
                &[api::ValueSlot::Object(root.argument())],
                frame(),
                1000,
            )
            .unwrap();
        assert!(matches!(
            call.resume(u64::MAX),
            Err(api::Error::Runtime { number: 94, .. })
        ));
        let explain = session
            .bind_values(
                "Api",
                "ExplainNull",
                &[api::ValueType::HostObject],
                api::ValueType::String,
            )
            .unwrap();
        let mut call = session
            .begin_values(
                &explain,
                &[api::ValueSlot::Object(root.argument())],
                frame(),
                1000,
            )
            .unwrap();
        let api::Value::String(description) = finish(&mut *call, 1) else {
            panic!("description expected");
        };
        assert_eq!(description.as_str(), "Invalid use of Null");
        let address = session
            .bind_values(
                "Api",
                "ReadAddress",
                &[api::ValueType::HostObject],
                api::ValueType::String,
            )
            .unwrap();
        let mut call = session
            .begin_values(&address, &[api::ValueSlot::Nothing], frame(), 1000)
            .unwrap();
        assert!(matches!(
            call.resume(u64::MAX),
            Err(api::Error::Runtime { number: 91, .. })
        ));
        let witness = Rc::new(RefCell::new(None));
        let seen = witness.clone();
        let root = api::HostObject::new(
            IID,
            vec![
                api::HostSlot::double(|_| Ok(0.0)),
                api::HostSlot::double(|_| Ok(0.0)),
                api::HostSlot::typed(
                    vec![api::ValueType::Long, api::ValueType::Long],
                    api::ValueType::HostObject,
                    move |_, _| {
                        let owner = Rc::new(());
                        *seen.borrow_mut() = Some(Rc::downgrade(&owner));
                        Ok(api::Value::Object(Some(api::HostObject::new(
                            [99; 16],
                            vec![api::HostSlot::double(move |_| {
                                let _keep = &owner;
                                Ok(0.0)
                            })],
                        ))))
                    },
                )
                .unwrap(),
            ],
        );
        let mut call = session
            .begin_values(
                &address,
                &[api::ValueSlot::Object(root.argument())],
                frame(),
                1000,
            )
            .unwrap();
        assert!(matches!(
            call.resume(u64::MAX),
            Err(api::Error::Runtime { number: 13, .. })
        ));
        assert!(
            witness.borrow().as_ref().unwrap().upgrade().is_none(),
            "rejected native result must release its transferred owner"
        );
    }
}

#[test]
fn minimal_bstr_and_variant_slots_own_their_exact_storage() {
    for backend in backends() {
        let project = api::Project::module(
            "Api",
            "Public Function Text(ByVal value As String) As String\nText = value\nEnd Function\nPublic Function Exact(ByVal value As Variant) As Variant\nExact = value\nEnd Function\n",
        );
        let mut session = Provider::default()
            .load(&project, backend, api::HOST_API_ABI)
            .unwrap();
        let text = api::BStr::from_bytes(&[0x61, 0, 0x55]).unwrap();
        let echo = session
            .bind_values(
                "Api",
                "Text",
                &[api::ValueType::String],
                api::ValueType::String,
            )
            .unwrap();
        let mut call = session
            .begin_values(&echo, &[api::ValueSlot::String(&text)], frame(), 1000)
            .unwrap();
        drop(text);
        session
            .reload(&api::Project::module(
                "Api",
                "Public Function Text(ByVal value As Long) As Long\nText = value\nEnd Function\n",
            ))
            .unwrap();
        drop(session);
        let api::Value::String(value) = finish(&mut *call, 1) else {
            panic!("BSTR expected");
        };
        assert_eq!(value.payload_bytes(), &[0x61, 0, 0x55]);
        let session = Provider::default()
            .load(&project, backend, api::HOST_API_ABI)
            .unwrap();
        let echo = session
            .bind_values(
                "Api",
                "Exact",
                &[api::ValueType::Variant],
                api::ValueType::Variant,
            )
            .unwrap();
        for input in [
            api::Variant::from_f64(42.0),
            api::Variant::null(),
            api::Variant::from_error_code(2042),
        ] {
            let mut call = session
                .begin_values(&echo, &[api::ValueSlot::Variant(&input)], frame(), 1000)
                .unwrap();
            let api::Value::Variant(value) = finish(&mut *call, 1) else {
                panic!("Variant expected");
            };
            assert_eq!(value, input);
        }
    }
}

#[test]
fn ignored_owning_member_result_is_released_before_the_next_callback() {
    for backend in backends() {
        let iid = [90; 16];
        let library = api::TypedObjectLibrary {
            name: "Objects".into(),
            interfaces: vec![api::TypedObjectInterface {
                name: "Root".into(),
                iid,
                members: vec![api::TypedObjectMember {
                    name: "Next".into(),
                    slot: 0,
                    kind: api::TypedMemberKind::Method,
                    parameters: vec![],
                    result: api::ValueType::Interface(iid),
                    is_default: false,
                }],
            }],
        };
        let session = Provider::default().load_with_object_libraries(
            &api::Project::module("Api", "Public Sub Ignore(ByVal root As Objects.Root)\nroot.Next\nroot.Next\nEnd Sub\n"),
            &[library], backend, api::HOST_API_ABI,
        ).unwrap();
        let observed = Rc::new(RefCell::new(Vec::<std::rc::Weak<()>>::new()));
        let seen = observed.clone();
        let root = api::HostObject::new(
            iid,
            vec![
                api::HostSlot::typed(vec![], api::ValueType::Interface(iid), move |_, _| {
                    assert!(seen.borrow().iter().all(|owner| owner.upgrade().is_none()));
                    let owner = Rc::new(());
                    seen.borrow_mut().push(Rc::downgrade(&owner));
                    let object = api::HostObject::new(
                        iid,
                        vec![api::HostSlot::double(move |_| {
                            let _ = &owner;
                            Ok(0.0)
                        })],
                    );
                    Ok(api::Value::Object(Some(object)))
                })
                .unwrap(),
            ],
        );
        let bound = session
            .bind_values(
                "Api",
                "Ignore",
                &[api::ValueType::HostObject],
                api::ValueType::Void,
            )
            .unwrap();
        let mut call = session
            .begin_values(
                &bound,
                &[api::ValueSlot::Object(root.argument())],
                frame(),
                1000,
            )
            .unwrap();
        assert!(matches!(finish(&mut *call, 1), api::Value::Empty));
        assert_eq!(observed.borrow().len(), 2);
        assert!(
            observed
                .borrow()
                .iter()
                .all(|owner| owner.upgrade().is_none())
        );
    }
}

#[test]
fn rich_callback_fault_fields_survive_handlers_slices_and_unhandled_return() {
    for backend in backends() {
        let mut library = library();
        library.interfaces[0].members = vec![api::TypedObjectMember {
            name: "Value".into(),
            slot: 0,
            kind: api::TypedMemberKind::PropertyGet,
            parameters: vec![],
            result: api::ValueType::Long,
            is_default: false,
        }];
        let project = api::Project::module(
            "Api",
            r#"
Public Function Direct(ByVal root As HostValues.Range) As Long
    Direct = root.Value
End Function
Public Function Describe(ByVal root As HostValues.Range) As String
    Dim value As Long
    On Error GoTo DescribeFault
    value = root.Value
    Exit Function
DescribeFault:
    Describe = Err.Source & "|" & Err.Description & "|" & Err.HelpFile
End Function
Public Function Help(ByVal root As HostValues.Range) As Long
    On Error GoTo DescribeFault
    Help = root.Value
    Exit Function
DescribeFault:
    Help = Err.HelpContext
End Function
"#,
        );
        let session = Provider::default()
            .load_with_object_libraries(&project, &[library], backend, api::HOST_API_ABI)
            .unwrap();
        let object = api::HostObject::new(
            IID,
            vec![
                api::HostSlot::typed(vec![], api::ValueType::Long, |context, _| {
                    Err(api::TypedFault {
                        number: 57,
                        description: "exact server description".into(),
                        source: Some("Server.Source".into()),
                        help_file: Some("server.chm".into()),
                        help_context: Some(42),
                    }
                    .publish(context))
                })
                .unwrap(),
            ],
        );
        for (name, result) in [
            ("Describe", api::ValueType::String),
            ("Help", api::ValueType::Long),
            ("Direct", api::ValueType::Long),
        ] {
            let bound = session
                .bind_values("Api", name, &[api::ValueType::HostObject], result)
                .unwrap();
            let mut call = session
                .begin_values(
                    &bound,
                    &[api::ValueSlot::Object(object.argument())],
                    frame(),
                    1000,
                )
                .unwrap();
            loop {
                match call.resume(1) {
                    Ok(api::ValueProgress::Suspended) => {}
                    Ok(api::ValueProgress::Returned(api::Value::String(value))) => {
                        assert_eq!(
                            value.as_str(),
                            "Server.Source|exact server description|server.chm"
                        );
                        break;
                    }
                    Ok(api::ValueProgress::Returned(api::Value::Long(value))) => {
                        assert_eq!(name, "Help");
                        assert_eq!(value, 42);
                        break;
                    }
                    Err(api::Error::Runtime {
                        number,
                        description,
                        vba_source,
                    }) => {
                        assert_eq!(name, "Direct");
                        assert_eq!(
                            (number, description.as_str(), vba_source.as_str()),
                            (57, "exact server description", "Server.Source")
                        );
                        break;
                    }
                    other => panic!("unexpected outcome {}", other.is_err()),
                }
            }
        }
    }
}

#[test]
fn variant_byref_aliases_keep_actual_cells_and_writes_after_child_fault() {
    for backend in backends() {
        let session = load(&Provider::default(), backend);
        let log = Rc::new(RefCell::new(Vec::new()));
        let root = range(
            Rc::new(RefCell::new(HashMap::from([(
                (1, 1),
                api::Variant::from_f64(42.0),
            )]))),
            log.clone(),
            1,
            1,
            1,
            1,
            1,
        );
        let bound = session
            .bind_values(
                "Api",
                "VariantAliases",
                &[api::ValueType::HostObject],
                api::ValueType::Variant,
            )
            .unwrap();
        let mut call = session
            .begin_values(
                &bound,
                &[api::ValueSlot::Object(root.argument())],
                frame(),
                1000,
            )
            .unwrap();
        let api::Value::Variant(value) = finish(&mut *call, 1) else {
            panic!("Variant expected")
        };
        assert_eq!(value, api::Variant::from_f64(44.0));
        assert_eq!(&*log.borrow(), &["Value2"]);
    }
}
