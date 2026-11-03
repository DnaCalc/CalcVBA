use calcvba_host::provider::Provider;
use calcvba_host_api::{self as api, RuntimeProvider};
use std::{cell::RefCell, rc::Rc};

const IID: api::InterfaceId = [101; 16];
fn library() -> api::TypedObjectLibrary {
    api::TypedObjectLibrary {
        name: "Defaults".into(),
        interfaces: vec![
            api::TypedObjectInterface {
                name: "ValueObject".into(),
                iid: IID,
                members: vec![api::TypedObjectMember {
                    name: "Read".into(),
                    slot: 0,
                    kind: api::TypedMemberKind::PropertyGet,
                    parameters: vec![],
                    result: api::ValueType::Variant,
                    is_default: true,
                }],
            },
            api::TypedObjectInterface {
                name: "Required".into(),
                iid: [103; 16],
                members: vec![api::TypedObjectMember {
                    name: "At".into(),
                    slot: 0,
                    kind: api::TypedMemberKind::PropertyGet,
                    parameters: vec![api::ValueType::Long],
                    result: api::ValueType::Variant,
                    is_default: true,
                }],
            },
        ],
    }
}
const SOURCE: &str = r#"
Public Function Add(ByVal x As Variant, ByVal y As Variant) As Variant
    Add = x + y
End Function
Public Function Text(ByVal x As Variant) As String
    Text = "v=" & x
End Function
Public Function Number(ByVal x As Variant) As Double
    Number = x
End Function
Public Function EqualValue(ByVal x As Variant) As Boolean
    EqualValue = x = 41#
End Function
Public Function Condition(ByVal x As Variant) As Long
    If x Then
        Condition = 1
    End If
End Function
Public Function Catch(ByVal x As Variant) As Long
    On Error GoTo Failed
    Dim n As Double
    n = x
    Exit Function
Failed:
    Catch = Err.Number
End Function
Public Function Echo(ByVal x As Variant) As Variant
    Echo = x
End Function
Public Function UntypedObject(ByVal x As Object) As Variant
    UntypedObject = x + 1#
End Function
Public Function Divide(ByVal x As Variant) As Double
    Divide = x / 2#
End Function
Public Function Negative(ByVal x As Variant) As Variant
    Negative = -x
End Function
Private Sub Mutate(ByRef x As Variant)
    x = x + 1#
End Sub
Public Function ThroughAlias(ByVal x As Variant) As Variant
    Mutate x
    ThroughAlias = x
End Function
"#;
#[test]
fn runtime_defaults_preserve_identity_order_faults_and_slice_effects() {
    for backend in [api::Backend::Interpreter, api::Backend::Jit] {
        if backend == api::Backend::Jit && !cfg!(feature = "jit") {
            continue;
        }
        let session = Provider::default()
            .load_with_object_libraries(
                &api::Project::module("Api", SOURCE),
                &[library()],
                backend,
                api::HOST_API_ABI,
            )
            .unwrap();
        let log = Rc::new(RefCell::new(Vec::new()));
        let object = |id, value: api::Variant| {
            let log = log.clone();
            api::HostObject::new(
                IID,
                vec![
                    api::HostSlot::typed(vec![], api::ValueType::Variant, move |_, _| {
                        assert_eq!(api::active_invocation().unwrap().id, 10);
                        log.borrow_mut().push(id);
                        Ok(api::Value::Variant(value.clone()))
                    })
                    .unwrap(),
                ],
            )
        };
        let left = object(1, api::Variant::from_f64(41.0));
        let right = object(2, api::Variant::from_f64(1.0));
        let x = api::Variant::from_object_ref(left.object_ref());
        let y = api::Variant::from_object_ref(right.object_ref());
        let invoke = |name: &str, inputs: &[api::ValueSlot<'_>], result| {
            let types = inputs
                .iter()
                .map(|_| api::ValueType::Variant)
                .collect::<Vec<_>>();
            let bound = session.bind_values("Api", name, &types, result).unwrap();
            let mut call = session
                .begin_values(
                    &bound,
                    inputs,
                    api::InvocationFrame::new(10, None, Rc::new(())),
                    1000,
                )
                .unwrap();
            loop {
                match call.resume(1).unwrap() {
                    api::ValueProgress::Suspended => {}
                    api::ValueProgress::Returned(value) => break value,
                }
            }
        };
        let result = invoke(
            "Add",
            &[api::ValueSlot::Variant(&x), api::ValueSlot::Variant(&y)],
            api::ValueType::Variant,
        );
        assert!(matches!(result, api::Value::Variant(v) if v.as_f64() == Some(42.0)));
        assert_eq!(*log.borrow(), vec![1, 2]);
        log.borrow_mut().clear();
        assert!(
            matches!(invoke("Text", &[api::ValueSlot::Variant(&x)], api::ValueType::String), api::Value::String(v) if v.as_str() == "v=41")
        );
        assert!(matches!(
            invoke(
                "Number",
                &[api::ValueSlot::Variant(&x)],
                api::ValueType::Double
            ),
            api::Value::Double(41.0)
        ));
        assert!(matches!(
            invoke(
                "EqualValue",
                &[api::ValueSlot::Variant(&x)],
                api::ValueType::Boolean
            ),
            api::Value::Boolean(true)
        ));
        assert!(matches!(
            invoke(
                "Condition",
                &[api::ValueSlot::Variant(&x)],
                api::ValueType::Long
            ),
            api::Value::Long(1)
        ));
        assert_eq!(*log.borrow(), vec![1, 1, 1, 1]);
        assert!(matches!(
            invoke(
                "Divide",
                &[api::ValueSlot::Variant(&x)],
                api::ValueType::Double
            ),
            api::Value::Double(20.5)
        ));
        assert!(
            matches!(invoke("Negative", &[api::ValueSlot::Variant(&x)], api::ValueType::Variant), api::Value::Variant(v) if v.as_f64() == Some(-41.0))
        );
        assert!(
            matches!(invoke("ThroughAlias", &[api::ValueSlot::Variant(&x)], api::ValueType::Variant), api::Value::Variant(v) if v.as_f64() == Some(42.0))
        );
        let outer = object(4, x.clone());
        let chain = api::Variant::from_object_ref(outer.object_ref());
        assert!(matches!(
            invoke(
                "Number",
                &[api::ValueSlot::Variant(&chain)],
                api::ValueType::Double
            ),
            api::Value::Double(41.0)
        ));
        let failure = api::HostObject::new(
            IID,
            vec![
                api::HostSlot::typed(vec![], api::ValueType::Variant, |context, _| {
                    Err(api::TypedFault {
                        number: 27,
                        description: "default failed".into(),
                        source: Some("default".into()),
                        help_file: None,
                        help_context: None,
                    }
                    .publish(context))
                })
                .unwrap(),
            ],
        );
        let failure = api::Variant::from_object_ref(failure.object_ref());
        assert!(matches!(
            invoke(
                "Catch",
                &[api::ValueSlot::Variant(&failure)],
                api::ValueType::Long
            ),
            api::Value::Long(27)
        ));
        let bound = session
            .bind_values(
                "Api",
                "UntypedObject",
                &[api::ValueType::HostObject],
                api::ValueType::Variant,
            )
            .unwrap();
        let mut call = session
            .begin_values(
                &bound,
                &[api::ValueSlot::Object(left.argument())],
                api::InvocationFrame::new(10, None, Rc::new(())),
                1000,
            )
            .unwrap();
        loop {
            match call.resume(1).unwrap() {
                api::ValueProgress::Suspended => {}
                api::ValueProgress::Returned(api::Value::Variant(v)) => {
                    assert_eq!(v.as_f64(), Some(42.0));
                    break;
                }
                _ => panic!("Variant expected"),
            }
        }
        drop(call);
        let before = log.borrow().len();
        assert!(
            matches!(invoke("Echo", &[api::ValueSlot::Variant(&x)], api::ValueType::Variant), api::Value::Variant(v) if v.as_object_ref().unwrap().raw() == left.object_ref().raw())
        );
        assert_eq!(log.borrow().len(), before);
        for (input, expected) in [(api::Variant::nothing(), 91), (api::Variant::null(), 94)] {
            assert!(
                matches!(invoke("Catch", &[api::ValueSlot::Variant(&input)], api::ValueType::Long), api::Value::Long(n) if n == expected)
            );
        }
        let bad = object(3, api::Variant::from_string("bad"));
        let bad = api::Variant::from_object_ref(bad.object_ref());
        assert!(matches!(
            invoke(
                "Catch",
                &[api::ValueSlot::Variant(&bad)],
                api::ValueType::Long
            ),
            api::Value::Long(13)
        ));
        let missing = api::HostObject::new([102; 16], vec![]);
        let missing = api::Variant::from_object_ref(missing.object_ref());
        assert!(matches!(
            invoke(
                "Catch",
                &[api::ValueSlot::Variant(&missing)],
                api::ValueType::Long
            ),
            api::Value::Long(438)
        ));
        let required = api::HostObject::new([103; 16], vec![]);
        let required = api::Variant::from_object_ref(required.object_ref());
        assert!(matches!(
            invoke(
                "Catch",
                &[api::ValueSlot::Variant(&required)],
                api::ValueType::Long
            ),
            api::Value::Long(450)
        ));
        let mut chain = x.clone();
        for depth in 1..=64 {
            let next = object(6, chain);
            chain = api::Variant::from_object_ref(next.object_ref());
            if depth == 63 {
                assert!(matches!(
                    invoke(
                        "Number",
                        &[api::ValueSlot::Variant(&chain)],
                        api::ValueType::Double
                    ),
                    api::Value::Double(41.0)
                ));
            }
        }
        assert!(matches!(
            invoke(
                "Catch",
                &[api::ValueSlot::Variant(&chain)],
                api::ValueType::Long
            ),
            api::Value::Long(28)
        ));
        for (payload, expected) in [
            (api::Variant::empty(), 0.0),
            (api::Variant::from_string("3"), 3.0),
        ] {
            let source = object(5, payload);
            let input = api::Variant::from_object_ref(source.object_ref());
            assert!(
                matches!(invoke("Number", &[api::ValueSlot::Variant(&input)], api::ValueType::Double), api::Value::Double(n) if n == expected)
            );
        }
    }
}
