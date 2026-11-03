use calcvba_host::provider::Provider;
use calcvba_host_api::{self as api, RuntimeProvider};
use std::{cell::Cell, rc::Rc};

const SOURCE: &str = r#"
Public Function Same(ByVal left As Variant, ByVal right As Variant) As Boolean
    Same = left Is right
End Function
Public Function Objects(ByVal left As Object, ByVal right As Object) As Boolean
    Objects = left Is right
End Function
Public Function CopySame(ByVal value As Object) As Boolean
    Dim other As Object
    Set other = value
    CopySame = other Is value
End Function
Public Function Catch(ByVal left As Variant, ByVal right As Variant) As Long
    On Error GoTo Failed
    Dim result As Boolean
    result = left Is right
    Catch = 1
    Exit Function
Failed:
    Catch = Err.Number
End Function
"#;

fn run(
    session: &dyn api::Session,
    name: &str,
    args: &[api::ValueSlot<'_>],
    ty: api::ValueType,
) -> api::Value {
    let types = args.iter().map(|a| a.ty()).collect::<Vec<_>>();
    let handle = session.bind_values("Api", name, &types, ty).unwrap();
    let mut call = session
        .begin_values(
            &handle,
            args,
            api::InvocationFrame::new(11, None, Rc::new(())),
            10000,
        )
        .unwrap();
    for _ in 0..10000 {
        match call.resume(1).unwrap() {
            api::ValueProgress::Suspended => {}
            api::ValueProgress::Returned(value) => return value,
        }
    }
    panic!("invocation did not complete")
}

#[test]
fn typed_is_preserves_identity_nothing_and_object_required_without_defaults() {
    for backend in [api::Backend::Interpreter, api::Backend::Jit] {
        if backend == api::Backend::Jit && !cfg!(feature = "jit") {
            continue;
        }
        let session = Provider::default()
            .load(
                &api::Project::module("Api", SOURCE),
                backend,
                api::HOST_API_ABI,
            )
            .unwrap();
        let defaults = Rc::new(Cell::new(0));
        let counter = defaults.clone();
        let first = api::HostObject::new(
            [11; 16],
            vec![api::HostSlot::double(move |_| {
                counter.set(counter.get() + 1);
                Ok(42.0)
            })],
        );
        let second = api::HostObject::new([11; 16], vec![]);
        let a = api::Variant::from_object_ref(first.object_ref());
        let b = api::Variant::from_object_ref(second.object_ref());
        let nothing = api::Variant::nothing();
        for (left, right, expected) in [
            (&a, &a, true),
            (&a, &b, false),
            (&a, &nothing, false),
            (&nothing, &a, false),
            (&nothing, &nothing, true),
        ] {
            let result = run(
                session.as_ref(),
                "Same",
                &[
                    api::ValueSlot::Variant(left),
                    api::ValueSlot::Variant(right),
                ],
                api::ValueType::Boolean,
            );
            assert!(
                matches!(result, api::Value::Boolean(v) if v == expected),
                "{backend:?}: {result:?}"
            );
        }
        for (args, expected) in [
            (
                [
                    api::ValueSlot::Object(first.argument()),
                    api::ValueSlot::Object(first.argument()),
                ],
                true,
            ),
            (
                [
                    api::ValueSlot::Object(first.argument()),
                    api::ValueSlot::Object(second.argument()),
                ],
                false,
            ),
            ([api::ValueSlot::Nothing, api::ValueSlot::Nothing], true),
        ] {
            assert!(
                matches!(run(session.as_ref(), "Objects", &args, api::ValueType::Boolean), api::Value::Boolean(v) if v == expected)
            );
        }
        assert!(matches!(
            run(
                session.as_ref(),
                "CopySame",
                &[api::ValueSlot::Object(first.argument())],
                api::ValueType::Boolean
            ),
            api::Value::Boolean(true)
        ));
        for bad in [
            api::Variant::empty(),
            api::Variant::null(),
            api::Variant::from_i32(0),
            api::Variant::from_string("x"),
            api::Variant::from_error_code(2042),
        ] {
            for args in [
                [api::ValueSlot::Variant(&bad), api::ValueSlot::Variant(&a)],
                [api::ValueSlot::Variant(&a), api::ValueSlot::Variant(&bad)],
            ] {
                assert!(matches!(
                    run(session.as_ref(), "Catch", &args, api::ValueType::Long),
                    api::Value::Long(424)
                ));
            }
        }
        assert_eq!(defaults.get(), 0);
    }
}
