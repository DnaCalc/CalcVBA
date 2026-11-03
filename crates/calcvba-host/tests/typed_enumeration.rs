use calcvba_host::provider::Provider;
use calcvba_host_api::{self as api, RuntimeProvider};
use std::{cell::Cell, rc::Rc};

const SOURCE_IID: api::InterfaceId = [111; 16];
const ITEM_IID: api::InterfaceId = [112; 16];
const SOURCE: &str = r#"
Public Function SumItems(ByVal items As Iter.Source) As Double
    Dim item As Iter.ValueItem
    For Each item In items
        SumItems = SumItems + item.Read
    Next
End Function
Public Function FirstItem(ByVal items As Iter.Source) As Double
    Dim item As Iter.ValueItem
    For Each item In items
        FirstItem = item.Read
        Exit For
    Next
End Function
Public Function CatchItems(ByVal items As Iter.Source) As Long
    On Error GoTo Failed
    Dim item As Iter.ValueItem
    For Each item In items
        Dim value As Double
        value = item.Read
    Next
    Exit Function
Failed:
    CatchItems = Err.Number
End Function
"#;
fn library() -> api::TypedObjectLibrary {
    api::TypedObjectLibrary {
        name: "Iter".into(),
        interfaces: vec![
            api::TypedObjectInterface {
                name: "Source".into(),
                iid: SOURCE_IID,
                members: vec![],
            },
            api::TypedObjectInterface {
                name: "ValueItem".into(),
                iid: ITEM_IID,
                members: vec![api::TypedObjectMember {
                    name: "Read".into(),
                    slot: 0,
                    kind: api::TypedMemberKind::PropertyGet,
                    parameters: vec![],
                    result: api::ValueType::Double,
                    is_default: false,
                }],
            },
        ],
    }
}
#[derive(Default)]
struct State {
    advances: Cell<usize>,
    released: Cell<usize>,
    second: Cell<f64>,
    fail: Cell<bool>,
}
struct Cursor(Rc<State>);
impl Drop for Cursor {
    fn drop(&mut self) {
        self.0.released.set(self.0.released.get() + 1);
    }
}
struct Services(Rc<State>);
impl api::HostObjectServices for Services {
    fn enumerator(&self, _: &mut api::TypedContext) -> Result<api::HostEnumerator, i32> {
        let cursor = Cursor(self.0.clone());
        let mut position = 0;
        Ok(api::HostEnumerator::new(move |_| {
            let state = cursor.0.clone();
            state.advances.set(state.advances.get() + 1);
            if state.fail.get() && position == 1 {
                return Err(9);
            }
            if position == 2 {
                return Ok(None);
            }
            let item_index = position;
            position += 1;
            Ok(Some(api::Value::Object(Some(api::HostObject::new(
                ITEM_IID,
                vec![
                    api::HostSlot::typed(vec![], api::ValueType::Double, move |_, _| {
                        if item_index == 0 {
                            state.second.set(20.0);
                            Ok(api::Value::Double(1.0))
                        } else {
                            Ok(api::Value::Double(state.second.get()))
                        }
                    })
                    .unwrap(),
                ],
            )))))
        }))
    }
}
fn object(state: Rc<State>) -> api::HostObject {
    let identity = Rc::as_ptr(&state).cast();
    let services = Rc::new(Services(state.clone()));
    // SAFETY: the payload and services both retain this exact unique Rc identity.
    unsafe { api::HostObject::new_with_services(SOURCE_IID, vec![], state, identity, services) }
}
#[test]
fn live_iterator_preserves_slices_exit_errors_and_cancel_ownership() {
    for backend in [api::Backend::Interpreter, api::Backend::Jit]
        .into_iter()
        .filter(|b| *b != api::Backend::Jit || cfg!(feature = "jit"))
    {
        let session = Provider::default()
            .load_with_object_libraries(
                &api::Project::module("Api", SOURCE),
                &[library()],
                backend,
                api::HOST_API_ABI,
            )
            .unwrap();
        let state = Rc::new(State::default());
        let source = object(state.clone());
        for (name, expected, advances) in [("SumItems", 21.0, 3), ("FirstItem", 1.0, 1)] {
            let handle = session
                .bind_values(
                    "Api",
                    name,
                    &[api::ValueType::HostObject],
                    api::ValueType::Double,
                )
                .unwrap();
            let before = state.advances.get();
            let releases = state.released.get();
            let mut invocation = session
                .begin_values(
                    &handle,
                    &[api::ValueSlot::Object(source.argument())],
                    api::InvocationFrame::new(1, None, Rc::new(())),
                    1000,
                )
                .unwrap();
            loop {
                if let api::ValueProgress::Returned(value) = invocation.resume(1).unwrap() {
                    assert!(matches!(value,api::Value::Double(n) if n==expected));
                    break;
                }
            }
            assert_eq!(state.advances.get() - before, advances);
            assert_eq!(state.released.get(), releases + 1);
        }
        state.fail.set(true);
        let handle = session
            .bind_values(
                "Api",
                "CatchItems",
                &[api::ValueType::HostObject],
                api::ValueType::Long,
            )
            .unwrap();
        let releases = state.released.get();
        let value = session
            .call_values(
                &handle,
                &[api::ValueSlot::Object(source.argument())],
                &mut api::TypedContext {
                    remaining_blocks: 1000,
                },
            )
            .unwrap();
        assert!(matches!(value, api::Value::Long(9)));
        assert_eq!(state.released.get(), releases + 1);
        state.fail.set(false);
        let handle = session
            .bind_values(
                "Api",
                "SumItems",
                &[api::ValueType::HostObject],
                api::ValueType::Double,
            )
            .unwrap();
        let before = state.advances.get();
        let releases = state.released.get();
        let mut invocation = session
            .begin_values(
                &handle,
                &[api::ValueSlot::Object(source.argument())],
                api::InvocationFrame::new(1, None, Rc::new(())),
                1000,
            )
            .unwrap();
        while state.advances.get() == before {
            invocation.resume(1).unwrap();
        }
        drop(invocation);
        assert_eq!(state.released.get(), releases + 1);
        drop(source);
        assert_eq!(Rc::strong_count(&state), 1);
    }
}
