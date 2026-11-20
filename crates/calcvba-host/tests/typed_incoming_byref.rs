use api::BStr;
use calcvba_host::provider::Provider;
use calcvba_host_api::{self as api, RuntimeProvider};
const SOURCE: &str = r#"
Public hits As Long
Public Sub Receive(ByRef Cancel As Boolean, ByVal number As Long, ByRef text As String)
    Cancel = True
    number = 99
    text = text & "!"
End Sub
Public Sub Fail(ByRef Cancel As Boolean)
    Cancel = True
    Err.Raise 5
End Sub
Public Sub Structural(ByRef value As Variant)
    hits = hits + 1
    value = hits
End Sub
Public Function ReadHits() As Long
    ReadHits = hits
End Function
"#;
#[test]
fn byref_writes_reach_the_caller_values_including_errors_and_byval_isolation() {
    for backend in [api::Backend::Interpreter, api::Backend::Jit]
        .into_iter()
        .filter(|b| *b != api::Backend::Jit || cfg!(feature = "jit"))
    {
        let session = Provider::default()
            .load(
                &api::Project::module("Api", SOURCE),
                backend,
                api::HOST_API_ABI,
            )
            .unwrap();
        let receive = session
            .bind_values(
                "Api",
                "Receive",
                &[
                    api::ValueType::Boolean,
                    api::ValueType::Long,
                    api::ValueType::String,
                ],
                api::ValueType::Void,
            )
            .unwrap();
        let mut args = [
            api::Value::Boolean(false),
            api::Value::Long(7),
            api::Value::String(BStr::from("before")),
        ];
        let mut context = api::TypedContext {
            remaining_blocks: 1000,
        };
        session
            .call_values_with_writebacks(&receive, &mut args, &mut context)
            .unwrap();
        assert!(matches!(args[0], api::Value::Boolean(true)));
        assert!(matches!(args[1], api::Value::Long(7)));
        assert!(matches!(&args[2],api::Value::String(s) if s.as_str()=="before!"));
        let fail = session
            .bind_values(
                "Api",
                "Fail",
                &[api::ValueType::Boolean],
                api::ValueType::Void,
            )
            .unwrap();
        let mut args = [api::Value::Boolean(false)];
        assert!(matches!(
            session.call_values_with_writebacks(&fail, &mut args, &mut context),
            Err(api::Error::Runtime { number: 5, .. })
        ));
        assert!(matches!(args[0], api::Value::Boolean(true)));
        let structural = session
            .bind_values(
                "Api",
                "Structural",
                &[api::ValueType::Variant],
                api::ValueType::Void,
            )
            .unwrap();
        // A ByRef Variant is the caller's cell too (was refused as structural).
        let mut args = [api::Value::Variant(calcvba_runtime::Variant::empty())];
        session
            .call_values_with_writebacks(&structural, &mut args, &mut context)
            .unwrap();
        assert!(matches!(&args[0], api::Value::Variant(v) if v.as_i32() == Some(1)));
        let read = session
            .bind_values("Api", "ReadHits", &[], api::ValueType::Long)
            .unwrap();
        assert!(matches!(
            session.call_values(&read, &[], &mut context).unwrap(),
            api::Value::Long(1)
        ));
    }
}

const PROBE_IID: api::InterfaceId = *b"ByRef-cell-probe";
const CELLS: &str = r#"
Public Function Observe(ByRef total As Double, ByRef text As String, ByRef v As Variant, ByVal probe As Probes.Probe) As Double
    total = 41.5
    Bump total
    text = "seen"
    v = Array(1, "two")
    Observe = probe.Look
    total = total + 1
End Function
Private Sub Bump(ByRef x As Double)
    x = x + 0.5
End Sub
Public Sub Twice(ByRef a As Long, ByRef b As Long)
    a = a + 1
    b = b + 10
End Sub
Public Sub Grow(ByRef items() As Variant, ByVal n As Long)
    ReDim Preserve items(n)
    items(n) = "last"
End Sub
Public Sub Copy(ByVal n As Long)
    n = 5
End Sub
"#;
fn cells_project() -> api::Project {
    let mut project = api::Project::module("Api", CELLS);
    project.libraries.push(api::ObjectLibrary {
        name: "Probes".into(),
        interfaces: vec![api::ObjectInterface {
            name: "Probe".into(),
            iid: PROBE_IID,
            members: vec![api::ObjectMember {
                name: "Look".into(),
                slot: 0,
                kind: api::MemberKind::PropertyGet,
                parameters: vec![],
                result: api::ScalarType::Double,
                is_default: false,
            }],
        }],
    });
    project
}
/// ByRef parameters given host cells are those cells: a host callback during the call
/// sees each write (also through a nested ByRef call), two parameters given one cell
/// share it, a ByRef array ReDim lands in the cell, and nothing is copied back.
#[test]
fn byref_cells_alias_host_storage_during_the_call() {
    use std::rc::Rc;
    for backend in [api::Backend::Interpreter, api::Backend::Jit]
        .into_iter()
        .filter(|b| *b != api::Backend::Jit || cfg!(feature = "jit"))
    {
        let session: Rc<dyn api::Session> = Provider::default()
            .load(&cells_project(), backend, api::HOST_API_ABI)
            .unwrap()
            .into();
        let copy = session
            .bind_values("Api", "Copy", &[api::ValueType::Long], api::ValueType::Void)
            .unwrap();
        let total = Rc::new(api::ByRefCell::new(api::Value::Double(0.0)));
        let text = Rc::new(api::ByRefCell::new(api::Value::String(BStr::from(
            "before",
        ))));
        let variant = Rc::new(api::ByRefCell::new(api::Value::Variant(
            calcvba_runtime::Variant::empty(),
        )));
        let seen = Rc::new(std::cell::RefCell::new(String::new()));
        let probe = {
            let (total, text, variant, seen) =
                (total.clone(), text.clone(), variant.clone(), seen.clone());
            let (weak, copy) = (Rc::downgrade(&session), copy.clone());
            api::HostObject::new(
                PROBE_IID,
                vec![api::HostSlot::double(move |context| {
                    let t = match total.get().unwrap() {
                        api::Value::Double(t) => t,
                        other => panic!("{other:?}"),
                    };
                    let s = match text.get().unwrap() {
                        api::Value::String(s) => s.as_str(),
                        other => panic!("{other:?}"),
                    };
                    let v = match variant.get().unwrap() {
                        api::Value::Variant(v) => format!(
                            "{}:{}",
                            v.safearray_element(0).unwrap().unwrap().as_i32().unwrap(),
                            v.safearray_element(1)
                                .unwrap()
                                .unwrap()
                                .as_bstr()
                                .unwrap()
                                .as_str()
                        ),
                        other => panic!("{other:?}"),
                    };
                    // The call holds the cells: the host cannot replace them meanwhile.
                    let refused = total.set(api::Value::Double(-1.0)).is_err();
                    // A call from this callback may use a fresh cell, not a held one.
                    let session = weak.upgrade().unwrap();
                    let fresh = api::ByRefCell::new(api::Value::Long(1));
                    let nested = session
                        .call_values_by_ref(&copy, &[api::ValueArg::ByRef(&fresh)], context)
                        .is_ok();
                    let held = matches!(
                        session.call_values_by_ref(&copy, &[api::ValueArg::ByRef(&total)], context),
                        Err(api::Error::Reentry)
                    );
                    *seen.borrow_mut() = format!("{t}|{s}|{v}|{refused}|{nested}|{held}");
                    Ok(7.0)
                })],
            )
        };
        let observe = session
            .bind_values(
                "Api",
                "Observe",
                &[
                    api::ValueType::Double,
                    api::ValueType::String,
                    api::ValueType::Variant,
                    api::ValueType::HostObject,
                ],
                api::ValueType::Double,
            )
            .unwrap();
        let mut context = api::TypedContext {
            remaining_blocks: 10_000,
        };
        let result = session
            .call_values_by_ref(
                &observe,
                &[
                    api::ValueArg::ByRef(&total),
                    api::ValueArg::ByRef(&text),
                    api::ValueArg::ByRef(&variant),
                    api::ValueArg::Value(api::ValueSlot::Object(probe.argument())),
                ],
                &mut context,
            )
            .unwrap();
        assert!(
            matches!(result, api::Value::Double(r) if r == 7.0),
            "{backend:?}"
        );
        assert_eq!(
            *seen.borrow(),
            "42|seen|1:two|true|true|true",
            "{backend:?}"
        );
        assert!(matches!(total.get().unwrap(), api::Value::Double(t) if t == 43.0));
        assert!(!total.is_lent());
        assert!(total.set(api::Value::Double(0.0)).is_ok());

        // F x, x: both parameters are one storage, as in VBA (copy-back gives 10 or 1).
        let twice = session
            .bind_values(
                "Api",
                "Twice",
                &[api::ValueType::Long, api::ValueType::Long],
                api::ValueType::Void,
            )
            .unwrap();
        let shared = api::ByRefCell::new(api::Value::Long(0));
        session
            .call_values_by_ref(
                &twice,
                &[api::ValueArg::ByRef(&shared), api::ValueArg::ByRef(&shared)],
                &mut context,
            )
            .unwrap();
        assert!(
            matches!(shared.get().unwrap(), api::Value::Long(11)),
            "{backend:?}"
        );

        let grow = session
            .bind_values(
                "Api",
                "Grow",
                &[api::ValueType::Array, api::ValueType::Long],
                api::ValueType::Void,
            )
            .unwrap();
        let items = api::ByRefCell::new(api::Value::Array(Some(api::SafeArray::from_variants(
            vec![api::Variant::from_i32(4)],
        ))));
        session
            .call_values_by_ref(
                &grow,
                &[
                    api::ValueArg::ByRef(&items),
                    api::ValueArg::Value(api::ValueSlot::Long(2)),
                ],
                &mut context,
            )
            .unwrap();
        let api::Value::Array(Some(array)) = items.into_inner() else {
            panic!("array cell")
        };
        let elements = array.variant_elements().unwrap();
        assert_eq!(elements.len(), 3, "{backend:?}");
        assert_eq!(elements[0].as_i32(), Some(4));
        assert_eq!(elements[2].as_bstr().unwrap().as_str(), "last");

        // A ByVal parameter given a cell takes a copy.
        let kept = api::ByRefCell::new(api::Value::Long(3));
        session
            .call_values_by_ref(&copy, &[api::ValueArg::ByRef(&kept)], &mut context)
            .unwrap();
        assert!(matches!(kept.get().unwrap(), api::Value::Long(3)));
    }
}
