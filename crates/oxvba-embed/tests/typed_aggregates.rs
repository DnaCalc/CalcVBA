use calcvba_host_api::{self as api, RuntimeProvider};
use oxvba_embed::provider::Provider;
use std::rc::Rc;

fn backends() -> Vec<api::Backend> {
    let mut values = vec![api::Backend::Interpreter];
    if cfg!(feature = "jit") {
        values.push(api::Backend::Jit);
    }
    values
}
fn run(
    session: &dyn api::Session,
    name: &str,
    args: &[api::ValueSlot<'_>],
    result: api::ValueType,
) -> api::Value {
    let types = args.iter().map(|a| a.ty()).collect::<Vec<_>>();
    let bound = session
        .bind_values("Api", name, &types, result)
        .unwrap_or_else(|e| panic!("bind {name}: {e}"));
    let mut call = session
        .begin_values(
            &bound,
            args,
            api::InvocationFrame::new(11, None, Rc::new(())),
            10000,
        )
        .unwrap_or_else(|e| panic!("begin {name}: {e}"));
    for _ in 0..10000 {
        match call
            .resume(1)
            .unwrap_or_else(|e| panic!("resume {name}: {e}"))
        {
            api::ValueProgress::Suspended => {}
            api::ValueProgress::Returned(value) => return value,
        }
    }
    panic!("invocation did not finish")
}
#[test]
fn records_and_paramarray_match_owned_vba_with_one_operation_slices() {
    let source = include_str!(
        "../../../docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-aggregates-run11.bas"
    );
    for backend in backends() {
        let session = Provider::default()
            .load(
                &api::Project::module("Api", source),
                backend,
                api::HOST_API_ABI,
            )
            .unwrap();
        let value = run(session.as_ref(), "RunProbe", &[], api::ValueType::String);
        assert!(
            matches!(&value,api::Value::String(s) if s.as_str()=="record=42:owned:7|copy=41|aliases=44|param=9:9:0|empty=0:-1"),
            "{backend:?}: {value:?}"
        );
        assert!(
            matches!(run(session.as_ref(),"RecordReturns",&[],api::ValueType::String),api::Value::String(s) if s.as_str()=="42:42:owned:7")
        );
        let record = run(session.as_ref(), "MakeSample", &[], api::ValueType::Record);
        let api::Value::Record(record) = record else {
            panic!("record expected")
        };
        let result = run(
            session.as_ref(),
            "AdvanceSample",
            &[api::ValueSlot::Record(&record)],
            api::ValueType::Record,
        );
        let api::Value::Record(result) = result else {
            panic!("record result expected")
        };
        assert_eq!(record.read_field_variant(0).unwrap().as_i32(), Some(41));
        assert_eq!(result.read_field_variant(0).unwrap().as_i32(), Some(42));
        let before = run(session.as_ref(), "EntryCount", &[], api::ValueType::Long);
        let bad_layout = oxvba_runtime::vba_record::VbaRecordLayout::new(vec![
            oxvba_runtime::vba_record::VbaRecordFieldSpec::anonymous(
                oxvba_runtime::vba_record::VbaRecordFieldKind::Long,
            ),
        ])
        .unwrap();
        let bad_record = api::VbaRecord::new_default(std::sync::Arc::new(bad_layout)).unwrap();
        let bound = session
            .bind_values(
                "Api",
                "AdvanceSample",
                &[api::ValueType::Record],
                api::ValueType::Record,
            )
            .unwrap();
        assert!(
            session
                .begin_values(
                    &bound,
                    &[api::ValueSlot::Record(&bad_record)],
                    api::InvocationFrame::new(12, None, Rc::new(())),
                    1000
                )
                .is_err()
        );
        assert_eq!(
            format!("{before:?}"),
            format!(
                "{:?}",
                run(session.as_ref(), "EntryCount", &[], api::ValueType::Long)
            )
        );
        let array = api::SafeArray::from_variants(vec![
            api::Variant::from_i32(2),
            api::Variant::from_i32(3),
        ]);
        let result = run(
            session.as_ref(),
            "AdvanceArray",
            &[api::ValueSlot::Array(&array)],
            api::ValueType::Variant,
        );
        assert_eq!(array.variant_elements().unwrap()[0].as_i32(), Some(2));
        assert!(
            matches!(result,api::Value::Variant(v) if v.safearray_element(0).unwrap().unwrap().as_i32()==Some(9))
        );
        let before = run(session.as_ref(), "EntryCount", &[], api::ValueType::Long);
        let wrong = api::SafeArray::from_shape_typed(
            vec![oxvba_runtime::safe_array::SafeArrayBound { lower: 0, count: 2 }],
            3,
        )
        .unwrap();
        let bound = session
            .bind_values(
                "Api",
                "AdvanceArray",
                &[api::ValueType::Array],
                api::ValueType::Variant,
            )
            .unwrap();
        assert!(
            session
                .begin_values(
                    &bound,
                    &[api::ValueSlot::Array(&wrong)],
                    api::InvocationFrame::new(12, None, Rc::new(())),
                    1000
                )
                .is_err()
        );
        assert_eq!(
            format!("{before:?}"),
            format!(
                "{:?}",
                run(session.as_ref(), "EntryCount", &[], api::ValueType::Long)
            )
        );
        let inputs = [api::ValueSlot::Long(1); 20];
        assert!(matches!(
            run(
                session.as_ref(),
                "SumFrom",
                &[
                    api::ValueSlot::Long(10),
                    api::ValueSlot::Long(2),
                    api::ValueSlot::Double(3.0)
                ],
                api::ValueType::Double
            ),
            api::Value::Double(15.0)
        ));
        assert!(matches!(
            run(session.as_ref(), "SumTail", &inputs, api::ValueType::Double),
            api::Value::Double(20.0)
        ));
        assert!(matches!(
            run(session.as_ref(), "SumTail", &[], api::ValueType::Double),
            api::Value::Double(0.0)
        ));
        assert!(
            matches!(run(session.as_ref(),"BoundsTail",&[],api::ValueType::String),api::Value::String(s) if s.as_str()=="0:-1")
        );
    }
}
#[test]
fn paramarray_private_repeated_object_aliases_and_detached_copy_match_vba() {
    let source = include_str!(
        "../../../docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-paramarray-edges-run11.bas"
    );
    for backend in backends() {
        let session = Provider::default()
            .load(
                &api::Project::module("Api", source),
                backend,
                api::HOST_API_ABI,
            )
            .unwrap();
        let payload = Rc::new(());
        let object = api::HostObject::new_with_payload([11; 16], vec![], payload.clone());
        let result = run(
            session.as_ref(),
            "Observe",
            &[api::ValueSlot::Object(object.argument())],
            api::ValueType::String,
        );
        assert!(
            matches!(&result,api::Value::String(s) if s.as_str()=="private=5|repeat=9:9|copy=9:27:9|bad=13:9:9|nested=0:0:True|object=True"),
            "{backend:?}: {result:?}"
        );
        assert_eq!(Rc::strong_count(&payload), 2);
        drop(object);
        assert_eq!(Rc::strong_count(&payload), 1);
    }
}
