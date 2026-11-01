use calcvba_host_api::{self as api, RuntimeProvider};
use oxvba_embed::provider::Provider;
use std::rc::Rc;
fn project() -> api::Project {
    let mut project = api::Project::module(
        "Api",
        include_str!(
            "../../../docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-classes-api-run11.bas"
        ),
    );
    project.modules.push(api::SourceModule { name:"Widget".into(),source:include_str!("../../../docs/evidence/programs/ideal-2026-07/windows-x64/WIN-4/typed-classes-widget-run11.cls").into() });
    project
}
fn run(
    session: &dyn api::Session,
    name: &str,
    args: &[api::ValueSlot<'_>],
    result: api::ValueType,
) -> api::Value {
    let types = args.iter().map(|a| a.ty()).collect::<Vec<_>>();
    let handle = session
        .bind_values("Api", name, &types, result)
        .unwrap_or_else(|e| panic!("bind {name}: {e}"));
    let mut call = session
        .begin_values(
            &handle,
            args,
            api::InvocationFrame::new(19, None, Rc::new(())),
            10000,
        )
        .unwrap_or_else(|e| panic!("begin {name}: {e}"));
    for _ in 0..10000 {
        match call
            .resume(1)
            .unwrap_or_else(|e| panic!("resume {name}: {e}"))
        {
            api::ValueProgress::Suspended => {}
            api::ValueProgress::Returned(v) => return v,
        }
    }
    panic!("did not finish")
}
#[test]
fn source_class_lifecycle_uses_original_object_in_both_typed_backends() {
    let mut backends = vec![api::Backend::Interpreter];
    if cfg!(feature = "jit") {
        backends.push(api::Backend::Jit);
    }
    for backend in backends {
        let provider = Provider::default().with_class_modules(vec!["Widget".into()]);
        let session = provider
            .load(&project(), backend, api::HOST_API_ABI)
            .unwrap();
        let result = run(session.as_ref(), "RunProbe", &[], api::ValueType::String);
        assert!(
            matches!(&result,api::Value::String(s) if s.as_str()=="same=True|start=40|next=42|distinct=False|journal=IIT42T40"),
            "{backend:?}: {result:?}"
        );
        let value = run(
            session.as_ref(),
            "MakeWidget",
            &[],
            api::ValueType::HostObject,
        );
        let api::Value::RuntimeObject(object) = value else {
            panic!("runtime object expected: {value:?}")
        };
        assert!(matches!(
            run(
                session.as_ref(),
                "ReadWidget",
                &[api::ValueSlot::RuntimeObject(&object)],
                api::ValueType::Long
            ),
            api::Value::Long(40)
        ));
        drop(object);
        let journal = run(session.as_ref(), "ReadJournal", &[], api::ValueType::String);
        assert!(
            matches!(&journal,api::Value::String(s) if s.as_str()=="IIT42T40IT40"),
            "{journal:?}"
        );
    }
}

#[test]
fn source_class_values_do_not_route_into_other_sessions_or_reset_generations() {
    let mut backends = vec![api::Backend::Interpreter];
    if cfg!(feature = "jit") {
        backends.push(api::Backend::Jit);
    }
    for backend in backends {
        let provider = Provider::default().with_class_modules(vec!["Widget".into()]);
        let first = provider
            .load(&project(), backend, api::HOST_API_ABI)
            .unwrap();
        let second = provider
            .load(&project(), backend, api::HOST_API_ABI)
            .unwrap();
        let api::Value::RuntimeObject(object) = run(
            first.as_ref(),
            "MakeWidget",
            &[],
            api::ValueType::HostObject,
        ) else {
            panic!("object")
        };
        let bound = second
            .bind_values(
                "Api",
                "ReadWidget",
                &[api::ValueType::HostObject],
                api::ValueType::Long,
            )
            .unwrap();
        assert!(
            second
                .begin_values(
                    &bound,
                    &[api::ValueSlot::RuntimeObject(&object)],
                    api::InvocationFrame::new(1, None, Rc::new(())),
                    1000
                )
                .is_err()
        );
        assert!(
            matches!(run(second.as_ref(),"ReadJournal",&[],api::ValueType::String),api::Value::String(s) if s.as_str().is_empty())
        );
        second.reset().unwrap();
        assert!(matches!(
            run(
                first.as_ref(),
                "ReadWidget",
                &[api::ValueSlot::RuntimeObject(&object)],
                api::ValueType::Long
            ),
            api::Value::Long(40)
        ));
        first.reset().unwrap();
        let bound = first
            .bind_values(
                "Api",
                "ReadWidget",
                &[api::ValueType::HostObject],
                api::ValueType::Long,
            )
            .unwrap();
        assert!(
            first
                .begin_values(
                    &bound,
                    &[api::ValueSlot::RuntimeObject(&object)],
                    api::InvocationFrame::new(1, None, Rc::new(())),
                    1000
                )
                .is_err()
        );
        drop(object);
        assert!(
            matches!(run(first.as_ref(),"ReadJournal",&[],api::ValueType::String),api::Value::String(s) if s.as_str().is_empty())
        );
    }
}
#[test]
fn cancelling_constructor_or_termination_slices_leaves_session_usable() {
    let mut backends = vec![api::Backend::Interpreter];
    if cfg!(feature = "jit") {
        backends.push(api::Backend::Jit);
    }
    for backend in backends {
        for slice in [0, 3, 9, 17, 29, 47, 81, 130, 201] {
            let session = Provider::default()
                .with_class_modules(vec!["Widget".into()])
                .load(&project(), backend, api::HOST_API_ABI)
                .unwrap();
            let bound = session
                .bind_values("Api", "RunProbe", &[], api::ValueType::String)
                .unwrap();
            let mut call = session
                .begin_values(
                    &bound,
                    &[],
                    api::InvocationFrame::new(1, None, Rc::new(())),
                    10000,
                )
                .unwrap();
            for _ in 0..slice {
                if matches!(call.resume(1).unwrap(), api::ValueProgress::Returned(_)) {
                    break;
                }
            }
            drop(call);
            session.reset().unwrap();
            assert!(
                matches!(run(session.as_ref(),"RunProbe",&[],api::ValueType::String),api::Value::String(s) if s.as_str()=="same=True|start=40|next=42|distinct=False|journal=IIT42T40")
            );
        }
    }
}
