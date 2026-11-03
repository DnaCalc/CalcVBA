use calcvba_host::{Backend, Module, ProcedureKind};
use calcvba_host_api::{
    self as api,
    object::{DynamicMemberCall, DynamicMemberKind as K},
};
use std::{cell::Cell, rc::Rc};
struct Services(Rc<Cell<usize>>);
impl api::HostObjectServices for Services {
    fn invoke_dynamic(
        &self,
        call: &DynamicMemberCall,
        args: &mut [api::Variant],
        _: &mut api::TypedContext,
    ) -> Result<api::Variant, api::TypedFault> {
        self.0.set(self.0.get() + 1);
        match call.name.as_deref() {
            Some("Bump") => {
                assert!(matches!(call.kind, K::Method | K::Get));
                assert_eq!(args.len(), 2);
                assert!(call.arguments.iter().all(|a| a.by_ref));
                assert_eq!(call.arguments[1].alias, 0);
                assert_eq!(args[0].as_i32(), Some(2));
                args[0] = api::Variant::from_i32(17);
                Ok(api::Variant::from_i32(34))
            }
            Some("Named") => {
                assert_eq!(call.arguments[0].name.as_deref(), Some("second"));
                assert_eq!(call.arguments[1].name.as_deref(), Some("first"));
                Ok(api::Variant::from_i32(
                    args[0].as_i32().unwrap() * 10 + args[1].as_i32().unwrap(),
                ))
            }
            Some("Omitted") => {
                assert_eq!(args[1].as_error_code(), Some(0x80020004u32 as i32));
                Ok(args[2].clone())
            }
            Some("Fail") => {
                args[0] = api::Variant::from_i32(23);
                Err(api::TypedFault {
                    number: 1004,
                    description: "fixture failure".into(),
                    source: Some("owned.dynamic".into()),
                    help_file: None,
                    help_context: None,
                })
            }
            None => Ok(api::Variant::from_f64(41.0)),
            _ => panic!("unexpected member {:?}", call.name),
        }
    }
}
#[test]
fn dynamic_arguments_keep_actual_aliases_names_omissions_and_error_copyback() {
    let module = Module::load(
        "Api",
        r#"
Public Function Run(ByVal obj As Object) As String
    Dim x As Long
    x = 2
    Run = obj.Bump(x, x) & ":" & x
    Run = Run & ":" & obj.Named(second:=3, first:=2)
    Run = Run & ":" & obj.Omitted(1, , 7)
    Run = Run & ":" & (obj + 1#)
    On Error Resume Next
    obj.Fail x
    Run = Run & ":" & x & ":" & Err.Number & ":" & Err.Description & ":" & Err.Source
End Function
"#,
    )
    .unwrap();
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let calls = Rc::new(Cell::new(0));
        let payload = Rc::new(());
        // SAFETY: payload retains the stable identity for the services' lifetime.
        let object = unsafe {
            api::HostObject::new_with_services(
                api::object::IDISPATCH,
                vec![],
                payload.clone(),
                Rc::as_ptr(&payload).cast(),
                Rc::new(Services(calls.clone())),
            )
        };
        let session = module.clone().into_session(backend).unwrap();
        let handle = session.bind("Api", "Run", ProcedureKind::Function).unwrap();
        let result = session
            .call_values(
                &handle,
                &[api::ValueSlot::Object(object.argument())],
                &mut api::TypedContext {
                    remaining_blocks: 1000,
                },
            )
            .unwrap();
        let api::Value::String(result) = result else {
            panic!("string")
        };
        assert_eq!(
            result.as_str(),
            "34:17:32:7:42:23:1004:fixture failure:owned.dynamic",
            "{backend:?}"
        );
        assert_eq!(calls.get(), 5);
    }
}
