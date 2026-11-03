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
End Sub
Public Function ReadHits() As Long
    ReadHits = hits
End Function
"#;
#[test]
fn synchronous_primitive_copyback_preserves_errors_and_byval_isolation() {
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
        let mut args = [api::Value::Variant(calcvba_runtime::Variant::empty())];
        assert!(matches!(
            session.call_values_with_writebacks(&structural, &mut args, &mut context),
            Err(api::Error::UnsupportedSignature(_))
        ));
        let read = session
            .bind_values("Api", "ReadHits", &[], api::ValueType::Long)
            .unwrap();
        assert!(matches!(
            session.call_values(&read, &[], &mut context).unwrap(),
            api::Value::Long(0)
        ));
    }
}
