//! dc-8f3x: calls qualified with the VBA library name (`VBA.Join`, `VBA.CDate`) link in the
//! typed source-project route (the route DNA Calc's `load_project` uses).
use calcvba_host::provider::Provider;
use calcvba_host_api::{self as api, RuntimeProvider};
#[path = "support/typed_host.rs"]
#[allow(dead_code)]
mod support;

const SOURCE: &str = "Public Function JoinedLength() As Double
    JoinedLength = Len(VBA.Join(Array(\"ab\", \"cd\"), \"-\"))
End Function
Public Function QualifiedDate() As Double
    QualifiedDate = CDbl(VBA.CDate(\"2020-01-02\"))
End Function
Public Function PlainDate() As Double
    PlainDate = CDbl(CDate(\"2020-01-02\"))
End Function
";

fn call(session: &dyn api::Session, name: &str) -> f64 {
    let bound = session
        .bind_slots("Api", name, &[], api::ScalarType::Double)
        .unwrap_or_else(|e| panic!("{name}: {e:?}"));
    let frame = api::InvocationFrame::new(1, None, std::rc::Rc::new(0u64));
    let mut invocation = session
        .begin_invocation(&bound, &[], frame, 10_000)
        .unwrap();
    for _ in 0..100_000 {
        match invocation.resume(1).unwrap() {
            api::InvocationProgress::Suspended => {}
            api::InvocationProgress::Returned(api::OwnedValue::Double(v)) => return v,
            api::InvocationProgress::Returned(other) => panic!("{name}: {other:?}"),
        }
    }
    panic!("{name}: did not return")
}

#[test]
fn library_qualified_join_and_cdate_link() {
    for backend in support::backends() {
        let session = Provider::default()
            .load(
                &api::Project::module("Api", SOURCE),
                backend,
                api::HOST_API_ABI,
            )
            .expect("project loads");
        assert_eq!(call(&*session, "JoinedLength"), 5.0, "{backend:?}");
        let qualified = call(&*session, "QualifiedDate");
        assert_eq!(qualified, call(&*session, "PlainDate"), "{backend:?}");
        assert_eq!(qualified, 43832.0, "{backend:?}");
    }
}
