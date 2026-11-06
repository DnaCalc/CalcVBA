use calcvba_com::{
    DynamicCallArg, TEST_DISPID_COUNT, TypeLibMemberInvokeKind, TypeLibMemberMetadata,
    known_typelib_identity_for_prog_id_name,
    platform::portable::{
        PortableComProjection, PortableDispatch, PortableObjectFactory, PortableTypedCallResult,
    },
};
use calcvba_engine::{Engine, HostConfig, RuntimeProfileId};
use calcvba_runtime::{RuntimeCallResult, Variant};
use calcvba_symbol::manifest::ProjectReference;
use std::sync::{
    Arc,
    atomic::{AtomicUsize, Ordering},
};

struct Counter(Arc<[AtomicUsize; 2]>);
impl PortableDispatch for Counter {
    fn invoke_early_bound(
        &self,
        member: &TypeLibMemberMetadata,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        if member.token != TEST_DISPID_COUNT
            || member.invoke_kind != TypeLibMemberInvokeKind::PropertyGet
            || !args.is_empty()
        {
            return Err("unsupported typed fixture member".into());
        }
        self.0[1].fetch_add(1, Ordering::SeqCst);
        Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
            Variant::from_i32(7),
        )))
    }
    fn get(&self, member: &str) -> Result<Variant, String> {
        if member != "Count" {
            return Err("unknown dynamic member".into());
        }
        self.0[0].fetch_add(1, Ordering::SeqCst);
        Ok(Variant::from_i32(7))
    }
    fn invoke(&self, member: &str, args: &[Variant]) -> Result<Variant, String> {
        if member.eq_ignore_ascii_case("Echo") && args.len() == 1 {
            Ok(args[0].clone())
        } else {
            Err("no methods".into())
        }
    }
    fn put(&self, _: &str, _: Variant) -> Result<(), String> {
        Err("read-only".into())
    }
    fn member_names(&self) -> Vec<String> {
        vec!["Count".into()]
    }
}

// Deliberately implements only the dynamic methods, leaving the typed default.
struct LateOnly(Counter);
impl PortableDispatch for LateOnly {
    fn get(&self, member: &str) -> Result<Variant, String> {
        self.0.get(member)
    }
    fn invoke(&self, member: &str, args: &[Variant]) -> Result<Variant, String> {
        self.0.invoke(member, args)
    }
    fn put(&self, member: &str, value: Variant) -> Result<(), String> {
        self.0.put(member, value)
    }
    fn member_names(&self) -> Vec<String> {
        self.0.member_names()
    }
}

struct Factory {
    calls: Arc<[AtomicUsize; 2]>,
    typed: bool,
}
impl PortableObjectFactory for Factory {
    fn create(&self) -> Box<dyn PortableDispatch> {
        let counter = Counter(self.calls.clone());
        if self.typed {
            Box::new(counter)
        } else {
            Box::new(LateOnly(counter))
        }
    }
}

fn engine(jit: bool, typed: bool) -> (Engine, Arc<[AtomicUsize; 2]>) {
    let calls = Arc::new([AtomicUsize::new(0), AtomicUsize::new(0)]);
    let projection = Arc::new(PortableComProjection::new());
    projection.register_object(
        "OxVba.TestDispatch",
        Arc::new(Factory {
            calls: calls.clone(),
            typed,
        }),
    );
    (
        Engine::new(if jit {
            HostConfig::jit()
        } else {
            HostConfig::vm3()
        })
        .with_runtime_profile(RuntimeProfileId::WindowsHeadless)
        .with_portable_com_projection(projection),
        calls,
    )
}

fn test_dispatch_ref() -> ProjectReference {
    let identity =
        known_typelib_identity_for_prog_id_name("OxVba.TestDispatch").expect("fixture typelib");

    ProjectReference::TypeLibrary {
        name: identity.reference_name,
        guid: identity.libid,
        version_major: Some(identity.major_version),
        version_minor: Some(identity.minor_version),
        lcid: identity.lcid,
        import_lib: Some(identity.importlib),
    }
}

fn long_at(values: &[Variant], index: usize) -> i32 {
    values
        .get(index)
        .and_then(Variant::as_i32)
        .unwrap_or_else(|| panic!("expected Long at snapshot slot {index}: {values:?}"))
}

#[test]
fn portable_projection_callbyname_lowers_through_typelib_metadata() {
    let source = r#"
Public result As Long
Public lateValue As Long
Public earlyValue As Long

Sub Main()
    Dim obj As OxVba.TestDispatch
    Set obj = New OxVba.TestDispatch

    lateValue = CallByName(obj, "Count", VbGet)
    earlyValue = obj.Count
    result = lateValue - earlyValue
End Sub
"#;

    for jit in [false, true] {
        let (engine, calls) = engine(jit, true);
        let values = engine
            .execute_source_with_references_and_snapshot(source, vec![test_dispatch_ref()])
            .expect("portable COM projection should support CallByName by typelib metadata");

        let result = long_at(&values, 0);
        let late_value = long_at(&values, 1);
        let early_value = long_at(&values, 2);

        assert_eq!(
            result, 0,
            "CallByName and direct property dispatch should agree"
        );
        assert!(late_value > 0, "projection fixture should return a value");
        assert_eq!(late_value, early_value);
        assert_eq!(early_value, 7);
        assert_eq!(
            calls[0].load(Ordering::SeqCst),
            1,
            "CallByName alone uses names"
        );
        assert_eq!(
            calls[1].load(Ordering::SeqCst),
            1,
            "typed call uses the explicit adapter"
        );
    }
}

#[test]
fn portable_typed_call_never_retries_a_late_only_object() {
    let source = "Sub Main()\nDim obj As OxVba.TestDispatch\nSet obj = New OxVba.TestDispatch\nDim n As Long\nn = obj.Count\nEnd Sub\n";
    for jit in [false, true] {
        let (engine, calls) = engine(jit, false);
        assert!(
            engine
                .execute_source_with_references_and_snapshot(source, vec![test_dispatch_ref()])
                .is_err()
        );
        assert_eq!(calls[0].load(Ordering::SeqCst), 0, "no dynamic retry");
        assert_eq!(calls[1].load(Ordering::SeqCst), 0);
    }
}

struct ByRefReceiver;
impl PortableDispatch for ByRefReceiver {
    fn invoke_dynamic(
        &self,
        member: &str,
        _: Option<calcvba_com::DynamicCallKind>,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        if member.eq_ignore_ascii_case("Change") {
            assert_eq!(args.len(), 3);
            assert_eq!(args[0].by_ref, args[1].by_ref, "duplicate caller aliases");
            let mut result = RuntimeCallResult::value(Variant::from_i32(321));
            result.add_writeback(calcvba_runtime::RuntimeByRefWriteback::new(
                args[0].by_ref.unwrap(),
                Variant::from_i32(11),
            ));
            result.add_writeback(calcvba_runtime::RuntimeByRefWriteback::new(
                args[2].by_ref.unwrap(),
                Variant::from_bool(true),
            ));
            return Ok(PortableTypedCallResult::Value(result));
        }
        if member.eq_ignore_ascii_case("Observe") {
            assert_eq!(args[0].name.as_deref(), Some("flag"));
            assert_eq!(args[1].name.as_deref(), Some("value"));
            assert_eq!(
                args[0].value.as_ref().unwrap().variant().as_bool(),
                Some(true)
            );
            assert_eq!(args[1].value.as_ref().unwrap().variant().as_i32(), Some(11));
            return Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                Variant::from_i32(22),
            )));
        }
        Err("unknown method".into())
    }
    fn invoke(&self, _: &str, _: &[Variant]) -> Result<Variant, String> {
        Err("dynamic arguments required".into())
    }
    fn get(&self, _: &str) -> Result<Variant, String> {
        Err("no properties".into())
    }
    fn put(&self, _: &str, _: Variant) -> Result<(), String> {
        Err("no properties".into())
    }
    fn member_names(&self) -> Vec<String> {
        vec!["Change".into(), "Observe".into()]
    }
}
struct ByRefFactory;
impl PortableObjectFactory for ByRefFactory {
    fn create(&self) -> Box<dyn PortableDispatch> {
        Box::new(ByRefReceiver)
    }
}

#[test]
fn portable_dynamic_calls_preserve_aliases_names_and_boolean_writebacks() {
    let source = r#"
Public observed As String
Sub Main()
    Dim obj As Object, number As Variant, cancel As Boolean, n As Long
    Set obj=CreateObject("CalcForms.ByRef")
    number=3
    n=obj.Change(number,number,cancel)
    observed=CStr(number) & ":" & CStr(cancel) & ":" & CStr(n)
    n=obj.Observe(flag:=cancel,value:=number)
    observed=observed & ":" & CStr(n)
End Sub
"#;
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let projection = Arc::new(PortableComProjection::new());
        projection.register_object("CalcForms.ByRef", Arc::new(ByRefFactory));
        let engine = Engine::new(config)
            .with_runtime_profile(RuntimeProfileId::WindowsHeadless)
            .with_portable_com_projection(projection);
        let values = engine
            .execute_source_with_variant_snapshot_clean(source)
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(Variant::as_bstr)
                .any(|value| value.to_string() == "11:True:321:22"),
            "{values:?}"
        );
    }
}

#[test]
fn legacy_portable_methods_accept_input_only_byref_arguments() {
    let source = "Public observed As Long\nSub Main()\nDim obj As Object, number As Long\nSet obj=CreateObject(\"OxVba.TestDispatch\")\nnumber=7\nobserved=obj.Echo(number)+number\nEnd Sub\n";
    for jit in [false, true] {
        let (engine, _) = engine(jit, false);
        let values = engine
            .execute_source_with_variant_snapshot_clean(source)
            .unwrap();
        assert_eq!(long_at(&values, 0), 14);
    }
}
