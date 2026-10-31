//! Report-6 idle Excel observations establish these value-context conversions.
//! This host fixture deliberately supplies a Range-like default; the engine
//! never recognizes a spreadsheet object or flattens its argument at entry.
use oxvba_com::{
    PortableComProjection,
    platform::portable::{PortableDispatch, PortableObjectFactory},
};
use oxvba_embed::{Backend, Error, Module, ProcedureKind};
use oxvba_runtime::{
    Variant,
    safe_array::{SafeArray, SafeArrayBound},
};
use std::sync::{
    Arc,
    atomic::{AtomicUsize, Ordering},
};

#[derive(Clone, Copy)]
enum DefaultValue {
    Number,
    Text,
    Blank,
    BadText,
    Array,
    Fault,
}
struct RangeFactory {
    value: DefaultValue,
    reads: Arc<AtomicUsize>,
}
impl PortableObjectFactory for RangeFactory {
    fn create(&self) -> Box<dyn PortableDispatch> {
        Box::new(Self {
            value: self.value,
            reads: self.reads.clone(),
        })
    }
}
impl PortableDispatch for RangeFactory {
    fn invoke(&self, _: &str, _: &[Variant]) -> Result<Variant, String> {
        Err("method path must not be used".into())
    }
    fn get(&self, member: &str) -> Result<Variant, String> {
        assert_eq!(member, "Item");
        self.reads.fetch_add(1, Ordering::SeqCst);
        Ok(match self.value {
            DefaultValue::Number => Variant::from_f64(41.0),
            DefaultValue::Text => Variant::from_string("3"),
            DefaultValue::Blank => Variant::empty(),
            DefaultValue::BadText => Variant::from_string("bad"),
            DefaultValue::Array => Variant::from_safearray(
                SafeArray::from_shape_and_variants(
                    vec![SafeArrayBound { lower: 1, count: 2 }],
                    vec![Variant::from_f64(41.0), Variant::from_f64(3.0)],
                )
                .unwrap(),
            ),
            DefaultValue::Fault => return Err("owned host default failed".into()),
        })
    }
    fn put(&self, _: &str, _: Variant) -> Result<(), String> {
        Err("read-only fixture".into())
    }
    fn member_names(&self) -> Vec<String> {
        vec!["Item".into()]
    }
}
const SOURCE: &str = r#"
Public Function Plus(x) As Double
    Dim saved As Object
    Set saved = x
    If TypeName(x) <> "Range" Then Err.Raise 13
    Plus = x + 1
    If Not IsObject(x) Then Err.Raise 13
    If Not (x Is saved) Then Err.Raise 13
    If TypeName(x) <> "Range" Then Err.Raise 13
End Function
Public Function Times(x) As Double
    Times = x * 2
    If Not IsObject(x) Then Err.Raise 13
End Function
Public Function Divide(x) As Double
    Divide = x / 2
End Function
Public Function Negative(x) As Double
    Negative = -x
End Function
Public Function CapturedError(x) As Long
    Dim result As Double
    On Error Resume Next
    result = x + 1
    CapturedError = Err.Number
End Function
"#;

#[test]
fn untyped_range_keeps_object_identity_and_uses_its_default_in_arithmetic() {
    let mut backends = vec![Backend::Interpreter];
    if cfg!(feature = "jit") {
        backends.push(Backend::Jit);
    }
    for backend in backends {
        for (kind, expected) in [
            (DefaultValue::Number, Some(42.0)),
            (DefaultValue::Text, Some(4.0)),
            (DefaultValue::Blank, Some(1.0)),
            (DefaultValue::BadText, None),
            (DefaultValue::Array, None),
        ] {
            let reads = Arc::new(AtomicUsize::new(0));
            let registry = Arc::new(PortableComProjection::new());
            registry.register_object(
                "Fixture.Range",
                Arc::new(RangeFactory {
                    value: kind,
                    reads: reads.clone(),
                }),
            );
            let host = oxvba_hal::adapters::builder::HostBuilder::new()
                .profile(oxvba_hal::model::HalProfileId::Windows)
                .portable_objects(registry)
                .build();
            let object = host
                .com()
                .create_object_variant(Variant::from_string("Fixture.Range"))
                .unwrap();
            let identity = object.as_object_ref().unwrap();
            let session = Module::load("Api", SOURCE)
                .unwrap()
                .with_host(host)
                .into_session(backend)
                .unwrap();
            let plus = session
                .bind("Api", "Plus", ProcedureKind::Function)
                .unwrap();
            let result = session.call_bound(&plus, std::slice::from_ref(&object), 10000);
            if let Some(expected) = expected {
                assert_eq!(result.unwrap().as_f64(), Some(expected), "{backend:?}");
            } else {
                assert!(
                    matches!(result, Err(Error::Runtime { number: 13, .. })),
                    "{backend:?}: {result:?}"
                );
            }
            assert_eq!(reads.load(Ordering::SeqCst), 1);
            assert_eq!(object.as_object_ref().unwrap(), identity);
            if matches!(kind, DefaultValue::Number) {
                for (name, expected) in [("Times", 82.0), ("Divide", 20.5), ("Negative", -41.0)] {
                    let handle = session.bind("Api", name, ProcedureKind::Function).unwrap();
                    assert_eq!(
                        session
                            .call_bound(&handle, std::slice::from_ref(&object), 10000)
                            .unwrap()
                            .as_f64(),
                        Some(expected)
                    );
                }
            }
            let handler = session
                .bind("Api", "CapturedError", ProcedureKind::Function)
                .unwrap();
            let code = session
                .call_bound(&handler, &[object], 10000)
                .unwrap()
                .as_i32()
                .unwrap();
            assert_eq!(code, if expected.is_some() { 0 } else { 13 });
        }
    }
}

#[test]
fn default_member_fault_and_nothing_do_not_turn_into_scalar_inputs() {
    for backend in [Backend::Interpreter, Backend::Jit] {
        if backend == Backend::Jit && !cfg!(feature = "jit") {
            continue;
        }
        let registry = Arc::new(PortableComProjection::new());
        registry.register_object(
            "Fixture.Range",
            Arc::new(RangeFactory {
                value: DefaultValue::Fault,
                reads: Arc::new(AtomicUsize::new(0)),
            }),
        );
        let host = oxvba_hal::adapters::builder::HostBuilder::new()
            .profile(oxvba_hal::model::HalProfileId::Windows)
            .portable_objects(registry)
            .build();
        let object = host
            .com()
            .create_object_variant(Variant::from_string("Fixture.Range"))
            .unwrap();
        let session = Module::load(
            "Api",
            "Public Function Plus(x) As Double\nPlus=x+1\nEnd Function\n",
        )
        .unwrap()
        .with_host(host)
        .into_session(backend)
        .unwrap();
        let plus = session
            .bind("Api", "Plus", ProcedureKind::Function)
            .unwrap();
        assert!(matches!(
            session.call_bound(&plus, &[object], 10000),
            Err(Error::Runtime { .. })
        ));
        assert!(matches!(
            session.call_bound(&plus, &[Variant::nothing()], 10000),
            Err(Error::Runtime { number: 91, .. })
        ));
    }
}
