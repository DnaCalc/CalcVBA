use oxvba_embed::{Backend, Module};
use std::{
    alloc::{GlobalAlloc, Layout, System},
    cell::Cell,
};
struct CountAlloc;
thread_local! { static ALLOCATIONS: Cell<Option<usize>> = const { Cell::new(None) }; }
fn count() {
    let _ = ALLOCATIONS.try_with(|n| {
        if let Some(value) = n.get() {
            n.set(Some(value + 1));
        }
    });
}
// SAFETY: every allocation and deallocation delegates unchanged to System.
unsafe impl GlobalAlloc for CountAlloc {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        count();
        // SAFETY: forwarded caller allocation contract.
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, p: *mut u8, layout: Layout) {
        // SAFETY: exact original allocation/layout forwarded.
        unsafe { System.dealloc(p, layout) }
    }
    unsafe fn realloc(&self, p: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        count();
        // SAFETY: exact caller reallocation contract forwarded.
        unsafe { System.realloc(p, layout, size) }
    }
}
#[global_allocator]
static ALLOCATOR: CountAlloc = CountAlloc;

#[test]
fn declared_variant_host_result_uses_inline_storage_without_per_callback_heap_boxing() {
    use calcvba_host_api::{self as api, RuntimeProvider};
    use std::rc::Rc;
    for backend in typed_host::backends() {
        let iid = [44; 16];
        let library = api::TypedObjectLibrary {
            name: "Values".into(),
            interfaces: vec![api::TypedObjectInterface {
                name: "Object".into(),
                iid,
                members: vec![api::TypedObjectMember {
                    name: "Value2".into(),
                    slot: 0,
                    kind: api::TypedMemberKind::PropertyGet,
                    parameters: vec![],
                    result: api::ValueType::Variant,
                    is_default: false,
                }],
            }],
        };
        let project = api::Project::module(
            "Api",
            "Public Function Read(ByVal root As Values.Object) As Double\nRead = root.Value2\nEnd Function\n",
        );
        let session = oxvba_embed::provider::Provider::default()
            .load_with_object_libraries(&project, &[library], backend, api::HOST_API_ABI)
            .unwrap();
        let object = api::HostObject::new(
            iid,
            vec![
                api::HostSlot::typed(vec![], api::ValueType::Variant, |_, _| {
                    Ok(api::Value::Variant(api::Variant::from_f64(42.0)))
                })
                .unwrap(),
            ],
        );
        let bound = session
            .bind_values(
                "Api",
                "Read",
                &[api::ValueType::HostObject],
                api::ValueType::Double,
            )
            .unwrap();
        let mut call = session
            .begin_values(
                &bound,
                &[api::ValueSlot::Object(object.argument())],
                api::InvocationFrame::new(1, None, Rc::new(())),
                1000,
            )
            .unwrap();
        ALLOCATIONS.with(|n| n.set(Some(0)));
        let result = call.resume(u64::MAX).unwrap();
        let allocations = ALLOCATIONS.with(|n| n.replace(None).unwrap());
        assert!(matches!(
            result,
            api::ValueProgress::Returned(api::Value::Double(42.0))
        ));
        assert_eq!(
            allocations, 0,
            "{backend:?} declared Variant output must use its actual carrier, without a heap box"
        );
    }
}
#[path = "support/typed_host.rs"]
mod typed_host;

#[test]
fn successful_typed_calls_allocate_no_heap_frames_or_other_storage() {
    let mut backends = vec![Backend::Interpreter];
    if cfg!(feature = "jit") {
        backends.push(Backend::Jit);
    }
    for backend in backends {
        let session = Module::load("Math", "Public Function Add2(ByVal x As Double, ByVal y As Double) As Double\nAdd2 = x + y\nEnd Function\n").unwrap().into_session(backend).unwrap();
        let entry = session
            .bind_typed::<(f64, f64), f64>("Math", "Add2")
            .unwrap();
        ALLOCATIONS.with(|n| n.set(Some(0)));
        for _ in 0..1000 {
            assert_eq!(session.call_typed(&entry, (19.0, 23.0), 100).unwrap(), 42.0);
        }
        let allocations = ALLOCATIONS.with(|n| n.replace(None).unwrap());
        assert_eq!(
            allocations, 0,
            "{backend:?} typed calls must be allocation-free"
        );
    }
}

#[test]
fn api_slot_calls_and_reentrant_range_callbacks_allocate_no_heap_or_variants() {
    use calcvba_host_api::{self as api, RuntimeProvider, SessionExt};
    use std::rc::Rc;
    for backend in typed_host::backends() {
        let session: Rc<Box<dyn api::Session>> = Rc::new(
            oxvba_embed::provider::Provider::default()
                .load(&typed_host::project(), backend, api::HOST_API_ABI)
                .unwrap(),
        );
        let inner = session.bind_typed::<(f64,), f64>("Api", "Twice").unwrap();
        let weak = Rc::downgrade(&session);
        let object = api::HostObject::new(
            typed_host::IID,
            vec![api::HostSlot::double_long(move |context, offset| {
                weak.upgrade()
                    .unwrap()
                    .call_typed_in_context(&inner, (f64::from(offset),), context)
                    .map_err(|_| 5)
            })],
        );
        let handle = session
            .bind_typed::<(api::ObjectArg<'_>, i32), f64>("Api", "ReadRange")
            .unwrap();
        assert_eq!(
            session
                .call_typed(&handle, (object.argument(), 20), 100)
                .unwrap(),
            41.0
        );
        let before = oxvba_runtime::current_thread_live_handle_counts();
        let variants = oxvba_runtime::live_counters::current_thread_variant_constructions();
        ALLOCATIONS.with(|n| n.set(Some(0)));
        for _ in 0..1000 {
            assert_eq!(
                session
                    .call_typed(&handle, (object.argument(), 20), 100)
                    .unwrap(),
                41.0
            );
        }
        let allocations = ALLOCATIONS.with(|n| n.replace(None).unwrap());
        assert_eq!(
            allocations, 0,
            "{backend:?} API/host/reentry path must remain unboxed"
        );
        assert_eq!(
            oxvba_runtime::live_counters::current_thread_variant_constructions(),
            variants,
            "{backend:?} typed host callback/reentry must not construct owning Variants"
        );
        assert!(
            before
                .balance_to(oxvba_runtime::current_thread_live_handle_counts())
                .is_zero()
        );
    }
}
