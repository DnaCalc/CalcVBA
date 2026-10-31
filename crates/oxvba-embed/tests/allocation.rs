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
