use super::*;
use crate::{TypedContext, with_native_call, with_native_callback_context};
use std::cell::RefCell;

struct State {
    caller: Cell<i32>,
    native: Cell<i32>,
    before_error: Cell<i32>,
    after_error: Cell<i32>,
    events: RefCell<Vec<&'static str>>,
}
impl State {
    fn new() -> Self {
        Self {
            caller: Cell::new(2),
            native: Cell::new(12),
            before_error: Cell::new(0),
            after_error: Cell::new(0),
            events: RefCell::new(vec![]),
        }
    }
    fn data(&self) -> *mut () {
        std::ptr::from_ref(self).cast_mut().cast()
    }
}
// SAFETY CONTRACT: each test registers one live same-thread State whose Cell and
// RefCell fields are accessed only by these sequential, nonpanicking hooks.
unsafe fn publish(data: *mut ()) -> Result<(), i32> {
    // SAFETY: every registration below lends this exact stable State.
    let state = unsafe { &*data.cast::<State>() };
    state.events.borrow_mut().push("publish");
    if state.before_error.get() != 0 {
        return Err(state.before_error.get());
    }
    state.caller.set(state.native.get());
    Ok(())
}
unsafe fn refresh(data: *mut ()) -> Result<(), i32> {
    // SAFETY: same registered State, after callback body has completed.
    let state = unsafe { &*data.cast::<State>() };
    state.events.borrow_mut().push("refresh");
    if state.after_error.get() != 0 {
        return Err(state.after_error.get());
    }
    state.native.set(state.caller.get());
    Ok(())
}
#[test]
fn sync_publishes_native_changes_and_refreshes_callback_writes_without_nested_replay() {
    let state = State::new();
    let mut context = TypedContext {
        remaining_blocks: 20,
    };
    with_native_call(&mut context, || {
        // SAFETY: hooks borrow this stable State for exactly the native body.
        unsafe {
            with_native_callback_sync(state.data(), publish, refresh, || {
                with_native_callback_context(|context| {
                    assert_eq!(state.caller.get(), 12);
                    state.caller.set(13);
                    with_native_call(context, || {
                        with_native_callback_context(|_| {
                            assert_eq!(
                                state.caller.get(),
                                13,
                                "outer native value must not republish in a nested call"
                            );
                            state.caller.set(14);
                        })
                        .unwrap();
                    });
                    assert_eq!(state.native.get(), 12, "outer native code is still paused");
                })
                .unwrap();
                assert_eq!(state.native.get(), 14);
                state.native.set(24);
                with_native_callback_context(|_| {
                    assert_eq!(state.caller.get(), 24);
                    state.caller.set(25);
                })
                .unwrap();
                assert_eq!(state.native.get(), 25);
            })
        }
        .unwrap();
    });
    assert_eq!(
        *state.events.borrow(),
        ["publish", "refresh", "publish", "refresh"]
    );
    assert!(matches!(
        with_native_callback_context(|_| ()),
        Err(Error::Reentry)
    ));
}
#[test]
fn synchronization_faults_refuse_entry_and_remain_visible_when_native_ignores_them() {
    let state = State::new();
    let mut context = TypedContext {
        remaining_blocks: 20,
    };
    with_native_call(&mut context, || {
        state.before_error.set(13);
        // SAFETY: registered State and hooks span this exact synchronous body.
        let result = unsafe {
            with_native_callback_sync(state.data(), publish, refresh, || {
                assert!(matches!(
                    with_native_callback_context(|_| panic!("must not enter")),
                    Err(Error::Runtime { number: 13, .. })
                ));
                7
            })
        };
        assert!(matches!(result, Err(13)));
        state.before_error.set(0);
        state.after_error.set(13);
        // SAFETY: same live same-thread State and nonpanicking hooks.
        let result = unsafe {
            with_native_callback_sync(state.data(), publish, refresh, || {
                assert!(matches!(
                    with_native_callback_context(|_| state.caller.set(13)),
                    Err(Error::Runtime { number: 13, .. })
                ));
                7
            })
        };
        assert!(matches!(result, Err(13)));
        state.after_error.set(0);
        with_native_callback_context(|_| ()).unwrap();
    });
}
#[test]
fn callback_unwind_refreshes_writes_and_restores_all_scopes() {
    let state = State::new();
    let mut context = TypedContext {
        remaining_blocks: 20,
    };
    with_native_call(&mut context, || {
        // SAFETY: registered State and hooks span the body and callback unwind.
        unsafe {
            with_native_callback_sync(state.data(), publish, refresh, || {
                assert!(
                    std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                        with_native_callback_context(|_| {
                            state.caller.set(13);
                            panic!("owned test unwind");
                        })
                    }))
                    .is_err()
                );
                assert_eq!(state.native.get(), 13);
                with_native_callback_context(|_| state.caller.set(14)).unwrap();
                assert_eq!(state.native.get(), 14);
            })
        }
        .unwrap();
    });
    with_native_call(&mut context, || {
        with_native_callback_context(|_| ()).unwrap();
    });
    assert_eq!(
        *state.events.borrow(),
        ["publish", "refresh", "publish", "refresh"]
    );
}
