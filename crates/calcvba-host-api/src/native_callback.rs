//! Same-thread context borrowing across a synchronous native call.
use crate::{Error, TypedContext};
use std::cell::Cell;
mod sync;
pub use sync::{NativeCallbackSync, with_native_callback_sync};
thread_local! {
    static CONTEXT: Cell<*mut TypedContext> = const { Cell::new(std::ptr::null_mut()) };
    static BORROWED: Cell<bool> = const { Cell::new(false) };
}
struct NativeScope {
    previous: *mut TypedContext,
    borrowed: bool,
}
impl Drop for NativeScope {
    fn drop(&mut self) {
        CONTEXT.with(|slot| slot.set(self.previous));
        BORROWED.with(|slot| slot.set(self.borrowed));
    }
}
/// Lend the current typed budget to same-thread callbacks during exactly this
/// synchronous native operation. The scope restores on return and unwinding.
/// Native callbacks cannot suspend the foreign call or retain a context reference.
pub fn with_native_call<R>(context: &mut TypedContext, call: impl FnOnce() -> R) -> R {
    let _synchronization_scope = sync::native_call_scope();
    let _scope = NativeScope {
        previous: CONTEXT.with(|slot| slot.replace(context)),
        borrowed: BORROWED.with(|slot| slot.replace(false)),
    };
    call()
}
struct CallbackBorrow;
impl Drop for CallbackBorrow {
    fn drop(&mut self) {
        BORROWED.with(|slot| slot.set(false));
    }
}
/// Reborrow the active native call's typed context for a synchronous host callback.
/// Missing scopes, other threads and overlapping borrows reject as Reentry.
/// Pass this context to SessionExt::call_typed_in_context or Session::call_values;
/// normal session generation, budget and recursion checks continue to apply.
pub fn with_native_callback_context<R>(
    callback: impl FnOnce(&mut TypedContext) -> R,
) -> Result<R, Error> {
    let pointer = CONTEXT.with(Cell::get);
    if pointer.is_null() || BORROWED.with(|slot| slot.replace(true)) {
        return Err(Error::Reentry);
    }
    let _borrow = CallbackBorrow;
    let synchronization = sync::begin_callback()?;
    // SAFETY: with_native_call retains the exclusive live input borrow until
    // its private scope drops. This same-thread flag excludes another mutable
    // borrower; nested native calls install their own scoped reborrow. The
    // callback's higher-ranked input cannot escape in R.
    let result = callback(unsafe { &mut *pointer });
    synchronization.finish()?;
    Ok(result)
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn native_scopes_share_budget_reject_overlap_and_restore_after_nested_calls() {
        assert!(matches!(
            with_native_callback_context(|_| ()),
            Err(Error::Reentry)
        ));
        let mut context = TypedContext {
            remaining_blocks: 5,
        };
        with_native_call(&mut context, || {
            with_native_callback_context(|context| {
                context.step().unwrap();
                assert!(matches!(
                    with_native_callback_context(|_| ()),
                    Err(Error::Reentry)
                ));
                with_native_call(context, || {
                    with_native_callback_context(|inner| inner.step().unwrap()).unwrap()
                });
                assert_eq!(context.remaining_blocks, 3);
            })
            .unwrap();
            assert!(
                std::thread::spawn(|| matches!(
                    with_native_callback_context(|_| ()),
                    Err(Error::Reentry)
                ))
                .join()
                .unwrap()
            );
        });
        assert_eq!(context.remaining_blocks, 3);
        with_native_call(&mut context, || {
            with_native_callback_context(|context| {
                for _ in 0..3 {
                    context.step().unwrap();
                }
                assert_eq!(context.step(), Err(-1));
            })
            .unwrap()
        });
        assert_eq!(context.remaining_blocks, 0);
        assert!(matches!(
            with_native_callback_context(|_| ()),
            Err(Error::Reentry)
        ));
    }
    #[test]
    fn native_scope_and_borrow_restore_on_unwind() {
        let mut context = TypedContext {
            remaining_blocks: 5,
        };
        assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| with_native_call(
                &mut context,
                || with_native_callback_context::<()>(|_| panic!("fixture"))
            )))
            .is_err()
        );
        assert!(matches!(
            with_native_callback_context(|_| ()),
            Err(Error::Reentry)
        ));
        with_native_call(&mut context, || {
            with_native_callback_context(|context| context.step().unwrap()).unwrap()
        });
        assert_eq!(context.remaining_blocks, 4);
    }
}
