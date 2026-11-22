use calcvba_rt_abi::native_export::WRONG_THREAD;
use std::{
    cell::Cell,
    sync::atomic::{AtomicBool, Ordering},
};

static CLAIMED: AtomicBool = AtomicBool::new(false);
thread_local! {
    static IS_OWNER: Cell<bool> = const { Cell::new(false) };
}

#[unsafe(no_mangle)]
pub extern "C" fn cvb_enter_scalar() -> i32 {
    let saved = super::thread::last_error();
    // A TLS flag has the lifetime of the actual thread. A later thread
    // with a recycled pthread ID cannot inherit it.
    if IS_OWNER.with(|owner| {
        if owner.get() {
            return true;
        }
        if CLAIMED.swap(true, Ordering::AcqRel) {
            return false;
        }
        owner.set(true);
        true
    }) {
        super::thread::set_last_error(saved);
        0
    } else {
        super::thread::set_last_error(WRONG_THREAD as u32);
        WRONG_THREAD
    }
}
