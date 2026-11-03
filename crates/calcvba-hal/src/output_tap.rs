use std::{cell::RefCell, marker::PhantomData, rc::Rc, sync::Arc};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HostOutputChannel {
    Stdout,
    Stderr,
    Host,
}

pub trait HostOutputTap {
    fn on_output(&self, channel: HostOutputChannel, text: &str);
}

impl<F> HostOutputTap for F
where
    F: Fn(HostOutputChannel, &str),
{
    fn on_output(&self, channel: HostOutputChannel, text: &str) {
        self(channel, text);
    }
}

thread_local! {
    static OUTPUT_TAPS: RefCell<Vec<Arc<dyn HostOutputTap>>> = RefCell::new(Vec::new());
}

/// Removes the current thread's output tap when dropped.
///
/// ```compile_fail
/// use calcvba_hal::install_thread_output_tap;
/// let guard = install_thread_output_tap(std::sync::Arc::new(|_: calcvba_hal::HostOutputChannel, _: &str| {}));
/// std::thread::spawn(move || drop(guard));
/// ```
pub struct ThreadOutputTapGuard(PhantomData<Rc<()>>);

impl Drop for ThreadOutputTapGuard {
    fn drop(&mut self) {
        OUTPUT_TAPS.with(|taps| {
            taps.borrow_mut().pop();
        });
    }
}

pub fn install_thread_output_tap(tap: Arc<dyn HostOutputTap>) -> ThreadOutputTapGuard {
    OUTPUT_TAPS.with(|taps| {
        taps.borrow_mut().push(tap);
    });
    ThreadOutputTapGuard(PhantomData)
}

pub(crate) fn emit_thread_output_tap(channel: HostOutputChannel, text: &str) {
    OUTPUT_TAPS.with(|taps| {
        for tap in taps.borrow().iter() {
            tap.on_output(channel, text);
        }
    });
}
