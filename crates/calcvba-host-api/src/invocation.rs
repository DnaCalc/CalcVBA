//! Host-owned invocation identity and a resumable, owning typed result contract.
use crate::{Error, HostObject};
use std::{any::Any, rc::Rc};

/// The engine retains this view; only the host interprets its contents.
#[derive(Clone)]
pub struct InvocationFrame {
    pub id: u64,
    pub parent: Option<u64>,
    pub read_view: Rc<dyn Any>,
}
impl InvocationFrame {
    pub fn new(id: u64, parent: Option<u64>, read_view: Rc<dyn Any>) -> Self {
        Self {
            id,
            parent,
            read_view,
        }
    }
}

pub enum OwnedValue {
    Long(i32),
    Double(f64),
    Object(Option<HostObject>),
}
impl std::fmt::Debug for OwnedValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Long(v) => f.debug_tuple("Long").field(v).finish(),
            Self::Double(v) => f.debug_tuple("Double").field(v).finish(),
            Self::Object(v) => f
                .debug_tuple("Object")
                .field(&v.as_ref().map(|o| o.argument().as_ptr()))
                .finish(),
        }
    }
}
#[derive(Debug)]
pub enum InvocationProgress {
    Suspended,
    Returned(OwnedValue),
}
/// Same-thread, single-use continuation. A slice pauses only between completed
/// instructions; callbacks are synchronous and cannot be interrupted mid-effect.
/// Dropping the invocation cancels it and releases its local owners and leases.
pub trait Invocation {
    fn frame(&self) -> &InvocationFrame;
    fn remaining_blocks(&self) -> u64;
    fn resume(&mut self, operations: u64) -> Result<InvocationProgress, Error>;
}

thread_local! {
    static ACTIVE_FRAME: std::cell::RefCell<Option<Rc<InvocationFrame>>> = const { std::cell::RefCell::new(None) };
}
/// Callback-scoped metadata. Retaining it retains the host's read view, not an
/// engine stack pointer. Nested synchronous calls inherit the active invocation.
pub fn active_invocation() -> Option<Rc<InvocationFrame>> {
    ACTIVE_FRAME.with(|slot| slot.borrow().clone())
}
pub struct InvocationScope(Option<Rc<InvocationFrame>>);
impl InvocationScope {
    pub fn enter(frame: Rc<InvocationFrame>) -> Self {
        Self(ACTIVE_FRAME.with(|slot| slot.replace(Some(frame))))
    }
}
impl Drop for InvocationScope {
    fn drop(&mut self) {
        ACTIVE_FRAME.with(|slot| {
            slot.replace(self.0.take());
        });
    }
}
