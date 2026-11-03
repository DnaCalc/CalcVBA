//! Same-thread host services for retained object interface views and iterators.
use super::{HostObject, InterfaceId};
use crate::{TypedContext, Value};

/// One retained, live iterator. Each advance completes synchronously; its owner
/// survives instruction-boundary suspension and is released on exhaustion/drop.
pub struct HostEnumerator {
    next: Box<Next>,
}
type Next = dyn FnMut(&mut TypedContext) -> Result<Option<Value>, i32>;
impl HostEnumerator {
    pub fn new(
        next: impl FnMut(&mut TypedContext) -> Result<Option<Value>, i32> + 'static,
    ) -> Self {
        Self {
            next: Box::new(next),
        }
    }
    pub fn advance(&mut self, context: &mut TypedContext) -> Result<Option<Value>, i32> {
        (self.next)(context)
    }
}

/// Optional capabilities supplied by an object library. Engine code knows no
/// native pointers or application member names through this contract.
pub trait HostObjectServices {
    fn resolve_interface(&self, _iid: InterfaceId) -> Result<HostObject, i32> {
        Err(13)
    }
    fn enumerator(&self, _context: &mut TypedContext) -> Result<HostEnumerator, i32> {
        Err(438)
    }
}
