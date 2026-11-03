//! Same-thread host services for retained object interface views and iterators.
use super::{HostObject, InterfaceId};
use crate::{TypedContext, Value};

#[derive(Clone, Copy, Debug)]
pub enum DynamicMemberKind {
    Method,
    Get,
    Let,
    Set,
}
pub struct DynamicArgument {
    pub name: Option<String>,
    pub by_ref: bool,
    pub alias: usize,
}
/// A genuinely untyped VBA call. None selects DISPID_VALUE.
pub struct DynamicMemberCall {
    pub name: Option<String>,
    pub kind: DynamicMemberKind,
    pub arguments: Vec<DynamicArgument>,
}

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
    fn subscribe_event(
        &self,
        _descriptor: &super::EventDescriptor,
        _callback: super::EventCallback,
    ) -> Result<super::EventSubscription, i32> {
        Err(438)
    }
    fn invoke_dynamic(
        &self,
        _call: &DynamicMemberCall,
        _arguments: &mut [oxvba_runtime::Variant],
        _context: &mut TypedContext,
    ) -> Result<oxvba_runtime::Variant, crate::TypedFault> {
        Err(crate::TypedFault {
            number: 438,
            description: oxvba_runtime::default_error_message(438).into(),
            source: None,
            help_file: None,
            help_context: None,
        })
    }
    fn resolve_interface(&self, _iid: InterfaceId) -> Result<HostObject, i32> {
        Err(13)
    }
    fn enumerator(&self, _context: &mut TypedContext) -> Result<HostEnumerator, i32> {
        Err(438)
    }
}
