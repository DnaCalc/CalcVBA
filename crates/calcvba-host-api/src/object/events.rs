//! Same-thread incoming event services supplied by the host object library.
use super::InterfaceId;
use crate::Value;
use std::{any::Any, rc::Rc};

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum EventParameter {
    Object(Option<InterfaceId>),
    BooleanByRef,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EventDescriptor {
    pub name: String,
    pub interface: InterfaceId,
    pub dispatch_id: i32,
    pub parameters: Vec<EventParameter>,
}
pub type EventCallback = Rc<dyn Fn(&mut [Value]) -> Result<(), i32>>;
/// Dropping the library's retained owner unsubscribes before releasing its source.
pub struct EventSubscription {
    _owner: Box<dyn Any>,
}
impl EventSubscription {
    pub fn new(owner: impl Any) -> Self {
        Self {
            _owner: Box::new(owner),
        }
    }
}
