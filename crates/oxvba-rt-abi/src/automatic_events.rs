//! Canonical owning subscription state, retained by the session event fabric.
use calcvba_host_api::object::EventSubscription;
use oxvba_runtime::{ObjectRef, Variant};
use std::collections::HashMap;

pub type BindingKey = (usize, i32, i32);
pub struct AutomaticBinding {
    /// Released before Class_Terminate; source/subscriptions survive that body.
    pub owner: Option<ObjectRef>,
    pub source: Variant,
    pub subscriptions: Vec<EventSubscription>,
    pub generation: u64,
}
#[derive(Default)]
pub struct AutomaticEventState {
    pub bindings: HashMap<BindingKey, AutomaticBinding>,
    pub next_generation: u64,
}
