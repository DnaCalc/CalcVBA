//! Automatic source subscriptions, using compiler routes and the canonical fabric.
use super::*;
use crate::cache::SessionCell;
use calcvba_host_api::{
    self as api, Value,
    object::{EventDescriptor, EventParameter},
};
use oxvba_rt_abi::automatic_events::{AutomaticBinding, AutomaticEventState, BindingKey};
use std::{collections::HashMap, rc::Weak};

pub(crate) struct AutomaticEvents {
    next_generation: Cell<u64>,
    pub unhandled: RefCell<Option<Error>>,
    pub session: RefCell<Weak<SessionCell>>,
    pub state: RefCell<Rc<RefCell<AutomaticEventState>>>,
    pub documents: RefCell<HashMap<(usize, usize), api::HostObject>>,
}
impl AutomaticEvents {
    pub fn new(state: Rc<RefCell<AutomaticEventState>>) -> Rc<Self> {
        Rc::new(Self {
            next_generation: Cell::new(0),
            unhandled: RefCell::new(None),
            session: RefCell::new(Weak::new()),
            state: RefCell::new(state),
            documents: RefCell::default(),
        })
    }
    pub fn clear(&self) {
        self.unhandled.borrow_mut().take();
        let state = self.state.borrow().clone();
        let old = std::mem::take(&mut state.borrow_mut().bindings);
        drop(old);
    }
    pub fn collect(&self) {
        let state = self.state.borrow().clone();
        let (retired, terminating) = {
            let state = state.borrow();
            let mut counts = HashMap::<i32, u32>::new();
            for binding in state.bindings.values() {
                if let Some(owner) = &binding.owner {
                    *counts.entry(owner.raw()).or_default() += 1;
                }
            }
            let mut retired = Vec::new();
            let mut terminating = Vec::new();
            for (key, binding) in &state.bindings {
                if let Some(owner) = &binding.owner
                    && owner.strong_count() <= counts[&owner.raw()]
                {
                    if owner.terminates_on_release() {
                        terminating.push(*key)
                    } else {
                        retired.push(*key)
                    }
                }
            }
            (retired, terminating)
        };
        let old = retired
            .into_iter()
            .filter_map(|key| state.borrow_mut().bindings.remove(&key))
            .collect::<Vec<_>>();
        drop(old);
        let owners = terminating
            .into_iter()
            .filter_map(|key| {
                state
                    .borrow_mut()
                    .bindings
                    .get_mut(&key)
                    .and_then(|binding| binding.owner.take())
            })
            .collect::<Vec<_>>();
        // Dropping these final private references parks the canonical instance
        // for its normal finalizer drain, while WithEvents remains readable.
        drop(owners);
    }
    pub fn get(&self, unit: usize, owner: &oxvba_runtime::ObjectRef, binding: i32) -> Variant {
        let state = self.state.borrow().clone();
        state
            .borrow()
            .bindings
            .get(&(unit, owner.raw(), binding))
            .map(|b| b.source.clone())
            .unwrap_or_else(Variant::nothing)
    }
    pub fn clear_owner(&self, unit: usize, owner: &oxvba_runtime::ObjectRef) {
        self.clear_owner_id(unit, owner.raw());
    }
    pub fn clear_owner_id(&self, unit: usize, owner: i32) {
        let state = self.state.borrow().clone();
        let keys = state
            .borrow()
            .bindings
            .keys()
            .filter(|key| key.0 == unit && key.1 == owner)
            .copied()
            .collect::<Vec<_>>();
        let old = keys
            .into_iter()
            .filter_map(|key| state.borrow_mut().bindings.remove(&key))
            .collect::<Vec<_>>();
        drop(old);
    }
    pub fn set(
        self: &Rc<Self>,
        unit: usize,
        owner: oxvba_runtime::ObjectRef,
        binding: i32,
        source: Variant,
        program: &OxProgram,
        entries: &[Result<oxvba_vm3::typed::TypedInterpreter<'_>, String>],
    ) -> Result<(), i32> {
        let key = (unit, owner.raw(), binding);
        let state = self.state.borrow().clone();
        if source.vtype() != oxvba_runtime::VarType::Object {
            return Err(424);
        }
        let Some(reference) = source.as_object_ref() else {
            let old = state.borrow_mut().bindings.remove(&key);
            drop(old);
            return Ok(());
        };
        let routes = program
            .event_routes
            .iter()
            .filter(|route| route.binding == binding)
            .collect::<Vec<_>>();
        // Validate every matching handler before publishing any subscription.
        let mut plans = Vec::new();
        for route in routes {
            let typed = entries
                .get(route.handler)
                .and_then(|e| e.as_ref().ok())
                .ok_or(5)?;
            let Some(native) = &route.native else {
                continue;
            };
            if native.dispatch_path != oxvba_com::TypeLibEventDispatchPath::Dispatch {
                return Err(438);
            }
            let iid = crate::events::parse_iid(native.connection_point_iid.as_deref().ok_or(13)?)?;
            let parameters = native
                .parameter_types
                .iter()
                .enumerate()
                .map(|(i, p)| match p {
                    oxvba_com::TypeLibParamType::Object => Ok(EventParameter::Object(
                        typed.layout().parameter_iids.get(i + 1).copied().flatten(),
                    )),
                    oxvba_com::TypeLibParamType::ByRefBoolean => Ok(EventParameter::BooleanByRef),
                    oxvba_com::TypeLibParamType::Boolean => Ok(EventParameter::Boolean),
                    _ => Err(13),
                })
                .collect::<Result<Vec<_>, _>>()?;
            if typed.layout().parameters.len() != parameters.len() + 1 || typed.layout().has_result
            {
                return Err(13);
            }
            let descriptor = EventDescriptor {
                name: native.name.clone(),
                interface: provider::interface_id(iid),
                dispatch_id: native.dispatch_member_id.ok_or(13)?,
                parameters,
            };
            plans.push((route.handler, descriptor));
        }
        if !plans.is_empty() && self.session.borrow().upgrade().is_none() {
            return Err(5);
        }
        let host = if reference.is_project_instance() {
            None
        } else {
            Some(api::HostObject::from_object_ref(&reference).ok_or(13)?)
        };
        let generation = {
            // Session reset replaces the canonical VM state. This counter must
            // survive that replacement so a queued old callback cannot match a
            // newly allocated document with the same numeric instance identity.
            let generation = self.next_generation.get().checked_add(1).ok_or(7)?;
            self.next_generation.set(generation);
            state.borrow_mut().next_generation = generation;
            generation
        };
        let old = state.borrow_mut().bindings.remove(&key);
        drop(old);
        state.borrow_mut().bindings.insert(
            key,
            AutomaticBinding {
                owner: Some(owner),
                source,
                subscriptions: vec![],
                generation,
            },
        );
        for (handler, descriptor) in plans {
            let weak = Rc::downgrade(self);
            let callback = Rc::new(move |arguments: &mut [Value]| {
                weak.upgrade().map_or(Ok(()), |events| {
                    events.deliver(key, generation, handler, arguments)
                })
            });
            let subscription = host
                .as_ref()
                .ok_or(13)?
                .subscribe_event(&descriptor, callback);
            let subscription = match subscription {
                Ok(subscription) => subscription,
                Err(error) => {
                    let old = {
                        let mut state = state.borrow_mut();
                        if state
                            .bindings
                            .get(&key)
                            .is_some_and(|binding| binding.generation == generation)
                        {
                            state.bindings.remove(&key)
                        } else {
                            None
                        }
                    };
                    drop(old);
                    return Err(error);
                }
            };
            let mut state = state.borrow_mut();
            if let Some(binding) = state
                .bindings
                .get_mut(&key)
                .filter(|b| b.generation == generation)
            {
                binding.subscriptions.push(subscription);
            } else {
                drop(state);
                drop(subscription);
            }
        }
        Ok(())
    }
    pub(crate) fn deliver(
        &self,
        key: BindingKey,
        generation: u64,
        handler: usize,
        arguments: &mut [Value],
    ) -> Result<(), i32> {
        let state = self.state.borrow().clone();
        let owner = {
            let state = state.borrow();
            let Some(binding) = state
                .bindings
                .get(&key)
                .filter(|b| b.generation == generation)
            else {
                return Ok(());
            };
            binding
                .owner
                .clone()
                .or_else(|| oxvba_runtime::retained_parked_termination_object(key.1))
                .ok_or(5)?
        };
        let Some(cell) = self.session.borrow().upgrade() else {
            return Ok(());
        };
        let called = Cell::new(false);
        let result = api::with_native_callback_context(|context| {
            called.set(true);
            if api::has_unhandled_event_fault(context) {
                return Err(api::UNHANDLED_EVENT_ABORT);
            }
            cell.with_dependent(|_, session| {
                session.call_event(key.0, FuncId(handler), &owner, arguments, context)
            })
        });
        match result {
            Ok(result) => result,
            Err(_) if called.get() => Err(5),
            Err(_) => cell.with_dependent(|_, session| {
                if session.typed_depth.get() != 0 || session.suspended.get() != 0 {
                    return Err(5);
                }
                session.call_event(
                    key.0,
                    FuncId(handler),
                    &owner,
                    arguments,
                    &mut api::TypedContext {
                        remaining_blocks: INITIALIZATION_BUDGET,
                    },
                )
            }),
        }
    }
}
// The native IID is data from the imported source descriptor; no registration,
// object-name lookup or type discovery occurs at delivery.
fn parse_iid(text: &str) -> Result<oxvba_com::ComInterfaceIid, i32> {
    let id = oxvba_runtime::RuntimeGuid::parse(text).ok_or(13)?;
    Ok(oxvba_com::ComInterfaceIid {
        data1: id.data1,
        data2: id.data2,
        data3: id.data3,
        data4: id.data4,
    })
}
