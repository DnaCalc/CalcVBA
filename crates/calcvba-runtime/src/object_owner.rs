//! Owner-scoped multi-producer termination queues for project objects (dc-yqq0.25).
//!
//! Each [`RuntimeObjectOwner`] owns a shared [`OwnerTerminationChannel`] (`Arc`, `Send + Sync`)
//! with an atomic `has_pending` fast-path flag checked at statement boundaries and a mutex-guarded
//! queue holding:
//! - parked instances awaiting `Class_Terminate` (`pending` + `parked`), and
//! - non-terminating instances whose final `Release` arrived on a foreign thread (`pending_drops`),
//!   so their fields, COM references, and any cascaded child `Class_Terminate` run strictly on the
//!   owning VBA thread.

use core::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::cell::RefCell;
use std::collections::BTreeMap;
use std::rc::Rc;
use std::sync::{Arc, Mutex, Weak};

use crate::object_ref::CompatObjectBase;

/// Raw pointer to a `CompatObjectBase` transferred into the owner's multi-producer queue when its
/// final `Release` arrives on a foreign thread.
pub(crate) struct SendCompatPtr(pub(crate) *mut CompatObjectBase);

// SAFETY: a `CompatObjectBase` is placed in `SendCompatPtr` only after its atomic `ref_count`
// reaches 0 with `AcqRel` ordering, transferring sole ownership back to the owning VBA thread's
// queue; fields, `native_state`, `invoker`, and `Class_Terminate` are accessed only on the owner
// thread.
unsafe impl Send for SendCompatPtr {}

#[derive(Default)]
pub(crate) struct OwnerQueueState {
    /// Parked instances awaiting `Class_Terminate`: `(instance_id, bundle_id, route_key)`.
    pub(crate) pending: Vec<(i32, i32, i32)>,
    /// Parked boxes owned by the queue until `finish_pending_termination` or discard.
    pub(crate) parked: BTreeMap<i32, SendCompatPtr>,
    /// Boxes without `Class_Terminate` (or already terminated) whose final `Release` arrived on a
    /// foreign thread; drained and freed on the owning thread.
    pub(crate) pending_drops: Vec<SendCompatPtr>,
}

pub(crate) struct OwnerTerminationChannel {
    pub(crate) alive: AtomicBool,
    /// Cheap atomic flag checked at every VBA statement boundary (`has_pending_terminations`)
    /// without acquiring `queue`.
    pub(crate) has_pending: AtomicBool,
    pub(crate) queue: Mutex<OwnerQueueState>,
}

impl OwnerTerminationChannel {
    fn new() -> Arc<Self> {
        Arc::new(Self {
            alive: AtomicBool::new(true),
            has_pending: AtomicBool::new(false),
            queue: Mutex::new(OwnerQueueState::default()),
        })
    }
}

impl Drop for OwnerTerminationChannel {
    fn drop(&mut self) {
        self.alive.store(false, Ordering::Release);
        crate::object_ref::discard_pending_channel(self);
    }
}

#[derive(Clone)]
pub struct RuntimeObjectOwner {
    id: u64,
    lease: Option<Rc<OwnerLease>>,
    channel: Arc<OwnerTerminationChannel>,
}

type WeakMemberInvoker = std::rc::Weak<RefCell<dyn crate::events::MemberInvoker>>;

struct OwnerLease {
    channel: Arc<OwnerTerminationChannel>,
    member_invoker: RefCell<Option<WeakMemberInvoker>>,
}

impl Drop for OwnerLease {
    fn drop(&mut self) {
        self.channel.alive.store(false, Ordering::Release);
        self.member_invoker.borrow_mut().take();
        crate::object_ref::discard_pending_channel(&self.channel);
    }
}

pub(crate) struct WeakObjectOwner {
    pub id: u64,
    pub(crate) channel: Weak<OwnerTerminationChannel>,
}

pub struct ObjectOwnerScope {
    previous: Option<RuntimeObjectOwner>,
}

static NEXT_OWNER: AtomicU64 = AtomicU64::new(1);

thread_local! {
    static AMBIENT_CHANNEL: Arc<OwnerTerminationChannel> = OwnerTerminationChannel::new();
    static AMBIENT_INVOKER: RefCell<Option<WeakMemberInvoker>> = const { RefCell::new(None) };
    static CURRENT_OWNER: RefCell<Option<RuntimeObjectOwner>> = const { RefCell::new(None) };
}

impl RuntimeObjectOwner {
    fn ambient() -> Self {
        let channel = AMBIENT_CHANNEL.with(Arc::clone);
        Self {
            id: 0,
            lease: None,
            channel,
        }
    }

    pub fn fresh() -> Option<Self> {
        let id = NEXT_OWNER
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |current| {
                current.checked_add(1)
            })
            .ok()?;
        let channel = OwnerTerminationChannel::new();
        let lease = Some(Rc::new(OwnerLease {
            channel: Arc::clone(&channel),
            member_invoker: RefCell::new(None),
        }));
        Some(Self { id, lease, channel })
    }

    /// Install a weak between-runs [`crate::events::MemberInvoker`] inherited by every project
    /// instance created under this owner (used by COM session holders such as `calcvba-comhost`).
    pub fn set_member_invoker(&self, invoker: WeakMemberInvoker) {
        if let Some(lease) = &self.lease {
            *lease.member_invoker.borrow_mut() = Some(invoker);
        } else {
            AMBIENT_INVOKER.with(|slot| {
                *slot.borrow_mut() = Some(invoker);
            });
        }
    }

    pub(crate) fn member_invoker(&self) -> Option<WeakMemberInvoker> {
        if let Some(lease) = &self.lease {
            lease.member_invoker.borrow().clone()
        } else {
            AMBIENT_INVOKER.with(|slot| slot.borrow().clone())
        }
    }

    pub fn current() -> Self {
        CURRENT_OWNER.with(|owner| {
            owner
                .borrow()
                .as_ref()
                .cloned()
                .unwrap_or_else(Self::ambient)
        })
    }

    pub fn id(&self) -> u64 {
        self.id
    }

    pub(crate) fn channel(&self) -> &Arc<OwnerTerminationChannel> {
        &self.channel
    }

    /// Lock-free check for queued `Class_Terminate` callbacks or deferred foreign-thread
    /// drops on this owner's channel, without touching thread-local owner scopes.
    #[inline]
    pub fn has_pending(&self) -> bool {
        self.channel.has_pending.load(Ordering::Acquire)
    }

    pub fn enter(&self) -> ObjectOwnerScope {
        let previous = CURRENT_OWNER.with(|owner| owner.borrow_mut().replace(self.clone()));
        ObjectOwnerScope {
            previous: Some(previous.unwrap_or_else(Self::ambient)),
        }
    }

    pub(crate) fn downgrade(&self) -> WeakObjectOwner {
        WeakObjectOwner {
            id: self.id,
            channel: Arc::downgrade(&self.channel),
        }
    }
}

#[cfg(test)]
impl WeakObjectOwner {
    pub fn is_alive(&self) -> bool {
        self.channel
            .upgrade()
            .is_some_and(|ch| ch.alive.load(Ordering::Acquire))
    }
}

impl Drop for ObjectOwnerScope {
    fn drop(&mut self) {
        if let Some(previous) = self.previous.take() {
            CURRENT_OWNER.with(|owner| {
                let mut slot = owner.borrow_mut();
                if previous.id == 0 && previous.lease.is_none() {
                    *slot = None;
                } else {
                    *slot = Some(previous);
                }
            });
        }
    }
}

#[cfg(test)]
#[allow(clippy::undocumented_unsafe_blocks)]
mod tests {
    use super::*;
    use crate::{ObjectRef, ObjectRefExt, Variant};

    fn object(id: i32) -> ObjectRef {
        let descriptor = ObjectRef::from_compat_identity(9).class_descriptor();
        ObjectRef::from_project_instance(id, 0, 0, true, descriptor)
    }

    fn object_no_terminate(id: i32) -> ObjectRef {
        let descriptor = ObjectRef::from_compat_identity(9).class_descriptor();
        ObjectRef::from_project_instance(id, 0, 0, false, descriptor)
    }

    #[test]
    fn owners_isolate_identical_instance_ids_and_restore_nested_scope() {
        let original = RuntimeObjectOwner::current().id();
        let a = RuntimeObjectOwner::fresh().unwrap();
        let b = RuntimeObjectOwner::fresh().unwrap();
        {
            let _a = a.enter();
            let first = object(1);
            assert!(first.project_field_set(7, Variant::from_i32(11)));
            drop(first);
            {
                let _b = b.enter();
                let second = object(1);
                assert!(second.project_field_set(7, Variant::from_i32(22)));
                drop(second);
            }
            assert_eq!(RuntimeObjectOwner::current().id(), a.id());
            assert_eq!(crate::take_pending_terminations(), vec![(1, 0, 0)]);
            let retained = crate::retained_parked_termination_object(1).unwrap();
            assert_eq!(retained.project_field_get(7).unwrap().as_i32(), Some(11));
            drop(retained);
            assert!(crate::finish_pending_termination(1));
            crate::reset_pending_terminations();
        }
        {
            let _b = b.enter();
            assert_eq!(crate::take_pending_terminations(), vec![(1, 0, 0)]);
            let retained = crate::retained_parked_termination_object(1).unwrap();
            assert_eq!(retained.project_field_get(7).unwrap().as_i32(), Some(22));
            drop(retained);
            assert!(crate::finish_pending_termination(1));
        }
        assert_eq!(RuntimeObjectOwner::current().id(), original);
    }

    #[test]
    fn owner_expiry_discards_parked_cascades_and_does_not_queue_late_releases() {
        let owner = RuntimeObjectOwner::fresh().unwrap();
        let weak = owner.downgrade();
        let held;
        {
            let _scope = owner.enter();
            let parent = object(1);
            let child = object(2);
            assert!(parent.project_field_set(7, Variant::from_object_ref(child)));
            drop(parent);
            held = object(3);
            assert!(crate::has_pending_terminations());
        }
        drop(owner);
        assert!(!weak.is_alive());
        drop(held);
        assert!(weak.channel.upgrade().is_none());
        assert!(!crate::has_pending_terminations());
    }

    #[test]
    fn project_instance_marker_does_not_depend_on_identity_differing_from_class_id() {
        let descriptor = ObjectRef::from_compat_identity(9).class_descriptor();
        let instance = ObjectRef::from_project_instance(1, 1, 0, false, descriptor);
        assert!(instance.is_project_instance());
        assert!(instance.project_field_set(7, Variant::from_i32(42)));
        assert_eq!(instance.project_field_get(7).unwrap().as_i32(), Some(42));
        assert!(!ObjectRef::from_compat_identity(1).is_project_instance());
    }

    #[test]
    fn foreign_thread_final_release_routes_termination_to_owner_thread_queue() {
        struct SendUnknown(*mut crate::RawRuntimeIUnknown);
        unsafe impl Send for SendUnknown {}

        let owner = RuntimeObjectOwner::fresh().unwrap();
        let _scope = owner.enter();

        let obj = object(101);
        assert!(obj.project_field_set(7, Variant::from_i32(42)));
        let raw = obj.raw_iunknown();
        unsafe { ((*(*raw).vtbl).add_ref)(raw.cast()) };
        drop(obj);
        assert!(
            !crate::has_pending_terminations(),
            "held foreign reference must prevent termination queuing"
        );

        let foreign_ptr = SendUnknown(raw);
        let remaining = std::thread::spawn(move || {
            let ptr = foreign_ptr;
            unsafe { ((*(*ptr.0).vtbl).release)(ptr.0.cast()) }
        })
        .join()
        .expect("foreign release thread");
        assert_eq!(remaining, 0);

        assert!(
            crate::has_pending_terminations(),
            "foreign-thread final release must enqueue onto owner thread's queue"
        );
        assert_eq!(crate::take_pending_terminations(), vec![(101, 0, 0)]);
        let parked = crate::retained_parked_termination_object(101).expect("parked object");
        assert_eq!(parked.project_field_get(7).unwrap().as_i32(), Some(42));
        drop(parked);
        assert!(crate::finish_pending_termination(101));
        assert!(!crate::has_pending_terminations());
    }

    #[test]
    fn foreign_thread_final_release_without_terminate_drops_fields_and_cascades_on_owner_thread() {
        struct SendUnknown(*mut crate::RawRuntimeIUnknown);
        unsafe impl Send for SendUnknown {}

        let owner = RuntimeObjectOwner::fresh().unwrap();
        let _scope = owner.enter();

        let parent = object_no_terminate(201);
        let child = object(202);
        assert!(child.project_field_set(9, Variant::from_i32(99)));
        assert!(parent.project_field_set(7, Variant::from_object_ref(child)));

        let raw_parent = parent.raw_iunknown();
        unsafe { ((*(*raw_parent).vtbl).add_ref)(raw_parent.cast()) };
        drop(parent);
        assert!(!crate::has_pending_terminations());

        let foreign_ptr = SendUnknown(raw_parent);
        let remaining = std::thread::spawn(move || {
            let ptr = foreign_ptr;
            unsafe { ((*(*ptr.0).vtbl).release)(ptr.0.cast()) }
        })
        .join()
        .expect("foreign release thread");
        assert_eq!(remaining, 0);

        // Parent had no Class_Terminate, so foreign release queued its drop onto the owner
        // thread without running child drop on the foreign thread.
        assert!(crate::has_pending_terminations());
        // Draining on the owner thread drops parent's fields on the owner thread and surfaces
        // child's cascaded Class_Terminate in the same batch!
        assert_eq!(crate::take_pending_terminations(), vec![(202, 0, 0)]);
        let parked_child = crate::retained_parked_termination_object(202).expect("parked child");
        assert_eq!(
            parked_child.project_field_get(9).unwrap().as_i32(),
            Some(99)
        );
        drop(parked_child);
        assert!(crate::finish_pending_termination(202));
        assert!(!crate::has_pending_terminations());
    }
}
