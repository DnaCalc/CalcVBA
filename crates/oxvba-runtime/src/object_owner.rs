//! Same-thread project ownership for deferred class finalization.
use std::{
    cell::{Cell, RefCell},
    rc::{Rc, Weak},
};
#[derive(Clone)]
pub struct RuntimeObjectOwner {
    id: u64,
    lease: Option<Rc<OwnerLease>>,
}
struct OwnerLease {
    id: u64,
}
impl Drop for OwnerLease {
    fn drop(&mut self) {
        crate::object_ref::discard_pending_owner(self.id);
    }
}
pub(crate) struct WeakObjectOwner {
    pub id: u64,
    lease: Option<Weak<OwnerLease>>,
}
pub struct ObjectOwnerScope {
    previous: Option<RuntimeObjectOwner>,
}
thread_local! {
    static NEXT_OWNER: Cell<u64> = const { Cell::new(1) };
    static CURRENT_OWNER: RefCell<RuntimeObjectOwner> = const { RefCell::new(RuntimeObjectOwner { id:0, lease:None }) };
}
impl RuntimeObjectOwner {
    pub fn fresh() -> Option<Self> {
        NEXT_OWNER.with(|next| {
            let id = next.get();
            next.set(id.checked_add(1)?);
            Some(Self {
                id,
                lease: Some(Rc::new(OwnerLease { id })),
            })
        })
    }
    pub fn current() -> Self {
        CURRENT_OWNER.with(|owner| owner.borrow().clone())
    }
    pub fn id(&self) -> u64 {
        self.id
    }
    pub fn enter(&self) -> ObjectOwnerScope {
        ObjectOwnerScope {
            previous: Some(CURRENT_OWNER.with(|owner| owner.replace(self.clone()))),
        }
    }
    pub(crate) fn downgrade(&self) -> WeakObjectOwner {
        WeakObjectOwner {
            id: self.id,
            lease: self.lease.as_ref().map(Rc::downgrade),
        }
    }
}
impl WeakObjectOwner {
    pub fn is_alive(&self) -> bool {
        self.lease
            .as_ref()
            .is_none_or(|lease| lease.strong_count() != 0)
    }
}
impl Drop for ObjectOwnerScope {
    fn drop(&mut self) {
        if let Some(previous) = self.previous.take() {
            CURRENT_OWNER.with(|owner| {
                owner.replace(previous);
            });
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{ObjectRef, Variant};
    fn object(id: i32) -> ObjectRef {
        let descriptor = ObjectRef::from_compat_identity(9).class_descriptor();
        ObjectRef::from_project_instance(id, 0, 0, true, descriptor)
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
        // The expired owner is absent from both the pending and parked queues.
        crate::object_ref::discard_pending_owner(weak.id);
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
}
