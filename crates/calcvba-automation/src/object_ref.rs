//! The owned interface pointer a `VT_UNKNOWN`/`VT_DISPATCH` cell holds.
//!
//! [`ObjectRef`] is one COM reference: `AddRef` on clone, `Release` on drop, nothing else. What
//! the object is (a CalcVBA project instance, a built-in, a foreign COM object) is the business
//! of whoever made it; `calcvba-runtime` adds its project-object operations through its
//! `ObjectRefExt` trait.

use crate::com::IUnknownVtbl;
use core::ptr::NonNull;
use std::hash::{Hash, Hasher};

/// The first field of every COM object: its vtable pointer.
#[repr(C)]
pub struct RawRuntimeIUnknown {
    pub vtbl: *const IUnknownVtbl,
}

/// One owned COM reference (`IUnknown*` or any interface derived from it).
pub struct ObjectRef(NonNull<RawRuntimeIUnknown>);

impl ObjectRef {
    /// Takes over the caller's reference.
    ///
    /// # Safety
    /// `raw` is null or a live COM interface pointer whose reference the caller owns.
    pub unsafe fn from_raw_iunknown_owned(raw: *mut RawRuntimeIUnknown) -> Option<Self> {
        NonNull::new(raw).map(Self)
    }

    /// Takes a new reference (`AddRef`).
    ///
    /// # Safety
    /// `raw` is null or a live COM interface pointer.
    pub unsafe fn from_raw_iunknown_addref(raw: *mut RawRuntimeIUnknown) -> Option<Self> {
        let raw = NonNull::new(raw)?;
        // SAFETY: guaranteed by the caller; slot 1 of every COM vtable is AddRef.
        unsafe {
            let vtbl = (*raw.as_ptr()).vtbl;
            ((*vtbl).add_ref)(raw.as_ptr().cast());
        }
        Some(Self(raw))
    }

    pub fn raw_iunknown(&self) -> *mut RawRuntimeIUnknown {
        self.0.as_ptr()
    }

    /// The interface's vtable pointer (the first pointer of every COM object).
    pub fn vtbl(&self) -> *const IUnknownVtbl {
        // SAFETY: this reference keeps a live COM object whose first field is its vtable.
        unsafe { (*self.0.as_ptr()).vtbl }
    }

    /// Transfers this reference to a raw carrier, which must adopt it later with
    /// [`Self::from_raw_iunknown_owned`].
    pub fn into_raw_iunknown(self) -> *mut RawRuntimeIUnknown {
        let owner = core::mem::ManuallyDrop::new(self);
        owner.0.as_ptr()
    }
}

impl Clone for ObjectRef {
    fn clone(&self) -> Self {
        // SAFETY: this `ObjectRef` owns a reference, so `self.0` is a live IUnknown; the AddRef
        // gives the clone its own.
        unsafe {
            let vtbl = (*self.0.as_ptr()).vtbl;
            ((*vtbl).add_ref)(self.0.as_ptr().cast());
        }
        Self(self.0)
    }
}

impl Drop for ObjectRef {
    fn drop(&mut self) {
        // SAFETY: releases the one reference this `ObjectRef` owns; the pointer is not touched
        // afterwards even if `release` frees the object.
        unsafe {
            let vtbl = (*self.0.as_ptr()).vtbl;
            ((*vtbl).release)(self.0.as_ptr().cast());
        }
    }
}

impl core::fmt::Debug for ObjectRef {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        f.debug_struct("ObjectRef").field("ptr", &self.0).finish()
    }
}

impl PartialEq for ObjectRef {
    fn eq(&self, other: &Self) -> bool {
        self.0 == other.0
    }
}

impl Eq for ObjectRef {}

impl Hash for ObjectRef {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.0.hash(state);
    }
}

/// A reference-counted test object with an identity number, standing in for the runtime's
/// project objects in this crate's tests.
#[cfg(test)]
pub(crate) mod test_object {
    use super::{ObjectRef, RawRuntimeIUnknown};
    use crate::com::{Guid, IID_IUNKNOWN, IUnknownVtbl};
    use core::ffi::c_void;
    use core::sync::atomic::{AtomicU32, Ordering};

    #[repr(C)]
    struct TestObject {
        unknown: RawRuntimeIUnknown,
        refs: AtomicU32,
        identity: i32,
    }

    static VTBL: IUnknownVtbl = IUnknownVtbl {
        query_interface,
        add_ref,
        release,
    };

    unsafe extern "system" fn query_interface(
        this: *mut c_void,
        riid: *const Guid,
        ppv: *mut *mut c_void,
    ) -> i32 {
        // SAFETY: COM contract: `ppv` is writable, `riid` readable.
        unsafe {
            *ppv = core::ptr::null_mut();
            if *riid != IID_IUNKNOWN {
                return 0x8000_4002_u32 as i32;
            }
            *ppv = this;
            add_ref(this);
        }
        0
    }

    unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
        // SAFETY: `this` is a live `TestObject`.
        unsafe {
            (*this.cast::<TestObject>())
                .refs
                .fetch_add(1, Ordering::AcqRel)
                + 1
        }
    }

    unsafe extern "system" fn release(this: *mut c_void) -> u32 {
        // SAFETY: `this` is a live `TestObject` from `Box::into_raw`; the last release frees it.
        unsafe {
            let left = (*this.cast::<TestObject>())
                .refs
                .fetch_sub(1, Ordering::AcqRel)
                - 1;
            if left == 0 {
                drop(Box::from_raw(this.cast::<TestObject>()));
            }
            left
        }
    }

    impl ObjectRef {
        pub(crate) fn from_compat_identity(identity: i32) -> Self {
            let object = Box::into_raw(Box::new(TestObject {
                unknown: RawRuntimeIUnknown { vtbl: &VTBL },
                refs: AtomicU32::new(1),
                identity,
            }));
            // SAFETY: a fresh object with one reference, handed over.
            unsafe { Self::from_raw_iunknown_owned(object.cast()) }.expect("non-null")
        }

        pub(crate) fn compat_identity(&self) -> i32 {
            assert!(core::ptr::eq(self.vtbl(), &VTBL), "not a test object");
            // SAFETY: checked to be a live `TestObject` above.
            unsafe { (*self.raw_iunknown().cast::<TestObject>()).identity }
        }

        #[allow(dead_code)]
        pub(crate) fn strong_count(&self) -> u32 {
            assert!(core::ptr::eq(self.vtbl(), &VTBL), "not a test object");
            // SAFETY: checked to be a live `TestObject` above.
            unsafe {
                (*self.raw_iunknown().cast::<TestObject>())
                    .refs
                    .load(Ordering::Acquire)
            }
        }
    }
}
