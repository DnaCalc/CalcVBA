//! Same-thread host-owned object tables. This ABI is independent of native COM.
use crate::{ScalarBits, Slot, TypedContext};
use oxvba_runtime::{
    RUNTIME_GUID_IUNKNOWN, RawRuntimeIUnknown, RawRuntimeIUnknownVtbl, RuntimeGuid,
};
use std::{
    marker::PhantomData,
    panic::{AssertUnwindSafe, catch_unwind},
    rc::Rc,
};
mod iteration;
pub use iteration::{
    DynamicArgument, DynamicMemberCall, DynamicMemberKind, HostEnumerator, HostObjectServices,
};

/// Canonical Automation IDispatch IID, in the same byte order as InterfaceId.
pub const IDISPATCH: InterfaceId = [0, 4, 2, 0, 0, 0, 0, 0, 0xc0, 0, 0, 0, 0, 0, 0, 0x46];
pub mod planned_slots;
pub mod value_slots;

pub type InterfaceId = [u8; 16];
pub const SLOT_DOUBLE_0: u32 = 1;
pub const SLOT_DOUBLE_LONG: u32 = 2;

/// Data and entry are retained by HostObject; callers check signature before entry.
#[repr(C)]
pub struct NativeSlot {
    pub entry: *const u8,
    pub data: *const (),
    pub signature: u32,
}
#[repr(C)]
pub struct ObjectHeader {
    /// Exact runtime IUnknown prefix; the typed table is a separate, versioned
    /// host ABI and never masquerades as an Automation/dual COM vtable.
    pub unknown: RawRuntimeIUnknown,
    pub iid: InterfaceId,
    pub slots: *const NativeSlot,
    pub slot_count: usize,
}

type Double0 = dyn Fn(&mut TypedContext) -> Result<f64, i32>;
type DoubleLong = dyn Fn(&mut TypedContext, i32) -> Result<f64, i32>;
enum Callback {
    Double0(Box<Double0>),
    DoubleLong(Box<DoubleLong>),
    Typed(value_slots::TypedSlot),
    Planned(Box<planned_slots::PlannedCallback>),
}
pub struct HostSlot {
    callback: Box<Callback>,
}
impl HostSlot {
    pub fn double(callback: impl Fn(&mut TypedContext) -> Result<f64, i32> + 'static) -> Self {
        Self {
            callback: Box::new(Callback::Double0(Box::new(callback))),
        }
    }
    pub fn double_long(
        callback: impl Fn(&mut TypedContext, i32) -> Result<f64, i32> + 'static,
    ) -> Self {
        Self {
            callback: Box::new(Callback::DoubleLong(Box::new(callback))),
        }
    }
    fn native(&self) -> NativeSlot {
        let (entry, signature) = match &*self.callback {
            Callback::Double0(_) => (double0 as *const u8, SLOT_DOUBLE_0),
            Callback::DoubleLong(_) => (double_long as *const u8, SLOT_DOUBLE_LONG),
            Callback::Typed(slot) => slot.native(),
            Callback::Planned(_) => (
                planned_slots::planned_entry as *const u8,
                planned_slots::SLOT_PLANNED_MEMBER,
            ),
        };
        NativeSlot {
            entry,
            data: &*self.callback as *const Callback as *const (),
            signature,
        }
    }
}

#[repr(C)]
struct ObjectStorage {
    header: ObjectHeader,
    _native: Rc<[NativeSlot]>,
    _owners: Vec<HostSlot>,
    payload: Option<Rc<dyn std::any::Any>>,
    identity: *const (),
    services: Option<Rc<dyn HostObjectServices>>,
}
type InterfaceResolver = dyn Fn(InterfaceId) -> Result<HostObject, i32>;
struct ResolverService(Box<InterfaceResolver>);
impl HostObjectServices for ResolverService {
    fn resolve_interface(&self, iid: InterfaceId) -> Result<HostObject, i32> {
        self.0(iid)
    }
}
/// Clone retains the table and callback captures. Drop releases them on this thread.
/// Objects contain no borrowed invocation pointers and can outlive any one call.
/// Captures stored by their own session should use weak session handles; the host
/// is responsible for breaking any strong reference cycle it creates.
#[derive(Clone)]
pub struct HostObject {
    storage: Rc<ObjectStorage>,
}
impl std::fmt::Debug for HostObject {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("HostObject")
            .field("iid", &self.header().iid)
            .field("pointer", &self.argument().as_ptr())
            .finish()
    }
}
impl HostObject {
    /// Retain this exact portable host table in the canonical IUnknown carrier.
    pub fn object_ref(&self) -> oxvba_runtime::object_ref::ObjectRef {
        // SAFETY: argument() preserves complete-allocation provenance and this
        // owner spans the AddRef; the carrier releases on the creating thread.
        unsafe {
            oxvba_runtime::object_ref::ObjectRef::from_raw_iunknown_addref(
                self.argument().as_ptr().cast_mut().cast(),
            )
        }
        .expect("host object is nonnull")
    }
    /// Recover only a table created by this ABI, never an arbitrary COM object.
    pub fn from_object_ref(object: &oxvba_runtime::object_ref::ObjectRef) -> Option<Self> {
        let raw = object.raw_iunknown();
        // SAFETY: ObjectRef owns a valid IUnknown prefix. The private vtable
        // identity proves the complete allocation came from this factory.
        if !std::ptr::eq(unsafe { (*raw).vtbl }, &UNKNOWN_VTABLE) {
            return None;
        }
        // SAFETY: object retains a factory-created allocation on this thread.
        Some(unsafe { Self::retain_header(raw.cast()) })
    }
    /// # Safety
    /// pointer is a live complete-allocation HostObject pointer on this thread.
    pub unsafe fn retain_header(pointer: *const ObjectHeader) -> Self {
        let raw = pointer.cast::<ObjectStorage>();
        // SAFETY: inherited factory allocation contract; mint one owned Rc.
        unsafe {
            Rc::increment_strong_count(raw);
        }
        Self {
            // SAFETY: consume exactly the strong reference just minted above.
            storage: unsafe { Rc::from_raw(raw) },
        }
    }
    pub fn into_header(self) -> *const ObjectHeader {
        Rc::into_raw(self.storage).cast()
    }
    /// # Safety
    /// pointer transfers exactly one native HostObject reference on this thread.
    pub unsafe fn release_header(pointer: *const ObjectHeader) {
        if !pointer.is_null() {
            // SAFETY: consume the one transferred reference; contain drop panics.
            unsafe {
                release(pointer.cast_mut().cast());
            }
        }
    }
    pub fn new(iid: InterfaceId, slots: Vec<HostSlot>) -> Self {
        Self::build(iid, slots, None, std::ptr::null(), None)
    }
    /// Attach a same-thread host owner without exposing native pointers in the
    /// portable header. Boundary crates may recover only their own payload type.
    pub fn new_with_payload(
        iid: InterfaceId,
        slots: Vec<HostSlot>,
        payload: Rc<dyn std::any::Any>,
    ) -> Self {
        Self::build(iid, slots, Some(payload), std::ptr::null(), None)
    }
    /// Construct a boundary table with an independently retained object identity.
    /// # Safety
    /// identity is a nonnull canonical identity pointer, unique among live objects
    /// on this thread (native COM uses QI(IID_IUnknown)). payload must retain that
    /// exact identity until this table and every clone have been released.
    pub unsafe fn new_with_payload_and_identity(
        iid: InterfaceId,
        slots: Vec<HostSlot>,
        payload: Rc<dyn std::any::Any>,
        identity: *const (),
    ) -> Self {
        assert!(!identity.is_null());
        Self::build(iid, slots, Some(payload), identity, None)
    }
    /// Retain a boundary-owned interface view resolver alongside the object.
    /// It selects only interfaces declared by the host's retained library.
    /// # Safety
    /// The identity contract is identical to new_with_payload_and_identity.
    /// Every resolved view must retain that same canonical identity.
    pub unsafe fn new_with_interface_resolver(
        iid: InterfaceId,
        slots: Vec<HostSlot>,
        payload: Rc<dyn std::any::Any>,
        identity: *const (),
        resolver: impl Fn(InterfaceId) -> Result<HostObject, i32> + 'static,
    ) -> Self {
        assert!(!identity.is_null());
        Self::build(
            iid,
            slots,
            Some(payload),
            identity,
            Some(Rc::new(ResolverService(Box::new(resolver)))),
        )
    }
    /// Construct an object with retained host interface/iteration services.
    /// # Safety
    /// The payload and every service-produced view retain the supplied canonical
    /// identity under the same contract as new_with_payload_and_identity.
    pub unsafe fn new_with_services(
        iid: InterfaceId,
        slots: Vec<HostSlot>,
        payload: Rc<dyn std::any::Any>,
        identity: *const (),
        services: Rc<dyn HostObjectServices>,
    ) -> Self {
        assert!(!identity.is_null());
        Self::build(iid, slots, Some(payload), identity, Some(services))
    }
    pub fn enumerator(&self, context: &mut TypedContext) -> Result<HostEnumerator, i32> {
        self.storage
            .services
            .as_ref()
            .ok_or(438)?
            .enumerator(context)
    }
    pub fn invoke_dynamic(
        &self,
        call: &DynamicMemberCall,
        arguments: &mut [oxvba_runtime::Variant],
        context: &mut TypedContext,
    ) -> Result<oxvba_runtime::Variant, crate::TypedFault> {
        match &self.storage.services {
            Some(services) => services.invoke_dynamic(call, arguments, context),
            None => Err(crate::TypedFault {
                number: 438,
                description: oxvba_runtime::default_error_message(438).into(),
                source: None,
                help_file: None,
                help_context: None,
            }),
        }
    }
    /// Prepare a nominal view without changing the source object's table.
    /// Portable tables without a resolver accept only their declared IID.
    pub fn resolve_interface(&self, iid: InterfaceId) -> Result<Self, i32> {
        if self.header().iid == iid {
            return Ok(self.clone());
        }
        let view = self
            .storage
            .services
            .as_ref()
            .ok_or(13)?
            .resolve_interface(iid)?;
        if view.header().iid != iid || !view.same_identity(self) {
            return Err(13);
        }
        Ok(view)
    }
    /// Compare retained identities without calling a member or QueryInterface.
    pub fn same_identity(&self, other: &Self) -> bool {
        self.identity() == other.identity()
    }
    fn identity(&self) -> *const () {
        if self.storage.identity.is_null() {
            Rc::as_ptr(&self.storage).cast()
        } else {
            self.storage.identity
        }
    }
    pub fn payload(&self) -> Option<&dyn std::any::Any> {
        self.storage.payload.as_deref()
    }
    fn build(
        iid: InterfaceId,
        slots: Vec<HostSlot>,
        payload: Option<Rc<dyn std::any::Any>>,
        identity: *const (),
        services: Option<Rc<dyn HostObjectServices>>,
    ) -> Self {
        // Derive raw table pointers only after immutable Rc allocation. Moving
        // a Box after taking as_ptr would invalidate that pointer's borrow tag.
        let native: Rc<[_]> = slots.iter().map(HostSlot::native).collect();
        Self {
            storage: Rc::new(ObjectStorage {
                header: ObjectHeader {
                    unknown: RawRuntimeIUnknown {
                        vtbl: &UNKNOWN_VTABLE,
                    },
                    iid,
                    slots: native.as_ptr(),
                    slot_count: native.len(),
                },
                _native: native,
                _owners: slots,
                payload,
                identity,
                services,
            }),
        }
    }
    pub fn header(&self) -> &ObjectHeader {
        &self.storage.header
    }
    pub fn argument(&self) -> ObjectArg<'_> {
        ObjectArg {
            // Preserve complete-allocation provenance for the IUnknown prefix.
            header: Rc::as_ptr(&self.storage).cast::<ObjectHeader>(),
            _owner: PhantomData,
        }
    }
    pub fn borrow_slot(&self) -> Slot<'_> {
        Slot::Object(self.argument())
    }
}

/// The native argument is a lifetime-bound header pointer, never a forgeable address.
#[repr(transparent)]
#[derive(Clone, Copy)]
pub struct ObjectArg<'a> {
    header: *const ObjectHeader,
    _owner: PhantomData<(&'a HostObject, Rc<()>)>,
}
impl ObjectArg<'_> {
    pub fn header(&self) -> &ObjectHeader {
        // SAFETY: private construction ties the live Rc owner to this borrow.
        unsafe { &*self.header }
    }
    pub fn as_ptr(self) -> *const ObjectHeader {
        self.header
    }
}

static UNKNOWN_VTABLE: RawRuntimeIUnknownVtbl = RawRuntimeIUnknownVtbl {
    query_interface: query_unknown,
    add_ref,
    release,
};
// SAFETY CONTRACT: this is the complete-allocation pointer minted by the
// HostObject factory; native references remain on its creating thread.
unsafe extern "C" fn query_unknown(
    this: *mut core::ffi::c_void,
    iid: RuntimeGuid,
    output: *mut *mut core::ffi::c_void,
) -> i32 {
    if output.is_null() {
        return 0x8000_4003u32 as i32;
    }
    // SAFETY: native caller supplies one aligned writable pointer cell.
    unsafe {
        output.write(std::ptr::null_mut());
    }
    if this.is_null() || iid != RUNTIME_GUID_IUNKNOWN {
        return 0x8000_4002u32 as i32;
    }
    // SAFETY: inherited retained same-thread allocation contract.
    unsafe {
        add_ref(this);
        output.write(this);
    }
    0
}
// SAFETY CONTRACT: this is a live complete ObjectStorage allocation from Rc;
// the native caller owns a reference and all operations stay on this thread.
unsafe extern "C" fn add_ref(this: *mut core::ffi::c_void) -> u32 {
    let pointer = this.cast::<ObjectStorage>();
    // SAFETY: pointer was obtained from Rc::as_ptr and a strong owner is live.
    unsafe {
        Rc::increment_strong_count(pointer);
    }
    // SAFETY: inspect the newly retained reference without consuming it.
    let owner = std::mem::ManuallyDrop::new(unsafe { Rc::from_raw(pointer) });
    Rc::strong_count(&owner).min(u32::MAX as usize) as u32
}
// SAFETY CONTRACT: consume exactly one retained native reference, on the owner
// thread, using the complete pointer supplied by this object's IUnknown.
unsafe extern "C" fn release(this: *mut core::ffi::c_void) -> u32 {
    // SAFETY: caller transfers one native strong reference back into Rc.
    let owner = unsafe { Rc::from_raw(this.cast::<ObjectStorage>()) };
    let remaining = Rc::strong_count(&owner) - 1;
    // No host destructor panic may unwind through the native ABI.
    if let Err(payload) = catch_unwind(AssertUnwindSafe(|| drop(owner))) {
        discard_panic(payload);
    }
    remaining.min(u32::MAX as usize) as u32
}

/// # Safety
/// Data is one live Callback of the advertised signature; context and output
/// are nonaliasing live objects, and the retained HostObject spans the call.
unsafe extern "C" fn double0(data: *const (), context: *mut TypedContext, output: *mut f64) -> i32 {
    // SAFETY: table construction paired these exact pointers and signature.
    let Callback::Double0(call) = (unsafe { &*data.cast::<Callback>() }) else {
        return 5;
    };
    // SAFETY: entry contract provides a unique context for synchronous reborrowing.
    let result = catch_unwind(AssertUnwindSafe(|| call(unsafe { &mut *context })));
    finish(result, output)
}
/// # Safety
/// Same contract as double0, with one native Long argument.
unsafe extern "C" fn double_long(
    data: *const (),
    context: *mut TypedContext,
    arg: i32,
    output: *mut f64,
) -> i32 {
    // SAFETY: table construction paired these exact pointers and signature.
    let Callback::DoubleLong(call) = (unsafe { &*data.cast::<Callback>() }) else {
        return 5;
    };
    // SAFETY: entry contract provides unique context for synchronous reborrowing.
    let result = catch_unwind(AssertUnwindSafe(|| call(unsafe { &mut *context }, arg)));
    finish(result, output)
}
fn finish(
    result: Result<Result<f64, i32>, Box<dyn std::any::Any + Send>>,
    output: *mut f64,
) -> i32 {
    match result {
        Ok(Ok(value)) if value.is_finite() => {
            // SAFETY: private callers inherit the live aligned output contract.
            unsafe {
                output.write(value);
            }
            0
        }
        Ok(Ok(_)) => 6,
        Ok(Err(status)) if status == -1 || status > 0 => status,
        Ok(Err(_)) => 5,
        Err(payload) => {
            discard_panic(payload);
            5
        }
    }
}
fn discard_panic(payload: Box<dyn std::any::Any + Send>) {
    // A custom panic payload can itself panic when dropped. Free ordinary
    // payloads and keep even that secondary unwind inside the native boundary.
    if let Err(secondary) = catch_unwind(AssertUnwindSafe(|| drop(payload))) {
        std::mem::forget(secondary);
    }
}

/// Checked interpreter transport; generated JIT code emits the equivalent loads
/// and a call_indirect of the exact native signature instead of calling this helper.
/// # Safety
/// Header is null or a live retained HostObject header. Context is unique. The
/// object borrow and provider's admission proof must span the complete call.
pub unsafe fn invoke_double_slot(
    header: *const ObjectHeader,
    iid: InterfaceId,
    slot: usize,
    argument: Option<i32>,
    context: &mut TypedContext,
) -> Result<ScalarBits, i32> {
    // SAFETY: caller proves live header or null.
    let header = unsafe { header.as_ref() }.ok_or(91)?;
    if header.iid != iid || slot >= header.slot_count {
        return Err(438);
    }
    // SAFETY: HostObject retains slot_count initialized native slots.
    let native = unsafe { &*header.slots.add(slot) };
    let signature = if argument.is_some() {
        SLOT_DOUBLE_LONG
    } else {
        SLOT_DOUBLE_0
    };
    if native.signature != signature {
        return Err(13);
    }
    let mut output = 0.0;
    let status = if let Some(arg) = argument {
        // SAFETY: checked signature identifies this C ABI; data/code retained.
        let entry: unsafe extern "C" fn(*const (), *mut TypedContext, i32, *mut f64) -> i32 =
            unsafe { std::mem::transmute(native.entry) };
        // SAFETY: all borrows survive entry; only context/output are writable.
        unsafe { entry(native.data, context, arg, &mut output) }
    } else {
        // SAFETY: checked signature identifies this C ABI; data/code retained.
        let entry: unsafe extern "C" fn(*const (), *mut TypedContext, *mut f64) -> i32 =
            unsafe { std::mem::transmute(native.entry) };
        // SAFETY: all borrows survive entry; only context/output are writable.
        unsafe { entry(native.data, context, &mut output) }
    };
    if status == 0 {
        Ok(ScalarBits::double(output))
    } else {
        Err(status)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn native_unknown_prefix_retains_the_exact_object_until_last_release() {
        assert_eq!(std::mem::offset_of!(ObjectHeader, unknown), 0);
        assert_eq!(std::mem::offset_of!(ObjectStorage, header), 0);
        let capture = Rc::new(());
        let weak = Rc::downgrade(&capture);
        let object = HostObject::new(
            [1; 16],
            vec![HostSlot::double(move |_| {
                let _keep = &capture;
                Ok(41.0)
            })],
        );
        let raw = object.argument().as_ptr();
        // SAFETY: argument preserves full Rc allocation provenance and the
        // creating-thread owner remains live while the native reference is taken.
        let table = unsafe { &*(*raw).unknown.vtbl };
        // SAFETY: one live same-thread retained object, native call retains one.
        assert_eq!(unsafe { (table.add_ref)(raw.cast_mut().cast()) }, 2);
        drop(object);
        assert!(weak.upgrade().is_some());
        let mut identity = std::ptr::null_mut();
        // SAFETY: the explicit native reference owns the live full allocation;
        // identity is a nonaliasing writable pointer cell.
        assert_eq!(
            // SAFETY: retained native owner and initialized output cell span QI.
            unsafe {
                (table.query_interface)(raw.cast_mut().cast(), RUNTIME_GUID_IUNKNOWN, &mut identity)
            },
            0
        );
        assert_eq!(identity, raw.cast_mut().cast());
        // SAFETY: release exactly the QI reference, leaving the explicit owner.
        assert_eq!(unsafe { (table.release)(identity) }, 1);
        identity = raw.cast_mut().cast();
        // SAFETY: live explicit native owner and writable pointer output.
        assert_ne!(
            // SAFETY: explicit native owner is live; output cell is writable.
            unsafe {
                (table.query_interface)(
                    raw.cast_mut().cast(),
                    RuntimeGuid::new(9, 0, 0, [0; 8]),
                    &mut identity,
                )
            },
            0
        );
        assert!(identity.is_null());
        // SAFETY: consume the last explicit native reference; raw is unused after.
        assert_eq!(unsafe { (table.release)(raw.cast_mut().cast()) }, 0);
        assert!(weak.upgrade().is_none());
    }
}
