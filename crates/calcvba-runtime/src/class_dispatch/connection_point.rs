//! Portable COM connection points (`IConnectionPointContainer`, `IConnectionPoint`,
//! `IEnumConnectionPoints`, `IEnumConnections`) directly on `CompatObjectBase` (dc-yqq0.27).
//!
//! Any VBA project class that declares at least one externally visible `Event` exposes
//! `IConnectionPointContainer` via `QueryInterface` on `&primary` (sharing one COM identity and
//! one atomic refcount) and delivers `RaiseEvent` invocations to all advised COM `IDispatch`
//! sinks synchronously on the owning VBA thread.

use core::ffi::c_void;
use core::sync::atomic::{AtomicU32, Ordering};

use calcvba_automation::com::{
    CONNECT_E_ADVISELIMIT, CONNECT_E_CANNOTCONNECT, CONNECT_E_NOCONNECTION, ConnectData,
    DISPATCH_METHOD, DispParams, E_NOINTERFACE, E_POINTER, ExcepInfo, Guid, HResult,
    IConnectionPointContainerVtbl, IConnectionPointVtbl, IDispatchVtbl, IEnumConnectionPointsVtbl,
    IEnumConnectionsVtbl, IID_ICONNECTIONPOINT, IID_IDISPATCH, IID_IENUMCONNECTIONPOINTS,
    IID_IENUMCONNECTIONS, IID_IUNKNOWN, IID_NULL, IUnknown, IUnknownVtbl, RPC_E_WRONG_THREAD,
    S_FALSE, S_OK,
};
use calcvba_automation::oleaut;

use crate::object_ref::{CompatObjectBase, compat_add_ref, compat_query_interface, compat_release};
use crate::{Variant, VariantCore};

use super::current_thread_id;

/// One advised COM event sink retained by a `CompatObjectBase`'s connection point.
pub(crate) struct ConnectionPointSink {
    pub(crate) cookie: u32,
    /// Retained COM `IDispatch*` sink pointer (released on `Unadvise` or object teardown).
    pub(crate) dispatch: *mut c_void,
}

/// Embedded `IConnectionPointContainer` COM face on `CompatObjectBase`.
///
/// Shares the owning `CompatObjectBase`'s single COM identity (`QI(IUnknown)` returns `&primary`)
/// and single atomic reference count.
#[repr(C)]
pub(crate) struct ConnectionPointContainerFace {
    pub(crate) vtbl: *const IConnectionPointContainerVtbl,
    pub(crate) owner: *mut CompatObjectBase,
}

impl ConnectionPointContainerFace {
    pub(crate) const fn new_unlinked() -> Self {
        Self {
            vtbl: &CONNECTION_POINT_CONTAINER_VTBL,
            owner: core::ptr::null_mut(),
        }
    }
}

/// Embedded `IConnectionPoint` sub-object on `CompatObjectBase`.
///
/// Per OCIDL, `IConnectionPoint` is obtained via `FindConnectionPoint`/`EnumConnectionPoints`
/// (not `QI` on the source object) and answers `QI(IUnknown)`/`QI(IConnectionPoint)` with itself
/// while sharing `owner`'s reference count so an outstanding `IConnectionPoint*` keeps the source
/// object alive.
#[repr(C)]
pub(crate) struct ConnectionPointFace {
    pub(crate) vtbl: *const IConnectionPointVtbl,
    pub(crate) owner: *mut CompatObjectBase,
}

impl ConnectionPointFace {
    pub(crate) const fn new_unlinked() -> Self {
        Self {
            vtbl: &CONNECTION_POINT_VTBL,
            owner: core::ptr::null_mut(),
        }
    }
}

pub(crate) static CONNECTION_POINT_CONTAINER_VTBL: IConnectionPointContainerVtbl =
    IConnectionPointContainerVtbl {
        unknown: IUnknownVtbl {
            query_interface: cpc_query_interface,
            add_ref: cpc_add_ref,
            release: cpc_release,
        },
        enum_connection_points: cpc_enum_connection_points,
        find_connection_point: cpc_find_connection_point,
    };

static CONNECTION_POINT_VTBL: IConnectionPointVtbl = IConnectionPointVtbl {
    unknown: IUnknownVtbl {
        query_interface: cp_query_interface,
        add_ref: cp_add_ref,
        release: cp_release,
    },
    get_connection_interface: cp_get_connection_interface,
    get_connection_point_container: cp_get_connection_point_container,
    advise: cp_advise,
    unadvise: cp_unadvise,
    enum_connections: cp_enum_connections,
};

pub(crate) fn is_cpc_query_interface(
    qi: unsafe extern "system" fn(*mut c_void, *const Guid, *mut *mut c_void) -> HResult,
) -> bool {
    core::ptr::fn_addr_eq(qi, CONNECTION_POINT_CONTAINER_VTBL.unknown.query_interface)
}

unsafe fn add_ref_com(ptr: *mut c_void) {
    if ptr.is_null() {
        return;
    }
    // SAFETY: `ptr` is a live COM interface pointer whose first word is its `IUnknownVtbl`.
    unsafe {
        let vtbl = (*ptr.cast::<IUnknown>()).vtbl;
        ((*vtbl).add_ref)(ptr);
    }
}

unsafe fn release_com(ptr: *mut c_void) {
    if ptr.is_null() {
        return;
    }
    // SAFETY: `ptr` is a live COM interface pointer owned by the caller.
    unsafe {
        let vtbl = (*ptr.cast::<IUnknown>()).vtbl;
        ((*vtbl).release)(ptr);
    }
}

unsafe fn query_interface_com(ptr: *mut c_void, iid: &Guid, out: *mut *mut c_void) -> HResult {
    if ptr.is_null() || out.is_null() {
        return E_POINTER;
    }
    // SAFETY: `ptr` is a live COM interface pointer and `out` is a writable out-slot.
    unsafe {
        *out = core::ptr::null_mut();
        let vtbl = (*ptr.cast::<IUnknown>()).vtbl;
        ((*vtbl).query_interface)(ptr, iid, out)
    }
}

/// Release every advised sink on `owner` during `CompatObjectBase` teardown.
///
/// # Safety
/// Caller holds sole final ownership of `owner` during teardown on the owner thread.
pub(crate) unsafe fn release_all_event_sinks(owner: *mut CompatObjectBase) {
    // SAFETY: sole teardown access to `owner.event_sinks`.
    let sinks = unsafe { core::mem::take(&mut *(*owner).event_sinks.borrow_mut()) };
    for sink in sinks {
        // SAFETY: each sink holds one retained COM reference acquired in `Advise`.
        unsafe {
            release_com(sink.dispatch);
        }
    }
}

// ── IConnectionPointContainer ────────────────────────────────────────────────

unsafe extern "system" fn cpc_query_interface(
    this: *mut c_void,
    riid: *const Guid,
    ppv: *mut *mut c_void,
) -> HResult {
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `ConnectionPointContainerFace` embedded in `CompatObjectBase`.
    let owner = unsafe { (*this.cast::<ConnectionPointContainerFace>()).owner };
    // SAFETY: delegates to `owner`'s primary `QueryInterface` (single COM identity).
    unsafe { compat_query_interface(owner.cast(), riid, ppv) }
}

unsafe extern "system" fn cpc_add_ref(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `ConnectionPointContainerFace`.
    let owner = unsafe { (*this.cast::<ConnectionPointContainerFace>()).owner };
    // SAFETY: shares `owner`'s refcount.
    unsafe { compat_add_ref(owner.cast()) }
}

unsafe extern "system" fn cpc_release(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `ConnectionPointContainerFace`.
    let owner = unsafe { (*this.cast::<ConnectionPointContainerFace>()).owner };
    // SAFETY: shares `owner`'s refcount.
    unsafe { compat_release(owner.cast()) }
}

unsafe extern "system" fn cpc_enum_connection_points(
    this: *mut c_void,
    pp_enum: *mut *mut c_void,
) -> HResult {
    if pp_enum.is_null() {
        return E_POINTER;
    }
    // SAFETY: `pp_enum` is non-null writable out-slot.
    unsafe {
        *pp_enum = core::ptr::null_mut();
    }
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `ConnectionPointContainerFace`.
    let owner = unsafe { (*this.cast::<ConnectionPointContainerFace>()).owner };
    // SAFETY: `owner` is live; `allocate_enum_connection_points` takes one ref on `owner`.
    unsafe {
        *pp_enum = allocate_enum_connection_points(owner, 0);
    }
    S_OK
}

unsafe extern "system" fn cpc_find_connection_point(
    this: *mut c_void,
    riid: *const Guid,
    pp_cp: *mut *mut c_void,
) -> HResult {
    if pp_cp.is_null() {
        return E_POINTER;
    }
    // SAFETY: `pp_cp` is non-null writable out-slot.
    unsafe {
        *pp_cp = core::ptr::null_mut();
    }
    if this.is_null() || riid.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `ConnectionPointContainerFace` and `riid` is readable.
    let (owner, requested, source_iid) = unsafe {
        let owner = (*this.cast::<ConnectionPointContainerFace>()).owner;
        (
            owner,
            riid.read(),
            (*owner).class_descriptor.source_interface_iid,
        )
    };
    if source_iid == Some(requested) {
        // SAFETY: `owner` is live; retain one ref for the returned `IConnectionPoint*`.
        unsafe {
            compat_add_ref(owner.cast());
            *pp_cp = (&raw mut (*owner).cp).cast();
        }
        S_OK
    } else {
        CONNECT_E_NOCONNECTION
    }
}

// ── IConnectionPoint ─────────────────────────────────────────────────────────

unsafe extern "system" fn cp_query_interface(
    this: *mut c_void,
    riid: *const Guid,
    ppv: *mut *mut c_void,
) -> HResult {
    if ppv.is_null() {
        return E_POINTER;
    }
    // SAFETY: `ppv` is non-null writable out-slot.
    unsafe {
        *ppv = core::ptr::null_mut();
    }
    if this.is_null() || riid.is_null() {
        return E_POINTER;
    }
    // SAFETY: `riid` is a non-null readable `Guid`.
    let iid = unsafe { riid.read() };
    if iid == IID_IUNKNOWN || iid == IID_ICONNECTIONPOINT {
        // SAFETY: `this` is a live `ConnectionPointFace`; retain one ref on `owner`.
        unsafe {
            cp_add_ref(this);
            *ppv = this;
        }
        S_OK
    } else {
        E_NOINTERFACE
    }
}

unsafe extern "system" fn cp_add_ref(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `ConnectionPointFace`.
    let owner = unsafe { (*this.cast::<ConnectionPointFace>()).owner };
    // SAFETY: retains `owner`.
    unsafe { compat_add_ref(owner.cast()) }
}

unsafe extern "system" fn cp_release(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `ConnectionPointFace`.
    let owner = unsafe { (*this.cast::<ConnectionPointFace>()).owner };
    // SAFETY: releases `owner`.
    unsafe { compat_release(owner.cast()) }
}

unsafe extern "system" fn cp_get_connection_interface(
    this: *mut c_void,
    p_iid: *mut Guid,
) -> HResult {
    if this.is_null() || p_iid.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `ConnectionPointFace`.
    let source_iid = unsafe {
        let owner = (*this.cast::<ConnectionPointFace>()).owner;
        (*owner).class_descriptor.source_interface_iid
    };
    let Some(iid) = source_iid else {
        return CONNECT_E_NOCONNECTION;
    };
    // SAFETY: `p_iid` is a writable out-slot.
    unsafe {
        p_iid.write(iid);
    }
    S_OK
}

unsafe extern "system" fn cp_get_connection_point_container(
    this: *mut c_void,
    pp_cpc: *mut *mut c_void,
) -> HResult {
    if pp_cpc.is_null() {
        return E_POINTER;
    }
    // SAFETY: `pp_cpc` is non-null writable out-slot.
    unsafe {
        *pp_cpc = core::ptr::null_mut();
    }
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `ConnectionPointFace`; retain one ref on `owner`.
    unsafe {
        let owner = (*this.cast::<ConnectionPointFace>()).owner;
        compat_add_ref(owner.cast());
        *pp_cpc = (&raw mut (*owner).cpc).cast();
    }
    S_OK
}

unsafe extern "system" fn cp_advise(
    this: *mut c_void,
    punk_sink: *mut c_void,
    pdw_cookie: *mut u32,
) -> HResult {
    if pdw_cookie.is_null() {
        return E_POINTER;
    }
    // SAFETY: `pdw_cookie` is non-null writable out-slot.
    unsafe {
        *pdw_cookie = 0;
    }
    if this.is_null() || punk_sink.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `ConnectionPointFace`.
    let (owner, owner_thread, source_iid) = unsafe {
        let owner = (*this.cast::<ConnectionPointFace>()).owner;
        (
            owner,
            (*owner).owner_thread,
            (*owner).class_descriptor.source_interface_iid,
        )
    };
    if current_thread_id() != owner_thread {
        return RPC_E_WRONG_THREAD;
    }
    let Some(source_iid) = source_iid else {
        return CONNECT_E_CANNOTCONNECT;
    };

    // VBA/OLE dispinterface event sinks may answer the source dispinterface IID directly or
    // `IID_IDISPATCH`. Try `source_iid` first, then fall back to `IID_IDISPATCH`.
    let mut dispatch: *mut c_void = core::ptr::null_mut();
    // SAFETY: `punk_sink` is a non-null COM `IUnknown*` supplied by the caller.
    let hr = unsafe { query_interface_com(punk_sink, &source_iid, &mut dispatch) };
    if hr < 0 || dispatch.is_null() {
        // SAFETY: same caller-supplied `punk_sink`.
        let hr_disp = unsafe { query_interface_com(punk_sink, &IID_IDISPATCH, &mut dispatch) };
        if hr_disp < 0 || dispatch.is_null() {
            return CONNECT_E_CANNOTCONNECT;
        }
    }

    // Allocate a non-zero cookie not currently in use.
    // SAFETY: `owner` is live on its owning thread.
    let mut sinks = unsafe { (*owner).event_sinks.borrow_mut() };
    let mut cookie = 0u32;
    for _ in 0..=sinks.len() + 1 {
        // SAFETY: `next_event_cookie` is an atomic field on live `owner`.
        let candidate = unsafe { (*owner).next_event_cookie.fetch_add(1, Ordering::Relaxed) };
        if candidate != 0 && !sinks.iter().any(|sink| sink.cookie == candidate) {
            cookie = candidate;
            break;
        }
    }
    if cookie == 0 {
        drop(sinks);
        // SAFETY: release the sink reference acquired above.
        unsafe {
            release_com(dispatch);
        }
        return CONNECT_E_ADVISELIMIT;
    }
    sinks.push(ConnectionPointSink { cookie, dispatch });
    drop(sinks);
    // SAFETY: `pdw_cookie` is non-null writable out-slot.
    unsafe {
        *pdw_cookie = cookie;
    }
    S_OK
}

unsafe extern "system" fn cp_unadvise(this: *mut c_void, dw_cookie: u32) -> HResult {
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `ConnectionPointFace`.
    let (owner, owner_thread) = unsafe {
        let owner = (*this.cast::<ConnectionPointFace>()).owner;
        (owner, (*owner).owner_thread)
    };
    if current_thread_id() != owner_thread {
        return RPC_E_WRONG_THREAD;
    }
    // Remove while holding the `RefCell` borrow, then release outside the borrow so a sink's
    // final `Release` can safely re-enter.
    let removed = {
        // SAFETY: `owner` is live on its owning thread.
        let mut sinks = unsafe { (*owner).event_sinks.borrow_mut() };
        sinks
            .iter()
            .position(|sink| sink.cookie == dw_cookie)
            .map(|idx| sinks.remove(idx))
    };
    let Some(sink) = removed else {
        return CONNECT_E_NOCONNECTION;
    };
    // SAFETY: releases the COM reference acquired in `Advise`.
    unsafe {
        release_com(sink.dispatch);
    }
    S_OK
}

unsafe extern "system" fn cp_enum_connections(
    this: *mut c_void,
    pp_enum: *mut *mut c_void,
) -> HResult {
    if pp_enum.is_null() {
        return E_POINTER;
    }
    // SAFETY: `pp_enum` is non-null writable out-slot.
    unsafe {
        *pp_enum = core::ptr::null_mut();
    }
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `ConnectionPointFace`.
    let (owner, owner_thread) = unsafe {
        let owner = (*this.cast::<ConnectionPointFace>()).owner;
        (owner, (*owner).owner_thread)
    };
    if current_thread_id() != owner_thread {
        return RPC_E_WRONG_THREAD;
    }
    // SAFETY: `owner` is live on its owning thread; snapshots sinks with `AddRef`.
    let entries: Vec<ConnectionPointSink> = unsafe {
        (*owner)
            .event_sinks
            .borrow()
            .iter()
            .map(|sink| {
                add_ref_com(sink.dispatch);
                ConnectionPointSink {
                    cookie: sink.cookie,
                    dispatch: sink.dispatch,
                }
            })
            .collect()
    };
    // SAFETY: `pp_enum` is non-null writable out-slot.
    unsafe {
        *pp_enum = allocate_enum_connections(entries, 0);
    }
    S_OK
}

// ── IEnumConnectionPoints ────────────────────────────────────────────────────

#[repr(C)]
struct EnumConnectionPoints {
    vtbl: *const IEnumConnectionPointsVtbl,
    ref_count: AtomicU32,
    owner: *mut CompatObjectBase,
    index: usize,
}

static ENUM_CONNECTION_POINTS_VTBL: IEnumConnectionPointsVtbl = IEnumConnectionPointsVtbl {
    unknown: IUnknownVtbl {
        query_interface: enum_cp_query_interface,
        add_ref: enum_cp_add_ref,
        release: enum_cp_release,
    },
    next: enum_cp_next,
    skip: enum_cp_skip,
    reset: enum_cp_reset,
    clone: enum_cp_clone,
};

unsafe fn allocate_enum_connection_points(
    owner: *mut CompatObjectBase,
    index: usize,
) -> *mut c_void {
    // SAFETY: `owner` is live; retain one reference held by the enumerator until release.
    unsafe {
        compat_add_ref(owner.cast());
    }
    Box::into_raw(Box::new(EnumConnectionPoints {
        vtbl: &ENUM_CONNECTION_POINTS_VTBL,
        ref_count: AtomicU32::new(1),
        owner,
        index,
    }))
    .cast()
}

unsafe extern "system" fn enum_cp_query_interface(
    this: *mut c_void,
    riid: *const Guid,
    ppv: *mut *mut c_void,
) -> HResult {
    if ppv.is_null() {
        return E_POINTER;
    }
    // SAFETY: `ppv` is non-null writable out-slot.
    unsafe {
        *ppv = core::ptr::null_mut();
    }
    if this.is_null() || riid.is_null() {
        return E_POINTER;
    }
    // SAFETY: `riid` is readable.
    let iid = unsafe { riid.read() };
    if iid == IID_IUNKNOWN || iid == IID_IENUMCONNECTIONPOINTS {
        // SAFETY: `this` is a live `EnumConnectionPoints`.
        unsafe {
            enum_cp_add_ref(this);
            *ppv = this;
        }
        S_OK
    } else {
        E_NOINTERFACE
    }
}

unsafe extern "system" fn enum_cp_add_ref(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `EnumConnectionPoints`.
    unsafe {
        (*this.cast::<EnumConnectionPoints>())
            .ref_count
            .fetch_add(1, Ordering::AcqRel)
            + 1
    }
}

unsafe extern "system" fn enum_cp_release(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `EnumConnectionPoints`.
    let previous = unsafe {
        (*this.cast::<EnumConnectionPoints>())
            .ref_count
            .fetch_sub(1, Ordering::AcqRel)
    };
    let remaining = previous.saturating_sub(1);
    if remaining == 0 {
        // SAFETY: sole final reference; release retained `owner` and free enumerator.
        unsafe {
            let boxed = Box::from_raw(this.cast::<EnumConnectionPoints>());
            let owner = boxed.owner;
            drop(boxed);
            if !owner.is_null() {
                compat_release(owner.cast());
            }
        }
    }
    remaining
}

unsafe extern "system" fn enum_cp_next(
    this: *mut c_void,
    c_connections: u32,
    pp_cp: *mut *mut c_void,
    pc_fetched: *mut u32,
) -> HResult {
    if this.is_null() || pp_cp.is_null() || (pc_fetched.is_null() && c_connections != 1) {
        return E_POINTER;
    }
    for i in 0..c_connections as usize {
        // SAFETY: `pp_cp` has `c_connections` writable slots.
        unsafe {
            *pp_cp.add(i) = core::ptr::null_mut();
        }
    }
    if !pc_fetched.is_null() {
        // SAFETY: `pc_fetched` is non-null writable out-slot.
        unsafe {
            *pc_fetched = 0;
        }
    }
    // SAFETY: `this` is a live `EnumConnectionPoints`.
    let enumerator = unsafe { &mut *this.cast::<EnumConnectionPoints>() };
    let mut copied = 0u32;
    while copied < c_connections && enumerator.index == 0 {
        enumerator.index = 1;
        // SAFETY: `enumerator.owner` is retained by `enumerator`.
        let has_source = unsafe {
            (*enumerator.owner)
                .class_descriptor
                .source_interface_iid
                .is_some()
        };
        if has_source {
            // SAFETY: retain one reference on `owner` for the returned `IConnectionPoint*`.
            unsafe {
                compat_add_ref(enumerator.owner.cast());
                *pp_cp.add(copied as usize) = (&raw mut (*enumerator.owner).cp).cast();
            }
            copied += 1;
        }
    }
    if !pc_fetched.is_null() {
        // SAFETY: `pc_fetched` is non-null writable out-slot.
        unsafe {
            *pc_fetched = copied;
        }
    }
    if copied == c_connections {
        S_OK
    } else {
        S_FALSE
    }
}

unsafe extern "system" fn enum_cp_skip(this: *mut c_void, c_connections: u32) -> HResult {
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `EnumConnectionPoints`.
    let enumerator = unsafe { &mut *this.cast::<EnumConnectionPoints>() };
    // SAFETY: `enumerator.owner` is retained by `enumerator`.
    let has_remaining = enumerator.index == 0
        && unsafe {
            (*enumerator.owner)
                .class_descriptor
                .source_interface_iid
                .is_some()
        };
    let remaining = usize::from(has_remaining);
    let skipped = remaining.min(c_connections as usize);
    enumerator.index += skipped;
    if skipped == c_connections as usize {
        S_OK
    } else {
        S_FALSE
    }
}

unsafe extern "system" fn enum_cp_reset(this: *mut c_void) -> HResult {
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `EnumConnectionPoints`.
    unsafe {
        (*this.cast::<EnumConnectionPoints>()).index = 0;
    }
    S_OK
}

unsafe extern "system" fn enum_cp_clone(this: *mut c_void, pp_enum: *mut *mut c_void) -> HResult {
    if pp_enum.is_null() {
        return E_POINTER;
    }
    // SAFETY: `pp_enum` is non-null writable out-slot.
    unsafe {
        *pp_enum = core::ptr::null_mut();
    }
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `EnumConnectionPoints`.
    unsafe {
        let enumerator = &*this.cast::<EnumConnectionPoints>();
        *pp_enum = allocate_enum_connection_points(enumerator.owner, enumerator.index);
    }
    S_OK
}

// ── IEnumConnections ─────────────────────────────────────────────────────────

#[repr(C)]
struct EnumConnections {
    vtbl: *const IEnumConnectionsVtbl,
    ref_count: AtomicU32,
    entries: Vec<ConnectionPointSink>,
    index: usize,
}

static ENUM_CONNECTIONS_VTBL: IEnumConnectionsVtbl = IEnumConnectionsVtbl {
    unknown: IUnknownVtbl {
        query_interface: enum_connections_query_interface,
        add_ref: enum_connections_add_ref,
        release: enum_connections_release,
    },
    next: enum_connections_next,
    skip: enum_connections_skip,
    reset: enum_connections_reset,
    clone: enum_connections_clone,
};

fn allocate_enum_connections(entries: Vec<ConnectionPointSink>, index: usize) -> *mut c_void {
    Box::into_raw(Box::new(EnumConnections {
        vtbl: &ENUM_CONNECTIONS_VTBL,
        ref_count: AtomicU32::new(1),
        entries,
        index,
    }))
    .cast()
}

unsafe extern "system" fn enum_connections_query_interface(
    this: *mut c_void,
    riid: *const Guid,
    ppv: *mut *mut c_void,
) -> HResult {
    if ppv.is_null() {
        return E_POINTER;
    }
    // SAFETY: `ppv` is non-null writable out-slot.
    unsafe {
        *ppv = core::ptr::null_mut();
    }
    if this.is_null() || riid.is_null() {
        return E_POINTER;
    }
    // SAFETY: `riid` is readable.
    let iid = unsafe { riid.read() };
    if iid == IID_IUNKNOWN || iid == IID_IENUMCONNECTIONS {
        // SAFETY: `this` is a live `EnumConnections`.
        unsafe {
            enum_connections_add_ref(this);
            *ppv = this;
        }
        S_OK
    } else {
        E_NOINTERFACE
    }
}

unsafe extern "system" fn enum_connections_add_ref(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `EnumConnections`.
    unsafe {
        (*this.cast::<EnumConnections>())
            .ref_count
            .fetch_add(1, Ordering::AcqRel)
            + 1
    }
}

unsafe extern "system" fn enum_connections_release(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    // SAFETY: `this` is a live `EnumConnections`.
    let previous = unsafe {
        (*this.cast::<EnumConnections>())
            .ref_count
            .fetch_sub(1, Ordering::AcqRel)
    };
    let remaining = previous.saturating_sub(1);
    if remaining == 0 {
        // SAFETY: sole final reference; release every retained sink and free enumerator.
        unsafe {
            let mut boxed = Box::from_raw(this.cast::<EnumConnections>());
            for entry in boxed.entries.drain(..) {
                release_com(entry.dispatch);
            }
            drop(boxed);
        }
    }
    remaining
}

unsafe extern "system" fn enum_connections_next(
    this: *mut c_void,
    c_connections: u32,
    rgcd: *mut ConnectData,
    pc_fetched: *mut u32,
) -> HResult {
    if this.is_null() || rgcd.is_null() || (pc_fetched.is_null() && c_connections != 1) {
        return E_POINTER;
    }
    for i in 0..c_connections as usize {
        // SAFETY: `rgcd` has `c_connections` writable slots.
        unsafe {
            *rgcd.add(i) = ConnectData::default();
        }
    }
    if !pc_fetched.is_null() {
        // SAFETY: `pc_fetched` is non-null writable out-slot.
        unsafe {
            *pc_fetched = 0;
        }
    }
    // SAFETY: `this` is a live `EnumConnections`.
    let enumerator = unsafe { &mut *this.cast::<EnumConnections>() };
    let mut copied = 0u32;
    while copied < c_connections && enumerator.index < enumerator.entries.len() {
        let entry = &enumerator.entries[enumerator.index];
        enumerator.index += 1;
        // SAFETY: `entry.dispatch` is retained by `enumerator`; `AddRef` for the caller.
        unsafe {
            add_ref_com(entry.dispatch);
            *rgcd.add(copied as usize) = ConnectData {
                punk: entry.dispatch,
                dw_cookie: entry.cookie,
            };
        }
        copied += 1;
    }
    if !pc_fetched.is_null() {
        // SAFETY: `pc_fetched` is non-null writable out-slot.
        unsafe {
            *pc_fetched = copied;
        }
    }
    if copied == c_connections {
        S_OK
    } else {
        S_FALSE
    }
}

unsafe extern "system" fn enum_connections_skip(this: *mut c_void, c_connections: u32) -> HResult {
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `EnumConnections`.
    let enumerator = unsafe { &mut *this.cast::<EnumConnections>() };
    let remaining = enumerator.entries.len().saturating_sub(enumerator.index);
    let skipped = remaining.min(c_connections as usize);
    enumerator.index += skipped;
    if skipped == c_connections as usize {
        S_OK
    } else {
        S_FALSE
    }
}

unsafe extern "system" fn enum_connections_reset(this: *mut c_void) -> HResult {
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `EnumConnections`.
    unsafe {
        (*this.cast::<EnumConnections>()).index = 0;
    }
    S_OK
}

unsafe extern "system" fn enum_connections_clone(
    this: *mut c_void,
    pp_enum: *mut *mut c_void,
) -> HResult {
    if pp_enum.is_null() {
        return E_POINTER;
    }
    // SAFETY: `pp_enum` is non-null writable out-slot.
    unsafe {
        *pp_enum = core::ptr::null_mut();
    }
    if this.is_null() {
        return E_POINTER;
    }
    // SAFETY: `this` is a live `EnumConnections`; clone retained entries with `AddRef`.
    unsafe {
        let enumerator = &*this.cast::<EnumConnections>();
        let entries = enumerator
            .entries
            .iter()
            .map(|entry| {
                add_ref_com(entry.dispatch);
                ConnectionPointSink {
                    cookie: entry.cookie,
                    dispatch: entry.dispatch,
                }
            })
            .collect();
        *pp_enum = allocate_enum_connections(entries, enumerator.index);
    }
    S_OK
}

// ── RaiseEvent fan-out ───────────────────────────────────────────────────────

/// True when `owner` currently has at least one advised COM connection-point sink.
///
/// # Safety
/// `owner` is a live `CompatObjectBase` on its owning thread.
pub(crate) unsafe fn has_connection_point_sinks(owner: *mut CompatObjectBase) -> bool {
    // SAFETY: caller guarantees `owner` is live on its owning thread.
    unsafe { !(*owner).event_sinks.borrow().is_empty() }
}

/// Deliver a VBA `RaiseEvent` (with binder `event_id`) to all advised COM connection-point sinks
/// on `owner`. Only externally visible events on `class_descriptor.events` are dispatched; private
/// events are ignored.
///
/// # Safety
/// `owner` is a live `CompatObjectBase` on its owning thread.
pub(crate) unsafe fn fire_connection_point_event_for_owner(
    owner: *mut CompatObjectBase,
    event_id: i32,
    args: &[Variant],
) -> Result<(), String> {
    // SAFETY: `owner` is a live `CompatObjectBase` on its owning thread.
    let descriptor = unsafe { (*owner).class_descriptor };
    let Some(event_desc) = descriptor
        .events
        .iter()
        .find(|ev| ev.externally_visible && ev.event_id == event_id)
    else {
        return Ok(());
    };
    let dispid = event_desc.dispid;

    // Snapshot and `AddRef` all advised sinks before dropping the `RefCell` borrow so sinks can
    // re-entrantly `Advise`, `Unadvise`, or call back into `owner` during delivery.
    // SAFETY: `owner` is live on its owning thread.
    let sinks: Vec<*mut c_void> = unsafe {
        (*owner)
            .event_sinks
            .borrow()
            .iter()
            .map(|sink| {
                add_ref_com(sink.dispatch);
                sink.dispatch
            })
            .collect()
    };
    if sinks.is_empty() {
        return Ok(());
    }

    // COM `DISPPARAMS.rgvarg` orders arguments right-to-left (last argument at index 0).
    let mut rgvarg: Vec<Variant> = args.iter().rev().cloned().collect();
    let mut params = DispParams {
        rgvarg: if rgvarg.is_empty() {
            core::ptr::null_mut()
        } else {
            rgvarg.as_mut_ptr().cast::<VariantCore>()
        },
        rgdispid_named_args: core::ptr::null_mut(),
        c_args: rgvarg.len() as u32,
        c_named_args: 0,
    };

    let mut first_error: Option<String> = None;
    for sink in sinks {
        let mut result = Variant::empty();
        let mut excep = ExcepInfo::default();
        let mut arg_err = 0u32;
        // SAFETY: `sink` is a retained `IDispatch*` pointer; `params`, `result`, `excep`, and
        // `arg_err` stay live for the duration of `Invoke`.
        let hr = unsafe {
            let vtbl = &*(*sink.cast::<IUnknown>()).vtbl.cast::<IDispatchVtbl>();
            (vtbl.invoke)(
                sink,
                dispid,
                &IID_NULL,
                0,
                DISPATCH_METHOD,
                (&raw mut params).cast::<c_void>(),
                (&raw mut result).cast::<c_void>(),
                (&raw mut excep).cast::<c_void>(),
                &mut arg_err,
            )
        };
        // SAFETY: free any BSTRs the sink allocated into `excep`.
        unsafe {
            oleaut::sys_free_string(excep.bstr_source);
            oleaut::sys_free_string(excep.bstr_description);
            oleaut::sys_free_string(excep.bstr_help_file);
            release_com(sink);
        }
        drop(result);
        if hr < 0 && first_error.is_none() {
            first_error = Some(format!(
                "event sink Invoke(dispid={dispid}) failed with HRESULT 0x{:08X}",
                hr as u32
            ));
        }
    }
    drop(rgvarg);

    match first_error {
        Some(message) => Err(message),
        None => Ok(()),
    }
}
