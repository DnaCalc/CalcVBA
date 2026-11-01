//! Registration-free COM object with distinct interface pointers and one identity.
use oxvba_com::{RawIDispatch, RawIDispatchVtbl, RawIUnknownVtbl};
use std::{cell::Cell, ffi::c_void, rc::Rc};
use windows_sys::{
    Win32::System::{
        Com::{DISPPARAMS, EXCEPINFO},
        Variant::VARIANT,
    },
    core::GUID,
};

pub const OTHER_IID: oxvba_com::ComInterfaceIid = oxvba_com::ComInterfaceIid {
    data1: 0x67190011,
    data2: 0x2026,
    data3: 0x1101,
    data4: [11; 8],
};
#[derive(Default)]
pub struct Observation {
    pub references: Cell<u32>,
    pub queries: Cell<u32>,
    pub calls: Cell<u32>,
    pub dynamic: Cell<u32>,
    pub refuse_identity: Cell<bool>,
}
#[repr(C)]
struct Header {
    table: *const Table,
    owner: *mut Storage,
}
#[repr(C)]
struct Table {
    dispatch: RawIDispatchVtbl,
    count: unsafe extern "system" fn(*mut c_void, *mut i32) -> i32,
}
struct Storage {
    headers: [Header; 3],
    observed: Rc<Observation>,
}

pub fn create() -> (*mut RawIDispatch, Rc<Observation>) {
    let observed = Rc::new(Observation::default());
    observed.references.set(1);
    let raw = Box::into_raw(Box::new(Storage {
        headers: std::array::from_fn(|_| Header {
            table: &TABLE,
            owner: std::ptr::null_mut(),
        }),
        observed: observed.clone(),
    }));
    // SAFETY: raw owns the complete stable allocation; only private header
    // backpointers are initialized, before publishing any interface pointer.
    unsafe {
        for header in &mut (*raw).headers {
            header.owner = raw;
        }
        (std::ptr::addr_of_mut!((*raw).headers[0]).cast(), observed)
    }
}
// SAFETY CONTRACT: all native methods receive a retained Header from create;
// its owner/backpointer remains stable until the last Release on this thread.
unsafe fn owner(this: *mut c_void) -> *mut Storage {
    // SAFETY: inherited live fixture Header contract.
    unsafe { (*this.cast::<Header>()).owner }
}
unsafe extern "system" fn query(this: *mut c_void, iid: *const GUID, out: *mut *mut c_void) -> i32 {
    if iid.is_null() || out.is_null() {
        return 0x80004003u32 as i32;
    }
    // SAFETY: native caller supplies readable IID/writable out-cell; this is a
    // live fixture header whose owner spans the QI and newly acquired reference.
    unsafe {
        *out = std::ptr::null_mut();
        let storage = owner(this);
        let observed = &(*storage).observed;
        observed.queries.set(observed.queries.get() + 1);
        let iid = *iid;
        let same = |v: GUID| {
            iid.data1 == v.data1
                && iid.data2 == v.data2
                && iid.data3 == v.data3
                && iid.data4 == v.data4
        };
        let index = if same(oxvba_com::windows_client::IID_IUNKNOWN) {
            if observed.refuse_identity.get() {
                return 0x80004002u32 as i32;
            }
            0
        } else if same(oxvba_com::windows_client::IID_IDISPATCH) {
            0
        } else if same(oxvba_com::DUAL_FIXTURE_INTERFACE_IID.to_guid()) {
            1
        } else if same(OTHER_IID.to_guid()) {
            2
        } else {
            return 0x80004002u32 as i32;
        };
        *out = std::ptr::addr_of_mut!((*storage).headers[index]).cast();
        add_ref(*out);
    }
    0
}
unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
    // SAFETY: inherited retained same-thread fixture allocation contract.
    unsafe {
        let observed = &(*owner(this)).observed;
        let n = observed.references.get() + 1;
        observed.references.set(n);
        n
    }
}
unsafe extern "system" fn release(this: *mut c_void) -> u32 {
    // SAFETY: caller transfers exactly one same-thread native fixture owner.
    unsafe {
        let storage = owner(this);
        let observed = &(*storage).observed;
        let n = observed.references.get() - 1;
        observed.references.set(n);
        if n == 0 {
            drop(Box::from_raw(storage));
        }
        n
    }
}
unsafe extern "system" fn count(this: *mut c_void, out: *mut i32) -> i32 {
    // SAFETY: admitted Count ABI supplies a live receiver and writable LONG.
    unsafe {
        let observed = &(*owner(this)).observed;
        observed.calls.set(observed.calls.get() + 1);
        *out = 17;
    }
    0
}
unsafe fn dynamic(this: *mut c_void) -> i32 {
    // SAFETY: inherited fixture receiver contract; record forbidden transport.
    unsafe {
        let observed = &(*owner(this)).observed;
        observed.dynamic.set(observed.dynamic.get() + 1);
    }
    0x80004001u32 as i32
}
unsafe extern "system" fn info_count(this: *mut c_void, _: *mut u32) -> i32 {
    // SAFETY: forwards the unchanged native receiver.
    unsafe { dynamic(this) }
}
unsafe extern "system" fn info(this: *mut c_void, _: u32, _: u32, _: *mut *mut c_void) -> i32 {
    // SAFETY: forwards the unchanged native receiver.
    unsafe { dynamic(this) }
}
unsafe extern "system" fn names(
    this: *mut c_void,
    _: *const GUID,
    _: *mut *mut u16,
    _: u32,
    _: u32,
    _: *mut i32,
) -> i32 {
    // SAFETY: forwards the unchanged native receiver.
    unsafe { dynamic(this) }
}
unsafe extern "system" fn invoke(
    this: *mut c_void,
    _: i32,
    _: *const GUID,
    _: u32,
    _: u16,
    _: *mut DISPPARAMS,
    _: *mut VARIANT,
    _: *mut EXCEPINFO,
    _: *mut u32,
) -> i32 {
    // SAFETY: forwards the unchanged native receiver.
    unsafe { dynamic(this) }
}
static TABLE: Table = Table {
    dispatch: RawIDispatchVtbl {
        unknown: RawIUnknownVtbl {
            query_interface: query,
            add_ref,
            release,
        },
        get_type_info_count: info_count,
        get_type_info: info,
        get_ids_of_names: names,
        invoke,
    },
    count,
};
