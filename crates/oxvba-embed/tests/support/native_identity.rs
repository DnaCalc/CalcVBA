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
pub type NativeBooleanCallback = unsafe extern "system" fn(*mut c_void, *mut i16) -> i32;
pub type NativeDoubleCallback = unsafe extern "system" fn(*mut c_void, *mut f64) -> i32;
pub type NativeCallback = unsafe extern "system" fn(*mut c_void, *mut i32) -> i32;
#[derive(Default)]
pub struct Observation {
    pub references: Cell<u32>,
    pub queries: Cell<u32>,
    pub calls: Cell<u32>,
    pub dynamic: Cell<u32>,
    pub refuse_identity: Cell<bool>,
    pub callback: Cell<Option<NativeCallback>>,
    pub double_callback: Cell<Option<NativeDoubleCallback>>,
    pub boolean_callback: Cell<Option<NativeBooleanCallback>>,
    pub callback_context: Cell<*mut c_void>,
    pub prebound: Cell<bool>,
    pub invokes: Cell<u32>,
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
    count_by_ref: unsafe extern "system" fn(*mut c_void, *mut i32, *mut i32) -> i32,
    count_double: unsafe extern "system" fn(*mut c_void, *mut f64, *mut f64, *mut f64) -> i32,
    count_boolean: unsafe extern "system" fn(*mut c_void, *mut i16, *mut i16, *mut i16) -> i32,
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
        if let Some(callback) = observed.callback.get() {
            return callback(observed.callback_context.get(), out);
        }
        *out = 17;
    }
    0
}
unsafe extern "system" fn count_by_ref(this: *mut c_void, amount: *mut i32, out: *mut i32) -> i32 {
    // SAFETY: the declared native signature supplies a retained receiver and
    // distinct writable LONG argument/result cells until the callback returns.
    unsafe {
        *amount += 10;
        count(this, out)
    }
}
unsafe extern "system" fn count_double(
    this: *mut c_void,
    amount: *mut f64,
    alias: *mut f64,
    out: *mut f64,
) -> i32 {
    // SAFETY: the exact descriptor supplies initialized DOUBLE argument/result
    // authorities, with repeated caller aliases sharing one native pointer.
    unsafe {
        if amount != alias {
            return 0x80070057u32 as i32;
        }
        *amount += 10.0;
        let observed = &(*owner(this)).observed;
        observed.calls.set(observed.calls.get() + 1);
        if let Some(callback) = observed.double_callback.get() {
            return callback(observed.callback_context.get(), out);
        }
        *out = *amount;
        0
    }
}
unsafe extern "system" fn count_boolean(
    this: *mut c_void,
    amount: *mut i16,
    alias: *mut i16,
    out: *mut i16,
) -> i32 {
    // SAFETY: the exact descriptor supplies initialized VARIANT_BOOL cells;
    // repeated caller aliases share this same native pointer.
    unsafe {
        if amount != alias {
            return 0x80070057u32 as i32;
        }
        *amount = -1;
        let observed = &(*owner(this)).observed;
        observed.calls.set(observed.calls.get() + 1);
        if let Some(callback) = observed.boolean_callback.get() {
            return callback(observed.callback_context.get(), out);
        }
        *out = *amount;
        0
    }
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
    dispid: i32,
    _: *const GUID,
    _: u32,
    flags: u16,
    params: *mut DISPPARAMS,
    result: *mut VARIANT,
    _: *mut EXCEPINFO,
    _: *mut u32,
) -> i32 {
    // SAFETY: controlled prebound plans lend initialized wrappers and output;
    // checked tags/counts below determine the active Automation union members.
    unsafe {
        let observed = &(*owner(this)).observed;
        if observed.prebound.get() {
            observed.invokes.set(observed.invokes.get() + 1);
            (*result).Anonymous.Anonymous.vt = windows_sys::Win32::System::Variant::VT_I4;
            let out = std::ptr::addr_of_mut!((*result).Anonymous.Anonymous.Anonymous.lVal);
            if dispid == 7010 && flags == 1 && (*params).cArgs == 2 {
                use windows_sys::Win32::System::Variant::{VT_BOOL, VT_BYREF};
                let left = &*(*params).rgvarg.add(1);
                let right = &*(*params).rgvarg;
                if left.Anonymous.Anonymous.vt != VT_BYREF | VT_BOOL
                    || right.Anonymous.Anonymous.vt != VT_BYREF | VT_BOOL
                {
                    return 0x80020005u32 as i32;
                }
                (*result).Anonymous.Anonymous.vt = VT_BOOL;
                return count_boolean(
                    this,
                    left.Anonymous.Anonymous.Anonymous.pboolVal,
                    right.Anonymous.Anonymous.Anonymous.pboolVal,
                    std::ptr::addr_of_mut!((*result).Anonymous.Anonymous.Anonymous.boolVal),
                );
            }
            if dispid == 7009 && flags == 1 && (*params).cArgs == 2 {
                use windows_sys::Win32::System::Variant::{VT_BYREF, VT_R8};
                let left = &*(*params).rgvarg.add(1);
                let right = &*(*params).rgvarg;
                if left.Anonymous.Anonymous.vt != VT_BYREF | VT_R8
                    || right.Anonymous.Anonymous.vt != VT_BYREF | VT_R8
                {
                    return 0x80020005u32 as i32;
                }
                (*result).Anonymous.Anonymous.vt = VT_R8;
                return count_double(
                    this,
                    left.Anonymous.Anonymous.Anonymous.pdblVal,
                    right.Anonymous.Anonymous.Anonymous.pdblVal,
                    std::ptr::addr_of_mut!((*result).Anonymous.Anonymous.Anonymous.dblVal),
                );
            }
            if dispid == 7008 && flags == 1 && (*params).cArgs == 1 {
                let argument = &*(*params).rgvarg;
                if argument.Anonymous.Anonymous.vt
                    == windows_sys::Win32::System::Variant::VT_BYREF
                        | windows_sys::Win32::System::Variant::VT_I4
                {
                    return count_by_ref(this, argument.Anonymous.Anonymous.Anonymous.plVal, out);
                }
            } else if dispid == 7007 && flags == 2 && (*params).cArgs == 0 {
                return count(this, out);
            }
            return 0x80020003u32 as i32;
        }
    }
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
    count_by_ref,
    count_double,
    count_boolean,
};
