//! Seven-slot, registration-free dispinterface. Lookup entry points refuse.
use oxvba_com::{RawIDispatch, RawIDispatchVtbl, RawIUnknownVtbl};
use std::{
    cell::{Cell, RefCell},
    ffi::c_void,
    rc::Rc,
};
use windows_sys::{
    Win32::{
        Foundation::SysAllocStringLen,
        System::{
            Com::{DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO},
            Variant::*,
        },
    },
    core::GUID,
};

pub const IID: oxvba_com::ComInterfaceIid = oxvba_com::ComInterfaceIid {
    data1: 0x67190012,
    data2: 0x2026,
    data3: 0x1102,
    data4: [12; 8],
};
#[derive(Default)]
pub struct Observation {
    pub references: Cell<u32>,
    pub lookups: Cell<u32>,
    pub type_reads: Cell<u32>,
    pub allow_dynamic: Cell<bool>,
    pub queries: Cell<u32>,
    pub calls: RefCell<Vec<(i32, u16, u32)>>,
}
struct Stored(VARIANT);
impl Drop for Stored {
    fn drop(&mut self) {
        // SAFETY: this fixture owns exactly this initialized Automation cell.
        unsafe {
            VariantClear(&mut self.0);
        }
    }
}
#[repr(C)]
struct Object {
    table: *const RawIDispatchVtbl,
    observed: Rc<Observation>,
    value: RefCell<Stored>,
}
pub fn create() -> (*mut RawIDispatch, Rc<Observation>) {
    let observed = Rc::new(Observation::default());
    observed.references.set(1);
    // SAFETY: zero is VT_EMPTY; initialize its exact Double union before publish.
    let mut value: VARIANT = unsafe { std::mem::zeroed() };
    value.Anonymous.Anonymous.vt = VT_R8;
    value.Anonymous.Anonymous.Anonymous.dblVal = 41.0;
    (
        Box::into_raw(Box::new(Object {
            table: &TABLE,
            observed: observed.clone(),
            value: RefCell::new(Stored(value)),
        }))
        .cast(),
        observed,
    )
}
// SAFETY CONTRACT: every entry receives this retained same-thread Object; all
// ABI argument/output cells are supplied by the controlled typed marshaller.
unsafe extern "system" fn query(this: *mut c_void, iid: *const GUID, out: *mut *mut c_void) -> i32 {
    // SAFETY: inherited native input/output and live fixture contract.
    unsafe {
        *out = std::ptr::null_mut();
        let object = &*this.cast::<Object>();
        object
            .observed
            .queries
            .set(object.observed.queries.get() + 1);
        if [
            IID.to_guid(),
            oxvba_com::windows_client::IID_IUNKNOWN,
            oxvba_com::windows_client::IID_IDISPATCH,
        ]
        .iter()
        .any(|v| {
            v.data1 == (*iid).data1
                && v.data2 == (*iid).data2
                && v.data3 == (*iid).data3
                && v.data4 == (*iid).data4
        }) {
            *out = this;
            add_ref(this);
            0
        } else {
            0x80004002u32 as i32
        }
    }
}
unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
    // SAFETY: retained fixture receiver on its creating thread.
    unsafe {
        let observed = &(*this.cast::<Object>()).observed;
        let n = observed.references.get() + 1;
        observed.references.set(n);
        n
    }
}
unsafe extern "system" fn release(this: *mut c_void) -> u32 {
    // SAFETY: caller transfers one fixture reference; last Release owns Box.
    unsafe {
        let object = this.cast::<Object>();
        let observed = &(*object).observed;
        let n = observed.references.get() - 1;
        observed.references.set(n);
        if n == 0 {
            drop(Box::from_raw(object));
        }
        n
    }
}
unsafe fn lookup(this: *mut c_void) -> i32 {
    // SAFETY: inherited retained same-thread fixture receiver.
    unsafe {
        let observed = &(*this.cast::<Object>()).observed;
        observed.lookups.set(observed.lookups.get() + 1);
    }
    0x80004001u32 as i32
}
unsafe extern "system" fn info_count(this: *mut c_void, _: *mut u32) -> i32 {
    // SAFETY: unchanged live receiver.
    unsafe {
        let observed = &(*this.cast::<Object>()).observed;
        observed.type_reads.set(observed.type_reads.get() + 1);
        lookup(this)
    }
}
unsafe extern "system" fn info(this: *mut c_void, _: u32, _: u32, _: *mut *mut c_void) -> i32 {
    // SAFETY: unchanged live receiver.
    unsafe {
        let observed = &(*this.cast::<Object>()).observed;
        observed.type_reads.set(observed.type_reads.get() + 1);
        lookup(this)
    }
}
unsafe extern "system" fn names(
    this: *mut c_void,
    _: *const GUID,
    names: *mut *mut u16,
    count: u32,
    _: u32,
    ids: *mut i32,
) -> i32 {
    // SAFETY: unchanged live receiver.
    unsafe {
        let failure = lookup(this);
        let observed = &(*this.cast::<Object>()).observed;
        if !observed.allow_dynamic.get() {
            return failure;
        }
        if count != 1 || names.is_null() || ids.is_null() {
            return 0x80070057u32 as i32;
        }
        let name = *names;
        if name.is_null() {
            return 0x80070057u32 as i32;
        }
        let mut length = 0;
        while length < 128 && *name.add(length) != 0 {
            length += 1;
        }
        let name = String::from_utf16_lossy(std::slice::from_raw_parts(name, length));
        *ids = match name.to_ascii_lowercase().as_str() {
            "value" => 0,
            "bump" => 18,
            "same" => 15,
            "fail" => 14,
            _ => return 0x80020006u32 as i32,
        };
        0
    }
}
unsafe fn bstr(text: &str) -> *mut u16 {
    let units: Vec<_> = text.encode_utf16().collect();
    // SAFETY: initialized UTF-16 slice spans the explicit allocation length.
    unsafe { SysAllocStringLen(units.as_ptr(), units.len() as u32).cast_mut() }
}
unsafe extern "system" fn fill(exception: *mut EXCEPINFO) -> i32 {
    // SAFETY: COM caller lends its initialized EXCEPINFO; transfer owned BSTRs.
    unsafe {
        (*exception).wCode = 13;
        (*exception).bstrDescription = bstr("dispatch fixture mismatch");
        (*exception).bstrSource = bstr("Owned.Dispatch");
        (*exception).bstrHelpFile = bstr("dispatch-help.chm");
        (*exception).dwHelpContext = 71;
        (*exception).pfnDeferredFillIn = None;
    }
    0
}
unsafe extern "system" fn invoke(
    this: *mut c_void,
    dispid: i32,
    iid: *const GUID,
    _: u32,
    flags: u16,
    params: *mut DISPPARAMS,
    result: *mut VARIANT,
    exception: *mut EXCEPINFO,
    _: *mut u32,
) -> i32 {
    // SAFETY: controlled call lends initialized wrappers and writable output;
    // all union reads below follow exact checked tags and argument counts.
    unsafe {
        let object = &*this.cast::<Object>();
        object
            .observed
            .calls
            .borrow_mut()
            .push((dispid, flags, (*params).cArgs));
        if (*iid).data1 != 0 || (*iid).data2 != 0 || (*iid).data3 != 0 || (*iid).data4 != [0; 8] {
            return 0x80070057u32 as i32;
        }
        let args = std::slice::from_raw_parts((*params).rgvarg, (*params).cArgs as usize);
        let tag = |i: usize| args[i].Anonymous.Anonymous.vt;
        match (dispid, flags, args.len()) {
            (0, flags, 0)
                if flags == DISPATCH_PROPERTYGET
                    || (object.observed.allow_dynamic.get() && flags == 3) =>
            {
                VariantCopy(result, &object.value.borrow().0)
            }
            (0, DISPATCH_PROPERTYPUT, 1)
                if (*params).cNamedArgs == 1 && *(*params).rgdispidNamedArgs == -3 =>
            {
                let mut stored = object.value.borrow_mut();
                VariantClear(&mut stored.0);
                VariantCopy(&mut stored.0, &args[0])
            }
            (12, 1, 2) if tag(0) == VT_R8 && tag(1) == VT_I4 => {
                (*result).Anonymous.Anonymous.vt = VT_R8;
                (*result).Anonymous.Anonymous.Anonymous.dblVal =
                    args[0].Anonymous.Anonymous.Anonymous.dblVal
                        + f64::from(args[1].Anonymous.Anonymous.Anonymous.lVal);
                0
            }
            (13, 1, 2) if tag(0) == VT_BYREF | VT_I4 && tag(1) == VT_BYREF | VT_I4 => {
                let a = args[1].Anonymous.Anonymous.Anonymous.plVal;
                let b = args[0].Anonymous.Anonymous.Anonymous.plVal;
                if a != b {
                    return 0x80070057u32 as i32;
                }
                *a += 10;
                *b += 1;
                (*result).Anonymous.Anonymous.vt = VT_I4;
                (*result).Anonymous.Anonymous.Anonymous.lVal = *a;
                0
            }
            (14, 1..=3, 0) => {
                (*exception).pfnDeferredFillIn = Some(fill);
                0x80020009u32 as i32
            }
            (18, 1..=3, 2)
                if tag(0) == VT_BYREF | VT_VARIANT && tag(1) == VT_BYREF | VT_VARIANT =>
            {
                let a = args[0].Anonymous.Anonymous.Anonymous.pvarVal;
                let b = args[1].Anonymous.Anonymous.Anonymous.pvarVal;
                if a != b || (*a).Anonymous.Anonymous.vt != VT_I4 {
                    return 0x80070057u32 as i32;
                }
                (*a).Anonymous.Anonymous.Anonymous.lVal += 10;
                (*b).Anonymous.Anonymous.Anonymous.lVal += 1;
                VariantCopy(result, a)
            }
            (15, flags, 0) if flags == 2 || (object.observed.allow_dynamic.get() && flags == 3) => {
                (*result).Anonymous.Anonymous.vt = VT_DISPATCH;
                (*result).Anonymous.Anonymous.Anonymous.pdispVal = this;
                add_ref(this);
                0
            }
            (16, 1, 1) if tag(0) == VT_BSTR => VariantCopy(result, &args[0]),
            (17, 1, 1 | 2) if args.iter().all(|v| v.Anonymous.Anonymous.vt == VT_R8) => {
                (*result).Anonymous.Anonymous.vt = VT_R8;
                (*result).Anonymous.Anonymous.Anonymous.dblVal = args
                    .iter()
                    .map(|v| v.Anonymous.Anonymous.Anonymous.dblVal)
                    .sum();
                0
            }
            _ => 0x80020003u32 as i32,
        }
    }
}
static TABLE: RawIDispatchVtbl = RawIDispatchVtbl {
    unknown: RawIUnknownVtbl {
        query_interface: query,
        add_ref,
        release,
    },
    get_type_info_count: info_count,
    get_type_info: info,
    get_ids_of_names: names,
    invoke,
};
