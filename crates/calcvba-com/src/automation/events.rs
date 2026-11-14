//! WithEvents over COM: one sink per subscription, advised on the source's connection point,
//! that hands each event to the running VBA program (`calcvba_runtime::events`) while the call
//! that fired it is still running, as VBA does (design doc, step 8c).
use super::Fault;
use crate::windows_client::{
    IID_ICONNECTIONPOINTCONTAINER, IID_IDISPATCH, IID_IUNKNOWN, OwnedComInterface,
    RawIConnectionPoint, RawIConnectionPointContainer, RawIDispatchVtbl, RawIUnknownVtbl,
    query_interface_pointer,
};
use calcvba_runtime::{ObjectRef, Variant};
use core::ffi::c_void;
use std::sync::atomic::{AtomicU32, Ordering};
use windows_sys::Win32::System::{Com::*, Variant::*};
use windows_sys::core::GUID;

const S_OK: i32 = 0;
const E_NOINTERFACE: i32 = 0x8000_4002u32 as i32;
const E_NOTIMPL: i32 = 0x8000_4001u32 as i32;
const E_POINTER: i32 = 0x8000_4003u32 as i32;
const DISP_E_EXCEPTION: i32 = 0x8002_0009u32 as i32;
const DISP_E_PARAMNOTFOUND: i32 = 0x8002_0004u32 as i32;
const DISP_E_UNKNOWNNAME: i32 = 0x8002_0006u32 as i32;

/// The sink object: an `IDispatch` that also answers the source interface.
#[repr(C)]
struct EventSink {
    vtbl: *const RawIDispatchVtbl,
    references: AtomicU32,
    source: GUID,
    key: u64,
}

static SINK_VTBL: RawIDispatchVtbl = RawIDispatchVtbl {
    unknown: RawIUnknownVtbl {
        query_interface: sink_query_interface,
        add_ref: sink_add_ref,
        release: sink_release,
    },
    get_type_info_count: sink_get_type_info_count,
    get_type_info: sink_get_type_info,
    get_ids_of_names: sink_get_ids_of_names,
    invoke: sink_invoke,
};

fn same_guid(a: &GUID, b: &GUID) -> bool {
    a.data1 == b.data1 && a.data2 == b.data2 && a.data3 == b.data3 && a.data4 == b.data4
}

unsafe extern "system" fn sink_query_interface(
    this: *mut c_void,
    riid: *const GUID,
    ppv: *mut *mut c_void,
) -> i32 {
    if ppv.is_null() || riid.is_null() {
        return E_POINTER;
    }
    let sink = this.cast::<EventSink>();
    // SAFETY: COM passes a live sink (this vtable is installed only on EventSink boxes) and a
    // readable IID; `ppv` is the caller's out-slot.
    unsafe {
        let iid = &*riid;
        if same_guid(iid, &IID_IUNKNOWN)
            || same_guid(iid, &IID_IDISPATCH)
            || same_guid(iid, &(*sink).source)
        {
            *ppv = this;
            sink_add_ref(this);
            S_OK
        } else {
            *ppv = core::ptr::null_mut();
            E_NOINTERFACE
        }
    }
}

unsafe extern "system" fn sink_add_ref(this: *mut c_void) -> u32 {
    // SAFETY: a live sink.
    unsafe {
        (*this.cast::<EventSink>())
            .references
            .fetch_add(1, Ordering::AcqRel)
            + 1
    }
}

unsafe extern "system" fn sink_release(this: *mut c_void) -> u32 {
    let sink = this.cast::<EventSink>();
    // SAFETY: a live sink; the last reference frees the box `new_sink` made.
    unsafe {
        let left = (*sink).references.fetch_sub(1, Ordering::AcqRel) - 1;
        if left == 0 {
            drop(Box::from_raw(sink));
        }
        left
    }
}

unsafe extern "system" fn sink_get_type_info_count(_this: *mut c_void, count: *mut u32) -> i32 {
    if count.is_null() {
        return E_POINTER;
    }
    // SAFETY: the caller's out-slot.
    unsafe { *count = 0 };
    S_OK
}

unsafe extern "system" fn sink_get_type_info(
    _this: *mut c_void,
    _index: u32,
    _lcid: u32,
    _info: *mut *mut c_void,
) -> i32 {
    E_NOTIMPL
}

unsafe extern "system" fn sink_get_ids_of_names(
    _this: *mut c_void,
    _riid: *const GUID,
    _names: *mut *mut u16,
    _count: u32,
    _lcid: u32,
    _ids: *mut i32,
) -> i32 {
    // Sources call their sinks by DISPID.
    DISP_E_UNKNOWNNAME
}

unsafe extern "system" fn sink_invoke(
    this: *mut c_void,
    dispid: i32,
    _riid: *const GUID,
    _lcid: u32,
    _flags: u16,
    params: *mut DISPPARAMS,
    _result: *mut VARIANT,
    excepinfo: *mut EXCEPINFO,
    _arg_error: *mut u32,
) -> i32 {
    let sink = this.cast::<EventSink>();
    // SAFETY: a live sink.
    let key = unsafe { (*sink).key };
    // The arguments in source order (rgvarg is last to first). A reference stays the server's
    // storage; a value is copied, as VBA copies ByVal arguments.
    let mut args = Vec::new();
    if !params.is_null() {
        // SAFETY: COM passes a readable DISPPARAMS for the call.
        let params = unsafe { &*params };
        let total = params.cArgs as usize;
        let named = params.cNamedArgs as usize;
        if named > total || (named > 0 && params.rgdispidNamedArgs.is_null()) {
            return DISP_E_PARAMNOTFOUND;
        }
        // Source order. Named arguments come first in rgvarg, each at the position its DISPID
        // names (an event's parameters are numbered 0, 1, ...; Excel passes them all by name);
        // positional ones follow, last to first.
        let mut slots: Vec<Option<usize>> = vec![None; total];
        for (position, slot) in slots.iter_mut().enumerate().take(total - named) {
            *slot = Some(total - 1 - position);
        }
        for entry in 0..named {
            // SAFETY: rgdispidNamedArgs holds cNamedArgs DISPIDs.
            let position = unsafe { *params.rgdispidNamedArgs.add(entry) };
            let Some(slot) = usize::try_from(position)
                .ok()
                .and_then(|p| slots.get_mut(p))
            else {
                return DISP_E_PARAMNOTFOUND;
            };
            if slot.replace(entry).is_some() {
                return DISP_E_PARAMNOTFOUND;
            }
        }
        for slot in slots {
            let Some(index) = slot else {
                return DISP_E_PARAMNOTFOUND;
            };
            // SAFETY: rgvarg holds cArgs initialized VARIANTs, which have the runtime cell's
            // layout; the cell is only borrowed here.
            let cell = unsafe { &*params.rgvarg.add(index).cast::<Variant>() };
            let vt = cell.core().raw_vt();
            if calcvba_runtime::VarType::from_u16(vt).is_none() {
                // An Automation type VBA has no type for.
                return fail(
                    excepinfo,
                    &calcvba_runtime::events::EventFault {
                        number: 458,
                        description: calcvba_runtime::default_error_message(458).to_string(),
                        source: String::new(),
                    },
                );
            }
            args.push(if vt & VT_BYREF != 0 {
                cell.clone_reference()
            } else {
                cell.clone()
            });
        }
    }
    let delivered = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        calcvba_runtime::events::deliver(key, dispid, &args)
    }));
    drop(args);
    match delivered {
        Ok(Ok(_)) => S_OK,
        Ok(Err(fault)) => fail(excepinfo, &fault),
        // A panic in the handler run is an internal error (VBA 51).
        Err(_) => fail(
            excepinfo,
            &calcvba_runtime::events::EventFault {
                number: 51,
                description: calcvba_runtime::default_error_message(51).to_string(),
                source: String::new(),
            },
        ),
    }
}

/// `DISP_E_EXCEPTION` with VBA's error in `EXCEPINFO` (`scode` is the VBA error's HRESULT).
fn fail(excepinfo: *mut EXCEPINFO, fault: &calcvba_runtime::events::EventFault) -> i32 {
    if !excepinfo.is_null() {
        let bstr = |text: &str| {
            let wide: Vec<u16> = text.encode_utf16().collect();
            // SAFETY: a counted UTF-16 copy; the EXCEPINFO's receiver frees it.
            unsafe {
                windows_sys::Win32::Foundation::SysAllocStringLen(wide.as_ptr(), wide.len() as u32)
            }
        };
        // SAFETY: the caller's EXCEPINFO, which it zeroed; its BSTRs become the caller's.
        unsafe {
            (*excepinfo).wCode = 0;
            (*excepinfo).scode = (0x800A_0000u32 | (fault.number as u32 & 0xFFFF)) as i32;
            (*excepinfo).bstrDescription = bstr(&fault.description);
            (*excepinfo).bstrSource = bstr(&fault.source);
        }
    }
    DISP_E_EXCEPTION
}

/// An advised sink. Dropping it unadvises the sink (`Set x = Nothing`, rebinding, or the owner's
/// release).
pub struct EventConnection {
    point: OwnedComInterface,
    cookie: u32,
}

impl crate::call::ComEventConnection for EventConnection {}

impl Drop for EventConnection {
    fn drop(&mut self) {
        let point = self.point.as_ptr().cast::<RawIConnectionPoint>();
        // SAFETY: the retained connection point this connection advised on.
        unsafe {
            ((*(*point).vtbl).unadvise)(point.cast(), self.cookie);
        }
    }
}

/// Advises a sink for the source interface `interface` (an IID in registry form) on `source`'s
/// connection point; the sink delivers events for subscription `key`. An object without that
/// connection point: VBA error 459 ("Object or class does not support the set of events").
pub fn advise_events(
    source: &ObjectRef,
    interface: &str,
    key: u64,
) -> Result<EventConnection, Fault> {
    let unsupported = || Fault::number(459);
    let iid = crate::windows_client::parse_guid_canonical(interface).ok_or_else(unsupported)?;
    // SAFETY: the ObjectRef retains a live interface for the call.
    let container = unsafe {
        query_interface_pointer(source.raw_iunknown().cast(), &IID_ICONNECTIONPOINTCONTAINER)
    }
    .map_err(|_| unsupported())?;
    // SAFETY: one owned reference from QueryInterface.
    let container = unsafe { OwnedComInterface::from_owned(container) };
    let raw = container.as_ptr().cast::<RawIConnectionPointContainer>();
    let mut point: *mut c_void = core::ptr::null_mut();
    // SAFETY: a retained IConnectionPointContainer; one out-slot.
    let hr = unsafe { ((*(*raw).vtbl).find_connection_point)(raw.cast(), &iid, &mut point) };
    if hr < 0 || point.is_null() {
        return Err(unsupported());
    }
    // SAFETY: FindConnectionPoint returned one owned reference.
    let point = unsafe { OwnedComInterface::from_owned(point) };
    let sink = Box::into_raw(Box::new(EventSink {
        vtbl: &SINK_VTBL,
        references: AtomicU32::new(1),
        source: iid,
        key,
    }));
    let mut cookie = 0u32;
    let raw_point = point.as_ptr().cast::<RawIConnectionPoint>();
    // SAFETY: a retained connection point; the sink is a live COM object it may AddRef.
    let hr = unsafe { ((*(*raw_point).vtbl).advise)(raw_point.cast(), sink.cast(), &mut cookie) };
    // SAFETY: our own reference to the sink; the connection point holds its own on success.
    unsafe { sink_release(sink.cast()) };
    if hr < 0 {
        return Err(unsupported());
    }
    Ok(EventConnection { point, cookie })
}
