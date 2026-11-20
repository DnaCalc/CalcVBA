//! The one Automation call path over the runtime cell (step 8 of the VARIANT-layout programme,
//! `docs/design/runtime-variant-automation-layout-2026-11-13.md`).
//!
//! The runtime cell is the Automation `VARIANT`, so a call is built straight from it, as VBA
//! builds it: a ByVal argument is a bit copy of the caller's cell that the callee must not clear
//! (borrowed), a ByRef argument is the `VT_BYREF` reference step 7 makes (`VT_BYREF | VT_VARIANT`
//! to a Variant cell, `VT_BYREF | vt` to typed storage), and the result moves into the runtime
//! as it is. Objects are their interface pointers. VM3, the JIT and the typed host stack call
//! these functions; nothing here keeps per-object state.
use calcvba_runtime::ObjectRefExt;
pub mod early;
mod events;
pub use events::{EventConnection, advise_events};

use crate::call::{
    ComArg as Arg, ComArgValue as ArgValue, ComCallKind as CallKind, ComFault as Fault,
    ComMember as Member,
};
use crate::windows_client::{
    IID_IDISPATCH, IID_NULL, OwnedComInterface, RawIDispatch, query_interface_pointer,
};
use calcvba_runtime::{ObjectRef, Variant};
use windows_sys::Win32::System::{Com::*, Variant::*};

/// The locale VBA passes to `GetIDsOfNames` and `Invoke` (`LOCALE_USER_DEFAULT`).
pub const LOCALE_USER_DEFAULT: u32 = 0x0400;
const DISPID_PROPERTYPUT: i32 = -3;
const DISPID_NEWENUM: i32 = -4;
const DISP_E_PARAMNOTFOUND: i32 = 0x8002_0004u32 as i32;
const DISP_E_UNKNOWNNAME: i32 = 0x8002_0006u32 as i32;
const VT_ERROR_TAG: u16 = 0x000A;

/// The fault for a failed `HRESULT`, with the thread's `IErrorInfo` and an `EXCEPINFO`.
fn fault_from_hresult(
    hr: i32,
    excepinfo: Option<crate::windows_invoke::ComInvokeExceptionInfo>,
) -> Fault {
    let info = crate::windows_vtable::take_error_info();
    let failure = crate::windows_invoke::ComInvokeFailure {
        label: "automation",
        dispid: 0,
        hr: Some(hr),
        arg_err: None,
        excep: excepinfo,
        detail: None,
    };
    let number = failure.vba_error_number();
    let excep = failure.excep;
    Fault {
        number,
        description: excep
            .as_ref()
            .and_then(|e| e.description.clone())
            .or_else(|| info.as_ref().and_then(|e| e.description.clone())),
        source: excep
            .as_ref()
            .and_then(|e| e.source.clone())
            .or_else(|| info.as_ref().and_then(|e| e.source.clone())),
        help_file: excep
            .as_ref()
            .and_then(|e| e.help_file.clone())
            .or_else(|| info.as_ref().and_then(|e| e.help_file.clone())),
        help_context: excep
            .as_ref()
            .and_then(|e| e.help_context)
            .or_else(|| info.and_then(|e| e.help_context)),
        hresult: Some(hr),
    }
}

/// The `IDispatch` a project object crosses into COM as, supplied by the caller until project
/// objects carry their own `IDispatch` (step 10): a `VT_DISPATCH` cell owning one reference.
pub type ProjectDispatch<'a> = &'a dyn Fn(&ObjectRef) -> Result<Variant, Fault>;

/// A retained `IDispatch` of an object.
pub struct Dispatch(OwnedComInterface);

impl Dispatch {
    /// The object's `IDispatch` (`QueryInterface`). An object without one cannot be called late
    /// bound: VBA error 438.
    pub fn of(object: &ObjectRef) -> Result<Self, Fault> {
        // SAFETY: the ObjectRef retains a live IUnknown-derived pointer for the call.
        let pointer =
            unsafe { query_interface_pointer(object.raw_iunknown().cast(), &IID_IDISPATCH) }
                .map_err(|_| Fault::number(438))?;
        // SAFETY: a successful QueryInterface returned one owned reference.
        Ok(Self(unsafe { OwnedComInterface::from_owned(pointer) }))
    }

    fn raw(&self) -> *mut RawIDispatch {
        self.0.as_ptr().cast()
    }

    /// The DISPIDs of `names` (one `GetIDsOfNames` call: the member, then named arguments).
    fn ids_of_names(&self, names: &[&str]) -> Result<Vec<i32>, Fault> {
        let mut wide: Vec<Vec<u16>> = names
            .iter()
            .map(|name| name.encode_utf16().chain(Some(0)).collect())
            .collect();
        let mut pointers: Vec<*mut u16> = wide.iter_mut().map(|name| name.as_mut_ptr()).collect();
        let mut ids = vec![0; names.len()];
        let receiver = self.raw();
        // SAFETY: a retained IDispatch; the name and id arrays are initialized and bounded.
        let hr = unsafe {
            ((*(*receiver).vtbl).get_ids_of_names)(
                receiver.cast(),
                &IID_NULL,
                pointers.as_mut_ptr(),
                pointers.len() as u32,
                LOCALE_USER_DEFAULT,
                ids.as_mut_ptr(),
            )
        };
        if hr == DISP_E_UNKNOWNNAME {
            // "Object doesn't support this property or method", or a named argument the member
            // does not have.
            return Err(Fault::number(if names.len() > 1 && ids[0] != -1 {
                448
            } else {
                438
            }));
        }
        if hr < 0 {
            return Err(fault_from_hresult(hr, None));
        }
        Ok(ids)
    }

    /// The name of member `id`, from the object's type information.
    fn member_name(&self, id: i32) -> Option<String> {
        // SAFETY: a retained IDispatch for the call.
        unsafe { crate::windows_typelib_loader::live_member_name(self.raw(), id) }
    }

    /// One `IDispatch::Invoke` with `args` in source order.
    ///
    /// # Safety
    /// Every `Reference` argument is a `VT_BYREF` cell whose target is live, writable storage of
    /// its type for the whole call, and no Rust reference to that storage is held across it.
    pub unsafe fn invoke(
        &self,
        member: Member<'_>,
        kind: CallKind,
        args: &[Arg<'_>],
        project: ProjectDispatch<'_>,
    ) -> Result<Variant, Fault> {
        let put = matches!(kind, CallKind::Let | CallKind::Set);
        if put && args.is_empty() {
            return Err(Fault::number(450));
        }
        // Member and named-argument DISPIDs.
        let named: Vec<usize> = args
            .iter()
            .enumerate()
            .filter_map(|(i, arg)| arg.name.map(|_| i))
            .collect();
        // Named arguments resolve against the member's name; a member given by DISPID gets its
        // name from the object's type information.
        let member_name;
        let member = match member {
            Member::Id(id) if !named.is_empty() => {
                member_name = self.member_name(id).ok_or_else(|| Fault::number(448))?;
                Member::Name(member_name.as_str())
            }
            member => member,
        };
        let mut names: Vec<&str> = Vec::new();
        if let Member::Name(name) = member {
            names.push(name);
        }
        names.extend(named.iter().filter_map(|&i| args[i].name));
        let ids = if names.is_empty() {
            Vec::new()
        } else {
            self.ids_of_names(&names)?
        };
        let dispid = match member {
            Member::Name(_) => ids[0],
            Member::Id(id) => id,
        };
        // rgvarg order: the put value, then named arguments, then positional ones last to first.
        let mut order = Vec::with_capacity(args.len());
        let mut named_ids = Vec::new();
        if put {
            if args[args.len() - 1].name.is_some() {
                return Err(Fault::number(448));
            }
            order.push(args.len() - 1);
            named_ids.push(DISPID_PROPERTYPUT);
        }
        for (j, &i) in named.iter().enumerate() {
            order.push(i);
            named_ids.push(ids[j + 1]);
        }
        for i in (0..args.len()).rev() {
            if !order.contains(&i) {
                order.push(i);
            }
        }
        // Project objects cross as the caller's IDispatch for them; the cells holding those
        // stay alive until the call returns.
        let mut kept: Vec<Variant> = Vec::new();
        let mut rgvarg: Vec<VARIANT> = Vec::with_capacity(args.len());
        for &i in &order {
            let cell = match args[i].value {
                ArgValue::Missing => {
                    let mut missing = [0u8; calcvba_runtime::WIRE_SIZE];
                    missing[..2].copy_from_slice(&VT_ERROR_TAG.to_le_bytes());
                    missing[8..12].copy_from_slice(&DISP_E_PARAMNOTFOUND.to_le_bytes());
                    missing
                }
                ArgValue::Value(value) => match value.as_object_ref() {
                    Some(object) if object.is_compat_object() => {
                        let substitute = project(&object)?;
                        let bytes = substitute.to_wire_bytes();
                        kept.push(substitute);
                        bytes
                    }
                    _ => value.to_wire_bytes(),
                },
                ArgValue::Reference(reference) => reference.to_wire_bytes(),
            };
            // SAFETY: the runtime cell has the Automation VARIANT layout (`WIRE_SIZE` bytes).
            rgvarg.push(unsafe {
                core::mem::transmute::<[u8; calcvba_runtime::WIRE_SIZE], VARIANT>(cell)
            });
        }
        let mut params = DISPPARAMS {
            rgvarg: if rgvarg.is_empty() {
                core::ptr::null_mut()
            } else {
                rgvarg.as_mut_ptr()
            },
            rgdispidNamedArgs: if named_ids.is_empty() {
                core::ptr::null_mut()
            } else {
                named_ids.as_mut_ptr()
            },
            cArgs: rgvarg.len() as u32,
            cNamedArgs: named_ids.len() as u32,
        };
        let flags = match kind {
            // A value read may name a property or a Function: VBA passes both in one Invoke.
            CallKind::Method | CallKind::Get => DISPATCH_METHOD | DISPATCH_PROPERTYGET,
            CallKind::Let => DISPATCH_PROPERTYPUT,
            CallKind::Set => DISPATCH_PROPERTYPUTREF,
        };
        // SAFETY: zero is the documented empty state of both.
        let mut result: VARIANT = unsafe { core::mem::zeroed() };
        // SAFETY: as above.
        let mut exception: EXCEPINFO = unsafe { core::mem::zeroed() };
        let mut argument_error = u32::MAX;
        crate::windows_vtable::clear_thread_error_info();
        let receiver = self.raw();
        // SAFETY: a retained IDispatch; the argument cells are borrowed bit copies or references
        // the caller keeps live (this function's contract); the result and exception cells are
        // initialized.
        let hr = unsafe {
            ((*(*receiver).vtbl).invoke)(
                receiver.cast(),
                dispid,
                &IID_NULL,
                LOCALE_USER_DEFAULT,
                flags,
                &mut params,
                &mut result,
                &mut exception,
                &mut argument_error,
            )
        };
        drop(kept);
        if hr < 0 {
            if let Some(fill) = exception.pfnDeferredFillIn {
                // SAFETY: COM supplies the deferred fill for its own EXCEPINFO; called once,
                // before its BSTRs are taken.
                unsafe { fill(&mut exception) };
            }
            // SAFETY: Invoke (and the deferred fill) transfer each non-null BSTR once.
            let excepinfo = unsafe { crate::windows_invoke::take_excepinfo(&mut exception) };
            // SAFETY: a failed Invoke may still have written the result; it is ours to clear.
            unsafe { VariantClear(&mut result) };
            return Err(fault_from_hresult(hr, excepinfo));
        }
        // SAFETY: a successful Invoke hands the result over; adopt moves it.
        unsafe { adopt(result) }
    }
}

/// Moves a VARIANT a callee produced into the runtime (error 458 for a type VBA has no type
/// for, after clearing it).
///
/// # Safety
/// `value` is owned by the caller and not used afterwards.
unsafe fn adopt(mut value: VARIANT) -> Result<Variant, Fault> {
    // SAFETY: the VARIANT layout is the runtime cell's (`WIRE_SIZE` bytes).
    let bytes = unsafe { core::mem::transmute::<VARIANT, [u8; calcvba_runtime::WIRE_SIZE]>(value) };
    // SAFETY: the caller owns `value`, now moved into `bytes`.
    match unsafe { Variant::adopt_automation(bytes) } {
        Ok(adopted) => Ok(adopted),
        Err(_) => {
            // SAFETY: still owned here; clearing releases what it holds.
            value =
                unsafe { core::mem::transmute::<[u8; calcvba_runtime::WIRE_SIZE], VARIANT>(bytes) };
            // SAFETY: as above.
            unsafe { VariantClear(&mut value) };
            Err(Fault::number(458))
        }
    }
}

/// A `For Each` over an object: its `IEnumVARIANT`, asked for one item at a time, as VBA does.
pub struct Enumerator(OwnedComInterface);

impl Enumerator {
    /// The object's enumerator: `_NewEnum` (`DISPID_NEWENUM`) and `QueryInterface` for
    /// `IEnumVARIANT`. An object without one: VBA error 438.
    pub fn of(object: &ObjectRef, project: ProjectDispatch<'_>) -> Result<Self, Fault> {
        let dispatch = Dispatch::of(object)?;
        // SAFETY: no reference arguments.
        let source =
            unsafe { dispatch.invoke(Member::Id(DISPID_NEWENUM), CallKind::Get, &[], project) }?;
        let unknown = source.as_object_ref().ok_or_else(|| Fault::number(438))?;
        // SAFETY: the ObjectRef retains a live IUnknown for the call.
        let enumerator = unsafe {
            crate::windows_client::query_interface_pointer(
                unknown.raw_iunknown().cast(),
                &crate::windows_client::IID_IENUMVARIANT,
            )
        }
        .map_err(|_| Fault::number(438))?;
        // SAFETY: a successful QueryInterface returned one owned reference.
        Ok(Self(unsafe { OwnedComInterface::from_owned(enumerator) }))
    }

    /// The next item (`IEnumVARIANT::Next(1, ...)`), or `None` at the end.
    pub fn next_item(&mut self) -> Result<Option<Variant>, Fault> {
        let enumerator = self
            .0
            .as_ptr()
            .cast::<crate::windows_client::RawIEnumVARIANT>();
        // SAFETY: zero is the empty VARIANT.
        let mut item: VARIANT = unsafe { core::mem::zeroed() };
        let mut fetched = 0u32;
        // SAFETY: a retained IEnumVARIANT; one initialized output cell.
        let hr =
            unsafe { ((*(*enumerator).vtbl).next)(enumerator.cast(), 1, &mut item, &mut fetched) };
        if hr < 0 {
            // SAFETY: ours to clear on failure.
            unsafe { VariantClear(&mut item) };
            return Err(fault_from_hresult(hr, None));
        }
        if fetched == 0 {
            return Ok(None);
        }
        // SAFETY: Next hands the item over.
        unsafe { adopt(item) }.map(Some)
    }
}

/// `Is`: two objects are the same when their `IUnknown`s are (COM identity).
pub fn same_object(a: &ObjectRef, b: &ObjectRef) -> bool {
    if a.raw_iunknown() == b.raw_iunknown() {
        return true;
    }
    if a.is_compat_object() || b.is_compat_object() {
        return false;
    }
    let identity = |object: &ObjectRef| {
        // SAFETY: the ObjectRef retains a live IUnknown for the call.
        unsafe {
            query_interface_pointer(
                object.raw_iunknown().cast(),
                &crate::windows_client::IID_IUNKNOWN,
            )
        }
        .ok()
        // SAFETY: one owned reference from QueryInterface.
        .map(|unknown| unsafe { OwnedComInterface::from_owned(unknown) })
    };
    match (identity(a), identity(b)) {
        (Some(x), Some(y)) => x.as_ptr() == y.as_ptr(),
        _ => false,
    }
}

impl crate::call::ComEnumerator for Enumerator {
    fn next_item(&mut self) -> Result<Option<Variant>, Fault> {
        Enumerator::next_item(self)
    }
}

/// The `IDispatch` the bridge hands COM for a project object (step 10 gives project objects
/// their own): a `VT_DISPATCH` cell owning one reference to the wrapper.
pub fn project_dispatch(object: &ObjectRef) -> Result<Variant, Fault> {
    let dispatch = crate::windows_runtime_object::create_runtime_object_dispatch(object.clone());
    // SAFETY: the wrapper comes back with one reference, which the cell takes over.
    unsafe { ObjectRef::from_raw_iunknown_owned(dispatch.cast()) }
        .map(Variant::from_object_ref)
        .ok_or_else(|| Fault::number(91))
}
