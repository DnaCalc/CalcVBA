//! Exact typed-slot transport for a bounded family of imported dual members.
use crate::{
    OptionalParamDefault as D, SourceTypeKind, TypeLibMemberMetadata, TypeLibParamType as P,
    TypeLibWireType as W,
};
use calcvba_host_api::ValueType as T;

/// Validates a native signature before any receiver or VBA body is entered.
/// The supported ABI is HRESULT(this, [Long], [hidden LCID], [retval]).
pub fn signature(member: &TypeLibMemberMetadata) -> Option<(Vec<T>, T)> {
    if !member.callconv_is_stdcall
        || !member.is_dual
        || member.source_typekind != Some(SourceTypeKind::Interface)
        || member.interface_iid.is_none()
    {
        return None;
    }
    let slot = member.vtable_slot?;
    if slot < 7 || slot >= member.vtable_slot_bound? || member.vtable_slot_bound? > 4096 {
        return None;
    }
    let count = member.parameter_types.len();
    if member.parameter_wire_types.len() != count
        || member.parameter_optional.len() != count
        || member.parameter_optional_defaults.len() != count
    {
        return None;
    }
    let mut parameters = Vec::new();
    for (i, (&ty, wire)) in member
        .parameter_types
        .iter()
        .zip(&member.parameter_wire_types)
        .enumerate()
    {
        if wire != &W::Automation(ty) {
            return None;
        }
        if member.parameter_optional_defaults[i] == D::Lcid {
            if i + 1 != count || ty != P::Long {
                return None;
            }
        } else {
            if member.parameter_optional[i]
                || member.parameter_optional_defaults[i] != D::Required
                || ty != P::Long
            {
                return None;
            }
            parameters.push(T::Long);
        }
    }
    let result = match member.return_type {
        Some(P::Long) => T::Long,
        Some(P::Double) => T::Double,
        Some(P::Boolean) => T::Boolean,
        Some(P::String) => T::String,
        None => T::Void,
        _ => return None,
    };
    if member.return_type.map(W::Automation) != member.return_wire_type {
        return None;
    }
    if !parameters.is_empty() && parameters != [T::Long] {
        return None;
    }
    Some((parameters, result))
}

#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
mod native {
    use super::*;
    use crate::{ComObjectToken, WindowsComBridge, windows_client::OwnedComInterface};
    use calcvba_host_api::{HostObject, HostSlot, Value, ValueSlot};
    use oxvba_runtime::ObjectRef;
    use std::{ffi::c_void, rc::Rc};

    impl WindowsComBridge {
        /// Bind compiler-selected native members to the existing typed host table.
        /// The returned owner is same-apartment and cannot be sent between threads.
        /// # Safety
        /// Accepted members with a supported IID describe the receiver's actual
        /// slot ABI, from trusted compiler-retained type information.
        pub unsafe fn prepare_typed_object(
            &self,
            object: ObjectRef,
            members: &[TypeLibMemberMetadata],
        ) -> Result<HostObject, String> {
            let first = members.first().ok_or("typed COM member list is empty")?;
            let iid = first.interface_iid.ok_or("typed COM IID is missing")?;
            for member in members {
                if member.interface_iid != Some(iid) || signature(member).is_none() {
                    return Err("unsupported typed COM member signature or mixed interfaces".into());
                }
            }
            let dispatch = {
                let state = self
                    .shared_state()
                    .lock()
                    .map_err(|_| "COM state lock poisoned")?;
                let binding = state
                    .bindings
                    .get(&ComObjectToken::new(object.raw()))
                    .ok_or("unbound typed COM receiver")?;
                if binding.native_owner_thread != std::thread::current().id() {
                    return Err("typed COM receiver belongs to a different thread/apartment".into());
                }
                let pointer = binding.native_dispatch as *mut crate::RawIDispatch;
                if pointer.is_null() {
                    return Err("typed COM receiver has no native dispatch".into());
                }
                // SAFETY: the locked binding owns this same-apartment native reference.
                unsafe {
                    crate::add_ref_dispatch(pointer);
                    OwnedComInterface::from_owned(pointer.cast())
                }
            };
            // SAFETY: dispatch is retained above and lives in the current COM apartment.
            if unsafe { crate::windows_invoke::dispatch_is_marshaling_proxy(dispatch.as_ptr()) }
                && !crate::windows_invoke::proxy_interface_is_vtable_safe(iid, self.shared_state())
            {
                return Err("typed COM receiver proxy is not typelib-aligned".into());
            }
            // SAFETY: QI borrows the retained receiver and transfers one exact interface reference.
            let raw = unsafe { crate::query_interface_pointer(dispatch.as_ptr(), &iid.to_guid()) }?;
            // SAFETY: raw is the successful QI's exact owned reference in this apartment.
            let receiver = Rc::new(unsafe { OwnedComInterface::from_owned(raw) });
            let bound = members.iter().filter_map(|m| m.vtable_slot).max().unwrap() as usize + 1;
            let mut slots = Vec::with_capacity(bound);
            for index in 0..bound {
                let candidates: Vec<_> = members
                    .iter()
                    .filter(|m| m.vtable_slot == Some(index as u16))
                    .collect();
                if candidates.len() > 1 {
                    return Err("duplicate native typed slot".into());
                }
                let Some(member) = candidates.first() else {
                    slots.push(
                        HostSlot::typed(vec![], T::Void, |_, _| Err(438))
                            .map_err(|e| e.to_string())?,
                    );
                    continue;
                };
                let (parameters, result) =
                    signature(member).ok_or("unsupported native typed signature")?;
                let lcid = member.parameter_optional_defaults.last() == Some(&D::Lcid);
                let receiver = receiver.clone();
                let bridge = self.clone();
                slots.push(
                    HostSlot::typed(parameters, result, move |context, args| {
                        crate::windows_vtable::clear_thread_error_info();
                        // SAFETY: retained QI'd interface, authoritative slot/bound/signature,
                        // and checked typed callback arguments remain live throughout.
                        let value = unsafe { invoke(receiver.as_ptr(), index, result, lcid, args) }
                            .map_err(|fault| fault.publish(context))?;
                        bridge.record_typed_vtable_call();
                        Ok(value)
                    })
                    .map_err(|e| e.to_string())?,
                );
            }
            let mut bytes = [0; 16];
            bytes[..4].copy_from_slice(&iid.data1.to_le_bytes());
            bytes[4..6].copy_from_slice(&iid.data2.to_le_bytes());
            bytes[6..8].copy_from_slice(&iid.data3.to_le_bytes());
            bytes[8..].copy_from_slice(&iid.data4);
            Ok(HostObject::new(bytes, slots))
        }
    }

    // SAFETY CONTRACT: receiver is the QI'd retained interface for this exact
    // authoritative, bounded slot and getter signature. All calls use Windows x64.
    unsafe fn get<R: Default>(
        this: *mut c_void,
        entry: *const c_void,
        lcid: bool,
        args: &[ValueSlot<'_>],
    ) -> Result<R, calcvba_host_api::TypedFault> {
        let mut output = R::default();
        let hr = if let [ValueSlot::Long(value)] = args {
            if lcid {
                // SAFETY: exact imported HRESULT(this, LONG, LCID, R*) signature.
                let call: unsafe extern "system" fn(*mut c_void, i32, u32, *mut R) -> i32 =
                    unsafe { std::mem::transmute(entry) };
                // SAFETY: retained receiver, checked LONG and aligned result cell.
                unsafe { call(this, *value, 0, &mut output) }
            } else {
                // SAFETY: exact imported HRESULT(this, LONG, R*) signature.
                let call: unsafe extern "system" fn(*mut c_void, i32, *mut R) -> i32 =
                    unsafe { std::mem::transmute(entry) };
                // SAFETY: retained receiver, checked LONG and aligned result cell.
                unsafe { call(this, *value, &mut output) }
            }
        } else if lcid {
            // SAFETY: metadata declares exactly HRESULT(this, LCID, R*).
            let call: unsafe extern "system" fn(*mut c_void, u32, *mut R) -> i32 =
                unsafe { std::mem::transmute(entry) };
            // SAFETY: output is one aligned initialized exact result cell.
            unsafe { call(this, 0, &mut output) }
        } else {
            // SAFETY: metadata declares exactly HRESULT(this, R*).
            let call: unsafe extern "system" fn(*mut c_void, *mut R) -> i32 =
                unsafe { std::mem::transmute(entry) };
            // SAFETY: output is one aligned initialized exact result cell.
            unsafe { call(this, &mut output) }
        };
        if hr < 0 {
            Err(map_hresult(hr))
        } else {
            Ok(output)
        }
    }
    fn map_hresult(hr: i32) -> calcvba_host_api::TypedFault {
        let info = crate::windows_vtable::take_error_info();
        let number = crate::windows_invoke::map_com_hresult_vba_number(Some(hr as u32));
        calcvba_host_api::TypedFault {
            number,
            description: info
                .as_ref()
                .and_then(|e| e.description.clone())
                .unwrap_or_else(|| oxvba_runtime::default_error_message(number).into()),
            source: info.as_ref().and_then(|e| e.source.clone()),
            help_file: info.as_ref().and_then(|e| e.help_file.clone()),
            help_context: info.and_then(|e| e.help_context).map(|n| n as i32),
        }
    }
    unsafe fn invoke(
        this: *mut c_void,
        slot: usize,
        result: T,
        lcid: bool,
        args: &[ValueSlot<'_>],
    ) -> Result<Value, calcvba_host_api::TypedFault> {
        // SAFETY: factory validated the authoritative bound and QI'd exact IID.
        let entry = unsafe { *(*(this.cast::<*const *const c_void>())).add(slot) };
        Ok(match result {
            // SAFETY: the imported retval is exactly a Windows LONG out-cell.
            T::Long => Value::Long(unsafe { get::<i32>(this, entry, lcid, args) }?),
            // SAFETY: the imported retval is exactly a Windows DOUBLE out-cell.
            T::Double => Value::Double(unsafe { get::<f64>(this, entry, lcid, args) }?),
            // SAFETY: the imported retval is exactly a 16-bit VARIANT_BOOL out-cell.
            T::Boolean => Value::Boolean(unsafe { get::<i16>(this, entry, lcid, args) }? != 0),
            T::String => {
                // SAFETY: the imported retval is one BSTR pointer out-cell.
                let raw = unsafe { get::<*mut u16>(this, entry, lcid, args) }?;
                // SAFETY: successful [out,retval] transfers exactly one BSTR owner.
                Value::String(unsafe { calcvba_host_api::BStr::from_raw_bstr(raw) })
            }
            T::Void => {
                let hr = match (args, lcid) {
                    ([], false) => {
                        // SAFETY: exact HRESULT(this) method signature.
                        let call: unsafe extern "system" fn(*mut c_void) -> i32 =
                            unsafe { std::mem::transmute(entry) };
                        // SAFETY: retained exact receiver; no inputs or retval.
                        unsafe { call(this) }
                    }
                    ([], true) => {
                        // SAFETY: metadata declares HRESULT(this, LCID), without retval.
                        let call: unsafe extern "system" fn(*mut c_void, u32) -> i32 =
                            unsafe { std::mem::transmute(entry) };
                        // SAFETY: retained exact receiver and by-value neutral LCID.
                        unsafe { call(this, 0) }
                    }
                    ([ValueSlot::Long(value)], false) => {
                        // SAFETY: metadata declares HRESULT(this, LONG), without retval.
                        let call: unsafe extern "system" fn(*mut c_void, i32) -> i32 =
                            unsafe { std::mem::transmute(entry) };
                        // SAFETY: retained receiver and exact checked LONG input.
                        unsafe { call(this, *value) }
                    }
                    ([ValueSlot::Long(value)], true) => {
                        // SAFETY: metadata declares HRESULT(this, LONG, LCID).
                        let call: unsafe extern "system" fn(*mut c_void, i32, u32) -> i32 =
                            unsafe { std::mem::transmute(entry) };
                        // SAFETY: retained receiver, checked LONG, neutral hidden LCID.
                        unsafe { call(this, *value, 0) }
                    }
                    _ => {
                        return Err(calcvba_host_api::TypedFault {
                            number: 13,
                            description: oxvba_runtime::default_error_message(13).into(),
                            source: None,
                            help_file: None,
                            help_context: None,
                        });
                    }
                };
                if hr < 0 {
                    return Err(map_hresult(hr));
                }
                Value::Empty
            }
            _ => {
                return Err(calcvba_host_api::TypedFault {
                    number: 13,
                    description: oxvba_runtime::default_error_message(13).into(),
                    source: None,
                    help_file: None,
                    help_context: None,
                });
            }
        })
    }
}
