//! Retained typed plans for imported dual members and pure dispinterfaces.
use crate::{
    OptionalParamDefault as D, SourceTypeKind, TypeLibMemberMetadata, TypeLibParamType as P,
    TypeLibWireType as W,
};
use calcvba_host_api::ValueType as T;

#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
mod events;
#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
mod factory;
#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
mod library;
#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
mod native_values;
#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
pub use events::{TypedEventSubscription, advise_typed_event, prepare_typed_event_object};
pub mod plan;

pub fn signature(member: &TypeLibMemberMetadata) -> Option<(Vec<T>, T)> {
    plan::visible_signature(member, plan::visible_indices(member).len())
}

/// Validates a native signature before any receiver or VBA body is entered.
/// The supported ABI is HRESULT(this, [Long], [hidden LCID], [retval]).
pub fn scalar_signature(member: &TypeLibMemberMetadata) -> Option<(Vec<T>, T)> {
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
    use crate::WindowsComBridge;
    use calcvba_host_api::{HostObject, Value, ValueSlot};
    use calcvba_runtime::ObjectRef;
    use std::ffi::c_void;

    impl WindowsComBridge {
        /// # Safety
        /// Members are unchanged trusted compiler-retained ABI descriptors.
        pub unsafe fn prepare_typed_object(
            &self,
            object: ObjectRef,
            members: &[TypeLibMemberMetadata],
        ) -> Result<HostObject, String> {
            // SAFETY: forward the same trusted descriptor and apartment contract.
            unsafe { self.prepare_typed_object_with_interfaces(object, members, &[]) }
        }
    }

    // SAFETY CONTRACT: receiver is the QI'd retained interface for this exact
    // authoritative, bounded slot and getter signature. All calls use Windows x64.
    trait NativeRetval: Default {
        fn discard(self) {}
    }
    impl NativeRetval for i32 {}
    impl NativeRetval for f64 {}
    impl NativeRetval for i16 {}
    impl NativeRetval for *mut u16 {
        fn discard(self) {
            // SAFETY: a failed out-call may still transfer one initialized BSTR;
            // get initialized this cell to null before entering the server.
            unsafe {
                windows_sys::Win32::Foundation::SysFreeString(self);
            }
        }
    }
    unsafe fn get<R: NativeRetval>(
        this: *mut c_void,
        entry: *const c_void,
        lcid: bool,
        locale_id: u32,
        args: &[ValueSlot<'_>],
    ) -> Result<R, calcvba_host_api::TypedFault> {
        let mut output = R::default();
        let hr = if let [ValueSlot::Long(value)] = args {
            if lcid {
                // SAFETY: exact imported HRESULT(this, LONG, LCID, R*) signature.
                let call: unsafe extern "system" fn(*mut c_void, i32, u32, *mut R) -> i32 =
                    unsafe { std::mem::transmute(entry) };
                // SAFETY: retained receiver, checked LONG and aligned result cell.
                unsafe { call(this, *value, locale_id, &mut output) }
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
            unsafe { call(this, locale_id, &mut output) }
        } else {
            // SAFETY: metadata declares exactly HRESULT(this, R*).
            let call: unsafe extern "system" fn(*mut c_void, *mut R) -> i32 =
                unsafe { std::mem::transmute(entry) };
            // SAFETY: output is one aligned initialized exact result cell.
            unsafe { call(this, &mut output) }
        };
        if hr < 0 {
            let fault = map_hresult(hr);
            output.discard();
            Err(fault)
        } else {
            Ok(output)
        }
    }
    pub(super) fn map_hresult(hr: i32) -> calcvba_host_api::TypedFault {
        let info = crate::windows_vtable::take_error_info();
        let number = crate::windows_invoke::map_com_hresult_vba_number(Some(hr as u32));
        calcvba_host_api::TypedFault {
            number,
            description: info
                .as_ref()
                .and_then(|e| e.description.clone())
                .unwrap_or_else(|| calcvba_runtime::default_error_message(number).into()),
            source: info.as_ref().and_then(|e| e.source.clone()),
            help_file: info.as_ref().and_then(|e| e.help_file.clone()),
            help_context: info.and_then(|e| e.help_context).map(|n| n as i32),
        }
    }
    pub(super) unsafe fn invoke(
        this: *mut c_void,
        slot: usize,
        result: T,
        lcid: bool,
        locale_id: u32,
        args: &[ValueSlot<'_>],
    ) -> Result<Value, calcvba_host_api::TypedFault> {
        // SAFETY: factory validated the authoritative bound and QI'd exact IID.
        let entry = unsafe { *(*(this.cast::<*const *const c_void>())).add(slot) };
        Ok(match result {
            // SAFETY: the imported retval is exactly a Windows LONG out-cell.
            T::Long => Value::Long(unsafe { get::<i32>(this, entry, lcid, locale_id, args) }?),
            // SAFETY: the imported retval is exactly a Windows DOUBLE out-cell.
            T::Double => Value::Double(unsafe { get::<f64>(this, entry, lcid, locale_id, args) }?),
            // SAFETY: the imported retval is exactly a 16-bit VARIANT_BOOL out-cell.
            T::Boolean => {
                // SAFETY: exact imported VARIANT_BOOL retval and selected LCID.
                Value::Boolean(unsafe { get::<i16>(this, entry, lcid, locale_id, args) }? != 0)
            }
            T::String => {
                // SAFETY: the imported retval is one BSTR pointer out-cell.
                let raw = unsafe { get::<*mut u16>(this, entry, lcid, locale_id, args) }?;
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
                        // SAFETY: retained exact receiver and by-value selected LCID.
                        unsafe { call(this, locale_id) }
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
                        unsafe { call(this, *value, locale_id) }
                    }
                    _ => {
                        return Err(calcvba_host_api::TypedFault {
                            number: 13,
                            description: calcvba_runtime::default_error_message(13).into(),
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
                    description: calcvba_runtime::default_error_message(13).into(),
                    source: None,
                    help_file: None,
                    help_context: None,
                });
            }
        })
    }
}
