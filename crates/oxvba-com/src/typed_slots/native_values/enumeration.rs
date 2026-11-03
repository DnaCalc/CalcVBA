//! Execute a retained _NewEnum plan and keep IEnumVARIANT live across slices.
use super::*;
use crate::{
    typed_slots::library::NativeMemberPlan,
    windows_client::{RawIEnumVARIANT, query_enumvariant_from_unknown},
};
use calcvba_host_api::{HostEnumerator, TypedContext};
use std::rc::Rc;

/// # Safety
/// The exact trusted member is the declared zero-argument enumerator getter;
/// owner retains its QI'd interface on this same apartment.
unsafe fn start(
    owner: &NativeReceiver,
    member: &TypeLibMemberMetadata,
) -> Result<OwnedComInterface, TypedFault> {
    let enum_pointer = if plan::is_dispatch(member) {
        let mut output = NativeVariant::empty();
        // SAFETY: retained exact DISPID/getter with no supplied source arguments.
        if let Some(error) = unsafe { invoke_dispatch(owner, member, vec![], &mut output) } {
            return Err(error);
        }
        // SAFETY: successful Invoke initialized this native VARIANT discriminant.
        let tag = unsafe { output.0.Anonymous.Anonymous.vt };
        if !matches!(
            tag,
            windows_sys::Win32::System::Variant::VT_UNKNOWN
                | windows_sys::Win32::System::Variant::VT_DISPATCH
        ) {
            return Err(fault(438));
        }
        // SAFETY: both admitted tags hold one borrowed native interface pointer.
        let pointer = unsafe { output.0.Anonymous.Anonymous.Anonymous.punkVal };
        // SAFETY: the native result cell owns this borrowed IUnknown prefix.
        unsafe { query_enumvariant_from_unknown(pointer.cast()) }.map_err(|_| fault(438))?
    } else {
        let mut output = NativePointer {
            raw: std::ptr::null_mut(),
        };
        // SAFETY: exact imported slot, bounded during preparation.
        let entry = unsafe {
            *(*(owner.pointer.as_ptr().cast::<*const *const c_void>()))
                .add(member.vtable_slot.unwrap() as usize)
        };
        crate::windows_vtable::clear_thread_error_info();
        let hr = if member.parameter_types.is_empty() {
            // SAFETY: exact HRESULT(this, IUnknown**) enumerator getter.
            let call: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> i32 =
                unsafe { std::mem::transmute(entry) };
            // SAFETY: retained receiver and one initialized writable out-pointer.
            unsafe { call(owner.pointer.as_ptr(), &mut output.raw) }
        } else {
            // SAFETY: admission permits only one hidden LCID before the retval.
            let call: unsafe extern "system" fn(*mut c_void, u32, *mut *mut c_void) -> i32 =
                unsafe { std::mem::transmute(entry) };
            // SAFETY: exact selected locale and initialized owning pointer cell.
            unsafe {
                call(
                    owner.pointer.as_ptr(),
                    owner.library.locale_id.unwrap_or(0),
                    &mut output.raw,
                )
            }
        };
        let error = (hr < 0).then(|| map_hresult(hr));
        owner.bridge.record_typed_vtable_call();
        if let Some(error) = error {
            return Err(error);
        }
        // SAFETY: successful getter transferred this live native IUnknown.
        unsafe { query_enumvariant_from_unknown(output.raw.cast()) }.map_err(|_| fault(438))?
    };
    // SAFETY: successful QI transferred one owned IEnumVARIANT/IUnknown ref.
    Ok(unsafe { OwnedComInterface::from_owned(enum_pointer.cast()) })
}

pub(in crate::typed_slots) fn prepare(
    owner: Rc<NativeReceiver>,
    plan: &NativeMemberPlan,
    context: &mut TypedContext,
) -> Result<HostEnumerator, i32> {
    let member = &plan.member;
    let valid_parameters = member.parameter_types.is_empty()
        || (member.parameter_types == [P::Long] && member.parameter_optional_defaults == [D::Lcid]);
    if member.token != -4
        || member.return_type != Some(P::Object)
        || !valid_parameters
        || !matches!(
            member.invoke_kind,
            crate::TypeLibMemberInvokeKind::PropertyGet | crate::TypeLibMemberInvokeKind::Method
        )
    {
        return Err(438);
    }
    let iterator = calcvba_host_api::with_native_call(context, || {
        // SAFETY: exact immutable descriptor and retained same-apartment owner.
        unsafe { start(&owner, member) }
    })
    .map_err(|e| e.publish(context))?;
    Ok(HostEnumerator::new(move |context| {
        calcvba_host_api::with_native_call(context, || {
            let mut output = NativeVariant::empty();
            let mut fetched = 0;
            let pointer = iterator.as_ptr().cast::<RawIEnumVARIANT>();
            crate::windows_vtable::clear_thread_error_info();
            // SAFETY: QI proved IEnumVARIANT; output/fetched are initialized cells.
            let hr = unsafe {
                ((*(*pointer).vtbl).next)(pointer.cast(), 1, &mut output.0, &mut fetched)
            };
            let error = (hr < 0).then(|| map_hresult(hr));
            owner.bridge.record_typed_vtable_call();
            if let Some(error) = error {
                return Err(error);
            }
            match (hr, fetched) {
                (0, 1) => Ok(Some(Value::Variant(output.value(&owner)?))),
                (1, 0) => Ok(None),
                _ => Err(fault(13)),
            }
        })
        .map_err(|e| e.publish(context))
    }))
}
