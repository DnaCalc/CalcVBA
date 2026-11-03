//! VBA Object/Variant calls use IDispatch's own runtime name binding.
use super::*;
use calcvba_host_api::object::{DynamicMemberCall, DynamicMemberKind as K};
use windows_sys::Win32::System::{
    Com::*,
    Variant::{VT_BYREF, VT_VARIANT},
};

pub(in crate::typed_slots) fn invoke(
    owner: &NativeReceiver,
    call: &DynamicMemberCall,
    values: &mut [Variant],
    context: &mut calcvba_host_api::TypedContext,
) -> Result<Variant, TypedFault> {
    if call.arguments.len() != values.len() || values.len() > calcvba_host_api::MAX_TYPED_ARGS {
        return Err(fault(450));
    }
    for (i, argument) in call.arguments.iter().enumerate() {
        if argument.alias > i
            || call.arguments[argument.alias].alias != argument.alias
            || (argument.alias != i && (!argument.by_ref || !call.arguments[argument.alias].by_ref))
        {
            return Err(fault(13));
        }
    }
    // SAFETY: retained same-apartment native owner; QI returns one owned ref.
    let dispatch = unsafe {
        crate::query_interface_pointer(
            owner.pointer.as_ptr(),
            &crate::windows_client::IID_IDISPATCH,
        )
    }
    .map_err(|_| fault(438))?;
    // SAFETY: exact successful QI reference is now owned by this guard.
    let dispatch = unsafe { OwnedComInterface::from_owned(dispatch) };
    let receiver = dispatch.as_ptr().cast::<crate::RawIDispatch>();
    let mut names = Vec::new();
    if let Some(name) = &call.name {
        names.push(name.as_str());
    }
    let named_indices: Vec<_> = call
        .arguments
        .iter()
        .enumerate()
        .filter_map(|(i, a)| {
            a.name.as_deref().map(|name| {
                names.push(name);
                i
            })
        })
        .collect();
    if call.name.is_none() && !named_indices.is_empty() {
        return Err(fault(448));
    }
    let mut ids = vec![0; names.len()];
    if !names.is_empty() {
        let mut wide: Vec<Vec<u16>> = names
            .iter()
            .map(|name| name.encode_utf16().chain(Some(0)).collect())
            .collect();
        let mut pointers: Vec<_> = wide.iter_mut().map(|name| name.as_mut_ptr()).collect();
        // SAFETY: exact IDispatch prefix, initialized bounded names/output arrays.
        let hr = unsafe {
            ((*(*receiver).vtbl).get_ids_of_names)(
                receiver.cast(),
                &crate::windows_client::IID_NULL,
                pointers.as_mut_ptr(),
                pointers.len() as u32,
                owner.library.locale_id.unwrap_or(0x400),
                ids.as_mut_ptr(),
            )
        };
        if hr < 0 {
            return Err(fault(if hr == 0x80020006u32 as i32 { 438 } else { 5 }));
        }
    }
    let dispid = if call.name.is_some() { ids[0] } else { 0 };
    let put = matches!(call.kind, K::Let | K::Set);
    if put && values.is_empty() {
        return Err(fault(450));
    }
    let mut order = Vec::with_capacity(values.len());
    let mut named = Vec::new();
    if put {
        order.push(values.len() - 1);
        named.push(crate::COM_DISPID_PROPERTYPUT);
    }
    for (j, &i) in named_indices.iter().enumerate() {
        if put && i == values.len() - 1 {
            return Err(fault(448));
        }
        order.push(i);
        named.push(ids[j + 1]);
    }
    for i in (0..values.len()).rev() {
        if !order.contains(&i) {
            order.push(i);
        }
    }
    // Vec<Box<_>> preserves the native cell address for repeated ByRef aliases.
    let mut cells: Vec<_> = values
        .iter()
        .map(NativeVariant::input)
        .map(|r| r.map(Box::new))
        .collect::<Result<_, _>>()?;
    // Project each cell once. Repeated aliases reuse that raw authority without
    // creating a competing mutable borrow of the same native storage.
    let pointers: Vec<*mut VARIANT> = cells.iter_mut().map(|cell| &raw mut cell.0).collect();
    let mut arguments = Vec::with_capacity(values.len());
    for i in order {
        let a = &call.arguments[i];
        let canonical = if a.by_ref { a.alias } else { i };
        if a.by_ref {
            // SAFETY: all-zero VARIANT is empty; BYREF wrapper owns no payload.
            let mut view: VARIANT = unsafe { std::mem::zeroed() };
            view.Anonymous.Anonymous.vt = VT_BYREF | VT_VARIANT;
            view.Anonymous.Anonymous.Anonymous.pvarVal = pointers[canonical];
            arguments.push(view);
        } else {
            // SAFETY: borrowed wrapper remains backed by its retained owning cell.
            arguments.push(unsafe { std::ptr::read(pointers[i]) });
        }
    }
    let flags = match call.kind {
        K::Method => DISPATCH_METHOD | DISPATCH_PROPERTYGET,
        // An untyped value lookup may denote either a property or a Function.
        // VBA supplies both flags in one Invoke; there is no retry or type lookup.
        K::Get => DISPATCH_METHOD | DISPATCH_PROPERTYGET,
        K::Let => DISPATCH_PROPERTYPUT,
        K::Set => DISPATCH_PROPERTYPUTREF,
    };
    let mut output = NativeVariant::empty();
    let failure = calcvba_host_api::with_native_call(context, || {
        // SAFETY: exact dispatch plus retained native cells and ordered wrappers.
        unsafe {
            invoke_dispatch_arguments(
                owner,
                receiver,
                dispid,
                flags,
                arguments,
                &mut named,
                &mut output,
            )
        }
    });
    for (i, argument) in call.arguments.iter().enumerate() {
        if argument.by_ref && argument.alias == i {
            values[i] = cells[i].value(owner)?;
        }
    }
    if let Some(failure) = failure {
        return Err(failure);
    }
    output.value(owner)
}
