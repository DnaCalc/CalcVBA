//! Same-apartment ownership and immutable type-library plans for native tables.
use super::library::{NativeLibrary, NativeMemberPlan};
use super::*;
use crate::{
    ComObjectToken, TypeLibInterfaceMetadata, WindowsComBridge, windows_client::OwnedComInterface,
};
use calcvba_host_api::{HostObject, HostSlot};
use std::{ffi::c_void, rc::Rc};

pub(super) struct NativeReceiver {
    pub pointer: Rc<OwnedComInterface>,
    _identity: OwnedComInterface,
    pub library: Rc<NativeLibrary>,
    pub bridge: WindowsComBridge,
    enumerator: Option<NativeMemberPlan>,
}
struct NativeServices(Rc<NativeReceiver>);

fn native_input(
    bridge: &WindowsComBridge,
    object: &calcvba_runtime::ObjectRef,
) -> Result<OwnedComInterface, String> {
    if let Some(prepared) = HostObject::from_object_ref(object) {
        let pointer = super::retain_prepared_native_dispatch(&prepared)?;
        // SAFETY: the boundary helper transferred one native dispatch reference.
        return Ok(unsafe { OwnedComInterface::from_owned(pointer) });
    }
    if !object.is_compat_object() {
        // SAFETY: an ObjectRef owns a same-apartment IUnknown; clone transfers
        // exactly one reference. The selected imported IID is queried by build.
        return Ok(unsafe {
            OwnedComInterface::from_owned(object.clone().into_raw_iunknown().cast())
        });
    }
    let state = bridge
        .shared_state()
        .lock()
        .map_err(|_| "COM state lock poisoned")?;
    let binding = state
        .bindings
        .get(&ComObjectToken::new(object.raw()))
        .ok_or("unbound native COM receiver")?;
    if binding.native_owner_thread != std::thread::current().id() {
        return Err("native COM receiver belongs to a different thread/apartment".into());
    }
    let pointer = binding.native_dispatch as *mut crate::RawIDispatch;
    if pointer.is_null() {
        return Err("native COM receiver has no native dispatch".into());
    }
    // SAFETY: the locked binding retains this same-apartment native object.
    unsafe {
        crate::add_ref_dispatch(pointer);
        Ok(OwnedComInterface::from_owned(pointer.cast()))
    }
}

impl calcvba_host_api::HostObjectServices for NativeServices {
    fn subscribe_event(
        &self,
        descriptor: &calcvba_host_api::object::EventDescriptor,
        callback: calcvba_host_api::object::EventCallback,
    ) -> Result<calcvba_host_api::object::EventSubscription, i32> {
        super::events::subscribe_automatic(self.0.clone(), descriptor, callback)
    }
    fn invoke_dynamic(
        &self,
        call: &calcvba_host_api::object::DynamicMemberCall,
        arguments: &mut [calcvba_runtime::Variant],
        context: &mut calcvba_host_api::TypedContext,
    ) -> Result<calcvba_runtime::Variant, calcvba_host_api::TypedFault> {
        super::native_values::invoke_dynamic(&self.0, call, arguments, context)
    }
    fn resolve_interface(&self, wanted: calcvba_host_api::InterfaceId) -> Result<HostObject, i32> {
        let owner = &self.0;
        let index = owner
            .library
            .metadata
            .iter()
            .position(|meta| {
                meta.iid.is_some_and(|iid| {
                    let mut bytes = [0; 16];
                    bytes[..4].copy_from_slice(&iid.data1.to_le_bytes());
                    bytes[4..6].copy_from_slice(&iid.data2.to_le_bytes());
                    bytes[6..8].copy_from_slice(&iid.data3.to_le_bytes());
                    bytes[8..].copy_from_slice(&iid.data4);
                    bytes == wanted
                })
            })
            .ok_or(13)?;
        // SAFETY: exact retained library metadata and same-apartment native owner.
        unsafe {
            make_native_object(
                owner.bridge.clone(),
                owner.pointer.as_ptr(),
                Some(index),
                owner.library.clone(),
            )
        }
        .map_err(|_| 13)
    }
    fn enumerator(
        &self,
        context: &mut calcvba_host_api::TypedContext,
    ) -> Result<calcvba_host_api::HostEnumerator, i32> {
        let plan = self.0.enumerator.as_ref().ok_or(438)?;
        super::native_values::prepare_enumerator(self.0.clone(), plan, context)
    }
}
impl WindowsComBridge {
    /// Retain an activated IDispatch as a genuinely dynamic owning receiver.
    /// No nominal interface, member metadata or vtable signature is invented.
    pub fn prepare_dynamic_object(
        &self,
        object: calcvba_runtime::ObjectRef,
    ) -> Result<HostObject, String> {
        let dispatch = native_input(self, &object)?;
        let library = Rc::new(NativeLibrary {
            metadata: vec![],
            plans: vec![],
            locale_id: None,
        });
        // SAFETY: retained same-apartment IDispatch, with no nominal member plans.
        unsafe { make_native_object(self.clone(), dispatch.as_ptr(), None, library) }
    }
    /// # Safety
    /// All accepted descriptors are unchanged trusted compiler/type-library ABI
    /// metadata for their exact IIDs. Result identities are resolved before entry.
    /// For a pure dispinterface, interfaces includes the complete compiler member
    /// order used to assign typed table indices; these are not native slots.
    pub unsafe fn prepare_typed_object_with_interfaces(
        &self,
        object: calcvba_runtime::ObjectRef,
        members: &[TypeLibMemberMetadata],
        interfaces: &[TypeLibInterfaceMetadata],
    ) -> Result<HostObject, String> {
        // SAFETY: forward the exact caller metadata/apartment contract.
        unsafe { self.prepare_typed_object_with_locale(object, members, interfaces, None) }
    }
    /// Prepare a native table with an explicit invocation locale. It applies to
    /// prebound Invoke and hidden native LCID arguments, including result views.
    /// # Safety
    /// Same trusted descriptor and apartment contract as the default preparation.
    pub unsafe fn prepare_typed_object_with_locale(
        &self,
        object: calcvba_runtime::ObjectRef,
        members: &[TypeLibMemberMetadata],
        interfaces: &[TypeLibInterfaceMetadata],
        locale_id: Option<u32>,
    ) -> Result<HostObject, String> {
        let first = members.first().ok_or("typed COM member list is empty")?;
        let iid = first.interface_iid.ok_or("typed COM IID is missing")?;
        for member in members {
            if member.interface_iid != Some(iid) || !plan::validate(member) {
                return Err("unsupported typed COM member signature or mixed interfaces".into());
            }
        }
        if members.iter().any(plan::is_dispatch)
            && !interfaces
                .iter()
                .any(|interface| interface.iid == Some(iid))
        {
            return Err(
                "typed dispinterface preparation requires its complete compiler interface table"
                    .into(),
            );
        }
        let (library, root_index) = NativeLibrary::new(interfaces.to_vec(), members, locale_id)?;
        let library = Rc::new(library);
        let selected: Vec<_> = members
            .iter()
            .map(|member| {
                library.plans[root_index]
                    .iter()
                    .find(|p| p.member == *member)
                    .cloned()
                    .ok_or_else(|| {
                        format!(
                            "native member {} lacks a retained result/interface plan",
                            member.name
                        )
                    })
            })
            .collect::<Result<_, _>>()?;
        let dispatch = native_input(self, &object)?;
        // SAFETY: retained same-apartment dispatch and validated selected members.
        unsafe { build(self.clone(), dispatch.as_ptr(), iid, &selected, library) }
    }
}
/// # Safety
/// pointer borrows a live same-apartment native object. The optional interface
/// index was resolved at preparation; no names/type information are read here.
pub(super) unsafe fn make_native_object(
    bridge: WindowsComBridge,
    pointer: *mut c_void,
    interface: Option<usize>,
    library: Rc<NativeLibrary>,
) -> Result<HostObject, String> {
    if let Some(index) = interface {
        let iid = library
            .metadata
            .get(index)
            .and_then(|m| m.iid)
            .ok_or("native result IID is missing")?;
        let members = library
            .plans
            .get(index)
            .ok_or("native result plan is missing")?;
        // SAFETY: exact retained interface metadata; build retains its own QI.
        unsafe { build(bridge, pointer, iid, members, library.clone()) }
    } else {
        // Opaque Variant objects retain native ownership/round-trip only. No
        // guessed default/interface/member plan is created for an unknown IID.
        let iid = crate::ComInterfaceIid::from_guid(&crate::windows_client::IID_IDISPATCH);
        // SAFETY: decoder supplied IDispatch; no custom member is exposed.
        unsafe { build(bridge, pointer, iid, &[], library) }
    }
}
unsafe fn build(
    bridge: WindowsComBridge,
    pointer: *mut c_void,
    iid: crate::ComInterfaceIid,
    members: &[NativeMemberPlan],
    library: Rc<NativeLibrary>,
) -> Result<HostObject, String> {
    if members.iter().any(|m| !plan::is_dispatch(&m.member)) {
        // SAFETY: input borrows a retained same-apartment native dual interface.
        let proxy = unsafe { crate::windows_invoke::dispatch_is_marshaling_proxy(pointer) };
        if proxy
            && !crate::windows_invoke::proxy_interface_is_vtable_safe(iid, bridge.shared_state())
        {
            return Err("typed COM receiver proxy is not typelib-aligned".into());
        }
    }
    // SAFETY: exact imported IID; successful QI transfers one owned reference.
    let raw = unsafe { crate::query_interface_pointer(pointer, &iid.to_guid()) }?;
    // SAFETY: successful QI above transferred one reference; retain it before
    // the next fallible operation so failed identity QI cannot leak it.
    let pointer = Rc::new(unsafe { OwnedComInterface::from_owned(raw) });
    // SAFETY: this same-apartment interface is live. COM identity is the exact
    // pointer returned by QI(IID_IUnknown), never an arbitrary interface address.
    let identity = unsafe {
        crate::query_interface_pointer(pointer.as_ptr(), &crate::windows_client::IID_IUNKNOWN)
    }?;
    // SAFETY: successful QI transfers exactly one retained identity reference.
    let identity_owner = unsafe { OwnedComInterface::from_owned(identity) };
    let interface_members = library
        .metadata
        .iter()
        .find(|m| m.iid == Some(iid))
        .map(|m| m.members.as_slice())
        .unwrap_or(&[]);
    let indexed = members
        .iter()
        .map(|m| {
            plan::table_slot(&m.member, interface_members)
                .map(|slot| (slot, m))
                .ok_or("native member has no typed table index")
        })
        .collect::<Result<Vec<_>, _>>()?;
    let receiver = Rc::new(NativeReceiver {
        pointer,
        _identity: identity_owner,
        library,
        bridge,
        enumerator: members
            .iter()
            .find(|m| {
                m.member.token == -4
                    && matches!(
                        m.member.invoke_kind,
                        crate::TypeLibMemberInvokeKind::PropertyGet
                            | crate::TypeLibMemberInvokeKind::Method
                    )
            })
            .cloned(),
    });
    let bound = indexed
        .iter()
        .map(|(slot, _)| *slot)
        .max()
        .map_or(0, |slot| slot + 1);
    let mut slots = Vec::with_capacity(bound);
    for index in 0..bound {
        let mut candidates = indexed.iter().filter(|(slot, _)| *slot == index);
        let Some((_, plan)) = candidates.next() else {
            slots.push(
                HostSlot::typed(vec![], T::Void, |_, _| Err(438)).map_err(|e| e.to_string())?,
            );
            continue;
        };
        if candidates.next().is_some() {
            return Err("duplicate native typed slot".into());
        }
        let member = &plan.member;
        let owner = receiver.clone();
        if let Some((parameters, result)) = scalar_signature(member) {
            let lcid = member.parameter_optional_defaults.last() == Some(&D::Lcid);
            slots.push(
                HostSlot::typed(parameters, result, move |context, args| {
                    crate::windows_vtable::clear_thread_error_info();
                    let value = calcvba_host_api::with_native_call(context, || {
                        // SAFETY: retained QI'd receiver and exact checked scalar ABI.
                        unsafe {
                            super::native::invoke(
                                owner.pointer.as_ptr(),
                                index,
                                result,
                                lcid,
                                owner.library.locale_id.unwrap_or(0),
                                args,
                            )
                        }
                    });
                    if calcvba_host_api::has_unhandled_event_fault(context) {
                        return Err(calcvba_host_api::UNHANDLED_EVENT_ABORT);
                    }
                    let value = value.map_err(|e| e.publish(context))?;
                    owner.bridge.record_typed_vtable_call();
                    Ok(value)
                })
                .map_err(|e| e.to_string())?,
            );
        } else {
            let plan = (**plan).clone();
            slots.push(HostSlot::planned(move |context, args| {
                let value = calcvba_host_api::with_native_call(context, || {
                    // SAFETY: retained QI'd receiver and immutable admitted descriptor.
                    unsafe { super::native_values::invoke(&owner, &plan, args) }
                });
                if calcvba_host_api::has_unhandled_event_fault(context) {
                    return Err(calcvba_host_api::UNHANDLED_EVENT_ABORT);
                }
                value.map_err(|e| e.publish(context))
            }));
        }
    }
    let mut bytes = [0; 16];
    bytes[..4].copy_from_slice(&iid.data1.to_le_bytes());
    bytes[4..6].copy_from_slice(&iid.data2.to_le_bytes());
    bytes[6..8].copy_from_slice(&iid.data3.to_le_bytes());
    bytes[8..].copy_from_slice(&iid.data4);
    let services = Rc::new(NativeServices(receiver.clone()));
    // SAFETY: payload/services retain the canonical native IUnknown identity.
    Ok(unsafe { HostObject::new_with_services(bytes, slots, receiver, identity.cast(), services) })
}
