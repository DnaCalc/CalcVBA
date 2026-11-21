use core::ffi::c_void;
use core::sync::atomic::{AtomicBool, AtomicU32, Ordering};
use std::cell::RefCell;
use std::collections::{BTreeMap, BTreeSet};
use std::hash::Hash;

use crate::Variant;
pub use calcvba_automation::object_ref::{ObjectRef, RawRuntimeIUnknown};

pub const RUNTIME_S_OK: i32 = 0;
pub const RUNTIME_E_NOINTERFACE: i32 = 0x8000_4002u32 as i32;
/// `E_POINTER`: a required out or in pointer was null.
pub const RUNTIME_E_POINTER: i32 = 0x8000_4003u32 as i32;

#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum RuntimeInterfaceId {
    IUnknown,
    IDispatch,
    IConnectionPointContainer,
    IConnectionPoint,
    Unsupported,
}

pub use calcvba_automation::com::{Guid, IUnknownVtbl};

pub const RUNTIME_GUID_IUNKNOWN: Guid =
    Guid::new(0x0000_0000, 0x0000, 0x0000, [0xC0, 0, 0, 0, 0, 0, 0, 0x46]);
pub const RUNTIME_GUID_IDISPATCH: Guid =
    Guid::new(0x0002_0400, 0x0000, 0x0000, [0xC0, 0, 0, 0, 0, 0, 0, 0x46]);
pub const RUNTIME_GUID_ICONNECTIONPOINTCONTAINER: Guid = Guid::new(
    0xB196_B284,
    0xBAB4,
    0x101A,
    [0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07],
);
pub const RUNTIME_GUID_ICONNECTIONPOINT: Guid = Guid::new(
    0xB196_B286,
    0xBAB4,
    0x101A,
    [0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07],
);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum RuntimeInterfaceKind {
    Unknown,
    Dispatch,
    Dual,
    Source,
    ConnectionPointContainer,
    ConnectionPoint,
    Custom,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct RuntimeInterfaceIdentity {
    pub id: RuntimeInterfaceId,
    pub guid: Guid,
    pub name: &'static str,
    pub kind: RuntimeInterfaceKind,
    pub major_version: Option<u16>,
    pub minor_version: Option<u16>,
    pub lcid: Option<u32>,
}

impl RuntimeInterfaceIdentity {
    pub const fn new(
        id: RuntimeInterfaceId,
        guid: Guid,
        name: &'static str,
        kind: RuntimeInterfaceKind,
        major_version: Option<u16>,
        minor_version: Option<u16>,
        lcid: Option<u32>,
    ) -> Self {
        Self {
            id,
            guid,
            name,
            kind,
            major_version,
            minor_version,
            lcid,
        }
    }

    pub const fn custom(
        guid: Guid,
        name: &'static str,
        kind: RuntimeInterfaceKind,
        major_version: Option<u16>,
        minor_version: Option<u16>,
        lcid: Option<u32>,
    ) -> Self {
        Self::new(
            RuntimeInterfaceId::Unsupported,
            guid,
            name,
            kind,
            major_version,
            minor_version,
            lcid,
        )
    }
}

pub const RUNTIME_IUNKNOWN_INTERFACE_IDENTITY: RuntimeInterfaceIdentity =
    RuntimeInterfaceIdentity::new(
        RuntimeInterfaceId::IUnknown,
        RUNTIME_GUID_IUNKNOWN,
        "IUnknown",
        RuntimeInterfaceKind::Unknown,
        None,
        None,
        None,
    );
pub const RUNTIME_IDISPATCH_INTERFACE_IDENTITY: RuntimeInterfaceIdentity =
    RuntimeInterfaceIdentity::new(
        RuntimeInterfaceId::IDispatch,
        RUNTIME_GUID_IDISPATCH,
        "IDispatch",
        RuntimeInterfaceKind::Dispatch,
        None,
        None,
        None,
    );
pub const RUNTIME_ICONNECTIONPOINTCONTAINER_INTERFACE_IDENTITY: RuntimeInterfaceIdentity =
    RuntimeInterfaceIdentity::new(
        RuntimeInterfaceId::IConnectionPointContainer,
        RUNTIME_GUID_ICONNECTIONPOINTCONTAINER,
        "IConnectionPointContainer",
        RuntimeInterfaceKind::ConnectionPointContainer,
        None,
        None,
        None,
    );
pub const RUNTIME_ICONNECTIONPOINT_INTERFACE_IDENTITY: RuntimeInterfaceIdentity =
    RuntimeInterfaceIdentity::new(
        RuntimeInterfaceId::IConnectionPoint,
        RUNTIME_GUID_ICONNECTIONPOINT,
        "IConnectionPoint",
        RuntimeInterfaceKind::ConnectionPoint,
        None,
        None,
        None,
    );

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuntimeLifetimePolicy {
    RefCounted,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuntimeApartmentModel {
    Unknown,
    SingleThreaded,
    MultiThreaded,
    Both,
    Neutral,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeObjectIdentity {
    pub stable_object_id: u64,
    /// Per-instance identity key (unique per allocated object). For a fresh project-class
    /// instance this is a VM-allocated unique id; the object's true identity for `Is` is its
    /// IUnknown pointer, this is the internal lookup handle 1:1 with it.
    pub compat_identity: i32,
    /// Class/route key: which project-dynamic route (class dispatch table) this instance
    /// belongs to. Distinct from `compat_identity` once instances are per-`New` allocations â€”
    /// many instances share one `route_key` but each has its own `compat_identity`.
    pub route_key: i32,
    /// The bundle (".NET assembly") that owns this instance's class. `(bundle_id,
    /// route_key)` together select the class dispatch table for a multi-bundle VM;
    /// `0` for single-bundle runs. Identity (`Is`) remains the IUnknown pointer.
    pub bundle_id: i32,
    /// True when this instance's class has a `Class_Terminate` â€” so the last `Release`
    /// (`compat_release` at refcount 0) enqueues it for the VM to run `Class_Terminate`.
    pub terminates_on_release: bool,
    pub class_descriptor: &'static RuntimeClassDescriptor,
    pub lifetime_policy: RuntimeLifetimePolicy,
    pub apartment_model: RuntimeApartmentModel,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeInterfaceProjection {
    pub object_identity: RuntimeObjectIdentity,
    pub interface_identity: RuntimeInterfaceIdentity,
    pub interface_descriptor: &'static RuntimeInterfaceDescriptor,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum RuntimeMemberInvokeKind {
    Method,
    PropertyGet,
    PropertyLet,
    PropertySet,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum RuntimeValueType {
    Variant,
    Long,
    Integer,
    String,
    Boolean,
    Double,
    Single,
    Currency,
    Date,
    Decimal,
    Object,
    Byte,
    LongLong,
    LongPtr,
    Record,
    Array,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum RuntimeMemberBacking {
    /// A compiled/interpreted VBA procedure (`FuncId(proc_index)` in the object's program).
    Proc(usize),
    /// A public class field backed directly by `CompatObjectBase.fields[token]`.
    Field(i32),
    /// No direct runtime target (e.g. imported typelib or synthetic test descriptor).
    None,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeParamDescriptor {
    pub name: &'static str,
    pub value_type: RuntimeValueType,
    pub by_ref: bool,
    pub optional: bool,
    pub param_array: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeMemberDescriptor {
    pub name: &'static str,
    pub dispatch_id: i32,
    pub vtable_slot: Option<u16>,
    pub invoke_kind: RuntimeMemberInvokeKind,
    pub arity: usize,
    pub params: &'static [RuntimeParamDescriptor],
    pub return_type: Option<RuntimeValueType>,
    pub is_default_member: bool,
    pub is_enumerator_member: bool,
    pub externally_visible: bool,
    pub backing: RuntimeMemberBacking,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeClassFieldDescriptor {
    pub name: &'static str,
    pub token: i32,
    pub value_type: RuntimeValueType,
    pub array_element_type: Option<RuntimeValueType>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuntimeClassActivationDescriptor {
    ProjectClass { class_index: usize },
    ExternClass { import_index: usize },
    ComClass { prog_id: &'static str },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeClassAsNewFieldDescriptor {
    pub field_token: i32,
    pub activation: RuntimeClassActivationDescriptor,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RuntimeClassLifecycleDescriptor {
    pub has_initialize: bool,
    pub has_terminate: bool,
}

pub const RUNTIME_CLASS_LIFECYCLE_NONE: RuntimeClassLifecycleDescriptor =
    RuntimeClassLifecycleDescriptor {
        has_initialize: false,
        has_terminate: false,
    };

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeInterfaceDescriptor {
    pub id: RuntimeInterfaceId,
    pub identity: RuntimeInterfaceIdentity,
    pub name: &'static str,
    pub members: &'static [RuntimeMemberDescriptor],
    pub dual_dispatch: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeProjectClassIdentity {
    pub unit_name: &'static str,
    pub class_index: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeClassDescriptor {
    pub name: &'static str,
    /// Public type name of an engine-supplied class, separate from its nominal identity.
    pub runtime_type_name: Option<&'static str>,
    pub project_identity: Option<RuntimeProjectClassIdentity>,
    /// Default dual-interface IID for this class (`Guid::deterministic(unit, "I{Class}")`).
    pub default_interface_iid: Option<Guid>,
    pub predeclared: bool,
    pub lifecycle: RuntimeClassLifecycleDescriptor,
    pub fields: &'static [RuntimeClassFieldDescriptor],
    pub as_new_fields: &'static [RuntimeClassAsNewFieldDescriptor],
    pub implements: &'static [&'static str],
    pub interfaces: &'static [RuntimeInterfaceDescriptor],
}

/// Structural issue that makes a runtime descriptor unsafe or ambiguous for future COM export.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RuntimeDescriptorIssue {
    InterfaceIdentityMismatch {
        interface: &'static str,
        descriptor_id: RuntimeInterfaceId,
        identity_id: RuntimeInterfaceId,
    },
    DuplicateInterfaceGuid {
        class: &'static str,
        interface: &'static str,
        guid: Guid,
    },
    BadMemberArity {
        interface: &'static str,
        member: &'static str,
        arity: usize,
        params: usize,
    },
    BadDefaultMemberDispatchId {
        interface: &'static str,
        member: &'static str,
        dispatch_id: i32,
    },
    BadEnumeratorDispatchId {
        interface: &'static str,
        member: &'static str,
        dispatch_id: i32,
    },
    BadDualDispatchVtableSlot {
        interface: &'static str,
        member: &'static str,
        vtable_slot: u16,
    },
    BadPropertyGetReturn {
        interface: &'static str,
        member: &'static str,
    },
    BadPropertySetterShape {
        interface: &'static str,
        member: &'static str,
    },
}

impl RuntimeInterfaceDescriptor {
    /// Return the member at `index`, if in bounds.
    pub fn member_at(&self, index: usize) -> Option<&'static RuntimeMemberDescriptor> {
        self.members.get(index)
    }

    /// Look up a member by vtable slot (`>= 7` on a dual interface).
    ///
    /// On canonical descriptors built by [`crate::build_runtime_class_descriptor`], dual members
    /// occupy consecutive slots starting at `7`, so the lookup resolves in O(1) without scanning.
    pub fn member_by_vtable_slot(&self, slot: u16) -> Option<&'static RuntimeMemberDescriptor> {
        if self.dual_dispatch
            && slot >= 7
            && let Some(member) = self.members.get((slot - 7) as usize)
            && member.vtable_slot == Some(slot)
        {
            return Some(member);
        }
        self.members
            .iter()
            .find(|member| member.vtable_slot == Some(slot))
    }

    /// Resolve a member on this interface by DISPID and invoke kind using VBA accessor fallback
    /// rules (`PropertyGet <-> Method`, and `PropertySet` onto a Variant public field).
    pub fn resolve_member_by_dispid(
        &self,
        dispid: i32,
        invoke_kind: RuntimeMemberInvokeKind,
        args_empty: bool,
        public_only: bool,
    ) -> Option<&'static RuntimeMemberDescriptor> {
        let exact = self
            .members
            .iter()
            .find(|member| member.dispatch_id == dispid && member.invoke_kind == invoke_kind);
        let member = exact.or_else(|| match invoke_kind {
            RuntimeMemberInvokeKind::PropertyGet if args_empty || dispid == 0 => {
                self.members.iter().find(|member| {
                    member.dispatch_id == dispid
                        && member.invoke_kind == RuntimeMemberInvokeKind::Method
                })
            }
            RuntimeMemberInvokeKind::Method => self.members.iter().find(|member| {
                member.dispatch_id == dispid
                    && member.invoke_kind == RuntimeMemberInvokeKind::PropertyGet
            }),
            RuntimeMemberInvokeKind::PropertySet => self
                .members
                .iter()
                .find(|member| member.dispatch_id == dispid && is_variant_field_setter(member)),
            _ => None,
        });
        member.filter(|member| !public_only || member.externally_visible)
    }

    /// Validate only the COM-publication shape of this descriptor. Internal dispatch-cache test
    /// descriptors may use narrower synthetic slots and should not call this export check.
    pub fn validate_com_export_shape(&self) -> Result<(), RuntimeDescriptorIssue> {
        if self.id != self.identity.id {
            return Err(RuntimeDescriptorIssue::InterfaceIdentityMismatch {
                interface: self.name,
                descriptor_id: self.id,
                identity_id: self.identity.id,
            });
        }
        for member in self.members {
            if member.arity != member.params.len() {
                return Err(RuntimeDescriptorIssue::BadMemberArity {
                    interface: self.name,
                    member: member.name,
                    arity: member.arity,
                    params: member.params.len(),
                });
            }
            if member.is_default_member && member.dispatch_id != 0 {
                return Err(RuntimeDescriptorIssue::BadDefaultMemberDispatchId {
                    interface: self.name,
                    member: member.name,
                    dispatch_id: member.dispatch_id,
                });
            }
            if member.is_enumerator_member && member.dispatch_id != -4 {
                return Err(RuntimeDescriptorIssue::BadEnumeratorDispatchId {
                    interface: self.name,
                    member: member.name,
                    dispatch_id: member.dispatch_id,
                });
            }
            if self.dual_dispatch
                && let Some(vtable_slot) = member.vtable_slot
                && vtable_slot < 7
            {
                return Err(RuntimeDescriptorIssue::BadDualDispatchVtableSlot {
                    interface: self.name,
                    member: member.name,
                    vtable_slot,
                });
            }
            match member.invoke_kind {
                RuntimeMemberInvokeKind::PropertyGet if member.return_type.is_none() => {
                    return Err(RuntimeDescriptorIssue::BadPropertyGetReturn {
                        interface: self.name,
                        member: member.name,
                    });
                }
                RuntimeMemberInvokeKind::PropertyLet | RuntimeMemberInvokeKind::PropertySet
                    if member.return_type.is_some() || member.params.is_empty() =>
                {
                    return Err(RuntimeDescriptorIssue::BadPropertySetterShape {
                        interface: self.name,
                        member: member.name,
                    });
                }
                _ => {}
            }
        }
        Ok(())
    }
}

impl RuntimeClassDescriptor {
    pub fn type_name(&self) -> &'static str {
        self.runtime_type_name.unwrap_or(self.name)
    }

    /// The class's default dual `IDispatch` interface descriptor, if present.
    pub fn dispatch_interface(&self) -> Option<&'static RuntimeInterfaceDescriptor> {
        self.interfaces
            .iter()
            .find(|interface| interface.id == RuntimeInterfaceId::IDispatch)
    }

    /// Look up a dual-interface member by DISPID and invoke kind.
    pub fn member_by_dispid(
        &self,
        dispid: i32,
        invoke_kind: RuntimeMemberInvokeKind,
        public_only: bool,
    ) -> Option<&'static RuntimeMemberDescriptor> {
        let dispatch = self.dispatch_interface()?;
        dispatch.members.iter().find(|member| {
            member.dispatch_id == dispid
                && member.invoke_kind == invoke_kind
                && (!public_only || member.externally_visible)
        })
    }

    /// Resolve a dual-interface member by DISPID using VBA accessor fallback rules.
    pub fn resolve_member_by_dispid(
        &self,
        dispid: i32,
        invoke_kind: RuntimeMemberInvokeKind,
        args_empty: bool,
        public_only: bool,
    ) -> Option<&'static RuntimeMemberDescriptor> {
        self.dispatch_interface()?.resolve_member_by_dispid(
            dispid,
            invoke_kind,
            args_empty,
            public_only,
        )
    }

    /// Resolve the class's `For Each` enumerator member (`DISPID_NEWENUM = -4`).
    pub fn resolve_enumerator_member(&self) -> Option<&'static RuntimeMemberDescriptor> {
        self.resolve_member_by_dispid(-4, RuntimeMemberInvokeKind::PropertyGet, true, false)
    }

    /// Look up a dual-interface member by vtable slot (`>= 7`).
    pub fn member_by_vtable_slot(&self, slot: u16) -> Option<&'static RuntimeMemberDescriptor> {
        self.dispatch_interface()?.member_by_vtable_slot(slot)
    }

    /// Validate Linux-testable descriptor facts that future COM export code must be able to trust.
    pub fn validate_com_export_shape(&self) -> Result<(), RuntimeDescriptorIssue> {
        let mut seen_guids = BTreeSet::new();
        for interface in self.interfaces {
            if !seen_guids.insert(interface.identity.guid) {
                return Err(RuntimeDescriptorIssue::DuplicateInterfaceGuid {
                    class: self.name,
                    interface: interface.name,
                    guid: interface.identity.guid,
                });
            }
            interface.validate_com_export_shape()?;
        }
        Ok(())
    }
}

pub const RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR: RuntimeInterfaceDescriptor =
    RuntimeInterfaceDescriptor {
        id: RuntimeInterfaceId::IUnknown,
        identity: RUNTIME_IUNKNOWN_INTERFACE_IDENTITY,
        name: "IUnknown",
        members: &[],
        dual_dispatch: false,
    };

fn runtime_iunknown_interface_descriptor() -> &'static RuntimeInterfaceDescriptor {
    static DESCRIPTOR: RuntimeInterfaceDescriptor = RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR;
    &DESCRIPTOR
}

pub const COMPAT_OBJECT_CLASS_DESCRIPTOR: RuntimeClassDescriptor = RuntimeClassDescriptor {
    name: "OxVba.CompatObject",
    runtime_type_name: None,
    project_identity: None,
    default_interface_iid: None,
    predeclared: false,
    lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
    fields: &[],
    as_new_fields: &[],
    implements: &[],
    interfaces: &[RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR],
};

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
pub struct RuntimeDispatchCacheKey {
    pub interface_descriptor: usize,
    pub interface_id: RuntimeInterfaceId,
    pub normalized_member_name: String,
    pub invoke_kind: RuntimeMemberInvokeKind,
    pub arity: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuntimeDispatchPlan {
    pub interface_id: RuntimeInterfaceId,
    pub member_index: usize,
    pub dispatch_id: i32,
    pub vtable_slot: Option<u16>,
    pub invoke_kind: RuntimeMemberInvokeKind,
    pub is_default_member: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
struct RuntimeProjectDispatchKey {
    interface_descriptor: usize,
    normalized_member_name: String,
    invoke_kind: Option<RuntimeMemberInvokeKind>,
    args_empty: bool,
    public_only: bool,
    host_fallback: bool,
}

#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct RuntimeDispatchPlanCache {
    entries: BTreeMap<RuntimeDispatchCacheKey, RuntimeDispatchPlan>,
    project_entries: BTreeMap<RuntimeProjectDispatchKey, RuntimeDispatchPlan>,
}

impl RuntimeDispatchPlanCache {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn len(&self) -> usize {
        self.entries.len() + self.project_entries.len()
    }

    pub fn project_plan_count(&self) -> usize {
        self.project_entries.len()
    }

    pub fn is_empty(&self) -> bool {
        self.entries.is_empty() && self.project_entries.is_empty()
    }

    /// Resolve a VBA late-bound member on a project interface through the plan cache, using VBA's
    /// accessor fallback rules (`PropertyGet(args_empty) <-> Method`, and `PropertySet` onto a
    /// Variant public field). Subsequent calls with the same `(interface, name, kind, args_empty,
    /// public_only)` return the cached plan without scanning members.
    pub fn resolve_project_member(
        &mut self,
        interface: &RuntimeInterfaceDescriptor,
        member_name: &str,
        invoke_kind: RuntimeMemberInvokeKind,
        args_empty: bool,
        public_only: bool,
    ) -> Option<RuntimeDispatchPlan> {
        let key = RuntimeProjectDispatchKey {
            interface_descriptor: runtime_interface_descriptor_key(interface),
            normalized_member_name: normalize_runtime_member_name(member_name),
            invoke_kind: Some(invoke_kind),
            args_empty,
            public_only,
            host_fallback: false,
        };
        if let Some(plan) = self.project_entries.get(&key).copied() {
            return Some(plan);
        }
        let exact = interface.members.iter().enumerate().find(|(_, candidate)| {
            candidate.name.eq_ignore_ascii_case(member_name.trim())
                && candidate.invoke_kind == invoke_kind
        });
        let matched = exact.or_else(|| match invoke_kind {
            RuntimeMemberInvokeKind::PropertyGet if args_empty => {
                interface.members.iter().enumerate().find(|(_, candidate)| {
                    candidate.name.eq_ignore_ascii_case(member_name.trim())
                        && candidate.invoke_kind == RuntimeMemberInvokeKind::Method
                })
            }
            RuntimeMemberInvokeKind::Method => {
                interface.members.iter().enumerate().find(|(_, candidate)| {
                    candidate.name.eq_ignore_ascii_case(member_name.trim())
                        && candidate.invoke_kind == RuntimeMemberInvokeKind::PropertyGet
                })
            }
            RuntimeMemberInvokeKind::PropertySet => {
                interface.members.iter().enumerate().find(|(_, candidate)| {
                    candidate.name.eq_ignore_ascii_case(member_name.trim())
                        && is_variant_field_setter(candidate)
                })
            }
            _ => None,
        });
        let (member_index, member) =
            matched.filter(|(_, candidate)| !public_only || candidate.externally_visible)?;
        let plan = RuntimeDispatchPlan {
            interface_id: interface.id,
            member_index,
            dispatch_id: member.dispatch_id,
            vtable_slot: member.vtable_slot,
            invoke_kind: member.invoke_kind,
            is_default_member: member.is_default_member,
        };
        self.project_entries.insert(key, plan);
        Some(plan)
    }

    /// Resolve a host-initiated project member call (`kind_hint` preferred, falling back to any
    /// accessor with the same name) through the plan cache.
    pub fn resolve_project_member_for_host(
        &mut self,
        interface: &RuntimeInterfaceDescriptor,
        member_name: &str,
        kind_hint: Option<RuntimeMemberInvokeKind>,
    ) -> Option<RuntimeDispatchPlan> {
        let key = RuntimeProjectDispatchKey {
            interface_descriptor: runtime_interface_descriptor_key(interface),
            normalized_member_name: normalize_runtime_member_name(member_name),
            invoke_kind: kind_hint,
            args_empty: false,
            public_only: false,
            host_fallback: true,
        };
        if let Some(plan) = self.project_entries.get(&key).copied() {
            return Some(plan);
        }
        let trimmed = member_name.trim();
        let matched = interface
            .members
            .iter()
            .enumerate()
            .find(|(_, candidate)| {
                candidate.name.eq_ignore_ascii_case(trimmed)
                    && kind_hint.is_none_or(|kind| candidate.invoke_kind == kind)
            })
            .or_else(|| {
                if kind_hint == Some(RuntimeMemberInvokeKind::PropertySet) {
                    interface.members.iter().enumerate().find(|(_, candidate)| {
                        candidate.name.eq_ignore_ascii_case(trimmed)
                            && is_variant_field_setter(candidate)
                    })
                } else {
                    None
                }
            })
            .or_else(|| {
                interface
                    .members
                    .iter()
                    .enumerate()
                    .find(|(_, candidate)| candidate.name.eq_ignore_ascii_case(trimmed))
            });
        let (member_index, member) = matched?;
        let plan = RuntimeDispatchPlan {
            interface_id: interface.id,
            member_index,
            dispatch_id: member.dispatch_id,
            vtable_slot: member.vtable_slot,
            invoke_kind: member.invoke_kind,
            is_default_member: member.is_default_member,
        };
        self.project_entries.insert(key, plan);
        Some(plan)
    }

    pub fn resolve_member(
        &mut self,
        interface: &RuntimeInterfaceDescriptor,
        member_name: &str,
        invoke_kind: RuntimeMemberInvokeKind,
        arity: usize,
    ) -> Option<RuntimeDispatchPlan> {
        let key = RuntimeDispatchCacheKey {
            interface_descriptor: runtime_interface_descriptor_key(interface),
            interface_id: interface.id,
            normalized_member_name: normalize_runtime_member_name(member_name),
            invoke_kind,
            arity,
        };
        self.resolve_with_key(interface, key, |candidate, key| {
            normalize_runtime_member_name(candidate.name) == key.normalized_member_name
                && candidate.invoke_kind == key.invoke_kind
                && candidate.arity == key.arity
        })
    }

    pub fn resolve_default_member(
        &mut self,
        interface: &RuntimeInterfaceDescriptor,
        invoke_kind: RuntimeMemberInvokeKind,
        arity: usize,
    ) -> Option<RuntimeDispatchPlan> {
        let key = RuntimeDispatchCacheKey {
            interface_descriptor: runtime_interface_descriptor_key(interface),
            interface_id: interface.id,
            normalized_member_name: "<default>".to_string(),
            invoke_kind,
            arity,
        };
        self.resolve_with_key(interface, key, |candidate, key| {
            candidate.is_default_member
                && candidate.invoke_kind == key.invoke_kind
                && candidate.arity == key.arity
        })
    }

    pub fn resolve_member_unhinted(
        &mut self,
        interface: &RuntimeInterfaceDescriptor,
        member_name: &str,
        arity: usize,
    ) -> Option<RuntimeDispatchPlan> {
        let normalized_member_name = normalize_runtime_member_name(member_name);
        self.resolve_unhinted_with_name(
            interface,
            normalized_member_name,
            arity,
            |candidate, name| normalize_runtime_member_name(candidate.name) == *name,
        )
    }

    pub fn resolve_default_member_unhinted(
        &mut self,
        interface: &RuntimeInterfaceDescriptor,
        arity: usize,
    ) -> Option<RuntimeDispatchPlan> {
        self.resolve_unhinted_with_name(
            interface,
            "<default>".to_string(),
            arity,
            |candidate, _| candidate.is_default_member,
        )
    }

    fn resolve_unhinted_with_name(
        &mut self,
        interface: &RuntimeInterfaceDescriptor,
        normalized_member_name: String,
        arity: usize,
        matches_name: impl Fn(&RuntimeMemberDescriptor, &String) -> bool,
    ) -> Option<RuntimeDispatchPlan> {
        let mut matches = interface
            .members
            .iter()
            .enumerate()
            .filter(|(_, candidate)| {
                matches_name(candidate, &normalized_member_name) && candidate.arity == arity
            });
        let (member_index, member) = matches.next()?;
        if matches.next().is_some() {
            return None;
        }
        let key = RuntimeDispatchCacheKey {
            interface_descriptor: runtime_interface_descriptor_key(interface),
            interface_id: interface.id,
            normalized_member_name,
            invoke_kind: member.invoke_kind,
            arity,
        };
        self.resolve_with_key(interface, key, |candidate, key| {
            candidate.invoke_kind == key.invoke_kind
                && candidate.arity == key.arity
                && candidate.name == interface.members[member_index].name
                && candidate.dispatch_id == interface.members[member_index].dispatch_id
        })
    }

    fn resolve_with_key(
        &mut self,
        interface: &RuntimeInterfaceDescriptor,
        key: RuntimeDispatchCacheKey,
        matches: impl Fn(&RuntimeMemberDescriptor, &RuntimeDispatchCacheKey) -> bool,
    ) -> Option<RuntimeDispatchPlan> {
        if let Some(plan) = self.entries.get(&key).copied() {
            return Some(plan);
        }
        let mut matches = interface
            .members
            .iter()
            .enumerate()
            .filter(|(_, candidate)| matches(candidate, &key));
        let (member_index, member) = matches.next()?;
        if matches.next().is_some() {
            return None;
        }
        let plan = RuntimeDispatchPlan {
            interface_id: interface.id,
            member_index,
            dispatch_id: member.dispatch_id,
            vtable_slot: member.vtable_slot,
            invoke_kind: member.invoke_kind,
            is_default_member: member.is_default_member,
        };
        self.entries.insert(key, plan);
        Some(plan)
    }
}

fn is_variant_field_setter(member: &RuntimeMemberDescriptor) -> bool {
    matches!(member.backing, RuntimeMemberBacking::Field(_))
        && member.invoke_kind == RuntimeMemberInvokeKind::PropertyLet
        && member
            .params
            .last()
            .is_some_and(|param| param.value_type == RuntimeValueType::Variant)
}

fn normalize_runtime_member_name(member_name: &str) -> String {
    member_name.trim().to_ascii_lowercase()
}

fn runtime_interface_descriptor_key(interface: &RuntimeInterfaceDescriptor) -> usize {
    interface as *const RuntimeInterfaceDescriptor as usize
}

#[repr(C)]
pub(crate) struct CompatObjectBase {
    pub(crate) unknown: RawRuntimeIUnknown,
    pub(crate) ref_count: AtomicU32,
    pub(crate) owner_thread: u64,
    pub(crate) has_dual_vtable: bool,
    pub(crate) identity: RuntimeObjectIdentity,
    pub(crate) object_owner: crate::object_owner::WeakObjectOwner,
    pub(crate) project_instance: bool,
    pub(crate) class_descriptor: &'static RuntimeClassDescriptor,
    /// Optional between-runs member invoker (weak so the object never keeps a session alive).
    pub(crate) invoker: RefCell<Option<std::rc::Weak<RefCell<dyn crate::events::MemberInvoker>>>>,
    // Each field is boxed and kept once created, so its cell has a stable address for
    // `VarPtr` (as a VBA instance's fields live in its memory until it is freed).
    pub(crate) fields: RefCell<BTreeMap<i32, Box<Variant>>>,
    /// Built-in native-object state carried on the box itself (currently a `VBA.Collection`'s
    /// contents). Lazily created on first use via [`ObjectRef::with_native_collection`]; freed
    /// with the box, so there is no VM-side store and no leak. `None` for ordinary project
    /// instances that hold no native collection.
    pub(crate) native_state: RefCell<Option<crate::collection::CollectionData>>,
    pub(crate) parked_for_termination: AtomicBool,
    pub(crate) running_termination: AtomicBool,
    pub(crate) terminated: AtomicBool,
}

pub(crate) static COMPAT_OBJECT_VTBL: IUnknownVtbl = IUnknownVtbl {
    query_interface: compat_query_interface,
    add_ref: compat_add_ref,
    release: compat_release,
};

pub(crate) unsafe extern "system" fn compat_query_interface(
    this: *mut c_void,
    riid: *const Guid,
    ppv: *mut *mut c_void,
) -> i32 {
    if ppv.is_null() {
        return RUNTIME_E_POINTER;
    }
    // SAFETY: per the QueryInterface ABI, `ppv` (checked non-null above) points to a writable
    // out-slot owned by the caller for the duration of this call.
    unsafe {
        *ppv = core::ptr::null_mut();
    }
    if riid.is_null() {
        return RUNTIME_E_POINTER;
    }
    // SAFETY: a non-null REFIID points at the caller's GUID for the duration of the call.
    let iid = unsafe { riid.read() };
    let owner = compat_owner_from_this(this);
    // SAFETY: `this` is a live `CompatObjectBase` kept alive by the caller's COM reference.
    let (has_dual, default_iid) = unsafe {
        (
            (*owner).has_dual_vtable,
            (*owner).class_descriptor.default_interface_iid,
        )
    };
    // A plain three-slot compat object answers `IUnknown` only (dc-yqq0.15). A project instance
    // built with a `ClassDualVtbl` answers `IUnknown`, `IDispatch`, and its class's default dual
    // interface IID with the same `&primary` pointer (one COM identity, dc-yqq0.23).
    let supports_iid = iid == RUNTIME_GUID_IUNKNOWN
        || (has_dual && (iid == RUNTIME_GUID_IDISPATCH || default_iid == Some(iid)));
    if !supports_iid {
        return RUNTIME_E_NOINTERFACE;
    }
    // SAFETY: same caller-owned writable out-slot as above (QueryInterface ABI, non-null checked).
    unsafe {
        *ppv = this;
    }
    // SAFETY: `this` was validated above as a live `CompatObjectBase`, which is exactly
    // `compat_add_ref`'s requirement; the reference taken here is owned by the `*ppv` we
    // just handed to the caller.
    unsafe { compat_add_ref(this) };
    RUNTIME_S_OK
}

/// True when at least one project instance has reached refcount 0 and awaits `Class_Terminate`
/// or owner-thread deferred drop from a foreign-thread final release (dc-yqq0.25).
/// Lock-free atomic check at statement boundaries.
pub fn has_pending_terminations() -> bool {
    let owner = crate::object_owner::RuntimeObjectOwner::current();
    owner.channel().has_pending.load(Ordering::Acquire)
}

/// Free deferred non-terminating boxes whose final `Release` arrived on a foreign thread, running
/// their field drops strictly on the owning VBA thread (which may enqueue cascaded `Class_Terminate`
/// entries on the same channel).
fn drain_pending_drops(channel: &crate::object_owner::OwnerTerminationChannel) {
    loop {
        let drops = {
            let mut q = channel
                .queue
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner());
            let drops = std::mem::take(&mut q.pending_drops);
            channel.has_pending.store(
                !q.pending.is_empty() || !q.pending_drops.is_empty(),
                Ordering::Release,
            );
            drops
        };
        if drops.is_empty() {
            break;
        }
        for crate::object_owner::SendCompatPtr(owner) in drops {
            // SAFETY: sole ownership of this refcount-0 box was transferred into `pending_drops`
            // on final foreign-thread release; dropping fields outside the lock on the owner thread
            // safely cascades any child terminations onto this channel.
            unsafe {
                free_compat_box_on_owner_thread(owner);
            }
        }
    }
}

/// Reclaim a refcount-0 `CompatObjectBase` on the owning thread.
///
/// # Safety
/// The caller owns the sole final reference to `owner` (from `Box::into_raw` in
/// `from_compat_object`) and runs on the owning thread (or after owner expiry).
unsafe fn free_compat_box_on_owner_thread(owner: *mut CompatObjectBase) {
    // SAFETY: forwarded caller contract; clearing `invoker`/`fields`/`native_state` before
    // freeing the box ensures child drops never observe a freed parent.
    unsafe {
        (*owner).invoker.borrow_mut().take();
        (*owner).fields.borrow_mut().clear();
        (*owner).native_state.borrow_mut().take();
        crate::live_counters::object_box_freed();
        drop(Box::from_raw(owner));
    }
}

/// Drains and returns the queued `(instance_id, bundle_id, route_key)` terminations (first
/// draining any foreign-thread deferred drops on the owning thread so cascaded child
/// `Class_Terminate` runs in the same drain). The parked object boxes remain owned by the
/// termination queue until `finish_pending_termination` is called.
pub fn take_pending_terminations() -> Vec<(i32, i32, i32)> {
    let owner = crate::object_owner::RuntimeObjectOwner::current();
    let channel = owner.channel();
    drain_pending_drops(channel);
    let mut q = channel
        .queue
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    let selected = std::mem::take(&mut q.pending);
    channel.has_pending.store(
        !q.pending.is_empty() || !q.pending_drops.is_empty(),
        Ordering::Release,
    );
    selected
}

/// Returns a retained reference to the original parked object whose `Class_Terminate` is about to
/// run. The returned `ObjectRef` is the same boxed object, not a rematerialized identity clone.
pub fn retained_parked_termination_object(instance_id: i32) -> Option<ObjectRef> {
    let owner = crate::object_owner::RuntimeObjectOwner::current();
    let boxed_ptr = {
        let q = owner
            .channel()
            .queue
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        q.parked.get(&instance_id).map(|entry| entry.0)?
    };
    // SAFETY: `boxed_ptr` came from `parked`, which owns the box until `finish_pending_termination`
    // or `discard_pending_channel` removes it on this owning thread. `compat_add_ref` retains it
    // for the returned `ObjectRef`. `CompatObjectBase` is `#[repr(C)]` with `unknown` first.
    unsafe {
        (*boxed_ptr)
            .running_termination
            .store(true, Ordering::Release);
        let unknown = boxed_ptr.cast::<RawRuntimeIUnknown>();
        compat_add_ref(unknown.cast());
        ObjectRef::from_raw_iunknown_owned(unknown)
    }
}

/// Finishes teardown for a parked project object after `Class_Terminate` has returned.
///
/// If `Class_Terminate` resurrected `Me`, the object remains alive and its fields/subscriptions
/// stay intact. Otherwise the field store is dropped before the box is freed; releasing those
/// field Variants may enqueue cascaded child terminations.
pub fn finish_pending_termination(instance_id: i32) -> bool {
    let owner = crate::object_owner::RuntimeObjectOwner::current();
    let boxed_ptr = {
        let mut q = owner
            .channel()
            .queue
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        q.parked.remove(&instance_id).map(|entry| entry.0)
    };
    let Some(boxed_ptr) = boxed_ptr else {
        return false;
    };
    // SAFETY: removing from `parked` transfers sole queue ownership of this box. If
    // `Class_Terminate` resurrected the object (`ref_count > 0`) we un-park it; otherwise nothing
    // else can reach it and `free_compat_box_on_owner_thread` frees it once.
    unsafe {
        (*boxed_ptr)
            .running_termination
            .store(false, Ordering::Release);
        (*boxed_ptr).terminated.store(true, Ordering::Release);
        if (*boxed_ptr).ref_count.load(Ordering::Acquire) > 0 {
            (*boxed_ptr)
                .parked_for_termination
                .store(false, Ordering::Release);
            return false;
        }
        free_compat_box_on_owner_thread(boxed_ptr);
        true
    }
}

/// Resets all pending termination state for the current owner (used when a VM run starts/ends).
pub fn reset_pending_terminations() {
    let owner = crate::object_owner::RuntimeObjectOwner::current();
    discard_pending_channel(owner.channel());
}

pub(crate) fn discard_pending_channel(channel: &crate::object_owner::OwnerTerminationChannel) {
    loop {
        let (parked_owners, pending_drops) = {
            let mut q = channel
                .queue
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner());
            q.pending.clear();
            let parked = std::mem::take(&mut q.parked)
                .into_values()
                .map(|entry| entry.0)
                .collect::<Vec<_>>();
            let drops = std::mem::take(&mut q.pending_drops)
                .into_iter()
                .map(|entry| entry.0)
                .collect::<Vec<_>>();
            channel.has_pending.store(false, Ordering::Release);
            (parked, drops)
        };
        if parked_owners.is_empty() && pending_drops.is_empty() {
            break;
        }
        for owner in pending_drops {
            // SAFETY: sole ownership transferred out of `pending_drops` outside the lock.
            unsafe {
                free_compat_box_on_owner_thread(owner);
            }
        }
        for owner in parked_owners {
            // SAFETY: removing from `parked` transfers sole queue authority; releasing fields
            // outside the lock safely handles cascaded child releases.
            unsafe {
                (*owner).running_termination.store(false, Ordering::Release);
                (*owner).terminated.store(true, Ordering::Release);
                if (*owner).ref_count.load(Ordering::Acquire) > 0 {
                    (*owner)
                        .parked_for_termination
                        .store(false, Ordering::Release);
                    continue;
                }
                free_compat_box_on_owner_thread(owner);
            }
        }
    }
}

pub(crate) unsafe extern "system" fn compat_add_ref(this: *mut c_void) -> u32 {
    let owner = compat_owner_from_this(this);
    // SAFETY: this vtable entry only exists on `CompatObjectBase` boxes (`#[repr(C)]` puts
    // `unknown` first), and the caller holds a reference keeping the box alive; `ref_count` is
    // atomic and safe across threads.
    unsafe { (*owner).ref_count.fetch_add(1, Ordering::AcqRel) + 1 }
}

pub(crate) unsafe extern "system" fn compat_release(this: *mut c_void) -> u32 {
    let owner = compat_owner_from_this(this);
    // SAFETY: as in `compat_add_ref`, `this` is a live `CompatObjectBase` kept alive by the
    // reference the caller is releasing right now. `AcqRel` synchronizes with all prior releases
    // on any thread before inspecting fields below when `remaining == 0`.
    let previous = unsafe { (*owner).ref_count.fetch_sub(1, Ordering::AcqRel) };
    let remaining = previous.saturating_sub(1);
    if remaining == 0 {
        // SAFETY: even at refcount 0 the box is still allocated — this call alone decides whether
        // to park, defer to the owner thread, or free it. `identity`, `owner_thread`, and
        // `object_owner` are immutable for the lifetime of the box.
        let (identity, owner_thread, channel) = unsafe {
            (
                (*owner).identity,
                (*owner).owner_thread,
                (*owner).object_owner.channel.upgrade(),
            )
        };
        let on_owner_thread = crate::class_dispatch::current_thread_id() == owner_thread;
        let owner_alive = channel
            .as_ref()
            .is_some_and(|ch| ch.alive.load(Ordering::Acquire));
        // SAFETY: atomic loads on the still-allocated box; `AcqRel` on `ref_count` above
        // synchronized all prior writes.
        let (is_terminated, is_running) = unsafe {
            (
                (*owner).terminated.load(Ordering::Acquire),
                (*owner).running_termination.load(Ordering::Acquire),
            )
        };
        let should_park =
            identity.terminates_on_release && owner_alive && !is_terminated && !is_running;
        if should_park && let Some(ch) = channel.as_ref() {
            // SAFETY: atomic claim ensures the box is parked at most once across any thread.
            if !unsafe { (*owner).parked_for_termination.swap(true, Ordering::AcqRel) } {
                let mut q = ch
                    .queue
                    .lock()
                    .unwrap_or_else(|poisoned| poisoned.into_inner());
                // Re-check `alive` under the lock so a concurrent `OwnerLease::drop` cannot miss
                // this entry or leave it orphaned in a dead queue.
                if ch.alive.load(Ordering::Acquire) {
                    q.pending.push((
                        identity.compat_identity,
                        identity.bundle_id,
                        identity.route_key,
                    ));
                    q.parked.insert(
                        identity.compat_identity,
                        crate::object_owner::SendCompatPtr(owner),
                    );
                    ch.has_pending.store(true, Ordering::Release);
                    return remaining;
                }
                // SAFETY: owner expired before we locked; undo parking on the still-allocated box
                // so teardown proceeds below.
                unsafe {
                    (*owner)
                        .parked_for_termination
                        .store(false, Ordering::Release);
                }
            } else {
                return remaining;
            }
        }
        // SAFETY: if parked, the owner's termination queue owns the box and will free it.
        if unsafe { (*owner).parked_for_termination.load(Ordering::Acquire) } {
            return remaining;
        }
        // Foreign-thread final release of an object without `Class_Terminate` (or already
        // terminated): never drop fields, `native_state`, or `invoker` on the foreign thread.
        // Transfer the box onto the owner's multi-producer queue so teardown runs on the owner.
        if !on_owner_thread && let Some(ch) = channel.as_ref() {
            let mut q = ch
                .queue
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner());
            if ch.alive.load(Ordering::Acquire) {
                q.pending_drops
                    .push(crate::object_owner::SendCompatPtr(owner));
                ch.has_pending.store(true, Ordering::Release);
                return remaining;
            }
        }
        // SAFETY: on the owning thread (or after owner expiry), sole final ownership of `owner`.
        unsafe {
            free_compat_box_on_owner_thread(owner);
        }
    }
    remaining
}

fn compat_field_get(owner: *mut CompatObjectBase, token: i32) -> Option<Variant> {
    // SAFETY: the only caller (`ObjectRef::project_field_get`) guards with
    // `is_project_instance()` (a vtbl-identity check), so `owner` is a live `CompatObjectBase`
    // kept alive by that caller's retained `ObjectRef` for the duration of the borrow.
    unsafe {
        (*owner)
            .fields
            .borrow()
            .get(&token)
            .map(|cell| (**cell).clone())
            .or_else(|| compat_field_default(owner, token))
    }
}

fn compat_field_set(owner: *mut CompatObjectBase, token: i32, value: Variant) {
    // SAFETY: as in `compat_field_get`, the sole caller (`ObjectRef::project_field_set`)
    // verified via `is_project_instance()` that `owner` is a live `CompatObjectBase` kept
    // alive by the caller's retained `ObjectRef`.
    unsafe {
        let mut fields = (*owner).fields.borrow_mut();
        match fields.get_mut(&token) {
            // Replace in place: the cell keeps its address.
            Some(cell) => **cell = value,
            None if matches!(value.vtype(), crate::VarType::Empty) => {}
            None => {
                fields.insert(token, Box::new(value));
            }
        }
    }
}

/// The value an absent field reads as: `Nothing` for an object field, else Empty.
///
/// # Safety
/// `owner` is a live `CompatObjectBase`.
unsafe fn compat_field_default(owner: *mut CompatObjectBase, token: i32) -> Option<Variant> {
    // SAFETY: forwarded from the caller.
    unsafe {
        (*owner)
            .class_descriptor
            .fields
            .iter()
            .find(|field| field.token == token && field.value_type == RuntimeValueType::Object)
            .map(|_| Variant::nothing())
    }
}

pub(crate) fn compat_owner_from_unknown(unknown: *mut RawRuntimeIUnknown) -> *mut CompatObjectBase {
    unknown.cast::<CompatObjectBase>()
}

pub(crate) fn compat_owner_from_this(this: *mut c_void) -> *mut CompatObjectBase {
    compat_owner_from_unknown(this.cast::<RawRuntimeIUnknown>())
}

/// The runtime's operations on an object reference: project instances (identity, class,
/// fields, termination), built-in native state, and interface projections. A foreign COM
/// object answers the neutral value (`0`, `false`, `None`).
pub trait ObjectRefExt: Sized {
    /// The identity of a runtime object, or `<foreign>` for a foreign COM object (the text the
    /// runtime prints for an object reference).
    fn display_text(&self) -> String;

    fn from_compat_identity(compat_identity: i32) -> Self;

    fn from_compat_identity_with_descriptor(
        compat_identity: i32,
        class_descriptor: &'static RuntimeClassDescriptor,
    ) -> Self;

    /// Allocates a fresh project-class instance: a distinct `CompatObjectBase` (a distinct
    /// IUnknown, hence a distinct identity for `Is`) with its own per-instance `instance_id`
    /// and the class's `route_key` for member dispatch. Refcount starts at 1; the instance is
    /// not pinned by any route map, so its lifetime tracks real references. `has_terminate`
    /// marks classes with a `Class_Terminate`, so the last release enqueues it for the VM.
    fn from_project_instance(
        instance_id: i32,
        route_key: i32,
        bundle_id: i32,
        has_terminate: bool,
        class_descriptor: &'static RuntimeClassDescriptor,
    ) -> Self;

    fn query_iunknown(&self) -> Self;

    /// The identity of a runtime compat object; 0 for a foreign COM object (which is its own
    /// interface pointer and carries none).
    fn compat_identity(&self) -> i32;

    fn raw(&self) -> i32;

    /// Whether this is a runtime placeholder for `Nothing` (a compat object with identity 0). A
    /// foreign object is its interface pointer and never is (step 8a), whatever `raw()` says.
    fn is_nothing_placeholder(&self) -> bool;

    /// The identity VBA's `Is` compares: a runtime object's identity number, and for a foreign
    /// object its `IUnknown` address (COM identity: `QueryInterface` for `IUnknown` returns the
    /// same pointer from every interface of one object).
    fn is_identity(&self) -> usize;

    /// Class/route key for project-dynamic dispatch (which class's route this instance uses).
    /// Only valid for compat (project-class) objects, like `compat_identity`/`class_descriptor`.
    fn route_key(&self) -> i32;

    /// The bundle (".NET assembly") that owns this instance's class â€” `(bundle_id,
    /// route_key)` selects the class dispatch table in a multi-bundle VM. `0` for
    /// single-bundle runs.
    fn bundle_id(&self) -> i32;

    fn is_project_instance(&self) -> bool;

    fn object_owner_id(&self) -> Option<u64>;

    /// True when this `ObjectRef` wraps one of our `CompatObjectBase` boxes (a project-class or
    /// internal object), as opposed to a foreign COM interface obtained via
    /// `from_raw_iunknown_addref`. Reading the vtbl pointer is safe for any IUnknown (every COM
    /// object's first field is its vtbl); comparing it to our static vtbl tells the two apart.
    /// `compat_identity`/`route_key`/`class_descriptor`/`terminates_on_release` are only valid
    /// when this returns true.
    fn is_compat_object(&self) -> bool;

    /// True when this `ObjectRef` carries a dual `ClassDualVtbl` answering `IUnknown`,
    /// `IDispatch`, and its class default interface IID directly on `&primary`.
    fn has_dual_dispatch_vtable(&self) -> bool;

    /// The owning thread ID recorded when this compat object was allocated, or `None` for foreign
    /// COM objects.
    fn owner_thread_id(&self) -> Option<u64>;

    /// Attach a between-runs weak [`crate::events::MemberInvoker`] to a compat object so
    /// `IDispatch::Invoke` can reach its owning session holder when no active run is on the stack.
    fn set_member_invoker(
        &self,
        invoker: std::rc::Weak<RefCell<dyn crate::events::MemberInvoker>>,
    ) -> bool;

    /// True when releasing this instance's last reference should run a `Class_Terminate`. Safe for
    /// any `ObjectRef`: foreign COM objects (which carry no such flag) report `false`.
    fn terminates_on_release(&self) -> bool;

    /// Class_Terminate is not scheduled when Class_Initialize failed.
    fn mark_initialization_failed(&self);

    fn class_descriptor(&self) -> &'static RuntimeClassDescriptor;

    fn try_class_descriptor(&self) -> Option<&'static RuntimeClassDescriptor>;

    fn object_identity(&self) -> RuntimeObjectIdentity;

    fn try_object_identity(&self) -> Option<RuntimeObjectIdentity>;

    fn project_field_get(&self, token: i32) -> Option<Variant>;

    fn project_field_set(&self, token: i32, value: Variant) -> bool;

    /// Borrow instance field `token` IN PLACE (no clone) and run `f` against it
    /// (`None` if the field is absent). This is the O(1) field-element-access
    /// counterpart to [`Self::project_field_get`]: a field-held array can be indexed
    /// without cloning the whole array. Returns `None` if this is not a project
    /// instance. The closure must not call back into this object's fields (it holds a
    /// shared `RefCell` borrow for its duration).
    fn with_project_field<R>(&self, token: i32, f: impl FnOnce(Option<&Variant>) -> R)
    -> Option<R>;

    /// Mutably borrow instance field `token` IN PLACE (no clone / store-back) and run
    /// `f` against it â€” the write counterpart to [`Self::with_project_field`]. `None`
    /// if this is not a project instance, or if the field is absent (it must already
    /// hold the array â€” class init dimensions array fields before they are indexed).
    fn with_project_field_mut<R>(&self, token: i32, f: impl FnOnce(&mut Variant) -> R)
    -> Option<R>;

    /// The cell of instance field `token`, created (Empty, or Nothing for an object field)
    /// if absent: its address is stable until the instance releases its fields. `None` if
    /// this is not a project instance.
    fn project_field_cell(&self, token: i32) -> Option<*mut Variant>;

    /// Run `f` against this instance's built-in native-`Collection` state, lazily creating an
    /// empty [`crate::collection::CollectionData`] on first use. Returns `None` if this
    /// `ObjectRef` is not one of our compat boxes (a foreign COM object has no native state).
    ///
    /// The state rides the object box's `native_state` slot, so `Set c2 = c1` (which clones the
    /// `ObjectRef` to the same box) shares one collection â€” VBA reference semantics â€” and box
    /// teardown reclaims it, so there is no VM-side store and no leak. The shared keyed-method
    /// logic (`calcvba_eval::collection::dispatch_collection`) runs inside `f`.
    fn with_native_collection<R>(
        &self,
        f: impl FnOnce(&mut crate::collection::CollectionData) -> R,
    ) -> Option<R>;

    /// A snapshot of this instance's native `Collection` values in insertion order, or `None`
    /// if it has no native collection â€” a foreign COM object, or a compat object that never had
    /// one (a project instance, or a `Collection` with no items added yet). Unlike
    /// [`Self::with_native_collection`] this does NOT lazily create a collection, so it is safe
    /// to probe any object (e.g. `For Each`) without planting empty state on project instances.
    fn native_collection_snapshot(&self) -> Option<Vec<Variant>>;

    fn query_interface_descriptor(
        &self,
        iid: RuntimeInterfaceId,
    ) -> Option<&'static RuntimeInterfaceDescriptor>;

    fn query_interface_descriptor_by_guid(
        &self,
        guid: Guid,
    ) -> Option<&'static RuntimeInterfaceDescriptor>;

    fn query_interface_projection(
        &self,
        iid: RuntimeInterfaceId,
    ) -> Option<RuntimeInterfaceProjection>;

    fn query_interface_projection_by_guid(&self, guid: Guid) -> Option<RuntimeInterfaceProjection>;

    fn strong_count(&self) -> u32;

    fn strong_count_for_test(&self) -> u32;
}

impl ObjectRefExt for ObjectRef {
    fn display_text(&self) -> String {
        if self.is_compat_object() {
            self.compat_identity().to_string()
        } else {
            "<foreign>".to_string()
        }
    }

    fn from_compat_identity(compat_identity: i32) -> Self {
        Self::from_compat_identity_with_descriptor(compat_identity, &COMPAT_OBJECT_CLASS_DESCRIPTOR)
    }

    fn from_compat_identity_with_descriptor(
        compat_identity: i32,
        class_descriptor: &'static RuntimeClassDescriptor,
    ) -> Self {
        // Legacy/template path: identity key and route key coincide; no terminate hook.
        from_compat_object(
            compat_identity,
            compat_identity,
            0,
            false,
            class_descriptor,
            false,
        )
    }

    fn from_project_instance(
        instance_id: i32,
        route_key: i32,
        bundle_id: i32,
        has_terminate: bool,
        class_descriptor: &'static RuntimeClassDescriptor,
    ) -> Self {
        from_compat_object(
            instance_id,
            route_key,
            bundle_id,
            has_terminate,
            class_descriptor,
            true,
        )
    }

    fn query_iunknown(&self) -> Self {
        self.clone()
    }

    fn compat_identity(&self) -> i32 {
        if !self.is_compat_object() {
            return 0;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: `self.0` points at a live `CompatObjectBase`. Every `ObjectRef` in the
        // workspace is such a box: the safe constructors mint them, the COM bridge wraps even
        // foreign IDispatch/IUnknown objects in compat boxes keyed by a binding handle
        // (`retained_runtime_object` in calcvba-com), and the unsafe raw constructors only
        // round-trip pointers `raw_iunknown()` produced from those boxes (Variant payloads,
        // SAFEARRAY object slots). The two entry points that previously reached here without an
        // `is_compat_object` check â€” the `Debug`/`Display` impls below â€” now guard on it, so the
        // remaining callers (`raw`/`route_key`/`is_project_instance`/HAL/vm2 dispatch) all sit on
        // compat objects. This `ObjectRef`'s retained reference keeps the box alive for the read.
        unsafe { (*owner).identity.compat_identity }
    }

    fn raw(&self) -> i32 {
        self.compat_identity()
    }

    fn is_nothing_placeholder(&self) -> bool {
        self.is_compat_object() && self.compat_identity() == 0
    }

    fn is_identity(&self) -> usize {
        if self.is_compat_object() {
            if self.is_project_instance() {
                return self.raw_iunknown() as usize;
            }
            return self.compat_identity() as u32 as usize;
        }
        let this = self.raw_iunknown();
        let mut unknown: *mut c_void = core::ptr::null_mut();
        // SAFETY: this ObjectRef retains a live COM object whose first field is its IUnknown
        // vtable; the reference QueryInterface returns is released at once (only its address is
        // kept).
        unsafe {
            let vtbl = (*this).vtbl;
            let hr = ((*vtbl).query_interface)(this.cast(), &RUNTIME_GUID_IUNKNOWN, &mut unknown);
            if hr < 0 || unknown.is_null() {
                return this as usize;
            }
            // Released through its own vtable: an out-of-process proxy's IUnknown is a
            // different object from the interface proxy it came from.
            let identity = unknown.cast::<RawRuntimeIUnknown>();
            ((*(*identity).vtbl).release)(unknown);
        }
        unknown as usize
    }

    fn route_key(&self) -> i32 {
        if !self.is_compat_object() {
            return 0;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: documented contract (doc above and `is_compat_object`): callers only invoke
        // this on compat objects, so `owner` is the live `CompatObjectBase` (`#[repr(C)]`,
        // `unknown` first) kept alive by this `ObjectRef`'s retained reference.
        unsafe { (*owner).identity.route_key }
    }

    fn bundle_id(&self) -> i32 {
        if !self.is_compat_object() {
            return 0;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: documented contract as for `route_key`: only called on compat objects, so
        // `owner` is the live `CompatObjectBase` kept alive by this `ObjectRef`'s reference.
        unsafe { (*owner).identity.bundle_id }
    }

    fn is_project_instance(&self) -> bool {
        if !self.is_compat_object() {
            return false;
        }
        // SAFETY: private vtable proof establishes the complete retained allocation.
        unsafe { (*compat_owner_from_unknown(self.raw_iunknown())).project_instance }
    }

    fn object_owner_id(&self) -> Option<u64> {
        if !self.is_compat_object() {
            return None;
        }
        // SAFETY: private-vtable proof retains this complete compat allocation.
        Some(unsafe {
            (*compat_owner_from_unknown(self.raw_iunknown()))
                .object_owner
                .id
        })
    }

    fn is_compat_object(&self) -> bool {
        // SAFETY: every COM-shaped object's first pointer-sized field is its vtbl (see doc
        // above), so this read is valid for compat boxes and foreign IUnknowns alike; `self.0`
        // is non-null and the object is alive because this `ObjectRef` holds a retained
        // reference to it.
        let vtbl = unsafe { (*self.raw_iunknown()).vtbl };
        if std::ptr::eq(vtbl, &COMPAT_OBJECT_VTBL) {
            return true;
        }
        // SAFETY: any live COM interface's vtbl has `query_interface` at slot 0. Every
        // `ClassDualVtbl` copies `COMPAT_OBJECT_VTBL.query_interface` verbatim into slot 0, so
        // comparing against that exact function pointer recognizes every dual vtable in O(1)
        // cleanly under both native codegen and Miri.
        let qi = unsafe { (*vtbl).query_interface };
        core::ptr::fn_addr_eq(qi, COMPAT_OBJECT_VTBL.query_interface)
    }

    fn has_dual_dispatch_vtable(&self) -> bool {
        if !self.is_compat_object() {
            return false;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: guarded by `is_compat_object()` above; `owner` is a live `CompatObjectBase`.
        unsafe { (*owner).has_dual_vtable }
    }

    fn owner_thread_id(&self) -> Option<u64> {
        if !self.is_compat_object() {
            return None;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: guarded by `is_compat_object()` above; `owner` is a live `CompatObjectBase`.
        Some(unsafe { (*owner).owner_thread })
    }

    fn set_member_invoker(
        &self,
        invoker: std::rc::Weak<RefCell<dyn crate::events::MemberInvoker>>,
    ) -> bool {
        if !self.is_compat_object() {
            return false;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: guarded by `is_compat_object()` above; `owner` is a live `CompatObjectBase`.
        unsafe {
            *(*owner).invoker.borrow_mut() = Some(invoker);
        }
        true
    }

    fn terminates_on_release(&self) -> bool {
        if !self.is_compat_object() {
            return false;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: guarded by the `is_compat_object()` check above (vtbl identity), so `owner` is
        // the live `CompatObjectBase` kept alive by this `ObjectRef`'s retained reference.
        unsafe { (*owner).identity.terminates_on_release }
    }

    fn mark_initialization_failed(&self) {
        if self.is_project_instance() {
            let owner = compat_owner_from_unknown(self.raw_iunknown());
            // SAFETY: the private vtable proves this retained project box.
            unsafe {
                (*owner).terminated.store(true, Ordering::Release);
            }
        }
    }

    fn class_descriptor(&self) -> &'static RuntimeClassDescriptor {
        debug_assert!(
            self.is_compat_object(),
            "class_descriptor is only valid for runtime compat objects"
        );
        self.try_class_descriptor()
            .expect("class_descriptor is only valid for runtime compat objects")
    }

    fn try_class_descriptor(&self) -> Option<&'static RuntimeClassDescriptor> {
        if !self.is_compat_object() {
            return None;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: documented contract (`is_compat_object`): only valid on compat objects, so
        // `owner` is the live `CompatObjectBase` kept alive by this `ObjectRef`'s retained
        // reference; the descriptor read out of it is `&'static`.
        Some(unsafe { (*owner).class_descriptor })
    }

    fn object_identity(&self) -> RuntimeObjectIdentity {
        debug_assert!(
            self.is_compat_object(),
            "object_identity is only valid for runtime compat objects"
        );
        self.try_object_identity()
            .expect("object_identity is only valid for runtime compat objects")
    }

    fn try_object_identity(&self) -> Option<RuntimeObjectIdentity> {
        if !self.is_compat_object() {
            return None;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: documented contract as for `class_descriptor`: only called on compat objects,
        // so `owner` is the live `CompatObjectBase` kept alive by this `ObjectRef`'s reference.
        Some(unsafe { (*owner).identity })
    }

    fn project_field_get(&self, token: i32) -> Option<Variant> {
        if !self.is_project_instance() {
            return None;
        }
        compat_field_get(compat_owner_from_unknown(self.raw_iunknown()), token)
    }

    fn project_field_set(&self, token: i32, value: Variant) -> bool {
        if !self.is_project_instance() {
            return false;
        }
        compat_field_set(compat_owner_from_unknown(self.raw_iunknown()), token, value);
        true
    }

    fn with_project_field<R>(
        &self,
        token: i32,
        f: impl FnOnce(Option<&Variant>) -> R,
    ) -> Option<R> {
        if !self.is_project_instance() {
            return None;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: `is_project_instance()` (a vtbl-identity check) guards that `owner` is a
        // live `CompatObjectBase` kept alive by this `ObjectRef`'s retained reference; the
        // `RefCell` borrow is single-threaded (`ObjectRef` is neither `Send` nor `Sync`).
        unsafe {
            let guard = (*owner).fields.borrow();
            Some(f(guard.get(&token).map(|cell| &**cell)))
        }
    }

    fn with_project_field_mut<R>(
        &self,
        token: i32,
        f: impl FnOnce(&mut Variant) -> R,
    ) -> Option<R> {
        if !self.is_project_instance() {
            return None;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: as in `with_project_field` â€” `owner` is a live `CompatObjectBase` held
        // alive by this `ObjectRef`; the exclusive `RefCell` borrow is single-threaded.
        unsafe {
            (*owner)
                .fields
                .borrow_mut()
                .get_mut(&token)
                .map(|cell| f(cell))
        }
    }

    fn project_field_cell(&self, token: i32) -> Option<*mut Variant> {
        if !self.is_project_instance() {
            return None;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: as in `with_project_field_mut`; the returned pointer addresses the boxed
        // cell, not the map, so later inserts do not move it.
        unsafe {
            let default = compat_field_default(owner, token).unwrap_or_else(Variant::empty);
            let mut fields = (*owner).fields.borrow_mut();
            let cell = fields.entry(token).or_insert_with(|| Box::new(default));
            Some(&mut **cell as *mut Variant)
        }
    }

    fn with_native_collection<R>(
        &self,
        f: impl FnOnce(&mut crate::collection::CollectionData) -> R,
    ) -> Option<R> {
        if !self.is_compat_object() {
            return None;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: guarded by `is_compat_object()` (vtbl identity) above, so `owner` is the live
        // `CompatObjectBase` kept alive by this `ObjectRef`'s retained reference; the `RefCell`
        // borrow is single-threaded (`ObjectRef` is neither `Send` nor `Sync`).
        unsafe {
            let mut guard = (*owner).native_state.borrow_mut();
            let data = guard.get_or_insert_with(crate::collection::CollectionData::default);
            Some(f(data))
        }
    }

    fn native_collection_snapshot(&self) -> Option<Vec<Variant>> {
        if !self.is_compat_object() {
            return None;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: guarded by `is_compat_object()` (vtbl identity), so `owner` is the live
        // `CompatObjectBase` kept alive by this `ObjectRef`'s retained reference; the borrow is
        // single-threaded (`ObjectRef` is neither `Send` nor `Sync`).
        unsafe { (*owner).native_state.borrow().as_ref().map(|c| c.values()) }
    }

    fn query_interface_descriptor(
        &self,
        iid: RuntimeInterfaceId,
    ) -> Option<&'static RuntimeInterfaceDescriptor> {
        let class_descriptor = self.try_class_descriptor()?;
        if iid == RuntimeInterfaceId::IUnknown {
            return Some(runtime_iunknown_interface_descriptor());
        }
        class_descriptor
            .interfaces
            .iter()
            .find(|descriptor| descriptor.id == iid)
    }

    fn query_interface_descriptor_by_guid(
        &self,
        guid: Guid,
    ) -> Option<&'static RuntimeInterfaceDescriptor> {
        let class_descriptor = self.try_class_descriptor()?;
        if guid == RUNTIME_GUID_IUNKNOWN {
            return Some(runtime_iunknown_interface_descriptor());
        }
        if class_descriptor.default_interface_iid == Some(guid)
            && let Some(dispatch) = class_descriptor.dispatch_interface()
        {
            return Some(dispatch);
        }
        class_descriptor
            .interfaces
            .iter()
            .find(|descriptor| descriptor.identity.guid == guid)
    }

    fn query_interface_projection(
        &self,
        iid: RuntimeInterfaceId,
    ) -> Option<RuntimeInterfaceProjection> {
        let descriptor = self.query_interface_descriptor(iid)?;
        Some(RuntimeInterfaceProjection {
            object_identity: self.try_object_identity()?,
            interface_identity: descriptor.identity,
            interface_descriptor: descriptor,
        })
    }

    fn query_interface_projection_by_guid(&self, guid: Guid) -> Option<RuntimeInterfaceProjection> {
        let descriptor = self.query_interface_descriptor_by_guid(guid)?;
        Some(RuntimeInterfaceProjection {
            object_identity: self.try_object_identity()?,
            interface_identity: descriptor.identity,
            interface_descriptor: descriptor,
        })
    }

    fn strong_count(&self) -> u32 {
        if !self.is_compat_object() {
            return 0;
        }
        let owner = compat_owner_from_unknown(self.raw_iunknown());
        // SAFETY: documented compat-object contract as for `class_descriptor`; the box is alive
        // because this `ObjectRef` holds a retained reference, and `ref_count` is atomic.
        unsafe { (*owner).ref_count.load(Ordering::Acquire) }
    }

    fn strong_count_for_test(&self) -> u32 {
        self.strong_count()
    }
}

fn from_compat_object(
    compat_identity: i32,
    route_key: i32,
    bundle_id: i32,
    terminates_on_release: bool,
    class_descriptor: &'static RuntimeClassDescriptor,
    project_instance: bool,
) -> ObjectRef {
    let stable_object_id = compat_identity as u32 as u64;
    let has_dual_vtable = project_instance && class_descriptor.dispatch_interface().is_some();
    let vtbl = if has_dual_vtable {
        crate::class_dispatch::class_dual_vtbl_for_descriptor(class_descriptor)
    } else {
        &COMPAT_OBJECT_VTBL
    };
    let boxed = Box::new(CompatObjectBase {
        unknown: RawRuntimeIUnknown { vtbl },
        ref_count: AtomicU32::new(1),
        owner_thread: crate::class_dispatch::current_thread_id(),
        has_dual_vtable,
        object_owner: crate::object_owner::RuntimeObjectOwner::current().downgrade(),
        project_instance,
        identity: RuntimeObjectIdentity {
            stable_object_id,
            compat_identity,
            route_key,
            bundle_id,
            terminates_on_release,
            class_descriptor,
            lifetime_policy: RuntimeLifetimePolicy::RefCounted,
            apartment_model: if project_instance {
                RuntimeApartmentModel::SingleThreaded
            } else {
                RuntimeApartmentModel::Unknown
            },
        },
        class_descriptor,
        invoker: RefCell::new(None),
        fields: RefCell::new(BTreeMap::new()),
        native_state: RefCell::new(None),
        parked_for_termination: AtomicBool::new(false),
        running_termination: AtomicBool::new(false),
        terminated: AtomicBool::new(false),
    });
    let raw = Box::into_raw(boxed);
    crate::live_counters::object_box_allocated();
    // `raw` came from `Box::into_raw` of the freshly allocated `CompatObjectBase` above.
    // `repr(C)` puts `unknown` first, so this cast has the field's exact address while
    // retaining provenance for the complete allocation; AddRef/Release need that full-box
    // provenance when they recover `CompatObjectBase` and touch the adjacent refcount.
    let unknown = raw.cast::<RawRuntimeIUnknown>();
    // SAFETY: a fresh object whose one reference is handed over.
    unsafe { ObjectRef::from_raw_iunknown_owned(unknown) }
        .expect("compat object unknown pointer must be non-null")
}

#[cfg(test)]
// Test-support code exercising the documented production vtable paths above.
#[allow(clippy::undocumented_unsafe_blocks)]
mod tests {
    use core::ffi::c_void;
    use core::sync::atomic::{AtomicU32, Ordering};

    use super::{
        Guid, IUnknownVtbl, ObjectRef, ObjectRefExt, RUNTIME_CLASS_LIFECYCLE_NONE,
        RUNTIME_E_NOINTERFACE, RUNTIME_GUID_IDISPATCH, RUNTIME_GUID_IUNKNOWN,
        RUNTIME_IDISPATCH_INTERFACE_IDENTITY, RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR, RUNTIME_S_OK,
        RawRuntimeIUnknown, RuntimeApartmentModel, RuntimeClassDescriptor, RuntimeDescriptorIssue,
        RuntimeDispatchPlanCache, RuntimeInterfaceDescriptor, RuntimeInterfaceId,
        RuntimeInterfaceIdentity, RuntimeInterfaceKind, RuntimeLifetimePolicy,
        RuntimeMemberBacking, RuntimeMemberDescriptor, RuntimeMemberInvokeKind,
        RuntimeParamDescriptor, RuntimeValueType,
    };

    #[test]
    fn object_ref_clone_tracks_refcount_and_identity() {
        let object = ObjectRef::from_compat_identity(42);
        assert_eq!(object.compat_identity(), 42);
        assert_eq!(object.strong_count_for_test(), 1);

        let clone = object.clone();
        assert_eq!(object, clone);
        assert_eq!(object.compat_identity(), clone.compat_identity());
        assert_eq!(object.strong_count_for_test(), 2);

        drop(clone);
        assert_eq!(object.strong_count_for_test(), 1);
    }

    #[test]
    fn object_ref_query_iunknown_returns_same_identity() {
        let object = ObjectRef::from_compat_identity(77);
        let unknown = object.query_iunknown();
        assert_eq!(object, unknown);
        assert_eq!(object.compat_identity(), unknown.compat_identity());
        assert_eq!(object.strong_count_for_test(), 2);
    }

    #[test]
    fn object_ref_raw_iunknown_is_first_field_vtable_shape() {
        let object = ObjectRef::from_project_instance(
            501,
            17,
            3,
            false,
            &super::COMPAT_OBJECT_CLASS_DESCRIPTOR,
        );
        let raw_unknown = object.raw_iunknown();
        assert_eq!(
            core::mem::size_of::<RawRuntimeIUnknown>(),
            core::mem::size_of::<*const IUnknownVtbl>()
        );
        assert_eq!(
            core::mem::align_of::<RawRuntimeIUnknown>(),
            core::mem::align_of::<*const IUnknownVtbl>()
        );

        unsafe {
            let owner = raw_unknown.cast::<super::CompatObjectBase>();
            assert_eq!(
                core::ptr::addr_of_mut!((*owner).unknown),
                raw_unknown,
                "ObjectRef must expose a pointer to the first field of the compat object box"
            );
            let first_word = raw_unknown.cast::<*const IUnknownVtbl>().read();
            assert!(
                core::ptr::eq(first_word, &super::COMPAT_OBJECT_VTBL),
                "the first machine word of the exposed IUnknown must be the runtime vtable"
            );
            assert!(
                core::ptr::eq((*raw_unknown).vtbl, &super::COMPAT_OBJECT_VTBL),
                "the typed RawRuntimeIUnknown view must agree with the raw first-word proof"
            );
        }
    }

    #[test]
    fn compat_object_exposes_descriptor_backed_iunknown_interface() {
        let object = ObjectRef::from_compat_identity(9);
        let class_descriptor = object.class_descriptor();
        assert_eq!(class_descriptor.name, "OxVba.CompatObject");
        assert_eq!(class_descriptor.project_identity, None);
        assert!(!class_descriptor.predeclared);
        assert_eq!(class_descriptor.lifecycle, RUNTIME_CLASS_LIFECYCLE_NONE);
        assert!(class_descriptor.fields.is_empty());
        assert!(class_descriptor.as_new_fields.is_empty());
        assert!(class_descriptor.implements.is_empty());
        assert_eq!(class_descriptor.interfaces.len(), 1);
        let iunknown = object
            .query_interface_descriptor(RuntimeInterfaceId::IUnknown)
            .expect("compat object should expose internal IUnknown descriptor");
        assert_eq!(iunknown.name, "IUnknown");
        assert!(!iunknown.dual_dispatch);
        assert!(iunknown.members.is_empty());
        assert!(
            object
                .query_interface_descriptor(RuntimeInterfaceId::IDispatch)
                .is_none(),
            "compat object floor must not falsely claim dual dispatch support"
        );
    }

    #[test]
    fn runtime_guid_parses_and_formats_canonical_identity() {
        let parsed = Guid::parse("{00020400-0000-0000-c000-000000000046}")
            .expect("IDispatch GUID should parse with braces and mixed case");
        assert_eq!(parsed, RUNTIME_GUID_IDISPATCH);
        assert_eq!(parsed.to_string(), "00020400-0000-0000-C000-000000000046");
        assert_eq!(
            Guid::parse("00000000-0000-0000-C000-000000000046"),
            Some(RUNTIME_GUID_IUNKNOWN)
        );
        assert!(Guid::parse("00020400-0000-0000-C000").is_none());
        assert!(Guid::parse("not-a-guid").is_none());
    }

    #[test]
    fn runtime_interface_identity_carries_custom_iid_metadata() {
        const CUSTOM_GUID: Guid = Guid::new(
            0x1111_1111,
            0x2222,
            0x3333,
            [0x44, 0x44, 0x55, 0x55, 0x55, 0x55, 0x55, 0x55],
        );
        const CUSTOM_IDENTITY: RuntimeInterfaceIdentity = RuntimeInterfaceIdentity::custom(
            CUSTOM_GUID,
            "Project._Widget",
            RuntimeInterfaceKind::Dual,
            Some(1),
            Some(0),
            Some(1033),
        );
        static CUSTOM_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::Unsupported,
            identity: CUSTOM_IDENTITY,
            name: "Project._Widget",
            members: &[],
            dual_dispatch: true,
        };
        static TEST_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
            name: "Project.Widget",
            runtime_type_name: None,
            project_identity: None,
            default_interface_iid: None,
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
            fields: &[],
            as_new_fields: &[],
            implements: &[],
            interfaces: &[RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR, CUSTOM_INTERFACE],
        };

        let object = ObjectRef::from_compat_identity_with_descriptor(202, &TEST_CLASS);
        let descriptor = object
            .query_interface_descriptor_by_guid(CUSTOM_GUID)
            .expect("custom interface should be discoverable by GUID");
        assert_eq!(descriptor.identity.guid, CUSTOM_GUID);
        assert_eq!(descriptor.identity.name, "Project._Widget");
        assert_eq!(descriptor.identity.kind, RuntimeInterfaceKind::Dual);
        assert_eq!(descriptor.identity.major_version, Some(1));
        assert_eq!(descriptor.identity.minor_version, Some(0));
        assert_eq!(descriptor.identity.lcid, Some(1033));
    }

    #[test]
    fn native_collection_state_is_shared_across_clones() {
        use crate::Variant;
        // W1-a: a built-in Collection's CollectionData rides the object box's `native_state`
        // slot, reached via `with_native_collection`. Cloning the ObjectRef (`Set c2 = c1`)
        // points at the SAME box, so both names share one CollectionData â€” VBA reference
        // semantics â€” and box teardown reclaims it (no VM-side store, no leak).
        static TEST_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
            name: "VBA.Collection",
            runtime_type_name: None,
            project_identity: None,
            default_interface_iid: None,
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
            fields: &[],
            as_new_fields: &[],
            implements: &[],
            interfaces: &[RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR],
        };
        let c1 = ObjectRef::from_compat_identity_with_descriptor(404, &TEST_CLASS);
        c1.with_native_collection(|d| {
            d.add(Variant::from_i32(10), None, None, None).unwrap();
            d.add(Variant::from_i32(20), None, None, None).unwrap();
        })
        .expect("a compat object has native-collection state");
        let c2 = c1.clone();
        let (count, first) = c2
            .with_native_collection(|d| {
                (
                    d.count(),
                    d.item(&crate::collection::Selector::Index(1))
                        .ok()
                        .and_then(|v| v.as_i32()),
                )
            })
            .expect("the clone sees the same box's native state");
        assert_eq!(count, 2, "both ObjectRef clones share one CollectionData");
        assert_eq!(first, Some(10));
        drop(c2);
        drop(c1); // refcount-0 frees the box and, with it, native_state â€” no leak
    }

    #[test]
    fn interface_projections_share_runtime_object_identity() {
        const CUSTOM_GUID: Guid = Guid::new(
            0x2222_2222,
            0x3333,
            0x4444,
            [0x55, 0x55, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66],
        );
        const CUSTOM_IDENTITY: RuntimeInterfaceIdentity = RuntimeInterfaceIdentity::custom(
            CUSTOM_GUID,
            "Project._WidgetEvents",
            RuntimeInterfaceKind::Source,
            Some(1),
            Some(0),
            Some(1033),
        );
        static CUSTOM_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::Unsupported,
            identity: CUSTOM_IDENTITY,
            name: "Project._WidgetEvents",
            members: &[],
            dual_dispatch: false,
        };
        static DISPATCH_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "Project._Widget",
            members: &[],
            dual_dispatch: true,
        };
        static TEST_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
            name: "Project.Widget",
            runtime_type_name: None,
            project_identity: None,
            default_interface_iid: None,
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
            fields: &[],
            as_new_fields: &[],
            implements: &[],
            interfaces: &[
                RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR,
                DISPATCH_INTERFACE,
                CUSTOM_INTERFACE,
            ],
        };

        let object = ObjectRef::from_compat_identity_with_descriptor(303, &TEST_CLASS);
        let unknown = object
            .query_interface_projection(RuntimeInterfaceId::IUnknown)
            .expect("IUnknown projection");
        let dispatch = object
            .query_interface_projection(RuntimeInterfaceId::IDispatch)
            .expect("IDispatch projection");
        let source = object
            .query_interface_projection_by_guid(CUSTOM_GUID)
            .expect("custom source projection");

        assert_eq!(unknown.object_identity, dispatch.object_identity);
        assert_eq!(dispatch.object_identity, source.object_identity);
        assert_eq!(unknown.object_identity.compat_identity, 303);
        assert_eq!(unknown.object_identity.stable_object_id, 303);
        assert_eq!(
            unknown.object_identity.lifetime_policy,
            RuntimeLifetimePolicy::RefCounted
        );
        assert_eq!(
            unknown.object_identity.apartment_model,
            RuntimeApartmentModel::Unknown
        );
        assert_eq!(unknown.interface_identity.guid, RUNTIME_GUID_IUNKNOWN);
        assert_eq!(dispatch.interface_identity.guid, RUNTIME_GUID_IDISPATCH);
        assert_eq!(source.interface_identity.guid, CUSTOM_GUID);
        assert_eq!(object.strong_count_for_test(), 1);
    }

    #[test]
    fn descriptor_backed_object_can_advertise_dual_dispatch_shape() {
        static VALUE_MEMBER: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 0,
            vtable_slot: Some(7),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Variant),
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static DISPATCH_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "ITestDual",
            members: &[VALUE_MEMBER],
            dual_dispatch: true,
        };
        static TEST_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
            name: "Project.Widget",
            runtime_type_name: None,
            project_identity: None,
            default_interface_iid: None,
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
            fields: &[],
            as_new_fields: &[],
            implements: &[],
            interfaces: &[RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR, DISPATCH_INTERFACE],
        };

        let object = ObjectRef::from_compat_identity_with_descriptor(101, &TEST_CLASS);
        assert_eq!(object.class_descriptor().name, "Project.Widget");
        let dispatch = object
            .query_interface_descriptor(RuntimeInterfaceId::IDispatch)
            .expect("test class should advertise a dual dispatch descriptor");
        assert!(dispatch.dual_dispatch);
        assert_eq!(dispatch.members.len(), 1);
        assert_eq!(dispatch.members[0].name, "Value");
        assert_eq!(dispatch.members[0].vtable_slot, Some(7));
        assert!(dispatch.members[0].is_default_member);
    }

    #[test]
    fn runtime_class_descriptor_validates_com_export_member_shape() {
        static GET_VALUE: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 0,
            vtable_slot: Some(7),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Long),
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static LET_VALUE: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 0,
            vtable_slot: Some(8),
            invoke_kind: RuntimeMemberInvokeKind::PropertyLet,
            arity: 1,
            params: &[RuntimeParamDescriptor {
                name: "value",
                value_type: RuntimeValueType::Long,
                by_ref: false,
                optional: false,
                param_array: false,
            }],
            return_type: None,
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static NEW_ENUM: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "_NewEnum",
            dispatch_id: -4,
            vtable_slot: Some(9),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Object),
            is_default_member: false,
            is_enumerator_member: true,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static DISPATCH_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "IDispatch",
            members: &[GET_VALUE, LET_VALUE, NEW_ENUM],
            dual_dispatch: true,
        };
        static TEST_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
            name: "Project.Widget",
            runtime_type_name: None,
            project_identity: None,
            default_interface_iid: None,
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
            fields: &[],
            as_new_fields: &[],
            implements: &[],
            interfaces: &[RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR, DISPATCH_INTERFACE],
        };

        assert_eq!(TEST_CLASS.validate_com_export_shape(), Ok(()));
    }

    #[test]
    fn runtime_class_descriptor_rejects_bad_com_export_member_shape() {
        static BAD_DEFAULT: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 12,
            vtable_slot: Some(7),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Long),
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static BAD_SLOT: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "FastValue",
            dispatch_id: 2,
            vtable_slot: Some(6),
            invoke_kind: RuntimeMemberInvokeKind::Method,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Long),
            is_default_member: false,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static BAD_SETTER: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Child",
            dispatch_id: 3,
            vtable_slot: Some(9),
            invoke_kind: RuntimeMemberInvokeKind::PropertySet,
            arity: 0,
            params: &[],
            return_type: None,
            is_default_member: false,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static BAD_ENUM: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "_NewEnum",
            dispatch_id: 99,
            vtable_slot: Some(10),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Object),
            is_default_member: false,
            is_enumerator_member: true,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static BAD_ARITY: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Echo",
            dispatch_id: 4,
            vtable_slot: Some(11),
            invoke_kind: RuntimeMemberInvokeKind::Method,
            arity: 2,
            params: &[RuntimeParamDescriptor {
                name: "value",
                value_type: RuntimeValueType::Long,
                by_ref: false,
                optional: false,
                param_array: false,
            }],
            return_type: None,
            is_default_member: false,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static BAD_DEFAULT_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "IDispatch",
            members: &[BAD_DEFAULT],
            dual_dispatch: true,
        };
        static BAD_SLOT_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "IDispatch",
            members: &[BAD_SLOT],
            dual_dispatch: true,
        };
        static BAD_SETTER_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "IDispatch",
            members: &[BAD_SETTER],
            dual_dispatch: true,
        };
        static BAD_ENUM_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "IDispatch",
            members: &[BAD_ENUM],
            dual_dispatch: true,
        };
        static BAD_ARITY_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "IDispatch",
            members: &[BAD_ARITY],
            dual_dispatch: true,
        };
        static BAD_IDENTITY_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: super::RUNTIME_IUNKNOWN_INTERFACE_IDENTITY,
            name: "IDispatch",
            members: &[],
            dual_dispatch: true,
        };

        for (interface, expected) in [
            (
                &BAD_IDENTITY_INTERFACE,
                RuntimeDescriptorIssue::InterfaceIdentityMismatch {
                    interface: "IDispatch",
                    descriptor_id: RuntimeInterfaceId::IDispatch,
                    identity_id: RuntimeInterfaceId::IUnknown,
                },
            ),
            (
                &BAD_DEFAULT_INTERFACE,
                RuntimeDescriptorIssue::BadDefaultMemberDispatchId {
                    interface: "IDispatch",
                    member: "Value",
                    dispatch_id: 12,
                },
            ),
            (
                &BAD_SLOT_INTERFACE,
                RuntimeDescriptorIssue::BadDualDispatchVtableSlot {
                    interface: "IDispatch",
                    member: "FastValue",
                    vtable_slot: 6,
                },
            ),
            (
                &BAD_SETTER_INTERFACE,
                RuntimeDescriptorIssue::BadPropertySetterShape {
                    interface: "IDispatch",
                    member: "Child",
                },
            ),
            (
                &BAD_ENUM_INTERFACE,
                RuntimeDescriptorIssue::BadEnumeratorDispatchId {
                    interface: "IDispatch",
                    member: "_NewEnum",
                    dispatch_id: 99,
                },
            ),
            (
                &BAD_ARITY_INTERFACE,
                RuntimeDescriptorIssue::BadMemberArity {
                    interface: "IDispatch",
                    member: "Echo",
                    arity: 2,
                    params: 1,
                },
            ),
        ] {
            assert_eq!(interface.validate_com_export_shape(), Err(expected));
        }
    }

    #[test]
    fn runtime_class_descriptor_rejects_duplicate_export_interface_guids() {
        static FIRST_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "IDispatch",
            members: &[],
            dual_dispatch: true,
        };
        static SECOND_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "DuplicateDispatch",
            members: &[],
            dual_dispatch: true,
        };
        static TEST_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
            name: "Project.Widget",
            runtime_type_name: None,
            project_identity: None,
            default_interface_iid: None,
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
            fields: &[],
            as_new_fields: &[],
            implements: &[],
            interfaces: &[FIRST_INTERFACE, SECOND_INTERFACE],
        };

        assert_eq!(
            TEST_CLASS.validate_com_export_shape(),
            Err(RuntimeDescriptorIssue::DuplicateInterfaceGuid {
                class: "Project.Widget",
                interface: "DuplicateDispatch",
                guid: RUNTIME_GUID_IDISPATCH,
            })
        );
    }

    #[test]
    fn runtime_dispatch_plan_cache_normalizes_and_reuses_member_lookup() {
        static VALUE_MEMBER: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 0,
            vtable_slot: Some(3),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Variant),
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static SET_VALUE_MEMBER: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 0,
            vtable_slot: Some(4),
            invoke_kind: RuntimeMemberInvokeKind::PropertyLet,
            arity: 1,
            params: &[RuntimeParamDescriptor {
                name: "value",
                value_type: RuntimeValueType::Variant,
                by_ref: false,
                optional: false,
                param_array: false,
            }],
            return_type: None,
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "ITestDual",
            members: &[VALUE_MEMBER, SET_VALUE_MEMBER],
            dual_dispatch: true,
        };

        let mut cache = RuntimeDispatchPlanCache::new();
        let first = cache
            .resolve_member(
                &INTERFACE,
                " VALUE ",
                RuntimeMemberInvokeKind::PropertyGet,
                0,
            )
            .expect("property get should resolve");
        assert_eq!(first.dispatch_id, 0);
        assert_eq!(first.vtable_slot, Some(3));
        assert_eq!(first.member_index, 0);
        assert!(first.is_default_member);
        assert_eq!(cache.len(), 1);

        let second = cache
            .resolve_member(&INTERFACE, "value", RuntimeMemberInvokeKind::PropertyGet, 0)
            .expect("cached property get should resolve case-insensitively");
        assert_eq!(first, second);
        assert_eq!(cache.len(), 1, "second lookup should reuse the cached plan");

        let put = cache
            .resolve_member(&INTERFACE, "value", RuntimeMemberInvokeKind::PropertyLet, 1)
            .expect("property let should use a distinct call-kind/arity plan");
        assert_eq!(put.vtable_slot, Some(4));
        assert_eq!(put.member_index, 1);
        assert_eq!(cache.len(), 2);
        assert!(
            cache
                .resolve_member(&INTERFACE, "value", RuntimeMemberInvokeKind::PropertyGet, 1)
                .is_none(),
            "arity participates in descriptor plan resolution"
        );

        let default_get = cache
            .resolve_default_member(&INTERFACE, RuntimeMemberInvokeKind::PropertyGet, 0)
            .expect("default property get should resolve through the descriptor cache");
        assert_eq!(default_get.member_index, 0);
        assert_eq!(cache.len(), 3);
        let default_get_again = cache
            .resolve_default_member(&INTERFACE, RuntimeMemberInvokeKind::PropertyGet, 0)
            .expect("cached default property get should resolve");
        assert_eq!(default_get, default_get_again);
        assert_eq!(cache.len(), 3);
    }

    #[test]
    fn runtime_dispatch_plan_cache_is_scoped_to_descriptor_identity() {
        static FIRST_VALUE_MEMBER: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 10,
            vtable_slot: Some(3),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Variant),
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static SECOND_VALUE_MEMBER: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 20,
            vtable_slot: Some(9),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Variant),
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static FIRST_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "IFirst",
            members: &[FIRST_VALUE_MEMBER],
            dual_dispatch: true,
        };
        static SECOND_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "ISecond",
            members: &[SECOND_VALUE_MEMBER],
            dual_dispatch: true,
        };

        let mut cache = RuntimeDispatchPlanCache::new();
        let first = cache
            .resolve_member(
                &FIRST_INTERFACE,
                "Value",
                RuntimeMemberInvokeKind::PropertyGet,
                0,
            )
            .expect("first interface member should resolve");
        let second = cache
            .resolve_member(
                &SECOND_INTERFACE,
                "Value",
                RuntimeMemberInvokeKind::PropertyGet,
                0,
            )
            .expect("second interface member should resolve independently");

        assert_eq!(first.dispatch_id, 10);
        assert_eq!(first.vtable_slot, Some(3));
        assert_eq!(second.dispatch_id, 20);
        assert_eq!(second.vtable_slot, Some(9));
        assert_eq!(cache.len(), 2);
    }

    #[test]
    fn runtime_dispatch_plan_cache_caches_unhinted_unique_member_lookup() {
        static VALUE_MEMBER: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 0,
            vtable_slot: Some(3),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Variant),
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static PUT_VALUE_MEMBER: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 0,
            vtable_slot: Some(4),
            invoke_kind: RuntimeMemberInvokeKind::PropertyLet,
            arity: 1,
            params: &[RuntimeParamDescriptor {
                name: "value",
                value_type: RuntimeValueType::Variant,
                by_ref: false,
                optional: false,
                param_array: false,
            }],
            return_type: None,
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "ITestDual",
            members: &[VALUE_MEMBER, PUT_VALUE_MEMBER],
            dual_dispatch: true,
        };

        let mut cache = RuntimeDispatchPlanCache::new();
        let get = cache
            .resolve_member_unhinted(&INTERFACE, " value ", 0)
            .expect("unique unhinted get should resolve and cache");
        assert_eq!(get.member_index, 0);
        assert_eq!(get.invoke_kind, RuntimeMemberInvokeKind::PropertyGet);
        assert_eq!(cache.len(), 1);
        let default_get = cache
            .resolve_default_member_unhinted(&INTERFACE, 0)
            .expect("unique unhinted default get should resolve and cache separately");
        assert_eq!(default_get.member_index, 0);
        assert_eq!(cache.len(), 2);
        assert!(
            cache
                .resolve_member_unhinted(&INTERFACE, "Value", 1)
                .is_some(),
            "different arity can resolve to the put member"
        );
        assert_eq!(cache.len(), 3);
    }

    #[test]
    fn runtime_dispatch_plan_cache_rejects_unhinted_ambiguous_member_lookup() {
        static GET_VALUE: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 0,
            vtable_slot: Some(3),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Variant),
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static METHOD_VALUE: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 1,
            vtable_slot: Some(4),
            invoke_kind: RuntimeMemberInvokeKind::Method,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Variant),
            is_default_member: false,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "IAmbiguous",
            members: &[GET_VALUE, METHOD_VALUE],
            dual_dispatch: true,
        };

        let mut cache = RuntimeDispatchPlanCache::new();
        assert!(
            cache
                .resolve_member_unhinted(&INTERFACE, "Value", 0)
                .is_none()
        );
        assert!(
            cache
                .resolve_default_member_unhinted(&INTERFACE, 0)
                .is_some()
        );
        assert_eq!(
            cache.len(),
            1,
            "only the unambiguous default lookup should cache"
        );
    }

    #[test]
    fn runtime_dispatch_plan_cache_rejects_ambiguous_default_member() {
        static FIRST: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 0,
            vtable_slot: Some(3),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Variant),
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static SECOND: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Item",
            dispatch_id: 0,
            vtable_slot: Some(4),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Variant),
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "IAmbiguousDefault",
            members: &[FIRST, SECOND],
            dual_dispatch: true,
        };

        let mut cache = RuntimeDispatchPlanCache::new();
        assert!(
            cache
                .resolve_default_member(&INTERFACE, RuntimeMemberInvokeKind::PropertyGet, 0)
                .is_none(),
            "ambiguous default member metadata must not be cached as a single plan"
        );
        assert!(cache.is_empty());
    }

    #[test]
    fn descriptor_backed_object_serves_only_iunknown_through_its_vtable() {
        const CUSTOM_GUID: Guid = Guid::new(
            0x3333_3333,
            0x4444,
            0x5555,
            [0x66, 0x66, 0x77, 0x77, 0x77, 0x77, 0x77, 0x77],
        );
        const CUSTOM_IDENTITY: RuntimeInterfaceIdentity = RuntimeInterfaceIdentity::custom(
            CUSTOM_GUID,
            "Project._WidgetCustom",
            RuntimeInterfaceKind::Dual,
            Some(1),
            Some(0),
            Some(1033),
        );
        static VALUE_MEMBER: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 0,
            vtable_slot: Some(7),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Variant),
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static DISPATCH_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "ITestDual",
            members: &[VALUE_MEMBER],
            dual_dispatch: true,
        };
        static CUSTOM_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::Unsupported,
            identity: CUSTOM_IDENTITY,
            name: "Project._WidgetCustom",
            members: &[VALUE_MEMBER],
            dual_dispatch: true,
        };
        static TEST_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
            name: "Project.Widget",
            runtime_type_name: None,
            project_identity: None,
            default_interface_iid: None,
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
            fields: &[],
            as_new_fields: &[],
            implements: &[],
            interfaces: &[
                RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR,
                DISPATCH_INTERFACE,
                CUSTOM_INTERFACE,
            ],
        };

        let object = ObjectRef::from_compat_identity_with_descriptor(12, &TEST_CLASS);
        let mut dispatch_out = core::ptr::null_mut();
        let dispatch_hr = unsafe {
            ((*(*object.raw_iunknown()).vtbl).query_interface)(
                object.raw_iunknown().cast(),
                &RUNTIME_GUID_IDISPATCH,
                &mut dispatch_out,
            )
        };
        // dc-yqq0.15: the three-slot object is no IDispatch; the class's interfaces are
        // descriptor projections, not binary interfaces of this vtable.
        assert_eq!(dispatch_hr, RUNTIME_E_NOINTERFACE);
        assert!(dispatch_out.is_null());
        assert_eq!(object.strong_count_for_test(), 1);
        assert!(
            object
                .query_interface_descriptor(RuntimeInterfaceId::IDispatch)
                .is_some()
        );

        let mut custom_out = core::ptr::null_mut();
        let custom_hr = unsafe {
            ((*(*object.raw_iunknown()).vtbl).query_interface)(
                object.raw_iunknown().cast(),
                &CUSTOM_GUID,
                &mut custom_out,
            )
        };
        assert_eq!(custom_hr, RUNTIME_E_NOINTERFACE);
        assert!(custom_out.is_null());
        assert_eq!(object.strong_count_for_test(), 1);
        assert!(
            object
                .query_interface_descriptor_by_guid(CUSTOM_GUID)
                .is_some()
        );
    }

    #[test]
    fn iunknown_identity_is_intrinsic_even_when_descriptor_omits_it() {
        static VALUE_MEMBER: RuntimeMemberDescriptor = RuntimeMemberDescriptor {
            name: "Value",
            dispatch_id: 0,
            vtable_slot: Some(7),
            invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
            arity: 0,
            params: &[],
            return_type: Some(RuntimeValueType::Variant),
            is_default_member: true,
            is_enumerator_member: false,
            externally_visible: true,
            backing: RuntimeMemberBacking::None,
        };
        static DISPATCH_INTERFACE: RuntimeInterfaceDescriptor = RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::IDispatch,
            identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
            name: "ITestDual",
            members: &[VALUE_MEMBER],
            dual_dispatch: true,
        };
        static TEST_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
            name: "Project.Widget",
            runtime_type_name: None,
            project_identity: None,
            default_interface_iid: None,
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
            fields: &[],
            as_new_fields: &[],
            implements: &[],
            interfaces: &[DISPATCH_INTERFACE],
        };

        let object = ObjectRef::from_compat_identity_with_descriptor(23, &TEST_CLASS);
        let iunknown = object
            .query_interface_descriptor(RuntimeInterfaceId::IUnknown)
            .expect("IUnknown must be intrinsic to every runtime object");
        assert_eq!(iunknown.name, "IUnknown");
        let projection = object
            .query_interface_projection_by_guid(RUNTIME_GUID_IUNKNOWN)
            .expect("IUnknown projection must not depend on class descriptor rows");
        assert_eq!(projection.object_identity.compat_identity, 23);
        assert_eq!(projection.interface_identity.guid, RUNTIME_GUID_IUNKNOWN);
        assert!(
            object
                .query_interface_descriptor(RuntimeInterfaceId::IDispatch)
                .is_some(),
            "class-advertised interfaces still resolve normally"
        );

        let mut unknown_out = core::ptr::null_mut();
        let hr = unsafe {
            ((*(*object.raw_iunknown()).vtbl).query_interface)(
                object.raw_iunknown().cast(),
                &RUNTIME_GUID_IUNKNOWN,
                &mut unknown_out,
            )
        };
        assert_eq!(hr, RUNTIME_S_OK);
        assert_eq!(unknown_out, object.raw_iunknown().cast());
        assert_eq!(object.strong_count_for_test(), 2);
        unsafe {
            ((*(*object.raw_iunknown()).vtbl).release)(unknown_out);
        }
        assert_eq!(object.strong_count_for_test(), 1);
    }

    #[test]
    fn object_ref_only_supports_iunknown_query() {
        let object = ObjectRef::from_compat_identity(9);
        let mut out = core::ptr::null_mut();
        let vtbl = unsafe { (*object.raw_iunknown()).vtbl };
        let hr = unsafe {
            ((*vtbl).query_interface)(
                object.raw_iunknown().cast(),
                &RUNTIME_GUID_IUNKNOWN,
                &mut out,
            )
        };
        assert_eq!(hr, 0);
        assert_eq!(out, object.raw_iunknown().cast());
        unsafe {
            ((*vtbl).release)(out);
        }

        let hr = unsafe {
            ((*vtbl).query_interface)(
                object.raw_iunknown().cast(),
                &Guid::new(
                    0xFFFF_FFFF,
                    0xFFFF,
                    0xFFFF,
                    [0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF],
                ),
                &mut out,
            )
        };
        assert_eq!(hr, RUNTIME_E_NOINTERFACE);
        assert!(out.is_null());
        assert_eq!(object.strong_count_for_test(), 1);
    }

    #[repr(C)]
    struct FakeForeignObject {
        unknown: RawRuntimeIUnknown,
        ref_count: AtomicU32,
    }

    static FAKE_FOREIGN_VTBL: IUnknownVtbl = IUnknownVtbl {
        query_interface: fake_foreign_query_interface,
        add_ref: fake_foreign_add_ref,
        release: fake_foreign_release,
    };

    unsafe extern "system" fn fake_foreign_query_interface(
        _this: *mut c_void,
        _iid: *const Guid,
        ppv: *mut *mut c_void,
    ) -> i32 {
        if !ppv.is_null() {
            unsafe {
                *ppv = core::ptr::null_mut();
            }
        }
        RUNTIME_E_NOINTERFACE
    }

    unsafe extern "system" fn fake_foreign_add_ref(this: *mut c_void) -> u32 {
        let object = this.cast::<FakeForeignObject>();
        unsafe { (*object).ref_count.fetch_add(1, Ordering::AcqRel) + 1 }
    }

    unsafe extern "system" fn fake_foreign_release(this: *mut c_void) -> u32 {
        let object = this.cast::<FakeForeignObject>();
        let previous = unsafe { (*object).ref_count.fetch_sub(1, Ordering::AcqRel) };
        let remaining = previous.saturating_sub(1);
        if remaining == 0 {
            unsafe {
                drop(Box::from_raw(object));
            }
        }
        remaining
    }

    fn fake_foreign_object() -> ObjectRef {
        let boxed = Box::new(FakeForeignObject {
            unknown: RawRuntimeIUnknown {
                vtbl: &FAKE_FOREIGN_VTBL,
            },
            ref_count: AtomicU32::new(1),
        });
        let raw = Box::into_raw(boxed);
        // SAFETY: `FakeForeignObject` is `repr(C)` with `unknown` first. Casting the
        // allocation pointer preserves full-object provenance needed by the fake vtable's
        // adjacent refcount access, and ownership transfers to the returned ObjectRef.
        unsafe {
            ObjectRef::from_raw_iunknown_owned(raw.cast::<RawRuntimeIUnknown>())
                .expect("fake foreign object pointer is non-null")
        }
    }

    #[test]
    fn foreign_iunknown_has_no_runtime_descriptor_projection() {
        let object = fake_foreign_object();
        assert!(!object.is_compat_object());
        assert!(object.try_class_descriptor().is_none());
        assert!(object.try_object_identity().is_none());
        assert!(
            object
                .query_interface_descriptor(RuntimeInterfaceId::IUnknown)
                .is_none()
        );
        assert!(
            object
                .query_interface_descriptor_by_guid(RUNTIME_GUID_IUNKNOWN)
                .is_none()
        );
        assert!(
            object
                .query_interface_projection(RuntimeInterfaceId::IUnknown)
                .is_none()
        );
        assert!(
            object
                .query_interface_projection_by_guid(RUNTIME_GUID_IUNKNOWN)
                .is_none()
        );
    }

    #[test]
    fn dispatch_plan_cache_and_dispid_slot_lookup_resolve_project_members_and_fields() {
        use crate::class_descriptor::{
            RuntimeClassFieldSpec, RuntimeClassSpec, RuntimeMemberSpec,
            build_runtime_class_descriptor,
        };
        let spec = RuntimeClassSpec {
            unit_name: "DispatchCacheUnit".into(),
            class_index: 0,
            name: "Widget".into(),
            runtime_type_name: None,
            default_interface_iid: None,
            predeclared: false,
            lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
            fields: vec![
                RuntimeClassFieldSpec {
                    name: "Value".into(),
                    token: 10,
                    value_type: RuntimeValueType::Long,
                    array_element_type: None,
                    externally_visible: true,
                    dispatch_id: None,
                    get_vtable_slot: None,
                    put_vtable_slot: None,
                },
                RuntimeClassFieldSpec {
                    name: "Tag".into(),
                    token: 11,
                    value_type: RuntimeValueType::Variant,
                    array_element_type: None,
                    externally_visible: true,
                    dispatch_id: None,
                    get_vtable_slot: None,
                    put_vtable_slot: None,
                },
            ],
            as_new_fields: Vec::new(),
            implements: Vec::new(),
            methods: vec![
                RuntimeMemberSpec {
                    name: "DefaultVal".into(),
                    dispatch_id: Some(0),
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
                    params: Vec::new(),
                    return_type: Some(RuntimeValueType::Long),
                    is_default_member: true,
                    is_enumerator_member: false,
                    externally_visible: true,
                    backing: RuntimeMemberBacking::Proc(1),
                },
                RuntimeMemberSpec {
                    name: "NewEnum".into(),
                    dispatch_id: Some(-4),
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
                    params: Vec::new(),
                    return_type: Some(RuntimeValueType::Object),
                    is_default_member: false,
                    is_enumerator_member: true,
                    externally_visible: true,
                    backing: RuntimeMemberBacking::Proc(2),
                },
                RuntimeMemberSpec {
                    name: "Twice".into(),
                    dispatch_id: None,
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::Method,
                    params: Vec::new(),
                    return_type: Some(RuntimeValueType::Long),
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: true,
                    backing: RuntimeMemberBacking::Proc(3),
                },
                RuntimeMemberSpec {
                    name: "FriendStep".into(),
                    dispatch_id: None,
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::Method,
                    params: Vec::new(),
                    return_type: None,
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: false,
                    backing: RuntimeMemberBacking::Proc(4),
                },
            ],
            implemented_interfaces: Vec::new(),
        };
        let descriptor = build_runtime_class_descriptor(&spec);
        let dispatch = descriptor.dispatch_interface().expect("dual interface");

        // DISPID 0 default member and DISPID -4 enumerator resolve directly without name scan.
        let default_m = descriptor
            .resolve_member_by_dispid(0, RuntimeMemberInvokeKind::PropertyGet, true, true)
            .expect("default member");
        assert_eq!(default_m.backing, RuntimeMemberBacking::Proc(1));
        let enum_m = descriptor
            .resolve_enumerator_member()
            .expect("enumerator member");
        assert_eq!(enum_m.backing, RuntimeMemberBacking::Proc(2));

        // O(1) vtable slot lookup round-trips every dual member.
        for member in dispatch.members {
            let slot = member.vtable_slot.expect("dual slot");
            assert_eq!(descriptor.member_by_vtable_slot(slot), Some(member));
        }

        // Plan cache resolves public method (with PropertyGet fallback), public Variant field Set,
        // filters non-public members when public_only = true, and reuses cached entries.
        let mut cache = RuntimeDispatchPlanCache::new();
        let twice1 = cache
            .resolve_project_member(
                dispatch,
                "Twice",
                RuntimeMemberInvokeKind::PropertyGet,
                true,
                true,
            )
            .expect("PropertyGet(0) falls back to Method");
        let twice2 = cache
            .resolve_project_member(
                dispatch,
                "TWICE",
                RuntimeMemberInvokeKind::PropertyGet,
                true,
                true,
            )
            .expect("cached hit");
        assert_eq!(twice1, twice2);
        assert_eq!(cache.project_plan_count(), 1);
        assert_eq!(
            dispatch.member_at(twice1.member_index).map(|m| m.backing),
            Some(RuntimeMemberBacking::Proc(3))
        );

        let tag_set = cache
            .resolve_project_member(
                dispatch,
                "Tag",
                RuntimeMemberInvokeKind::PropertySet,
                false,
                true,
            )
            .expect("Variant public field accepts PropertySet");
        assert_eq!(
            dispatch.member_at(tag_set.member_index).map(|m| m.backing),
            Some(RuntimeMemberBacking::Field(11))
        );

        assert!(
            cache
                .resolve_project_member(
                    dispatch,
                    "FriendStep",
                    RuntimeMemberInvokeKind::Method,
                    true,
                    true,
                )
                .is_none(),
            "public_only must reject Friend members"
        );
        let friend_plan = cache
            .resolve_project_member(
                dispatch,
                "FriendStep",
                RuntimeMemberInvokeKind::Method,
                true,
                false,
            )
            .expect("same-project internal receiver resolves Friend member");
        assert_eq!(
            dispatch
                .member_at(friend_plan.member_index)
                .map(|m| m.backing),
            Some(RuntimeMemberBacking::Proc(4))
        );
    }
}
