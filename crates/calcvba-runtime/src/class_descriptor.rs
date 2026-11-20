//! Shared runtime class descriptor builder (dc-yqq0.21).
//!
//! VM3, rt-abi/JIT, and the native AOT template all construct `RuntimeClassDescriptor` through
//! [`build_runtime_class_descriptor`] from one [`RuntimeClassSpec`]. Descriptors are interned
//! once per unique class spec for the process lifetime so repeated activations and executors
//! share the same `&'static RuntimeClassDescriptor` (and no allocations leak past global roots).
//!
//! Provisional DISPID rule (until Excel's own `ITypeInfo` DISPIDs are measured on Windows):
//! - default member (`VB_UserMemId = 0`): `0`
//! - enumerator (`VB_UserMemId = -4`): `-4`
//! - other members: sequential `1, 2, ...` per unique member name (shared across Property
//!   Get/Let/Set and public field Get/Let/Set), public members numbered first matching
//!   `calcvba-symbol::surface`, followed by any non-public (`Friend`/`Private`) methods.
//! - dual vtable slots start at `7` (`IUnknown` 0..2, `IDispatch` 3..6), one slot per accessor,
//!   ordered by `vtable_slot`.

use std::collections::{BTreeSet, HashMap};
use std::sync::{Mutex, OnceLock};

use crate::object_ref::{
    Guid, RUNTIME_IDISPATCH_INTERFACE_IDENTITY, RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR,
    RuntimeClassActivationDescriptor, RuntimeClassAsNewFieldDescriptor, RuntimeClassDescriptor,
    RuntimeClassFieldDescriptor, RuntimeClassLifecycleDescriptor, RuntimeInterfaceDescriptor,
    RuntimeInterfaceId, RuntimeInterfaceIdentity, RuntimeInterfaceKind, RuntimeMemberBacking,
    RuntimeMemberDescriptor, RuntimeMemberInvokeKind, RuntimeParamDescriptor,
    RuntimeProjectClassIdentity, RuntimeValueType,
};

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RuntimeParamSpec {
    pub name: String,
    pub value_type: RuntimeValueType,
    pub by_ref: bool,
    pub optional: bool,
    pub param_array: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RuntimeMemberSpec {
    pub name: String,
    pub dispatch_id: Option<i32>,
    pub vtable_slot: Option<u16>,
    pub invoke_kind: RuntimeMemberInvokeKind,
    pub params: Vec<RuntimeParamSpec>,
    pub return_type: Option<RuntimeValueType>,
    pub is_default_member: bool,
    pub is_enumerator_member: bool,
    pub externally_visible: bool,
    pub backing: RuntimeMemberBacking,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RuntimeClassFieldSpec {
    pub name: String,
    pub token: i32,
    pub value_type: RuntimeValueType,
    pub array_element_type: Option<RuntimeValueType>,
    pub externally_visible: bool,
    pub dispatch_id: Option<i32>,
    pub get_vtable_slot: Option<u16>,
    pub put_vtable_slot: Option<u16>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum RuntimeClassActivationSpec {
    ProjectClass { class_index: usize },
    ExternClass { import_index: usize },
    ComClass { prog_id: String },
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RuntimeClassAsNewFieldSpec {
    pub field_token: i32,
    pub activation: RuntimeClassActivationSpec,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RuntimeImplementedInterfaceSpec {
    pub name: String,
    pub qualified_name: String,
    pub guid: Option<Guid>,
    pub members: Vec<RuntimeMemberSpec>,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RuntimeClassSpec {
    pub unit_name: String,
    pub class_index: usize,
    pub name: String,
    pub runtime_type_name: Option<String>,
    pub default_interface_iid: Option<Guid>,
    pub predeclared: bool,
    pub lifecycle: RuntimeClassLifecycleDescriptor,
    pub fields: Vec<RuntimeClassFieldSpec>,
    pub as_new_fields: Vec<RuntimeClassAsNewFieldSpec>,
    pub implements: Vec<String>,
    pub methods: Vec<RuntimeMemberSpec>,
    pub implemented_interfaces: Vec<RuntimeImplementedInterfaceSpec>,
}

/// Deterministic project-interface GUID for a class-defined `Implements` interface.
pub fn runtime_project_interface_guid(unit_name: &str, interface_name: &str) -> Guid {
    const OFFSET: u64 = 0xcbf2_9ce4_8422_2325;
    const PRIME: u64 = 0x0000_0100_0000_01b3;
    fn fnv1a64(seed: u64, input: &str) -> u64 {
        let mut hash = seed;
        for byte in input.bytes() {
            hash ^= u64::from(byte);
            hash = hash.wrapping_mul(PRIME);
        }
        hash
    }
    let key = format!(
        "oxvba:project-interface:{}:{}",
        unit_name.to_ascii_lowercase(),
        interface_name.to_ascii_lowercase()
    );
    let hi = fnv1a64(OFFSET, &key);
    let lo = fnv1a64(!OFFSET, &key);
    Guid::new(
        (hi >> 32) as u32,
        (hi >> 16) as u16,
        hi as u16,
        lo.to_be_bytes(),
    )
}

/// Deterministic default dual-interface IID for a project class (matches `calcvba-build`'s
/// `deterministic_uuid(unit_name, &format!("I{class_name}"))` and class-typed SAFEARRAY IIDs).
pub fn default_class_interface_iid(unit_name: &str, class_name: &str) -> Guid {
    Guid::deterministic(unit_name, &format!("I{class_name}"))
}

fn descriptor_cache() -> &'static Mutex<HashMap<RuntimeClassSpec, &'static RuntimeClassDescriptor>>
{
    static CACHE: OnceLock<Mutex<HashMap<RuntimeClassSpec, &'static RuntimeClassDescriptor>>> =
        OnceLock::new();
    CACHE.get_or_init(|| Mutex::new(HashMap::new()))
}

/// Build (or return the cached) `'static RuntimeClassDescriptor` for `spec`.
///
/// Every executor (VM3, rt-abi/JIT, native AOT template) builds project class descriptors
/// through this single function.
pub fn build_runtime_class_descriptor(spec: &RuntimeClassSpec) -> &'static RuntimeClassDescriptor {
    let mut cache = descriptor_cache()
        .lock()
        .expect("runtime class descriptor cache lock");
    if let Some(&descriptor) = cache.get(spec) {
        return descriptor;
    }
    let descriptor = build_descriptor_uncached(spec);
    cache.insert(spec.clone(), descriptor);
    descriptor
}

fn leak_str(value: &str) -> &'static str {
    Box::leak(value.to_string().into_boxed_str())
}

fn finalize_members(raw: Vec<RuntimeMemberSpec>) -> &'static [RuntimeMemberDescriptor] {
    let mut used_dispids = BTreeSet::new();
    let mut name_to_dispid: HashMap<String, i32> = HashMap::new();
    let mut used_slots = BTreeSet::new();

    for member in &raw {
        if let Some(dispid) = member.dispatch_id {
            used_dispids.insert(dispid);
            name_to_dispid
                .entry(member.name.to_ascii_lowercase())
                .or_insert(dispid);
        }
        if let Some(slot) = member.vtable_slot {
            used_slots.insert(slot);
        }
    }

    let mut next_dispid = used_dispids
        .iter()
        .copied()
        .filter(|&d| d > 0)
        .max()
        .unwrap_or(0)
        .saturating_add(1);
    let mut next_slot = used_slots
        .iter()
        .copied()
        .max()
        .unwrap_or(6)
        .saturating_add(1)
        .max(7);

    let mut resolved_dispids = vec![0i32; raw.len()];
    let mut resolved_slots = vec![0u16; raw.len()];

    // Public members first so any unslotted non-public methods land strictly after the public
    // dual interface slots.
    for pass_public in [true, false] {
        for (index, member) in raw.iter().enumerate() {
            if member.externally_visible != pass_public {
                continue;
            }
            let folded = member.name.to_ascii_lowercase();
            let dispid = if let Some(dispid) = member.dispatch_id {
                dispid
            } else if member.is_default_member {
                0
            } else if member.is_enumerator_member {
                -4
            } else if let Some(&existing) = name_to_dispid.get(&folded) {
                existing
            } else {
                while used_dispids.contains(&next_dispid) {
                    next_dispid = next_dispid.saturating_add(1);
                }
                let assigned = next_dispid;
                next_dispid = next_dispid.saturating_add(1);
                used_dispids.insert(assigned);
                name_to_dispid.insert(folded, assigned);
                assigned
            };
            resolved_dispids[index] = dispid;

            let slot = if let Some(slot) = member.vtable_slot {
                slot
            } else {
                while used_slots.contains(&next_slot) {
                    next_slot = next_slot.saturating_add(1);
                }
                let assigned = next_slot;
                next_slot = next_slot.saturating_add(1);
                used_slots.insert(assigned);
                assigned
            };
            resolved_slots[index] = slot;
        }
    }

    let mut indexed: Vec<(usize, u16, i32, RuntimeMemberSpec)> = raw
        .into_iter()
        .enumerate()
        .map(|(index, member)| {
            (
                index,
                resolved_slots[index],
                resolved_dispids[index],
                member,
            )
        })
        .collect();
    indexed.sort_by_key(|(index, slot, _, _)| (*slot, *index));

    let descriptors: Vec<RuntimeMemberDescriptor> = indexed
        .into_iter()
        .map(|(_, slot, dispid, member)| {
            let params: &'static [RuntimeParamDescriptor] = Box::leak(
                member
                    .params
                    .iter()
                    .map(|param| RuntimeParamDescriptor {
                        name: leak_str(&param.name),
                        value_type: param.value_type,
                        by_ref: param.by_ref,
                        optional: param.optional,
                        param_array: param.param_array,
                    })
                    .collect::<Vec<_>>()
                    .into_boxed_slice(),
            );
            RuntimeMemberDescriptor {
                name: leak_str(&member.name),
                dispatch_id: dispid,
                vtable_slot: Some(slot),
                invoke_kind: member.invoke_kind,
                arity: params.len(),
                params,
                return_type: member.return_type,
                is_default_member: member.is_default_member,
                is_enumerator_member: member.is_enumerator_member,
                externally_visible: member.externally_visible,
                backing: member.backing,
            }
        })
        .collect();
    Box::leak(descriptors.into_boxed_slice())
}

fn build_descriptor_uncached(spec: &RuntimeClassSpec) -> &'static RuntimeClassDescriptor {
    let name = leak_str(&spec.name);
    let runtime_type_name = spec.runtime_type_name.as_deref().map(leak_str);
    let project_identity = Some(RuntimeProjectClassIdentity {
        unit_name: leak_str(&spec.unit_name),
        class_index: spec.class_index,
    });
    let default_interface_iid = Some(
        spec.default_interface_iid
            .unwrap_or_else(|| default_class_interface_iid(&spec.unit_name, &spec.name)),
    );

    let fields: &'static [RuntimeClassFieldDescriptor] = Box::leak(
        spec.fields
            .iter()
            .map(|field| RuntimeClassFieldDescriptor {
                name: leak_str(&field.name),
                token: field.token,
                value_type: field.value_type,
                array_element_type: field.array_element_type,
            })
            .collect::<Vec<_>>()
            .into_boxed_slice(),
    );
    let as_new_fields: &'static [RuntimeClassAsNewFieldDescriptor] = Box::leak(
        spec.as_new_fields
            .iter()
            .map(|field| RuntimeClassAsNewFieldDescriptor {
                field_token: field.field_token,
                activation: match &field.activation {
                    RuntimeClassActivationSpec::ProjectClass { class_index } => {
                        RuntimeClassActivationDescriptor::ProjectClass {
                            class_index: *class_index,
                        }
                    }
                    RuntimeClassActivationSpec::ExternClass { import_index } => {
                        RuntimeClassActivationDescriptor::ExternClass {
                            import_index: *import_index,
                        }
                    }
                    RuntimeClassActivationSpec::ComClass { prog_id } => {
                        RuntimeClassActivationDescriptor::ComClass {
                            prog_id: leak_str(prog_id),
                        }
                    }
                },
            })
            .collect::<Vec<_>>()
            .into_boxed_slice(),
    );
    let implements: &'static [&'static str] = Box::leak(
        spec.implements
            .iter()
            .map(|iface| leak_str(iface))
            .collect::<Vec<_>>()
            .into_boxed_slice(),
    );

    // Combine explicit methods with property accessors for Public fields (matching
    // `calcvba-symbol::surface`: each public field surfaces as PropertyGet + PropertyLet/Set
    // sharing one DISPID and occupying two consecutive vtable slots).
    let mut dual_members = spec.methods.clone();
    for field in &spec.fields {
        if !field.externally_visible {
            continue;
        }
        let folded = field.name.to_ascii_lowercase();
        let has_get = dual_members.iter().any(|m| {
            m.invoke_kind == RuntimeMemberInvokeKind::PropertyGet
                && m.name.to_ascii_lowercase() == folded
        });
        if !has_get {
            dual_members.push(RuntimeMemberSpec {
                name: field.name.clone(),
                dispatch_id: field.dispatch_id,
                vtable_slot: field.get_vtable_slot,
                invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
                params: Vec::new(),
                return_type: Some(field.value_type),
                is_default_member: false,
                is_enumerator_member: false,
                externally_visible: true,
                backing: RuntimeMemberBacking::Field(field.token),
            });
        }
        let put_kind = if field.value_type == RuntimeValueType::Object {
            RuntimeMemberInvokeKind::PropertySet
        } else {
            RuntimeMemberInvokeKind::PropertyLet
        };
        let has_put = dual_members
            .iter()
            .any(|m| m.invoke_kind == put_kind && m.name.to_ascii_lowercase() == folded);
        if !has_put {
            dual_members.push(RuntimeMemberSpec {
                name: field.name.clone(),
                dispatch_id: field.dispatch_id,
                vtable_slot: field.put_vtable_slot,
                invoke_kind: put_kind,
                params: vec![RuntimeParamSpec {
                    name: "Value".to_string(),
                    value_type: field.value_type,
                    by_ref: false,
                    optional: false,
                    param_array: false,
                }],
                return_type: None,
                is_default_member: false,
                is_enumerator_member: false,
                externally_visible: true,
                backing: RuntimeMemberBacking::Field(field.token),
            });
        }
    }

    let members = finalize_members(dual_members);
    let dispatch = RuntimeInterfaceDescriptor {
        id: RuntimeInterfaceId::IDispatch,
        identity: RUNTIME_IDISPATCH_INTERFACE_IDENTITY,
        name: "IDispatch",
        members,
        dual_dispatch: true,
    };

    let mut interface_descriptors = vec![RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR, dispatch];
    for iface in &spec.implemented_interfaces {
        let guid = iface
            .guid
            .unwrap_or_else(|| runtime_project_interface_guid(&spec.unit_name, &iface.name));
        let iface_members = finalize_members(iface.members.clone());
        interface_descriptors.push(RuntimeInterfaceDescriptor {
            id: RuntimeInterfaceId::Unsupported,
            identity: RuntimeInterfaceIdentity::custom(
                guid,
                leak_str(&iface.qualified_name),
                RuntimeInterfaceKind::Custom,
                None,
                None,
                None,
            ),
            name: leak_str(&iface.name),
            members: iface_members,
            dual_dispatch: false,
        });
    }
    let interfaces: &'static [RuntimeInterfaceDescriptor] =
        Box::leak(interface_descriptors.into_boxed_slice());

    Box::leak(Box::new(RuntimeClassDescriptor {
        name,
        runtime_type_name,
        project_identity,
        default_interface_iid,
        predeclared: spec.predeclared,
        lifecycle: spec.lifecycle,
        fields,
        as_new_fields,
        implements,
        interfaces,
    }))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{ObjectRef, ObjectRefExt};

    fn widget_spec() -> RuntimeClassSpec {
        RuntimeClassSpec {
            unit_name: "Demo".into(),
            class_index: 0,
            name: "Widget".into(),
            runtime_type_name: None,
            default_interface_iid: None,
            predeclared: false,
            lifecycle: RuntimeClassLifecycleDescriptor {
                has_initialize: false,
                has_terminate: true,
            },
            fields: vec![RuntimeClassFieldSpec {
                name: "Value".into(),
                token: 0,
                value_type: RuntimeValueType::Long,
                array_element_type: None,
                externally_visible: true,
                dispatch_id: Some(1),
                get_vtable_slot: Some(7),
                put_vtable_slot: Some(8),
            }],
            as_new_fields: Vec::new(),
            implements: Vec::new(),
            methods: vec![
                RuntimeMemberSpec {
                    name: "Twice".into(),
                    dispatch_id: Some(2),
                    vtable_slot: Some(9),
                    invoke_kind: RuntimeMemberInvokeKind::Method,
                    params: Vec::new(),
                    return_type: Some(RuntimeValueType::Long),
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: true,
                    backing: RuntimeMemberBacking::Proc(1),
                },
                RuntimeMemberSpec {
                    name: "InternalHelper".into(),
                    dispatch_id: None,
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::Method,
                    params: Vec::new(),
                    return_type: None,
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: false,
                    backing: RuntimeMemberBacking::Proc(2),
                },
            ],
            implemented_interfaces: Vec::new(),
        }
    }

    #[test]
    fn shared_builder_caches_descriptor_and_assigns_real_dispids_and_slots() {
        let spec = widget_spec();
        let first = build_runtime_class_descriptor(&spec);
        let second = build_runtime_class_descriptor(&spec);
        assert!(
            core::ptr::eq(first, second),
            "identical class specs must share one built-once descriptor"
        );
        assert_eq!(first.validate_com_export_shape(), Ok(()));
        assert_eq!(
            first.default_interface_iid,
            Some(Guid::deterministic("Demo", "IWidget"))
        );

        let dispatch = first.dispatch_interface().expect("dual IDispatch");
        assert!(dispatch.dual_dispatch);
        assert_eq!(dispatch.members.len(), 4);

        // Ordered by vtable slot: Value Get (7), Value Let (8), Twice (9), InternalHelper (10).
        assert_eq!(dispatch.members[0].name, "Value");
        assert_eq!(dispatch.members[0].dispatch_id, 1);
        assert_eq!(dispatch.members[0].vtable_slot, Some(7));
        assert_eq!(
            dispatch.members[0].invoke_kind,
            RuntimeMemberInvokeKind::PropertyGet
        );
        assert!(dispatch.members[0].externally_visible);
        assert_eq!(dispatch.members[0].backing, RuntimeMemberBacking::Field(0));

        assert_eq!(dispatch.members[1].name, "Value");
        assert_eq!(dispatch.members[1].dispatch_id, 1);
        assert_eq!(dispatch.members[1].vtable_slot, Some(8));
        assert_eq!(
            dispatch.members[1].invoke_kind,
            RuntimeMemberInvokeKind::PropertyLet
        );
        assert_eq!(dispatch.members[1].backing, RuntimeMemberBacking::Field(0));

        assert_eq!(dispatch.members[2].name, "Twice");
        assert_eq!(dispatch.members[2].dispatch_id, 2);
        assert_eq!(dispatch.members[2].vtable_slot, Some(9));
        assert_eq!(dispatch.members[2].backing, RuntimeMemberBacking::Proc(1));

        assert_eq!(dispatch.members[3].name, "InternalHelper");
        assert_eq!(dispatch.members[3].dispatch_id, 3);
        assert_eq!(dispatch.members[3].vtable_slot, Some(10));
        assert!(!dispatch.members[3].externally_visible);
        assert_eq!(dispatch.members[3].backing, RuntimeMemberBacking::Proc(2));

        // Default interface IID resolves through ObjectRefExt::query_interface_descriptor_by_guid.
        let instance = ObjectRef::from_project_instance(1, 0, 0, false, first);
        let by_iid = instance
            .query_interface_descriptor_by_guid(Guid::deterministic("Demo", "IWidget"))
            .expect("class default interface IID resolves");
        assert_eq!(by_iid.id, RuntimeInterfaceId::IDispatch);
    }

    #[test]
    fn provisional_dispid_rule_shares_property_dispids_and_honors_default_and_newenum() {
        let spec = RuntimeClassSpec {
            unit_name: "Demo".into(),
            class_index: 1,
            name: "Bag".into(),
            runtime_type_name: None,
            default_interface_iid: None,
            predeclared: false,
            lifecycle: RuntimeClassLifecycleDescriptor {
                has_initialize: false,
                has_terminate: false,
            },
            fields: Vec::new(),
            as_new_fields: Vec::new(),
            implements: Vec::new(),
            methods: vec![
                RuntimeMemberSpec {
                    name: "Item".into(),
                    dispatch_id: None,
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
                    params: Vec::new(),
                    return_type: Some(RuntimeValueType::Variant),
                    is_default_member: true,
                    is_enumerator_member: false,
                    externally_visible: true,
                    backing: RuntimeMemberBacking::Proc(0),
                },
                RuntimeMemberSpec {
                    name: "Item".into(),
                    dispatch_id: None,
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::PropertyLet,
                    params: vec![RuntimeParamSpec {
                        name: "v".into(),
                        value_type: RuntimeValueType::Variant,
                        by_ref: false,
                        optional: false,
                        param_array: false,
                    }],
                    return_type: None,
                    is_default_member: true,
                    is_enumerator_member: false,
                    externally_visible: true,
                    backing: RuntimeMemberBacking::Proc(1),
                },
                RuntimeMemberSpec {
                    name: "NewEnum".into(),
                    dispatch_id: None,
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
                    name: "Label".into(),
                    dispatch_id: None,
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::PropertyGet,
                    params: Vec::new(),
                    return_type: Some(RuntimeValueType::String),
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: true,
                    backing: RuntimeMemberBacking::Proc(3),
                },
                RuntimeMemberSpec {
                    name: "Label".into(),
                    dispatch_id: None,
                    vtable_slot: None,
                    invoke_kind: RuntimeMemberInvokeKind::PropertyLet,
                    params: vec![RuntimeParamSpec {
                        name: "v".into(),
                        value_type: RuntimeValueType::String,
                        by_ref: false,
                        optional: false,
                        param_array: false,
                    }],
                    return_type: None,
                    is_default_member: false,
                    is_enumerator_member: false,
                    externally_visible: true,
                    backing: RuntimeMemberBacking::Proc(4),
                },
            ],
            implemented_interfaces: Vec::new(),
        };
        let desc = build_runtime_class_descriptor(&spec);
        assert_eq!(desc.validate_com_export_shape(), Ok(()));
        let members = desc.dispatch_interface().unwrap().members;
        assert_eq!(members[0].dispatch_id, 0);
        assert_eq!(members[0].vtable_slot, Some(7));
        assert_eq!(members[1].dispatch_id, 0);
        assert_eq!(members[1].vtable_slot, Some(8));
        assert_eq!(members[2].dispatch_id, -4);
        assert_eq!(members[2].vtable_slot, Some(9));
        assert_eq!(members[3].dispatch_id, 1);
        assert_eq!(members[3].vtable_slot, Some(10));
        assert_eq!(members[4].dispatch_id, 1);
        assert_eq!(members[4].vtable_slot, Some(11));
    }
}
