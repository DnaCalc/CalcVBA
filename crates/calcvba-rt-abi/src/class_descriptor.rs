//! Single OxIR/native adapter feeding `calcvba_runtime::build_runtime_class_descriptor`.
//!
//! VM3, rt-abi/JIT, and the native AOT template all build their `RuntimeClassDescriptor`s through
//! this module and `calcvba_runtime::build_runtime_class_descriptor`, guaranteeing identical
//! DISPIDs, vtable slots, default interface IIDs, fields, and implemented interfaces across all
//! three executors (dc-yqq0.21).

use calcvba_bundle::{ArrayElementType, NumericCoerceTarget, ProjectMemberKind};
use calcvba_com::{TypeLibMemberInvokeKind, TypeLibMemberMetadata};
use calcvba_eval::arith;
use calcvba_oxir::{OxProgram, OxTy, inst::OxAsNew, program::OxClassMethod};
use calcvba_runtime::{
    Guid, RuntimeClassActivationSpec, RuntimeClassAsNewFieldSpec, RuntimeClassDescriptor,
    RuntimeClassFieldSpec, RuntimeClassLifecycleDescriptor, RuntimeClassSpec,
    RuntimeDispatchPlanCache, RuntimeImplementedInterfaceSpec, RuntimeMemberBacking,
    RuntimeMemberDescriptor, RuntimeMemberInvokeKind, RuntimeMemberSpec, RuntimeParamSpec,
    RuntimeValueType, VarType, Variant, build_runtime_class_descriptor, default_error_message,
    runtime_project_interface_guid,
};

const NATIVE_METADATA_MAGIC: [u8; 4] = *b"CVBD";
const NATIVE_METADATA_VERSION_V1: u32 = 1;
const NATIVE_METADATA_VERSION: u32 = 2;

/// Convert a bundle [`ProjectMemberKind`] into [`RuntimeMemberInvokeKind`].
pub fn runtime_invoke_kind(kind: ProjectMemberKind) -> RuntimeMemberInvokeKind {
    match kind {
        ProjectMemberKind::Method => RuntimeMemberInvokeKind::Method,
        ProjectMemberKind::PropertyGet => RuntimeMemberInvokeKind::PropertyGet,
        ProjectMemberKind::PropertyLet => RuntimeMemberInvokeKind::PropertyLet,
        ProjectMemberKind::PropertySet => RuntimeMemberInvokeKind::PropertySet,
    }
}

/// Convert a typelib [`TypeLibMemberInvokeKind`] into [`RuntimeMemberInvokeKind`].
pub fn runtime_invoke_kind_from_typelib(kind: TypeLibMemberInvokeKind) -> RuntimeMemberInvokeKind {
    match kind {
        TypeLibMemberInvokeKind::Method => RuntimeMemberInvokeKind::Method,
        TypeLibMemberInvokeKind::PropertyGet => RuntimeMemberInvokeKind::PropertyGet,
        TypeLibMemberInvokeKind::PropertyPut => RuntimeMemberInvokeKind::PropertyLet,
        TypeLibMemberInvokeKind::PropertyPutRef => RuntimeMemberInvokeKind::PropertySet,
    }
}

/// Resolve a project-instance member call without scanning `OxClass::methods`:
///
/// - Early-bound calls (`early_member.is_some()`) dispatch by vtable slot or DISPID on the
///   matching class/implemented interface descriptor when its IID matches, falling back to
///   `RuntimeDispatchPlanCache` for imported COM interfaces implemented by a VBA class.
/// - Default-member calls (`name.is_empty()`) dispatch directly by DISPID `0`.
/// - Late-bound named calls resolve exclusively through `RuntimeDispatchPlanCache` and index
///   directly into the class descriptor's dual interface member table.
pub fn resolve_project_member_for_call(
    cache: &mut RuntimeDispatchPlanCache,
    descriptor: &'static RuntimeClassDescriptor,
    early_member: Option<&TypeLibMemberMetadata>,
    name: &str,
    invoke_kind: RuntimeMemberInvokeKind,
    args_empty: bool,
    public_only: bool,
) -> Option<&'static RuntimeMemberDescriptor> {
    if let Some(meta) = early_member {
        if let Some(iid) = meta.interface_iid {
            let guid = Guid::new(iid.data1, iid.data2, iid.data3, iid.data4);
            if descriptor.default_interface_iid == Some(guid) {
                if let Some(slot) = meta.vtable_slot
                    && let Some(member) = descriptor.member_by_vtable_slot(slot)
                    && (!public_only || member.externally_visible)
                {
                    return Some(member);
                }
                if let Some(member) = descriptor.resolve_member_by_dispid(
                    meta.token,
                    invoke_kind,
                    args_empty,
                    public_only,
                ) {
                    return Some(member);
                }
            }
            if let Some(iface) = descriptor
                .interfaces
                .iter()
                .find(|iface| iface.identity.guid == guid)
            {
                if let Some(slot) = meta.vtable_slot
                    && let Some(member) = iface.member_by_vtable_slot(slot)
                    && (!public_only || member.externally_visible)
                {
                    return Some(member);
                }
                if let Some(member) =
                    iface.resolve_member_by_dispid(meta.token, invoke_kind, args_empty, public_only)
                {
                    return Some(member);
                }
            }
        }
        let dispatch = descriptor.dispatch_interface()?;
        let lookup_name = if name.is_empty() {
            meta.name.as_str()
        } else {
            name
        };
        let plan = cache.resolve_project_member(
            dispatch,
            lookup_name,
            invoke_kind,
            args_empty,
            public_only,
        )?;
        return dispatch.member_at(plan.member_index);
    }

    if name.is_empty() {
        return descriptor.resolve_member_by_dispid(0, invoke_kind, args_empty, public_only);
    }

    let dispatch = descriptor.dispatch_interface()?;
    let plan =
        cache.resolve_project_member(dispatch, name, invoke_kind, args_empty, public_only)?;
    dispatch.member_at(plan.member_index)
}

/// Coerce `value` to a class field's runtime value type (used for late-bound / COM access to
/// `Public` class fields backed by `RuntimeMemberBacking::Field`).
pub fn coerce_value_for_runtime_type(
    value: &Variant,
    ty: RuntimeValueType,
) -> Result<Variant, crate::Fault> {
    let numeric = match ty {
        RuntimeValueType::Byte => Some(NumericCoerceTarget::Byte),
        RuntimeValueType::Integer => Some(NumericCoerceTarget::Integer),
        RuntimeValueType::Long => Some(NumericCoerceTarget::Long),
        RuntimeValueType::LongLong => Some(NumericCoerceTarget::LongLong),
        RuntimeValueType::LongPtr => {
            #[cfg(target_pointer_width = "64")]
            {
                Some(NumericCoerceTarget::LongLong)
            }
            #[cfg(not(target_pointer_width = "64"))]
            {
                Some(NumericCoerceTarget::Long)
            }
        }
        RuntimeValueType::Single => Some(NumericCoerceTarget::Single),
        RuntimeValueType::Double => Some(NumericCoerceTarget::Double),
        RuntimeValueType::Currency => Some(NumericCoerceTarget::Currency),
        RuntimeValueType::Date => Some(NumericCoerceTarget::Date),
        RuntimeValueType::Boolean => Some(NumericCoerceTarget::Boolean),
        _ => None,
    };
    if let Some(target) = numeric {
        return arith::coerce_numeric(value, target).map_err(crate::Fault::from_arith);
    }
    match ty {
        RuntimeValueType::String => arith::coerce_string(value).map_err(crate::Fault::from_arith),
        RuntimeValueType::Object => {
            if value.vtype() == VarType::Object {
                Ok(value.clone())
            } else {
                Err(crate::Fault::new(424, default_error_message(424)))
            }
        }
        RuntimeValueType::Variant
        | RuntimeValueType::Decimal
        | RuntimeValueType::Record
        | RuntimeValueType::Array => Ok(value.clone()),
        RuntimeValueType::Byte
        | RuntimeValueType::Integer
        | RuntimeValueType::Long
        | RuntimeValueType::LongLong
        | RuntimeValueType::LongPtr
        | RuntimeValueType::Single
        | RuntimeValueType::Double
        | RuntimeValueType::Currency
        | RuntimeValueType::Date
        | RuntimeValueType::Boolean => unreachable!("handled above"),
    }
}

/// Build (or return cached) `'static RuntimeClassDescriptor`s for every class in `program`.
pub fn runtime_class_descriptors_for_program(
    program: &OxProgram,
) -> Vec<&'static RuntimeClassDescriptor> {
    program
        .classes
        .iter()
        .enumerate()
        .map(|(class_index, _)| {
            let spec = runtime_class_spec_for_program(program, class_index);
            build_runtime_class_descriptor(&spec)
        })
        .collect()
}

/// Convert `program.classes[class_index]` into the shared [`RuntimeClassSpec`].
pub fn runtime_class_spec_for_program(program: &OxProgram, class_index: usize) -> RuntimeClassSpec {
    let class = &program.classes[class_index];
    let (unit_name, display_name) = match class.name.rsplit_once('.') {
        Some((unit, display)) if !unit.is_empty() => (unit.to_string(), display.to_string()),
        _ => (program.unit_name.clone(), class.name.clone()),
    };

    let fields = class
        .fields
        .iter()
        .map(|field| RuntimeClassFieldSpec {
            name: field.name.clone(),
            token: field.token,
            value_type: runtime_value_type(&field.ty),
            array_element_type: field.array_element.as_ref().map(runtime_array_element_type),
            externally_visible: field.externally_visible,
            dispatch_id: field.dispid,
            get_vtable_slot: field.get_vtable_slot,
            put_vtable_slot: field.put_vtable_slot,
            putref_vtable_slot: field.putref_vtable_slot,
        })
        .collect();

    let as_new_fields = class
        .as_new_fields
        .iter()
        .map(|field| RuntimeClassAsNewFieldSpec {
            field_token: field.field,
            activation: match &field.binding {
                OxAsNew::ProjectClass { class } => RuntimeClassActivationSpec::ProjectClass {
                    class_index: class.0,
                },
                OxAsNew::ExternClass { import } => RuntimeClassActivationSpec::ExternClass {
                    import_index: import.0,
                },
                OxAsNew::ComClass { prog_id } => RuntimeClassActivationSpec::ComClass {
                    prog_id: prog_id.clone(),
                },
            },
        })
        .collect();

    let methods = class
        .methods
        .iter()
        .map(|method| {
            member_spec(
                program,
                method,
                &method.name,
                method.dispid,
                method.vtable_slot,
                method.is_default_member,
                method.is_enumerator_member,
                method.externally_visible,
            )
        })
        .collect();

    let implemented_interfaces = class
        .implements
        .iter()
        .filter_map(|interface| implemented_interface_spec(program, &unit_name, class, interface))
        .collect();

    RuntimeClassSpec {
        unit_name,
        class_index,
        name: display_name,
        runtime_type_name: class.runtime_type_name.clone(),
        default_interface_iid: None,
        predeclared: class.predeclared,
        lifecycle: RuntimeClassLifecycleDescriptor {
            has_initialize: class.initialize.is_some(),
            has_terminate: class.terminate.is_some(),
        },
        fields,
        as_new_fields,
        implements: class.implements.clone(),
        methods,
        implemented_interfaces,
    }
}

#[allow(clippy::too_many_arguments)]
fn member_spec(
    program: &OxProgram,
    method: &OxClassMethod,
    display_name: &str,
    dispatch_id: Option<i32>,
    vtable_slot: Option<u16>,
    is_default_member: bool,
    is_enumerator_member: bool,
    externally_visible: bool,
) -> RuntimeMemberSpec {
    let proc = program.funcs.get(method.proc.0);
    RuntimeMemberSpec {
        name: display_name.to_string(),
        dispatch_id,
        vtable_slot,
        invoke_kind: runtime_member_invoke_kind(method.kind),
        params: proc.map(runtime_member_params).unwrap_or_default(),
        return_type: proc.and_then(runtime_return_type),
        is_default_member,
        is_enumerator_member,
        externally_visible,
        func_flags: 0,
        backing: RuntimeMemberBacking::Proc(method.proc.0),
    }
}

fn implemented_interface_spec(
    program: &OxProgram,
    unit_name: &str,
    class: &calcvba_oxir::OxClass,
    interface_name: &str,
) -> Option<RuntimeImplementedInterfaceSpec> {
    let bare_interface = interface_name.rsplit('.').next().unwrap_or(interface_name);
    let interface_class = program.classes.iter().find(|candidate| {
        let candidate_bare = candidate.name.rsplit('.').next().unwrap_or(&candidate.name);
        candidate_bare.eq_ignore_ascii_case(bare_interface)
    })?;
    let mut members = Vec::new();
    for interface_method in &interface_class.methods {
        let mangled = format!("{bare_interface}_{}", interface_method.name);
        let Some(implementation_method) = class.methods.iter().find(|method| {
            method.kind == interface_method.kind && method.name.eq_ignore_ascii_case(&mangled)
        }) else {
            continue;
        };
        members.push(member_spec(
            program,
            implementation_method,
            &interface_method.name,
            interface_method.dispid.or(implementation_method.dispid),
            interface_method
                .vtable_slot
                .or(implementation_method.vtable_slot),
            interface_method.is_default_member || implementation_method.is_default_member,
            interface_method.is_enumerator_member || implementation_method.is_enumerator_member,
            true,
        ));
    }
    let qualified_name = if unit_name.is_empty() {
        bare_interface.to_string()
    } else {
        format!("{unit_name}.{bare_interface}")
    };
    Some(RuntimeImplementedInterfaceSpec {
        name: bare_interface.to_string(),
        qualified_name,
        guid: Some(runtime_project_interface_guid(unit_name, bare_interface)),
        members,
    })
}

fn runtime_member_invoke_kind(kind: ProjectMemberKind) -> RuntimeMemberInvokeKind {
    match kind {
        ProjectMemberKind::Method => RuntimeMemberInvokeKind::Method,
        ProjectMemberKind::PropertyGet => RuntimeMemberInvokeKind::PropertyGet,
        ProjectMemberKind::PropertyLet => RuntimeMemberInvokeKind::PropertyLet,
        ProjectMemberKind::PropertySet => RuntimeMemberInvokeKind::PropertySet,
    }
}

pub(crate) fn hidden_me_receiver_param_count(func: &calcvba_oxir::OxFunc) -> usize {
    usize::from(
        func.param_count > 0
            && func.locals.first().is_some_and(|local| {
                local.param.is_some() && local.name.eq_ignore_ascii_case("Me")
            }),
    )
}

pub fn runtime_member_params(func: &calcvba_oxir::OxFunc) -> Vec<RuntimeParamSpec> {
    func.locals
        .iter()
        .take(func.param_count)
        .skip(hidden_me_receiver_param_count(func))
        .map(|local| {
            let param = local.param.as_ref();
            RuntimeParamSpec {
                name: local.name.clone(),
                value_type: runtime_value_type(&local.ty),
                by_ref: param.map(|p| p.by_ref).unwrap_or(false),
                optional: param.map(|p| p.optional).unwrap_or(false),
                param_array: param.map(|p| p.variadic).unwrap_or(false),
            }
        })
        .collect()
}

fn runtime_return_type(func: &calcvba_oxir::OxFunc) -> Option<RuntimeValueType> {
    func.return_local
        .and_then(|local| func.locals.get(local.0))
        .map(|local| runtime_value_type(&local.ty))
}

fn runtime_value_type(ty: &OxTy) -> RuntimeValueType {
    match ty {
        OxTy::Bool => RuntimeValueType::Boolean,
        OxTy::Byte => RuntimeValueType::Byte,
        OxTy::Integer => RuntimeValueType::Integer,
        OxTy::Long => RuntimeValueType::Long,
        OxTy::LongLong => RuntimeValueType::LongLong,
        OxTy::Currency => RuntimeValueType::Currency,
        OxTy::Single => RuntimeValueType::Single,
        OxTy::Double => RuntimeValueType::Double,
        OxTy::Date => RuntimeValueType::Date,
        OxTy::Decimal => RuntimeValueType::Decimal,
        OxTy::Str | OxTy::FixedStr(_) => RuntimeValueType::String,
        OxTy::Object(_) => RuntimeValueType::Object,
        OxTy::Record(_) => RuntimeValueType::Record,
        OxTy::Array(_, _) => RuntimeValueType::Array,
        OxTy::Variant | OxTy::ProcRef => RuntimeValueType::Variant,
    }
}

fn runtime_array_element_type(element: &ArrayElementType) -> RuntimeValueType {
    match element {
        ArrayElementType::Variant => RuntimeValueType::Variant,
        ArrayElementType::Integer => RuntimeValueType::Integer,
        ArrayElementType::Long => RuntimeValueType::Long,
        ArrayElementType::LongLong => RuntimeValueType::LongLong,
        ArrayElementType::Byte => RuntimeValueType::Byte,
        ArrayElementType::Single => RuntimeValueType::Single,
        ArrayElementType::Double => RuntimeValueType::Double,
        ArrayElementType::Currency => RuntimeValueType::Currency,
        ArrayElementType::Date => RuntimeValueType::Date,
        ArrayElementType::String | ArrayElementType::FixedString(_) => RuntimeValueType::String,
        ArrayElementType::Boolean => RuntimeValueType::Boolean,
        ArrayElementType::Record(_) => RuntimeValueType::Record,
        ArrayElementType::FixedArray { .. } | ArrayElementType::DynamicArray(_) => {
            RuntimeValueType::Array
        }
        ArrayElementType::Object | ArrayElementType::Interface(_) => RuntimeValueType::Object,
    }
}

// ── Native AOT class metadata encoding/decoding ──────────────────────────────
//
// Preserves the v1 prefix (`class_id`, `name_len`, `implements_count`, `field_count`, `name`,
// `implements`, `(token, kind)` pairs) so prebuilt v1 native templates keep working unchanged,
// and appends a `CVBD` v1 trailer carrying the full `RuntimeClassSpec` so rebuilt native
// templates build the exact same `RuntimeClassDescriptor` through the shared builder.

fn encode_u8(out: &mut Vec<u8>, v: u8) {
    out.push(v);
}

fn encode_u16(out: &mut Vec<u8>, v: u16) {
    out.extend_from_slice(&v.to_le_bytes());
}

fn encode_u32(out: &mut Vec<u8>, v: u32) {
    out.extend_from_slice(&v.to_le_bytes());
}

fn encode_i32(out: &mut Vec<u8>, v: i32) {
    out.extend_from_slice(&v.to_le_bytes());
}

fn encode_str(out: &mut Vec<u8>, s: &str) {
    encode_u32(out, s.len() as u32);
    out.extend_from_slice(s.as_bytes());
}

fn encode_opt_str(out: &mut Vec<u8>, s: Option<&str>) {
    match s {
        Some(value) => {
            encode_u8(out, 1);
            encode_str(out, value);
        }
        None => encode_u8(out, 0),
    }
}

fn encode_opt_i32(out: &mut Vec<u8>, value: Option<i32>) {
    match value {
        Some(v) => {
            encode_u8(out, 1);
            encode_i32(out, v);
        }
        None => encode_u8(out, 0),
    }
}

fn encode_opt_u16(out: &mut Vec<u8>, value: Option<u16>) {
    match value {
        Some(v) => {
            encode_u8(out, 1);
            encode_u16(out, v);
        }
        None => encode_u8(out, 0),
    }
}

fn value_type_tag(ty: RuntimeValueType) -> u8 {
    match ty {
        RuntimeValueType::Variant => 0,
        RuntimeValueType::Long => 1,
        RuntimeValueType::Integer => 2,
        RuntimeValueType::String => 3,
        RuntimeValueType::Boolean => 4,
        RuntimeValueType::Double => 5,
        RuntimeValueType::Single => 6,
        RuntimeValueType::Currency => 7,
        RuntimeValueType::Date => 8,
        RuntimeValueType::Decimal => 9,
        RuntimeValueType::Object => 10,
        RuntimeValueType::Byte => 11,
        RuntimeValueType::LongLong => 12,
        RuntimeValueType::LongPtr => 13,
        RuntimeValueType::Record => 14,
        RuntimeValueType::Array => 15,
    }
}

fn decode_value_type(tag: u8) -> RuntimeValueType {
    match tag {
        1 => RuntimeValueType::Long,
        2 => RuntimeValueType::Integer,
        3 => RuntimeValueType::String,
        4 => RuntimeValueType::Boolean,
        5 => RuntimeValueType::Double,
        6 => RuntimeValueType::Single,
        7 => RuntimeValueType::Currency,
        8 => RuntimeValueType::Date,
        9 => RuntimeValueType::Decimal,
        10 => RuntimeValueType::Object,
        11 => RuntimeValueType::Byte,
        12 => RuntimeValueType::LongLong,
        13 => RuntimeValueType::LongPtr,
        14 => RuntimeValueType::Record,
        15 => RuntimeValueType::Array,
        _ => RuntimeValueType::Variant,
    }
}

fn invoke_kind_tag(kind: RuntimeMemberInvokeKind) -> u8 {
    match kind {
        RuntimeMemberInvokeKind::Method => 0,
        RuntimeMemberInvokeKind::PropertyGet => 1,
        RuntimeMemberInvokeKind::PropertyLet => 2,
        RuntimeMemberInvokeKind::PropertySet => 3,
    }
}

fn decode_invoke_kind(tag: u8) -> RuntimeMemberInvokeKind {
    match tag {
        1 => RuntimeMemberInvokeKind::PropertyGet,
        2 => RuntimeMemberInvokeKind::PropertyLet,
        3 => RuntimeMemberInvokeKind::PropertySet,
        _ => RuntimeMemberInvokeKind::Method,
    }
}

fn encode_member_spec(out: &mut Vec<u8>, member: &RuntimeMemberSpec) {
    encode_str(out, &member.name);
    encode_opt_i32(out, member.dispatch_id);
    encode_opt_u16(out, member.vtable_slot);
    encode_u8(out, invoke_kind_tag(member.invoke_kind));
    let flags = u8::from(member.is_default_member)
        | (u8::from(member.is_enumerator_member) << 1)
        | (u8::from(member.externally_visible) << 2);
    encode_u8(out, flags);
    encode_u16(out, member.func_flags);
    match member.return_type {
        Some(ty) => {
            encode_u8(out, 1);
            encode_u8(out, value_type_tag(ty));
        }
        None => encode_u8(out, 0),
    }
    match member.backing {
        RuntimeMemberBacking::None => encode_u8(out, 0),
        RuntimeMemberBacking::Proc(proc_index) => {
            encode_u8(out, 1);
            encode_u32(out, proc_index as u32);
        }
        RuntimeMemberBacking::Field(token) => {
            encode_u8(out, 2);
            encode_i32(out, token);
        }
    }
    encode_u32(out, member.params.len() as u32);
    for param in &member.params {
        encode_str(out, &param.name);
        encode_u8(out, value_type_tag(param.value_type));
        let pflags = u8::from(param.by_ref)
            | (u8::from(param.optional) << 1)
            | (u8::from(param.param_array) << 2);
        encode_u8(out, pflags);
    }
}

/// Encode class metadata for AOT (`cvb_class_new` / `cvb_as_new`).
pub fn encode_native_class_metadata(
    program: &OxProgram,
    class_index: usize,
    field_kinds: &[u32],
) -> Vec<u8> {
    let schema = &program.classes[class_index];
    let spec = runtime_class_spec_for_program(program, class_index);
    let mut bytes = Vec::new();
    // v1 backward-compatible prefix:
    for value in [
        class_index as u32,
        schema.name.len() as u32,
        schema.implements.len() as u32,
        schema.fields.len() as u32,
    ] {
        encode_u32(&mut bytes, value);
    }
    bytes.extend_from_slice(schema.name.as_bytes());
    for name in &schema.implements {
        encode_str(&mut bytes, name);
    }
    for (field, &kind) in schema.fields.iter().zip(field_kinds.iter()) {
        encode_i32(&mut bytes, field.token);
        encode_u32(&mut bytes, kind);
    }

    // Extended shared-descriptor trailer (`CVBD` v1):
    bytes.extend_from_slice(&NATIVE_METADATA_MAGIC);
    encode_u32(&mut bytes, NATIVE_METADATA_VERSION);
    encode_str(&mut bytes, &spec.unit_name);
    encode_str(&mut bytes, &spec.name);
    encode_opt_str(&mut bytes, spec.runtime_type_name.as_deref());
    encode_u8(&mut bytes, u8::from(spec.predeclared));
    encode_u32(&mut bytes, spec.fields.len() as u32);
    for field in &spec.fields {
        encode_str(&mut bytes, &field.name);
        encode_i32(&mut bytes, field.token);
        encode_u8(&mut bytes, value_type_tag(field.value_type));
        match field.array_element_type {
            Some(elem) => {
                encode_u8(&mut bytes, 1);
                encode_u8(&mut bytes, value_type_tag(elem));
            }
            None => encode_u8(&mut bytes, 0),
        }
        encode_u8(&mut bytes, u8::from(field.externally_visible));
        encode_opt_i32(&mut bytes, field.dispatch_id);
        encode_opt_u16(&mut bytes, field.get_vtable_slot);
        encode_opt_u16(&mut bytes, field.put_vtable_slot);
        encode_opt_u16(&mut bytes, field.putref_vtable_slot);
    }
    encode_u32(&mut bytes, spec.as_new_fields.len() as u32);
    for field in &spec.as_new_fields {
        encode_i32(&mut bytes, field.field_token);
        match &field.activation {
            RuntimeClassActivationSpec::ProjectClass { class_index } => {
                encode_u8(&mut bytes, 0);
                encode_u32(&mut bytes, *class_index as u32);
            }
            RuntimeClassActivationSpec::ExternClass { import_index } => {
                encode_u8(&mut bytes, 1);
                encode_u32(&mut bytes, *import_index as u32);
            }
            RuntimeClassActivationSpec::ComClass { prog_id } => {
                encode_u8(&mut bytes, 2);
                encode_str(&mut bytes, prog_id);
            }
        }
    }
    encode_u32(&mut bytes, spec.methods.len() as u32);
    for method in &spec.methods {
        encode_member_spec(&mut bytes, method);
    }
    encode_u32(&mut bytes, spec.implemented_interfaces.len() as u32);
    for iface in &spec.implemented_interfaces {
        encode_str(&mut bytes, &iface.name);
        encode_str(&mut bytes, &iface.qualified_name);
        match iface.guid {
            Some(guid) => {
                encode_u8(&mut bytes, 1);
                bytes.extend_from_slice(&guid.to_le_bytes());
            }
            None => encode_u8(&mut bytes, 0),
        }
        encode_u32(&mut bytes, iface.members.len() as u32);
        for member in &iface.members {
            encode_member_spec(&mut bytes, member);
        }
    }
    bytes
}

struct Cursor {
    at: *const u8,
}

impl Cursor {
    unsafe fn u8(&mut self) -> u8 {
        // SAFETY: caller guarantees `self.at` points within the encoded immutable schema.
        unsafe {
            let v = self.at.read();
            self.at = self.at.add(1);
            v
        }
    }

    unsafe fn u16(&mut self) -> u16 {
        // SAFETY: caller guarantees `self.at` points within the encoded immutable schema.
        unsafe {
            let v = self.at.cast::<u16>().read_unaligned();
            self.at = self.at.add(2);
            u16::from_le(v)
        }
    }

    unsafe fn u32(&mut self) -> u32 {
        // SAFETY: caller guarantees `self.at` points within the encoded immutable schema.
        unsafe {
            let v = self.at.cast::<u32>().read_unaligned();
            self.at = self.at.add(4);
            u32::from_le(v)
        }
    }

    unsafe fn i32(&mut self) -> i32 {
        // SAFETY: caller guarantees `self.at` points within the encoded immutable schema.
        unsafe { self.u32() as i32 }
    }

    unsafe fn bytes(&mut self, len: usize) -> Vec<u8> {
        // SAFETY: caller guarantees `len` bytes are readable from `self.at`.
        unsafe {
            let slice = std::slice::from_raw_parts(self.at, len);
            self.at = self.at.add(len);
            slice.to_vec()
        }
    }

    unsafe fn str(&mut self) -> String {
        // SAFETY: strings were encoded as valid UTF-8 with a u32 byte length.
        unsafe {
            let len = self.u32() as usize;
            let bytes = self.bytes(len);
            String::from_utf8(bytes).unwrap_or_default()
        }
    }

    unsafe fn opt_str(&mut self) -> Option<String> {
        // SAFETY: forwarded from caller.
        unsafe { (self.u8() != 0).then(|| self.str()) }
    }

    unsafe fn opt_i32(&mut self) -> Option<i32> {
        // SAFETY: forwarded from caller.
        unsafe { (self.u8() != 0).then(|| self.i32()) }
    }

    unsafe fn opt_u16(&mut self) -> Option<u16> {
        // SAFETY: forwarded from caller.
        unsafe { (self.u8() != 0).then(|| self.u16()) }
    }

    unsafe fn member_spec(&mut self, version: u32) -> RuntimeMemberSpec {
        // SAFETY: forwarded from caller.
        unsafe {
            let name = self.str();
            let dispatch_id = self.opt_i32();
            let vtable_slot = self.opt_u16();
            let invoke_kind = decode_invoke_kind(self.u8());
            let flags = self.u8();
            let func_flags = if version >= NATIVE_METADATA_VERSION {
                self.u16()
            } else {
                0
            };
            let return_type = (self.u8() != 0).then(|| decode_value_type(self.u8()));
            let backing = match self.u8() {
                1 => RuntimeMemberBacking::Proc(self.u32() as usize),
                2 => RuntimeMemberBacking::Field(self.i32()),
                _ => RuntimeMemberBacking::None,
            };
            let param_count = self.u32() as usize;
            let mut params = Vec::with_capacity(param_count);
            for _ in 0..param_count {
                let pname = self.str();
                let value_type = decode_value_type(self.u8());
                let pflags = self.u8();
                params.push(RuntimeParamSpec {
                    name: pname,
                    value_type,
                    by_ref: pflags & 1 != 0,
                    optional: pflags & 2 != 0,
                    param_array: pflags & 4 != 0,
                });
            }
            RuntimeMemberSpec {
                name,
                dispatch_id,
                vtable_slot,
                invoke_kind,
                params,
                return_type,
                is_default_member: flags & 1 != 0,
                is_enumerator_member: flags & 2 != 0,
                externally_visible: flags & 4 != 0,
                func_flags,
                backing,
            }
        }
    }
}

/// Decode compiler-emitted native class metadata (from [`encode_native_class_metadata`]) into the
/// native field `(token, kind)` list and the shared [`RuntimeClassSpec`].
///
/// # Safety
/// `metadata` must point to an immutable byte buffer emitted by [`encode_native_class_metadata`].
pub unsafe fn decode_native_class_metadata(
    metadata: *const u8,
    has_initialize: bool,
    has_terminate: bool,
) -> (Vec<(i32, u32)>, RuntimeClassSpec) {
    // SAFETY: `metadata` was emitted by `encode_native_class_metadata`.
    unsafe {
        let mut cursor = Cursor { at: metadata };
        let class_index = cursor.u32() as usize;
        let name_len = cursor.u32() as usize;
        let implements_count = cursor.u32() as usize;
        let field_count = cursor.u32() as usize;
        let raw_name = String::from_utf8(cursor.bytes(name_len)).unwrap_or_default();
        let mut implements = Vec::with_capacity(implements_count);
        for _ in 0..implements_count {
            implements.push(cursor.str());
        }
        let mut field_kinds = Vec::with_capacity(field_count);
        for _ in 0..field_count {
            let token = cursor.i32();
            let kind = cursor.u32();
            field_kinds.push((token, kind));
        }

        let magic = cursor.bytes(4);
        let version = if magic == NATIVE_METADATA_MAGIC {
            cursor.u32()
        } else {
            0
        };
        if version != NATIVE_METADATA_VERSION_V1 && version != NATIVE_METADATA_VERSION {
            let (unit, display) = raw_name.rsplit_once('.').unwrap_or(("", &raw_name));
            return (
                field_kinds,
                RuntimeClassSpec {
                    unit_name: unit.to_string(),
                    class_index,
                    name: display.to_string(),
                    runtime_type_name: None,
                    default_interface_iid: None,
                    predeclared: false,
                    lifecycle: RuntimeClassLifecycleDescriptor {
                        has_initialize,
                        has_terminate,
                    },
                    fields: Vec::new(),
                    as_new_fields: Vec::new(),
                    implements,
                    methods: Vec::new(),
                    implemented_interfaces: Vec::new(),
                },
            );
        }

        let unit_name = cursor.str();
        let name = cursor.str();
        let runtime_type_name = cursor.opt_str();
        let predeclared = cursor.u8() != 0;

        let spec_field_count = cursor.u32() as usize;
        let mut fields = Vec::with_capacity(spec_field_count);
        for _ in 0..spec_field_count {
            let fname = cursor.str();
            let token = cursor.i32();
            let value_type = decode_value_type(cursor.u8());
            let array_element_type = (cursor.u8() != 0).then(|| decode_value_type(cursor.u8()));
            let externally_visible = cursor.u8() != 0;
            let dispatch_id = cursor.opt_i32();
            let get_vtable_slot = cursor.opt_u16();
            let put_vtable_slot = cursor.opt_u16();
            let putref_vtable_slot = if version >= NATIVE_METADATA_VERSION {
                cursor.opt_u16()
            } else {
                None
            };
            fields.push(RuntimeClassFieldSpec {
                name: fname,
                token,
                value_type,
                array_element_type,
                externally_visible,
                dispatch_id,
                get_vtable_slot,
                put_vtable_slot,
                putref_vtable_slot,
            });
        }

        let as_new_count = cursor.u32() as usize;
        let mut as_new_fields = Vec::with_capacity(as_new_count);
        for _ in 0..as_new_count {
            let field_token = cursor.i32();
            let activation = match cursor.u8() {
                0 => RuntimeClassActivationSpec::ProjectClass {
                    class_index: cursor.u32() as usize,
                },
                1 => RuntimeClassActivationSpec::ExternClass {
                    import_index: cursor.u32() as usize,
                },
                _ => RuntimeClassActivationSpec::ComClass {
                    prog_id: cursor.str(),
                },
            };
            as_new_fields.push(RuntimeClassAsNewFieldSpec {
                field_token,
                activation,
            });
        }

        let method_count = cursor.u32() as usize;
        let mut methods = Vec::with_capacity(method_count);
        for _ in 0..method_count {
            methods.push(cursor.member_spec(version));
        }

        let iface_count = cursor.u32() as usize;
        let mut implemented_interfaces = Vec::with_capacity(iface_count);
        for _ in 0..iface_count {
            let iname = cursor.str();
            let qualified_name = cursor.str();
            let guid = if cursor.u8() != 0 {
                let raw = cursor.bytes(16);
                let mut arr = [0u8; 16];
                arr.copy_from_slice(&raw);
                Some(Guid::from_le_bytes(arr))
            } else {
                None
            };
            let imember_count = cursor.u32() as usize;
            let mut imembers = Vec::with_capacity(imember_count);
            for _ in 0..imember_count {
                imembers.push(cursor.member_spec(version));
            }
            implemented_interfaces.push(RuntimeImplementedInterfaceSpec {
                name: iname,
                qualified_name,
                guid,
                members: imembers,
            });
        }

        (
            field_kinds,
            RuntimeClassSpec {
                unit_name,
                class_index,
                name,
                runtime_type_name,
                default_interface_iid: None,
                predeclared,
                lifecycle: RuntimeClassLifecycleDescriptor {
                    has_initialize,
                    has_terminate,
                },
                fields,
                as_new_fields,
                implements,
                methods,
                implemented_interfaces,
            },
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use calcvba_bundle::ProcedureKind;
    use calcvba_oxir::{
        BlockId, FuncId, LocalId, OxBlock, OxClass, OxClassField, OxFunc, OxLocal, OxParamInfo,
        OxTerminator,
    };

    #[test]
    fn native_class_metadata_roundtrip_builds_identical_descriptor_as_vm3_and_jit() {
        let twice = OxFunc {
            name: "Twice".into(),
            kind: ProcedureKind::Function,
            locals: vec![
                OxLocal {
                    name: "Me".into(),
                    ty: OxTy::Variant,
                    array_element: None,
                    param: Some(OxParamInfo {
                        optional: false,
                        by_ref: false,
                        variadic: false,
                        default: None,
                    }),
                    escaped: false,
                },
                OxLocal {
                    name: "Twice".into(),
                    ty: OxTy::Long,
                    array_element: None,
                    param: None,
                    escaped: false,
                },
            ],
            temps: Vec::new(),
            param_count: 1,
            return_local: Some(LocalId(1)),
            blocks: vec![OxBlock::new(BlockId(0), OxTerminator::Return)],
            entry: BlockId(0),
        };
        let terminate = OxFunc {
            name: "Class_Terminate".into(),
            kind: ProcedureKind::Sub,
            locals: vec![OxLocal {
                name: "Me".into(),
                ty: OxTy::Variant,
                array_element: None,
                param: Some(OxParamInfo {
                    optional: false,
                    by_ref: false,
                    variadic: false,
                    default: None,
                }),
                escaped: false,
            }],
            temps: Vec::new(),
            param_count: 1,
            return_local: None,
            blocks: vec![OxBlock::new(BlockId(0), OxTerminator::Return)],
            entry: BlockId(0),
        };
        let program = OxProgram {
            unit_name: "ProbeProj".into(),
            funcs: vec![twice, terminate],
            classes: vec![OxClass {
                name: "Widget".into(),
                runtime_type_name: None,
                predeclared: false,
                initialize: None,
                terminate: Some(FuncId(1)),
                fields: vec![OxClassField {
                    externally_visible: true,
                    name: "Value".into(),
                    token: 0,
                    ty: OxTy::Long,
                    array_element: None,
                    dispid: Some(0x4003_0000),
                    get_vtable_slot: Some(7),
                    put_vtable_slot: Some(8),
                    putref_vtable_slot: None,
                }],
                methods: vec![OxClassMethod {
                    externally_visible: true,
                    name: "Twice".into(),
                    kind: ProjectMemberKind::Method,
                    proc: FuncId(0),
                    dispid: Some(0x6003_0000),
                    vtable_slot: Some(9),
                    is_default_member: false,
                    is_enumerator_member: false,
                }],
                as_new_fields: Vec::new(),
                implements: Vec::new(),
            }],
            ..OxProgram::empty()
        };

        let vm_jit_descriptors = runtime_class_descriptors_for_program(&program);
        assert_eq!(vm_jit_descriptors.len(), 1);
        let vm_jit = vm_jit_descriptors[0];

        let encoded = encode_native_class_metadata(&program, 0, &[3]);
        // SAFETY: `encoded` was just produced by `encode_native_class_metadata`.
        let (field_kinds, decoded_spec) =
            unsafe { decode_native_class_metadata(encoded.as_ptr(), false, true) };
        assert_eq!(field_kinds, vec![(0, 3)]);
        let native = build_runtime_class_descriptor(&decoded_spec);

        assert!(
            core::ptr::eq(vm_jit, native),
            "VM3, rt-abi/JIT and native template must share the exact same built-once descriptor"
        );
        assert_eq!(native.validate_com_export_shape(), Ok(()));
        assert_eq!(
            native.default_interface_iid,
            Some(Guid::deterministic("ProbeProj", "IWidget"))
        );
        let dispatch = native.dispatch_interface().unwrap();
        assert_eq!(dispatch.members.len(), 3);
        assert_eq!(dispatch.members[0].name, "Value");
        assert_eq!(dispatch.members[0].dispatch_id, 0x4003_0000);
        assert_eq!(dispatch.members[0].vtable_slot, Some(7));
        assert_eq!(dispatch.members[1].name, "Value");
        assert_eq!(dispatch.members[1].dispatch_id, 0x4003_0000);
        assert_eq!(dispatch.members[1].vtable_slot, Some(8));
        assert_eq!(dispatch.members[2].name, "Twice");
        assert_eq!(dispatch.members[2].dispatch_id, 0x6003_0000);
        assert_eq!(dispatch.members[2].vtable_slot, Some(9));
    }
}
