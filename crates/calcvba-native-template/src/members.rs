//! One descriptor-planned native member operation, using the JIT's COM cells.
use calcvba_host_api::{
    TypedContext, TypedFault, Value,
    object::{
        DynamicArgument, DynamicMemberCall, DynamicMemberKind,
        planned_slots::{CallerSyncPhase, MemberArguments},
    },
};
use calcvba_rt_abi::{
    host_entry::value::{from_variant, to_variant},
    native_export::{Context, NativeMemberBinding, NativeMemberCall},
};
use calcvba_runtime::{ObjectRef, Variant};
use std::{collections::HashMap, rc::Rc};

pub(super) type Calls = HashMap<usize, Rc<NativeMemberCall>>;

fn fault(number: i32) -> TypedFault {
    TypedFault {
        number,
        description: calcvba_runtime::default_error_message(number).into(),
        source: None,
        help_file: None,
        help_context: None,
    }
}

/// # Safety
/// Array owns an initialized Variant; kinds/indices span rank native cells. Value is an
/// initialized native cell of kind, writable for a get and borrowed for a store.
#[unsafe(no_mangle)]
#[allow(clippy::too_many_arguments)] // Explicit native cell/vector ABI.
pub unsafe extern "C" fn cvb_variant_index(
    context: *mut Context,
    array: *mut u8,
    kinds: *const u32,
    indices: *const *const u8,
    rank: u32,
    cell: *mut u8,
    kind: u32,
    store: u32,
) -> i32 {
    let result = (|| -> Result<(), TypedFault> {
        // SAFETY: all cells are retained by the active compiled frame; no state
        // borrow spans the host call or owning result conversion.
        unsafe {
            let native = super::values::borrowed(array, 12);
            let tag = native.Anonymous.Anonymous.vt;
            let mut actuals = (0..rank as usize)
                .map(|index| value(context, indices.add(index).read(), kinds.add(index).read()))
                .collect::<Result<Vec<_>, _>>()?;
            if tag & 0x2000 != 0 {
                let indices = actuals
                    .iter()
                    .map(|index| {
                        calcvba_eval::arith::int(index)
                            .map_err(|error| fault(error.code))
                            .and_then(|index| i32::try_from(index).map_err(|_| fault(6)))
                    })
                    .collect::<Result<Vec<_>, _>>()?;
                let pointer = native.Anonymous.Anonymous.Anonymous.parray;
                let element = u32::from(tag & 0x0fff);
                let status = if store == 0 {
                    super::values::cvb_array_get(
                        &pointer,
                        indices.as_ptr(),
                        rank,
                        element,
                        cell,
                        kind,
                    )
                } else {
                    super::values::cvb_array_set(
                        &pointer,
                        indices.as_ptr(),
                        rank,
                        element,
                        cell,
                        kind,
                    )
                };
                return if status == 0 {
                    Ok(())
                } else {
                    Err(fault(status))
                };
            }
            if !matches!(tag, 9 | 13) {
                return Err(fault(13));
            }
            let receiver = object(array, 12)?;
            let member_kind = if store == 0 {
                DynamicMemberKind::Get
            } else {
                let input = value(context, cell, kind)?;
                let member_kind = if input.vtype() == calcvba_runtime::VarType::Object {
                    DynamicMemberKind::Set
                } else {
                    DynamicMemberKind::Let
                };
                actuals.push(input);
                member_kind
            };
            let arguments = (0..actuals.len())
                .map(|alias| DynamicArgument {
                    name: None,
                    by_ref: false,
                    alias,
                })
                .collect();
            let call = DynamicMemberCall {
                name: None,
                kind: member_kind,
                arguments,
            };
            let result = if receiver.native_collection_snapshot().is_some() {
                calcvba_rt_abi::host_entry::collection::invoke(
                    &receiver,
                    "Item",
                    &actuals,
                    &call.arguments,
                )
                .map_err(fault)?
            } else if receiver.is_project_instance() {
                return Err(fault(438));
            } else {
                let host = super::services::host(context);
                let prepared = host
                    .com()
                    .prepare_dynamic_object(receiver)
                    .map_err(|error| TypedFault {
                        number: error.host_error_code.unwrap_or(13),
                        description: error.message,
                        source: None,
                        help_file: None,
                        help_context: None,
                    })?;
                prepared.invoke_dynamic(
                    &call,
                    &mut actuals,
                    &mut *context.cast::<TypedContext>(),
                )?
            };
            if store == 0 {
                super::services::write(context, cell, kind, &result).map_err(fault)?;
            }
            Ok(())
        }
    })();
    match result {
        Ok(()) => 0,
        // SAFETY: only the active owning frame's error state is updated.
        Err(error) => unsafe {
            super::services::typed_fault(context, error)
                .err()
                .unwrap_or(5)
        },
    }
}

/// # Safety
/// Receiver has its declared native kind and out is a writable class index.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_member_route(
    _context: *mut Context,
    cell: *const u8,
    kind: u32,
    out: *mut i32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: the caller retains the receiver through the temporary owner.
        unsafe {
            if kind == 12 && !matches!(cell.cast::<u16>().read(), 9 | 13) {
                out.write(-1);
                return Ok(());
            }
            let reference = object(cell, kind).map_err(|fault| fault.number)?;
            out.write(if reference.is_project_instance() {
                reference.route_key()
            } else {
                -1
            });
        }
        Ok(())
    })
}

fn snapshot(value: &Value) -> Result<Variant, TypedFault> {
    let Some(slot) = value.borrow_slot() else {
        return Ok(Variant::empty());
    };
    // SAFETY: the Value retains this exact carrier through the independent copy.
    unsafe { calcvba_host_api::value::retain_bits(slot.ty(), slot.bits()) }
        .map(to_variant)
        .map_err(fault)
}

unsafe fn object(cell: *const u8, kind: u32) -> Result<ObjectRef, TypedFault> {
    // SAFETY: the generated caller retains an exact native cell. QI/retention
    // needs only IUnknown, including for interfaces that have no IDispatch.
    unsafe {
        let native = super::values::borrowed(cell, kind);
        let tag = native.Anonymous.Anonymous.vt;
        if !matches!(tag, 9 | 13) {
            return Err(fault(424));
        }
        ObjectRef::from_raw_iunknown_addref(native.Anonymous.Anonymous.Anonymous.punkVal.cast())
            .ok_or_else(|| fault(91))
    }
}

unsafe fn value(context: *mut Context, cell: *const u8, kind: u32) -> Result<Variant, TypedFault> {
    // SAFETY: only object cells bypass Automation decoding, to preserve a raw
    // IUnknown-only interface until its imported IID has selected the view.
    unsafe {
        let native = super::values::borrowed(cell, kind);
        if matches!(native.Anonymous.Anonymous.vt, 9 | 13) {
            let pointer = native.Anonymous.Anonymous.Anonymous.punkVal;
            return Ok(ObjectRef::from_raw_iunknown_addref(pointer.cast())
                .map_or_else(Variant::nothing, Variant::from_object_ref));
        }
        super::services::read(context, cell, kind).map_err(fault)
    }
}

fn prepared_value(
    host: &dyn calcvba_hal::HostServices,
    value: Variant,
    interface: Option<[u8; 16]>,
    interfaces: &[calcvba_com::TypeLibInterfaceMetadata],
) -> Result<Variant, TypedFault> {
    if let Some(object) = value.as_object_ref() {
        if object.is_project_instance() || object.native_collection_snapshot().is_some() {
            return Err(fault(13));
        }
        if calcvba_host_api::HostObject::from_object_ref(&object).is_some() {
            return Ok(value);
        }
        let nominal = interface.and_then(|iid| {
            interfaces.iter().find(|meta| {
                meta.iid.is_some_and(|wanted| {
                    let mut bytes = [0; 16];
                    bytes[..4].copy_from_slice(&wanted.data1.to_le_bytes());
                    bytes[4..6].copy_from_slice(&wanted.data2.to_le_bytes());
                    bytes[6..8].copy_from_slice(&wanted.data3.to_le_bytes());
                    bytes[8..].copy_from_slice(&wanted.data4);
                    bytes == iid
                })
            })
        });
        let prepared = if let Some(nominal) = nominal {
            let members: Vec<_> = nominal
                .members
                .iter()
                .filter(|m| calcvba_com::typed_slots::plan::validate(m))
                .cloned()
                .collect();
            // SAFETY: unchanged imported interface and native ABI descriptors.
            unsafe {
                host.com()
                    .prepare_typed_object_with_interfaces(object, &members, interfaces)
            }
        } else {
            host.com().prepare_dynamic_object(object)
        }
        .map_err(|error| TypedFault {
            number: error.host_error_code.unwrap_or(13),
            description: error.message,
            source: None,
            help_file: None,
            help_context: None,
        })?;
        return Ok(Variant::from_object_ref(prepared.object_ref()));
    }
    if let Some(array) = value.as_safearray() {
        let bounds = array.bounds().ok_or_else(|| fault(13))?;
        let elements = array
            .variant_elements()
            .ok_or_else(|| fault(13))?
            .into_iter()
            .map(|v| prepared_value(host, v, None, interfaces))
            .collect::<Result<Vec<_>, _>>()?;
        return calcvba_runtime::safe_array::SafeArray::from_typed_variants_nd(
            bounds,
            value.array_element_vartype().unwrap_or(12),
            elements,
        )
        .map(Variant::from_safearray)
        .map_err(|_| fault(13));
    }
    Ok(value)
}

unsafe fn invoke(
    context: *mut Context,
    metadata: *const u8,
    receiver: *const u8,
    arguments: *const *mut u8,
    out: *mut u8,
) -> Result<(), TypedFault> {
    // SAFETY: compiler metadata and exact cells are retained throughout the call.
    unsafe {
        let runtime = super::services::runtime(context);
        let call = if let Some(call) = (*runtime).members.get(&metadata.addr()) {
            call.clone()
        } else {
            let len = metadata.cast::<u32>().read_unaligned() as usize;
            let call: NativeMemberCall =
                serde_json::from_slice(std::slice::from_raw_parts(metadata.add(4), len))
                    .map_err(|_| fault(5))?;
            let call = Rc::new(call);
            (*runtime).members.insert(metadata.addr(), call.clone());
            call
        };
        let receiver = object(receiver, call.receiver_kind)?;
        let host = super::services::host(context);
        let mut values = Vec::with_capacity(call.arguments.len());
        let mut specifications = Vec::with_capacity(call.arguments.len());
        let mut addresses = Vec::with_capacity(call.arguments.len());
        for (index, argument) in call.arguments.iter().enumerate() {
            let cell = arguments.add(index).read();
            let alias = if argument.by_ref {
                addresses
                    .iter()
                    .position(|old| *old == Some(cell))
                    .unwrap_or(index)
            } else {
                index
            };
            addresses.push(argument.by_ref.then_some(cell));
            specifications.push(DynamicArgument {
                name: argument.name.clone(),
                by_ref: argument.by_ref,
                alias,
            });
            values.push(if argument.omitted {
                Variant::from_error_code(0x80020004u32 as i32)
            } else {
                value(context, cell, argument.kind)?
            });
        }
        let typed_context = &mut *context.cast::<TypedContext>();
        let result = match &call.binding {
            NativeMemberBinding::Early { member, interfaces } => {
                let (parameters, _) =
                    calcvba_com::typed_slots::plan::visible_signature(member, values.len())
                        .ok_or_else(|| fault(450))?;
                // SAFETY: the compiler retained the exact selected member and
                // complete imported interfaces. This never retries IDispatch.
                let prepared = host
                    .com()
                    .prepare_typed_object_with_interfaces(
                        receiver,
                        std::slice::from_ref(member.as_ref()),
                        interfaces,
                    )
                    .map_err(|error| TypedFault {
                        number: error.host_error_code.unwrap_or(13),
                        description: error.message,
                        source: None,
                        help_file: None,
                        help_context: None,
                    })?;
                let indices = calcvba_com::typed_slots::plan::visible_indices(member);
                let mut typed = Vec::new();
                for (index, parameter) in parameters.iter().take(values.len()).enumerate() {
                    let input = if call.arguments[index].omitted {
                        calcvba_com::typed_slots::plan::default_value(member, indices[index])
                            .ok_or_else(|| fault(450))?
                    } else {
                        std::mem::replace(&mut values[index], Variant::empty())
                    };
                    let input =
                        prepared_value(host.as_ref(), input, parameter.interface(), interfaces)?;
                    typed.push(from_variant(parameter.carrier(), input).map_err(fault)?);
                }
                let aliases: Vec<_> = specifications.iter().map(|s| s.alias).collect();
                let mut synchronize = |phase, inputs: &mut [Value]| -> Result<(), i32> {
                    for (index, argument) in call.arguments.iter().enumerate() {
                        if !argument.by_ref
                            || aliases[index] != index
                            || !calcvba_com::typed_slots::plan::is_by_ref(member, indices[index])
                        {
                            continue;
                        }
                        if !matches!(
                            parameters[index].carrier(),
                            calcvba_host_api::ValueType::Long
                                | calcvba_host_api::ValueType::Double
                                | calcvba_host_api::ValueType::Boolean
                        ) {
                            continue;
                        }
                        let cell = arguments.add(index).read();
                        match phase {
                            CallerSyncPhase::WriteCaller => super::services::write(
                                context,
                                cell,
                                argument.kind,
                                &snapshot(&inputs[index]).map_err(|fault| fault.number)?,
                            )?,
                            CallerSyncPhase::ReadCaller => {
                                let input = value(context, cell, argument.kind)
                                    .map_err(|fault| fault.number)?;
                                inputs[index] = from_variant(parameters[index].carrier(), input)?;
                            }
                        }
                    }
                    Ok(())
                };
                let mut inputs = MemberArguments::new(&mut typed, &aliases)
                    .map_err(fault)?
                    .with_caller_sync(&mut synchronize);
                let result = calcvba_com::typed_slots::invoke_prepared_native_member(
                    &prepared,
                    member,
                    &mut inputs,
                    typed_context,
                );
                let mut copyback = None;
                for (index, argument) in call.arguments.iter().enumerate() {
                    if !argument.by_ref
                        || aliases[index] != index
                        || !calcvba_com::typed_slots::plan::is_by_ref(member, indices[index])
                    {
                        continue;
                    }
                    let input = snapshot(inputs.get(index).ok_or_else(|| fault(5))?)?;
                    if let Err(number) = super::services::write(
                        context,
                        arguments.add(index).read(),
                        argument.kind,
                        &input,
                    ) {
                        copyback.get_or_insert(number);
                    }
                }
                let result = result?;
                if let Some(number) = copyback {
                    return Err(fault(number));
                }
                to_variant(result)
            }
            binding => {
                let (name, kind, base) = match binding {
                    NativeMemberBinding::Late { name, kind } => (
                        name.clone(),
                        match kind {
                            calcvba_com::TypeLibMemberInvokeKind::Method => {
                                DynamicMemberKind::Method
                            }
                            calcvba_com::TypeLibMemberInvokeKind::PropertyGet => {
                                DynamicMemberKind::Get
                            }
                            calcvba_com::TypeLibMemberInvokeKind::PropertyPut => {
                                DynamicMemberKind::Let
                            }
                            calcvba_com::TypeLibMemberInvokeKind::PropertyPutRef => {
                                DynamicMemberKind::Set
                            }
                        },
                        0,
                    ),
                    NativeMemberBinding::CallByName => {
                        let name = calcvba_eval::arith::coerce_string(
                            values.first().ok_or_else(|| fault(5))?,
                        )
                        .map_err(|error| fault(error.code))?
                        .as_bstr()
                        .ok_or_else(|| fault(13))?
                        .to_string();
                        let kind = calcvba_eval::arith::coerce_numeric(
                            values.get(1).ok_or_else(|| fault(5))?,
                            calcvba_bundle::NumericCoerceTarget::Long,
                        )
                        .map_err(|error| fault(error.code))?
                        .as_i32()
                        .ok_or_else(|| fault(13))?;
                        let kind = match kind {
                            1 => DynamicMemberKind::Method,
                            2 => DynamicMemberKind::Get,
                            4 => DynamicMemberKind::Let,
                            8 => DynamicMemberKind::Set,
                            _ => return Err(fault(5)),
                        };
                        (Some(name), kind, 2)
                    }
                    _ => unreachable!(),
                };
                let mut specs = specifications.split_off(base);
                for specification in &mut specs {
                    specification.alias = specification.alias.saturating_sub(base);
                }
                let dynamic = DynamicMemberCall {
                    name,
                    kind,
                    arguments: specs,
                };
                let actuals = &mut values[base..];
                let result = if receiver.native_collection_snapshot().is_some() {
                    if matches!(binding, NativeMemberBinding::CallByName)
                        && !matches!(dynamic.kind, DynamicMemberKind::Method)
                    {
                        return Err(fault(438));
                    }
                    calcvba_rt_abi::host_entry::collection::invoke(
                        &receiver,
                        dynamic.name.as_deref().unwrap_or("Item"),
                        actuals,
                        &dynamic.arguments,
                    )
                    .map_err(fault)
                } else if receiver.is_project_instance() {
                    Err(fault(438))
                } else {
                    let prepared =
                        host.com()
                            .prepare_dynamic_object(receiver)
                            .map_err(|error| TypedFault {
                                number: error.host_error_code.unwrap_or(13),
                                description: error.message,
                                source: None,
                                help_file: None,
                                help_context: None,
                            })?;
                    for input in actuals.iter_mut() {
                        *input = prepared_value(
                            host.as_ref(),
                            std::mem::replace(input, Variant::empty()),
                            None,
                            &[],
                        )?;
                    }
                    prepared.invoke_dynamic(&dynamic, actuals, typed_context)
                };
                let mut copyback = None;
                for (index, argument) in call.arguments[base..].iter().enumerate() {
                    if argument.by_ref
                        && dynamic.arguments[index].alias == index
                        && let Err(number) = super::services::write(
                            context,
                            arguments.add(base + index).read(),
                            argument.kind,
                            &actuals[index],
                        )
                    {
                        copyback.get_or_insert(number);
                    }
                }
                let result = result?;
                if let Some(number) = copyback {
                    return Err(fault(number));
                }
                result
            }
        };
        super::services::write(context, out, call.result_kind, &result).map_err(fault)
    }
}

/// # Safety
/// Descriptor is retained compiler metadata; receiver, arguments and output
/// use its exact native kinds. Execution stays on the DLL's owning thread.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_member(
    context: *mut Context,
    descriptor: *const u8,
    receiver: *const u8,
    arguments: *const *mut u8,
    out: *mut u8,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: forward the generated frame/schema lease; publish rich faults
        // only after all temporary member argument owners have been released.
        unsafe {
            invoke(context, descriptor, receiver, arguments, out)
                .or_else(|fault| super::services::typed_fault(context, fault))
        }
    })
}
