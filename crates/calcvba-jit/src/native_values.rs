use super::*;

pub(super) const HELPERS: &[(&str, &[ir::Type])] = &[
    ("cvb_member", &[types::I64; 5]),
    (
        "cvb_member_route",
        &[types::I64, types::I64, types::I32, types::I64],
    ),
    ("cvb_collection_ensure", &[types::I64]),
    ("cvb_com_ensure", &[types::I64; 4]),
    ("cvb_class_new", &[types::I64; 5]),
    ("cvb_as_new", &[types::I64; 5]),
    ("cvb_class_route", &[types::I64; 2]),
    (
        "cvb_class_field",
        &[
            types::I64,
            types::I64,
            types::I32,
            types::I64,
            types::I32,
            types::I32,
        ],
    ),
    ("cvb_drain_terminations", &[types::I64]),
    ("cvb_collection_new", &[types::I64]),
    (
        "cvb_object_type",
        &[
            types::I64,
            types::I64,
            types::I32,
            types::I64,
            types::I64,
            types::I32,
            types::I64,
        ],
    ),
    ("cvb_array_erase", &[types::I64, types::I64, types::I32]),
    ("cvb_frame_enter", &[types::I64]),
    ("cvb_frame_exit", &[types::I64]),
    (
        "cvb_gosub",
        &[types::I64, types::I32, types::I32, types::I64],
    ),
    ("cvb_error_control", &[types::I64, types::I32, types::I32]),
    (
        "cvb_error_get",
        &[types::I64, types::I64, types::I32, types::I32],
    ),
    (
        "cvb_error_set",
        &[types::I64, types::I32, types::I64, types::I32],
    ),
    (
        "cvb_fault_route",
        &[types::I64, types::I32, types::I32, types::I64],
    ),
    (
        "cvb_resume",
        &[types::I64, types::I32, types::I32, types::I64],
    ),
    ("cvb_array_literal", &[types::I64; 3]),
    (
        "cvb_variant_index",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I64,
            types::I32,
            types::I32,
        ],
    ),
    (
        "cvb_variant_array_bound",
        &[types::I64, types::I32, types::I32, types::I64],
    ),
    ("cvb_paramarray_literal", &[types::I64; 5]),
    (
        "cvb_paramarray_set",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I64,
            types::I32,
        ],
    ),
    (
        "cvb_iter_init",
        &[types::I64, types::I64, types::I64, types::I32],
    ),
    (
        "cvb_iter_next",
        &[types::I64, types::I64, types::I64, types::I32, types::I64],
    ),
    ("cvb_iter_clear", &[types::I64; 2]),
    (
        "cvb_value_unary",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
    (
        "cvb_library",
        &[
            types::I64,
            types::I32,
            types::I32,
            types::I64,
            types::I64,
            types::I64,
        ],
    ),
    (
        "cvb_value_binary",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
    ("cvb_services_clear", &[types::I64]),
    ("cvb_value_type", &[types::I64, types::I64, types::I32]),
    (
        "cvb_fixed_string_coerce",
        &[types::I64, types::I64, types::I32, types::I32],
    ),
    (
        "cvb_declare_result",
        &[types::I64, types::I64, types::I32, types::I32, types::I64],
    ),
    (
        "cvb_record_array_get",
        &[types::I64, types::I64, types::I32, types::I64, types::I64],
    ),
    (
        "cvb_record_array_set",
        &[types::I64, types::I64, types::I32, types::I64, types::I64],
    ),
    (
        "cvb_record_array_redim",
        &[
            types::I64,
            types::I32,
            types::I64,
            types::I32,
            types::I32,
            types::I64,
        ],
    ),
    (
        "cvb_fixed_string_get",
        &[types::I64, types::I64, types::I32, types::I32],
    ),
    (
        "cvb_fixed_string_set",
        &[types::I64, types::I32, types::I64, types::I32],
    ),
    ("cvb_object_is", &[types::I64; 3]),
    ("cvb_get_error_help", &[types::I64; 3]),
    (
        "cvb_raise_dynamic",
        &[
            types::I64,
            types::I32,
            types::I64,
            types::I32,
            types::I64,
            types::I32,
            types::I64,
            types::I32,
            types::I64,
            types::I32,
            types::I32,
        ],
    ),
    ("cvb_declare", &[types::I64; 4]),
    ("cvb_declare_clear", &[types::I64]),
    ("cvb_value_truthy", &[types::I64, types::I64, types::I32]),
    (
        "cvb_raise",
        &[
            types::I64,
            types::I32,
            types::I64,
            types::I64,
            types::I64,
            types::I64,
            types::I32,
        ],
    ),
    ("cvb_last_dll_error", &[]),
    ("cvb_record_copy", &[types::I64; 3]),
    ("cvb_record_clear", &[types::I64; 2]),
    (
        "cvb_array_get",
        &[
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I64,
            types::I32,
        ],
    ),
    (
        "cvb_array_set",
        &[
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I64,
            types::I32,
        ],
    ),
    (
        "cvb_array_bound",
        &[types::I64, types::I32, types::I32, types::I64],
    ),
    (
        "cvb_array_redim_encoded",
        &[
            types::I64,
            types::I32,
            types::I32,
            types::I64,
            types::I32,
            types::I32,
            types::I64,
        ],
    ),
    (
        "cvb_value_copy",
        &[types::I64, types::I64, types::I32, types::I32],
    ),
    ("cvb_value_clear", &[types::I64, types::I32]),
    ("cvb_string_concat", &[types::I64; 3]),
    ("cvb_string_import", &[types::I64, types::I32, types::I64]),
    ("cvb_string_len", &[types::I64; 2]),
    (
        "cvb_value_add",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
    (
        "cvb_value_sub",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
    (
        "cvb_value_mul",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
    (
        "cvb_value_div",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
    (
        "cvb_value_concat",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
    (
        "cvb_value_compare",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
    (
        "cvb_value_neg",
        &[types::I64, types::I64, types::I32, types::I32],
    ),
    (
        "cvb_value_not",
        &[types::I64, types::I64, types::I32, types::I32],
    ),
    (
        "cvb_value_and",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
    (
        "cvb_value_or",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
    (
        "cvb_value_xor",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
    (
        "cvb_value_eqv",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
    (
        "cvb_value_imp",
        &[
            types::I64,
            types::I64,
            types::I64,
            types::I32,
            types::I32,
            types::I32,
        ],
    ),
];

pub(super) fn kind(ty: ScalarType) -> u32 {
    match ty {
        ScalarType::Byte => 17,
        ScalarType::Integer => 2,
        ScalarType::LongLong => 20,
        ScalarType::Single => 4,
        ScalarType::Currency => 6,
        ScalarType::Date => 7,
        ScalarType::Long => 3,
        ScalarType::Double => 5,
        ScalarType::Boolean => 11,
        ScalarType::String => 8,
        ScalarType::Variant => 12,
        ScalarType::HostObject | ScalarType::Interface(_) => 9,
        ScalarType::Array => 0x200c,
        ScalarType::Record => 36,
        ScalarType::Void => 0,
    }
}

pub(super) fn validate_library_operand(
    layout: &ScalarLayout,
    operand: &OxOperand,
) -> Result<(), &'static str> {
    if layout.operand_type(operand) == Ok(ScalarType::Record)
        || matches!(operand, OxOperand::Use(place) if layout.slot(*place).is_ok_and(|slot| layout.array_record_slots.contains_key(&slot)))
    {
        return Err("native record value requires a retained Automation record descriptor");
    }
    Ok(())
}
pub(super) fn owning(ty: ScalarType) -> bool {
    !numeric(ty) && ty != ScalarType::Void
}

pub(super) struct Values<'a> {
    pub lower: ScalarLower<'a>,
    pub functions: &'a HashMap<&'static str, ClifFuncId>,
    pub program: &'a OxProgram,
    pub iterator_slots: BTreeSet<usize>,
    pub class_entries: &'a HashMap<FuncId, ClifFuncId>,
    pub as_new: HashMap<OxPlace, calcvba_oxir::inst::OxAsNew>,
}
impl Values<'_> {
    pub fn member_call(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        inst: &OxInst,
    ) {
        use calcvba_rt_abi::native_export::{
            NativeMemberArgument, NativeMemberBinding, NativeMemberCall,
        };
        let (dst, receiver, binding, arguments) = match inst {
            OxInst::ComCallEarly {
                dst,
                recv,
                method,
                args,
                ..
            } => (
                *dst,
                recv,
                NativeMemberBinding::Early {
                    member: Box::new(
                        self.program
                            .com_method(*method)
                            .expect("admitted COM descriptor")
                            .clone(),
                    ),
                    interfaces: self
                        .program
                        .com_interfaces
                        .iter()
                        .filter_map(|interface| {
                            if let calcvba_oxir::ComInterface::Com(metadata) = interface {
                                Some(metadata.clone())
                            } else {
                                None
                            }
                        })
                        .collect(),
                },
                args.clone(),
            ),
            OxInst::ComCallLate {
                dst,
                recv,
                name,
                default_member,
                invoke_kind,
                args,
            } => (
                *dst,
                recv,
                NativeMemberBinding::Late {
                    name: (!default_member).then(|| name.clone()),
                    kind: *invoke_kind,
                },
                args.clone(),
            ),
            OxInst::CallByName {
                dst,
                object,
                name,
                calltype,
                args,
            } => {
                let mut arguments = vec![
                    OxCallArg::Operand(name.clone()),
                    OxCallArg::Operand(calltype.clone()),
                ];
                arguments.extend_from_slice(args);
                (*dst, object, NativeMemberBinding::CallByName, arguments)
            }
            _ => unreachable!("native member instruction"),
        };
        let receiver_cell = self.cell(module, b, receiver);
        let pointers = b.create_sized_stack_slot(StackSlotData::new(
            StackSlotKind::ExplicitSlot,
            (arguments.len().max(1) * 8) as u32,
            3,
        ));
        let mut specifications = Vec::new();
        for (index, argument) in arguments.iter().enumerate() {
            let (operand, by_ref, omitted, name) = match argument {
                OxCallArg::Operand(operand) => (operand.clone(), false, false, None),
                OxCallArg::ByRef(place) => (OxOperand::Use(*place), true, false, None),
                OxCallArg::Named { name, value } => {
                    (value.clone(), false, false, Some(name.clone()))
                }
                OxCallArg::Const(value) => {
                    (OxOperand::Const(OxConst::I32(*value)), false, false, None)
                }
                OxCallArg::Omitted => (OxOperand::Const(OxConst::Empty), false, true, None),
            };
            let cell = self.cell(module, b, &operand);
            b.ins().stack_store(cell, pointers, (index * 8) as i32);
            specifications.push(NativeMemberArgument {
                kind: kind_of_operand(self.lower.layout, &operand),
                by_ref,
                omitted,
                name,
            });
        }
        let result_kind = dst
            .map(|dst| self.slot_kind(self.lower.layout.slot(dst).unwrap()))
            .unwrap_or(0);
        let call = NativeMemberCall {
            receiver_kind: kind_of_operand(self.lower.layout, receiver),
            result_kind,
            binding,
            arguments: specifications,
        };
        let payload = serde_json::to_vec(&call).expect("native COM metadata serialization");
        let mut bytes = (payload.len() as u32).to_le_bytes().to_vec();
        bytes.extend(payload);
        let metadata = data(module, b, bytes, 4);
        let args = b.ins().stack_addr(types::I64, pointers, 0);
        let out = if let Some(dst) = dst {
            self.address(b, self.lower.layout.slot(dst).unwrap())
        } else {
            b.ins().iconst(types::I64, 0)
        };
        self.call(
            module,
            b,
            "cvb_member",
            &[self.lower.runtime, metadata, receiver_cell, args, out],
        );
    }
    fn object_type(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        operand: &OxOperand,
        name: &str,
        assignment: bool,
        dst: Option<OxPlace>,
    ) {
        let cell = self.cell(module, b, operand);
        let kind = b.ins().iconst(
            types::I32,
            i64::from(kind_of_operand(self.lower.layout, operand)),
        );
        let mut schema = name.as_bytes().to_vec();
        let iid = self
            .lower
            .layout
            .assignment_iids
            .get(name)
            .copied()
            .flatten()
            .or_else(|| self.program.named_interface_iid(name))
            .unwrap_or([0; 16]);
        schema.extend_from_slice(&iid);
        let text = data(module, b, schema, 1);
        let len = b.ins().iconst(types::I64, name.len() as i64);
        let mode = b.ins().iconst(types::I32, i64::from(assignment));
        let out = if let Some(dst) = dst {
            self.address(b, self.lower.layout.slot(dst).unwrap())
        } else {
            b.ins().iconst(types::I64, 0)
        };
        self.call(
            module,
            b,
            "cvb_object_type",
            &[self.lower.runtime, cell, kind, text, len, mode, out],
        );
    }
    pub fn class_new(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        class: calcvba_oxir::ClassId,
        out: Value,
        ensure: bool,
    ) {
        let schema = &self.program.classes[class.0];
        let mut bytes = Vec::new();
        for value in [
            class.0 as u32,
            schema.name.len() as u32,
            schema.implements.len() as u32,
            schema.fields.len() as u32,
        ] {
            bytes.extend_from_slice(&value.to_le_bytes());
        }
        bytes.extend_from_slice(schema.name.as_bytes());
        for name in &schema.implements {
            bytes.extend_from_slice(&(name.len() as u32).to_le_bytes());
            bytes.extend_from_slice(name.as_bytes());
        }
        for field in &schema.fields {
            bytes.extend_from_slice(&field.token.to_le_bytes());
            let kind = if let Some(element) = &field.array_element {
                u32::from(calcvba_bundle::safearray_vartype_for_element(element)) | 0x2000
            } else {
                kind(
                    calcvba_rt_abi::host_entry::native_scalar_type(&field.ty)
                        .expect("admitted class field"),
                )
            };
            bytes.extend_from_slice(&kind.to_le_bytes());
        }
        let metadata = data(module, b, bytes, 4);
        let mut entry = |procedure: Option<FuncId>| {
            if let Some(procedure) = procedure {
                let function = module.declare_func_in_func(self.class_entries[&procedure], b.func);
                b.ins().func_addr(types::I64, function)
            } else {
                b.ins().iconst(types::I64, 0)
            }
        };
        let initialize = entry(schema.initialize);
        let terminate = entry(schema.terminate);
        self.call(
            module,
            b,
            if ensure {
                "cvb_as_new"
            } else {
                "cvb_class_new"
            },
            &[self.lower.runtime, metadata, initialize, terminate, out],
        );
    }
    pub fn library_call(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        native: u32,
        string_alias: bool,
        args: &[OxArg],
        dst: Option<OxPlace>,
    ) {
        let inputs = crate::library::lower_arguments(b, args, |b, operand| {
            Ok((
                self.cell(module, b, operand),
                kind_of_operand(self.lower.layout, operand),
            ))
        })
        .expect("admitted library arguments");
        let (out, to) = if let Some(dst) = dst {
            let index = self.lower.layout.slot(dst).unwrap();
            (self.address(b, index), self.slot_kind(index))
        } else {
            (b.ins().iconst(types::I64, 0), 0)
        };
        let mut descriptor = vec![inputs.len() as u32, to];
        let pointers = b.create_sized_stack_slot(StackSlotData::new(
            StackSlotKind::ExplicitSlot,
            (inputs.len().max(1) * 8) as u32,
            3,
        ));
        for (index, (cell, kind)) in inputs.into_iter().enumerate() {
            b.ins().stack_store(cell, pointers, (index * 8) as i32);
            descriptor.push(kind);
        }
        let metadata = data(
            module,
            b,
            descriptor.into_iter().flat_map(u32::to_le_bytes).collect(),
            4,
        );
        let arguments = b.ins().stack_addr(types::I64, pointers, 0);
        let native = b.ins().iconst(types::I32, i64::from(native));
        let string_alias = b.ins().iconst(types::I32, i64::from(string_alias));
        self.call(
            module,
            b,
            "cvb_library",
            &[
                self.lower.runtime,
                native,
                string_alias,
                metadata,
                arguments,
                out,
            ],
        );
    }
    pub fn address(&self, b: &mut FunctionBuilder<'_>, index: usize) -> Value {
        let PhysicalSlots::Native(pointers) = self.lower.slots else {
            unreachable!()
        };
        b.ins().stack_load(types::I64, pointers, (index * 8) as i32)
    }
    pub fn call(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        name: &str,
        args: &[Value],
    ) {
        let helper = module.declare_func_in_func(self.functions[name], b.func);
        let call = b.ins().call(helper, args);
        let status = b.inst_results(call)[0];
        self.lower.check_status(b, status);
    }
    pub fn clear(&self, module: &mut impl Module, b: &mut FunctionBuilder<'_>, index: usize) {
        let address = self.address(b, index);
        if let Some(layout) = self.lower.layout.record_slots.get(&index) {
            let descriptor = data(
                module,
                b,
                record_descriptor(layout).expect("admitted record layout"),
                4,
            );
            self.call(module, b, "cvb_record_clear", &[address, descriptor]);
            return;
        }
        let kind = b.ins().iconst(types::I32, i64::from(self.slot_kind(index)));
        self.call(module, b, "cvb_value_clear", &[address, kind]);
    }
    pub fn slot_kind(&self, index: usize) -> u32 {
        if self.lower.layout.slots[index] == ScalarType::Array {
            0x2000 | u32::from(self.lower.layout.array_slots[&index])
        } else {
            kind(self.lower.layout.slots[index])
        }
    }
    pub fn cell(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        operand: &OxOperand,
    ) -> Value {
        if let OxOperand::Use(place) = operand {
            let cell = self.address(b, self.lower.layout.slot(*place).unwrap());
            if let Some(calcvba_oxir::inst::OxAsNew::ProjectClass { class }) =
                self.as_new.get(place)
            {
                self.class_new(module, b, *class, cell, true);
            } else if let Some(calcvba_oxir::inst::OxAsNew::ExternClass { import }) =
                self.as_new.get(place)
            {
                assert!(self.lower.layout.library_news.contains(import));
                self.call(module, b, "cvb_collection_ensure", &[cell]);
            } else if let Some(calcvba_oxir::inst::OxAsNew::ComClass { prog_id }) =
                self.as_new.get(place)
            {
                let name = data(module, b, prog_id.as_bytes().to_vec(), 1);
                let len = b.ins().iconst(types::I64, prog_id.len() as i64);
                self.call(
                    module,
                    b,
                    "cvb_com_ensure",
                    &[self.lower.runtime, name, len, cell],
                );
            }
            return cell;
        }
        let ty = self.lower.layout.operand_type(operand).unwrap();
        if matches!(operand, OxOperand::Const(OxConst::Empty | OxConst::Null)) {
            let mut bytes = vec![0; 24];
            bytes[0] = u8::from(matches!(operand, OxOperand::Const(OxConst::Null)));
            return data(module, b, bytes, 8);
        }
        let slot = b.create_sized_stack_slot(StackSlotData::new(StackSlotKind::ExplicitSlot, 8, 3));
        let empty = b.ins().iconst(types::I64, 0);
        b.ins().stack_store(empty, slot, 0);
        let value = if let OxOperand::Const(OxConst::Str(text)) = operand {
            let units: Vec<_> = text.encode_utf16().collect();
            let mut bytes = ((units.len() * 2) as u32).to_le_bytes().to_vec();
            for unit in units {
                bytes.extend_from_slice(&unit.to_le_bytes());
            }
            bytes.extend_from_slice(&[0, 0]);
            let address = data(module, b, bytes, 8);
            b.ins().iadd_imm(address, 4)
        } else {
            self.lower.operand(b, operand, ty)
        };
        let value = if ty == ScalarType::Boolean {
            b.ins().ireduce(types::I16, value)
        } else {
            value
        };
        b.ins().stack_store(value, slot, 0);
        b.ins().stack_addr(types::I64, slot, 0)
    }
    pub fn operand(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        operand: &OxOperand,
        to: ScalarType,
    ) -> Value {
        let from = self.lower.layout.operand_type(operand).unwrap();
        if from != to && numeric(from) && numeric(to) {
            let slot =
                b.create_sized_stack_slot(StackSlotData::new(StackSlotKind::ExplicitSlot, 8, 3));
            let zero = b.ins().iconst(types::I64, 0);
            b.ins().stack_store(zero, slot, 0);
            let output = b.ins().stack_addr(types::I64, slot, 0);
            let source = self.cell(module, b, operand);
            self.copy(module, b, output, kind(to), source, kind(from));
            return match to {
                ScalarType::Byte => {
                    let value = b.ins().stack_load(types::I8, slot, 0);
                    b.ins().uextend(types::I32, value)
                }
                ScalarType::Integer | ScalarType::Boolean => {
                    let value = b.ins().stack_load(types::I16, slot, 0);
                    b.ins().sextend(types::I32, value)
                }
                _ => b.ins().stack_load(machine_type(to), slot, 0),
            };
        }
        if matches!(operand, OxOperand::Use(place) if self.as_new.contains_key(place))
            || matches!(operand, OxOperand::Const(OxConst::Str(_)))
        {
            let cell = self.cell(module, b, operand);
            b.ins().load(types::I64, ir::MemFlagsData::new(), cell, 0)
        } else if matches!(operand, OxOperand::Const(OxConst::Empty | OxConst::Null)) {
            self.cell(module, b, operand)
        } else {
            self.lower.operand(b, operand, to)
        }
    }
    pub fn copy(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        dst: Value,
        to: u32,
        src: Value,
        from: u32,
    ) {
        let to = b.ins().iconst(types::I32, i64::from(to));
        let from = b.ins().iconst(types::I32, i64::from(from));
        self.call(module, b, "cvb_value_copy", &[dst, src, to, from]);
    }
    pub fn copy_record(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        index: usize,
        dst: Value,
        src: Value,
    ) {
        let layout = &self.lower.layout.record_slots[&index];
        self.copy_record_layout(module, b, layout, dst, src);
    }
    pub fn copy_record_layout(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        layout: &calcvba_runtime::vba_record::VbaRecordLayout,
        dst: Value,
        src: Value,
    ) -> Value {
        let descriptor = data(
            module,
            b,
            record_descriptor(layout).expect("admitted record layout"),
            4,
        );
        self.call(module, b, "cvb_record_copy", &[dst, src, descriptor]);
        descriptor
    }
    pub fn clear_record_layout(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        layout: &calcvba_runtime::vba_record::VbaRecordLayout,
        cell: Value,
    ) {
        let descriptor = data(
            module,
            b,
            record_descriptor(layout).expect("admitted record layout"),
            4,
        );
        self.call(module, b, "cvb_record_clear", &[cell, descriptor]);
    }
    pub fn instruction(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        inst: &OxInst,
        foreign_state: Option<DataId>,
    ) -> bool {
        let layout = self.lower.layout;
        match inst {
            OxInst::ComCallEarly { .. } => {
                self.member_call(module, b, inst);
                return true;
            }
            OxInst::ComCallLate { .. } | OxInst::CallByName { .. }
                if !layout.class_calls.iter().any(|(site, _, _)| site == inst) =>
            {
                self.member_call(module, b, inst);
                return true;
            }
            OxInst::DrainTerminations => {
                self.call(module, b, "cvb_drain_terminations", &[self.lower.runtime]);
                return true;
            }
            OxInst::NewObject { dst, class } => {
                let out = self.address(b, layout.slot(*dst).unwrap());
                self.class_new(module, b, *class, out, false);
                return true;
            }
            OxInst::NewExtern { dst, .. } => {
                let out = self.address(b, layout.slot(*dst).unwrap());
                self.call(module, b, "cvb_collection_new", &[out]);
                return true;
            }
            OxInst::AsNew { .. } => return true,
            OxInst::FieldGet { dst, object, field } => {
                let object = self.operand(module, b, object, ScalarType::HostObject);
                let slot = layout.slot(*dst).unwrap();
                let out = self.address(b, slot);
                let field = b.ins().iconst(types::I32, i64::from(*field));
                let kind = b.ins().iconst(types::I32, i64::from(self.slot_kind(slot)));
                let set = b.ins().iconst(types::I32, 0);
                self.call(
                    module,
                    b,
                    "cvb_class_field",
                    &[self.lower.runtime, object, field, out, kind, set],
                );
                return true;
            }
            OxInst::FieldSet {
                object,
                field,
                value,
            } => {
                let object = self.operand(module, b, object, ScalarType::HostObject);
                let cell = self.cell(module, b, value);
                let field = b.ins().iconst(types::I32, i64::from(*field));
                let kind = b
                    .ins()
                    .iconst(types::I32, i64::from(kind_of_operand(layout, value)));
                let set = b.ins().iconst(types::I32, 1);
                self.call(
                    module,
                    b,
                    "cvb_class_field",
                    &[self.lower.runtime, object, field, cell, kind, set],
                );
                return true;
            }
            OxInst::TypeOfIs {
                dst,
                object,
                type_name,
            } => {
                self.object_type(module, b, object, type_name, false, Some(*dst));
                return true;
            }
            OxInst::ForEachInit { iter, source } => {
                let key = self.address(b, layout.slot(*iter).unwrap());
                let source_cell = self.cell(module, b, source);
                let kind = b.ins().iconst(
                    types::I32,
                    i64::from(
                        kind_of_operand(layout, source)
                            | if matches!(
                                source,
                                OxOperand::Use(OxPlace::Local(_) | OxPlace::Global(_))
                            ) {
                                0x1000_0000
                            } else {
                                0
                            },
                    ),
                );
                self.call(
                    module,
                    b,
                    "cvb_iter_init",
                    &[self.lower.runtime, key, source_cell, kind],
                );
                return true;
            }
            OxInst::ForEachNext {
                iter,
                item,
                has_value,
            } => {
                let key = self.address(b, layout.slot(*iter).unwrap());
                let item_slot = layout.slot(*item).unwrap();
                let item = self.address(b, item_slot);
                let kind = b
                    .ins()
                    .iconst(types::I32, i64::from(self.slot_kind(item_slot)));
                let has_value = self.address(b, layout.slot(*has_value).unwrap());
                self.call(
                    module,
                    b,
                    "cvb_iter_next",
                    &[self.lower.runtime, key, item, kind, has_value],
                );
                return true;
            }
            OxInst::ArrayLiteral {
                dst,
                values,
                lower_bound,
                aliases,
            } => {
                let slot = layout.slot(*dst).unwrap();
                let out = self.address(b, slot);
                let mut metadata = vec![
                    values.len() as u32,
                    self.slot_kind(slot),
                    *lower_bound as u32,
                ];
                let arguments = b.create_sized_stack_slot(StackSlotData::new(
                    StackSlotKind::ExplicitSlot,
                    (values.len().max(1) * 8) as u32,
                    3,
                ));
                for (index, value) in values.iter().enumerate() {
                    let cell = self.cell(module, b, value);
                    b.ins().stack_store(cell, arguments, (index * 8) as i32);
                    metadata.push(kind_of_operand(layout, value));
                }
                let alias_cells = b.create_sized_stack_slot(StackSlotData::new(
                    StackSlotKind::ExplicitSlot,
                    (values.len().max(1) * 8) as u32,
                    3,
                ));
                for index in 0..values.len() {
                    let alias = aliases.get(index).copied().flatten();
                    let (cell, kind) = if let Some(place) = alias {
                        (
                            self.cell(module, b, &OxOperand::Use(place)),
                            self.slot_kind(layout.slot(place).unwrap()),
                        )
                    } else {
                        (b.ins().iconst(types::I64, 0), 0)
                    };
                    b.ins().stack_store(cell, alias_cells, (index * 8) as i32);
                    metadata.push(kind);
                }
                let metadata = data(
                    module,
                    b,
                    metadata.into_iter().flat_map(u32::to_le_bytes).collect(),
                    4,
                );
                let arguments = b.ins().stack_addr(types::I64, arguments, 0);
                let aliases = b.ins().stack_addr(types::I64, alias_cells, 0);
                self.call(
                    module,
                    b,
                    "cvb_paramarray_literal",
                    &[self.lower.runtime, metadata, arguments, aliases, out],
                );
                return true;
            }
            OxInst::CallNative {
                dst,
                callee: OxNativeCallee::Builtin(native),
                args,
            } => {
                let args: Vec<_> = args
                    .iter()
                    .map(|argument| match argument {
                        OxCallArg::Operand(operand) => OxArg::ByVal(operand.clone()),
                        OxCallArg::ByRef(place) => OxArg::ByRef(*place),
                        OxCallArg::Omitted => OxArg::Omitted,
                        OxCallArg::Const(value) => {
                            OxArg::ByVal(OxOperand::Const(OxConst::I32(*value)))
                        }
                        _ => unreachable!("admitted builtin argument"),
                    })
                    .collect();
                self.library_call(
                    module,
                    b,
                    native_impl_index(*native).unwrap(),
                    false,
                    &args,
                    *dst,
                );
                return true;
            }
            OxInst::Coerce {
                dst,
                src,
                target: calcvba_oxir::OxCoerceTarget::FixedStr(len),
            } => {
                let out = self.address(b, layout.slot(*dst).unwrap());
                let source = self.cell(module, b, src);
                let kind = b
                    .ins()
                    .iconst(types::I32, i64::from(kind_of_operand(layout, src)));
                let len = b.ins().iconst(types::I32, i64::from(*len));
                self.call(
                    module,
                    b,
                    "cvb_fixed_string_coerce",
                    &[out, source, kind, len],
                );
                return true;
            }
            OxInst::Ptr { dst, kind, src } => {
                let address = self.cell(module, b, src);
                let value = if matches!(
                    kind,
                    calcvba_oxir::PtrKind::Str | calcvba_oxir::PtrKind::Obj
                ) {
                    b.ins()
                        .load(types::I64, ir::MemFlagsData::new(), address, 0)
                } else {
                    address
                };
                self.lower
                    .store_slot(b, layout.slot(*dst).unwrap(), ScalarType::LongLong, value);
                return true;
            }
            OxInst::CallNative {
                dst,
                callee: OxNativeCallee::Declare { descriptor_id, .. },
                args,
            } => {
                let descriptor = self
                    .program
                    .external_calls
                    .iter()
                    .find(|d| d.descriptor_id == *descriptor_id)
                    .unwrap();
                let result_record = descriptor.return_record_fields.as_ref().map(|fields| {
                    vba_record_layout_for_fields(fields).expect("admitted return record")
                });
                let result = if result_record.is_some() {
                    36
                } else {
                    descriptor
                        .return_type
                        .as_ref()
                        .map(declare_kind)
                        .unwrap_or(Some(0))
                        .unwrap()
                };
                let mut words = vec![result, args.len() as u32, 0, 0, 0];
                let mut schemas: Vec<(usize, Vec<u32>)> = Vec::new();
                if let Some(record) = &result_record {
                    schemas.push((4, record.boundary_schema()));
                }
                let vector = b.create_sized_stack_slot(StackSlotData::new(
                    StackSlotKind::ExplicitSlot,
                    (args.len().max(1) * 8) as u32,
                    3,
                ));
                for (i, arg) in args.iter().enumerate() {
                    let source = match arg {
                        OxCallArg::Operand(op) => op.clone(),
                        OxCallArg::ByRef(place) => OxOperand::Use(*place),
                        _ => unreachable!(),
                    };
                    let from = kind_of_operand(layout, &source);
                    let target = declare_kind(&descriptor.param_types[i]).unwrap_or(from);
                    let by_ref = descriptor.param_by_ref[i];
                    let mut cell = self.cell(module, b, &source);
                    if !by_ref && target != from {
                        let private = b.create_sized_stack_slot(StackSlotData::new(
                            StackSlotKind::ExplicitSlot,
                            8,
                            3,
                        ));
                        let zero = b.ins().iconst(types::I64, 0);
                        b.ins().stack_store(zero, private, 0);
                        let out = b.ins().stack_addr(types::I64, private, 0);
                        self.copy(module, b, out, target, cell, from);
                        cell = out;
                    }
                    let writable = target == 8 && matches!(source, OxOperand::Use(_));
                    words.push(
                        target
                            | if by_ref { 0x10000 } else { 0 }
                            | if writable { 0x20000 } else { 0 },
                    );
                    words.push(0);
                    if from == 36 {
                        let OxOperand::Use(place) = &source else {
                            unreachable!("admitted record argument")
                        };
                        let record = &layout.record_slots[&layout.slot(*place).unwrap()];
                        schemas.push((6 + i * 2, record.boundary_schema()));
                    }
                    b.ins().stack_store(cell, vector, (i * 8) as i32);
                }
                let mut metadata = vec![0u8; words.len() * 4];
                for (slot, schema) in schemas {
                    words[slot] = metadata.len() as u32;
                    for word in schema {
                        metadata.extend(word.to_le_bytes());
                    }
                }
                words[2] = metadata.len() as u32;
                metadata.extend_from_slice(descriptor.library.as_bytes());
                metadata.push(0);
                words[3] = metadata.len() as u32;
                metadata.extend_from_slice(if descriptor.alias.is_empty() {
                    descriptor.declared_name.as_bytes()
                } else {
                    descriptor.alias.as_bytes()
                });
                metadata.push(0);
                for (i, word) in words.iter().enumerate() {
                    metadata[i * 4..i * 4 + 4].copy_from_slice(&word.to_le_bytes());
                }
                let metadata = data(module, b, metadata, 4);
                let state = module.declare_data_in_func(
                    foreign_state.expect("planned native call state"),
                    b.func,
                );
                let state = b.ins().global_value(types::I64, state);
                let arguments = b.ins().stack_addr(types::I64, vector, 0);
                let output_size = result_record.as_ref().map_or(24, |r| {
                    r.automation_shape()
                        .expect("admitted return record")
                        .0
                        .max(24)
                        .next_multiple_of(8)
                });
                let output = b.create_sized_stack_slot(StackSlotData::new(
                    StackSlotKind::ExplicitSlot,
                    output_size as u32,
                    3,
                ));
                let zero = b.ins().iconst(types::I64, 0);
                for offset in (0..output_size).step_by(8) {
                    b.ins().stack_store(zero, output, offset as i32);
                }
                let output = b.ins().stack_addr(types::I64, output, 0);
                self.call(
                    module,
                    b,
                    "cvb_declare",
                    &[state, metadata, arguments, output],
                );
                if result != 0 {
                    let (destination, to) = if let Some(dst) = dst {
                        let slot = layout.slot(*dst).unwrap();
                        (self.address(b, slot), self.slot_kind(slot))
                    } else {
                        (b.ins().iconst(types::I64, 0), 0)
                    };
                    let descriptor = if let Some(record) = &result_record {
                        data(
                            module,
                            b,
                            record_descriptor(record).expect("admitted record result"),
                            4,
                        )
                    } else {
                        b.ins().iconst(types::I64, 0)
                    };
                    let to = b.ins().iconst(types::I32, i64::from(to));
                    let from = b.ins().iconst(types::I32, i64::from(result));
                    self.call(
                        module,
                        b,
                        "cvb_declare_result",
                        &[destination, output, to, from, descriptor],
                    );
                }
                return true;
            }
            OxInst::ErrFieldGet { dst, field } => {
                let out = self.address(b, layout.slot(*dst).unwrap());
                let kind = b.ins().iconst(
                    types::I32,
                    i64::from(self.slot_kind(layout.slot(*dst).unwrap())),
                );
                let field = b.ins().iconst(types::I32, i64::from(raw_err_field(*field)));
                self.call(
                    module,
                    b,
                    "cvb_error_get",
                    &[self.lower.runtime, out, kind, field],
                );
                return true;
            }
            OxInst::ErlGet { dst } => {
                let out = self.address(b, layout.slot(*dst).unwrap());
                let kind = b.ins().iconst(
                    types::I32,
                    i64::from(self.slot_kind(layout.slot(*dst).unwrap())),
                );
                let field = b.ins().iconst(types::I32, 6);
                self.call(
                    module,
                    b,
                    "cvb_error_get",
                    &[self.lower.runtime, out, kind, field],
                );
                return true;
            }
            OxInst::ErrFieldSet { field, src } => {
                let input = self.cell(module, b, src);
                let kind = b
                    .ins()
                    .iconst(types::I32, i64::from(kind_of_operand(layout, src)));
                let field = b.ins().iconst(types::I32, i64::from(raw_err_field(*field)));
                self.call(
                    module,
                    b,
                    "cvb_error_set",
                    &[self.lower.runtime, field, input, kind],
                );
                return true;
            }
            OxInst::SetErrorHandler(_) | OxInst::ClearErr | OxInst::SetLineNumber { .. } => {
                let (action, argument) = match inst {
                    OxInst::SetErrorHandler(handler) => match handler {
                        calcvba_oxir::ErrorHandler::Goto0 => (0, 0),
                        calcvba_oxir::ErrorHandler::ResumeNext => (1, 0),
                        calcvba_oxir::ErrorHandler::GotoLabel(block) => (2, block.0 as i64),
                        calcvba_oxir::ErrorHandler::GotoMinus1 => (3, 0),
                    },
                    OxInst::ClearErr => (4, 0),
                    OxInst::SetLineNumber { line } => (5, i64::from(*line)),
                    _ => unreachable!(),
                };
                let action = b.ins().iconst(types::I32, action);
                let argument = b.ins().iconst(types::I32, argument);
                self.call(
                    module,
                    b,
                    "cvb_error_control",
                    &[self.lower.runtime, action, argument],
                );
                return true;
            }
            OxInst::ArrayErase { array, .. } => {
                let slot = layout.slot(*array).unwrap();
                let cell = self.address(b, slot);
                let kind = b.ins().iconst(types::I32, i64::from(self.slot_kind(slot)));
                self.call(
                    module,
                    b,
                    "cvb_array_erase",
                    &[self.lower.runtime, cell, kind],
                );
                return true;
            }
            OxInst::ArrayGet {
                dst,
                array,
                indices,
            } => {
                let OxOperand::Use(place) = array else {
                    unreachable!("native array operand")
                };
                let slot = layout.slot(*place).unwrap();
                let array = self.cell(module, b, array);
                let rank = b.ins().iconst(types::I32, indices.len() as i64);
                if layout.slots[slot] == ScalarType::Variant {
                    let (kinds, vector) = self.variant_indices(module, b, indices);
                    let dst = layout.slot(*dst).unwrap();
                    let out = self.address(b, dst);
                    let to = b.ins().iconst(types::I32, i64::from(self.slot_kind(dst)));
                    let store = b.ins().iconst(types::I32, 0);
                    self.call(
                        module,
                        b,
                        "cvb_variant_index",
                        &[
                            self.lower.runtime,
                            array,
                            kinds,
                            vector,
                            rank,
                            out,
                            to,
                            store,
                        ],
                    );
                    return true;
                }
                let vector = self.indices(module, b, indices);
                if let Some((_, record)) = layout.array_record_slots.get(&slot) {
                    let descriptor = data(
                        module,
                        b,
                        record_descriptor(record).expect("admitted record array"),
                        4,
                    );
                    let out = self.address(b, layout.slot(*dst).unwrap());
                    self.call(
                        module,
                        b,
                        "cvb_record_array_get",
                        &[array, vector, rank, out, descriptor],
                    );
                    return true;
                }
                let kind = b
                    .ins()
                    .iconst(types::I32, i64::from(layout.array_slots[&slot]));
                let dst = layout.slot(*dst).unwrap();
                let out = self.address(b, dst);
                let to = b.ins().iconst(types::I32, i64::from(self.slot_kind(dst)));
                self.call(
                    module,
                    b,
                    "cvb_array_get",
                    &[array, vector, rank, kind, out, to],
                );
                return true;
            }
            OxInst::ArraySet {
                array,
                indices,
                value,
            } => {
                let slot = layout.slot(*array).unwrap();
                let array = self.address(b, slot);
                let rank = b.ins().iconst(types::I32, indices.len() as i64);
                if layout.slots[slot] == ScalarType::Variant {
                    let (kinds, vector) = self.variant_indices(module, b, indices);
                    let input = self.cell(module, b, value);
                    let from = b
                        .ins()
                        .iconst(types::I32, i64::from(kind_of_operand(layout, value)));
                    let store = b.ins().iconst(types::I32, 1);
                    self.call(
                        module,
                        b,
                        "cvb_variant_index",
                        &[
                            self.lower.runtime,
                            array,
                            kinds,
                            vector,
                            rank,
                            input,
                            from,
                            store,
                        ],
                    );
                    return true;
                }
                let vector = self.indices(module, b, indices);
                if let Some((_, record)) = layout.array_record_slots.get(&slot) {
                    let descriptor = data(
                        module,
                        b,
                        record_descriptor(record).expect("admitted record array"),
                        4,
                    );
                    let source = self.cell(module, b, value);
                    self.call(
                        module,
                        b,
                        "cvb_record_array_set",
                        &[array, vector, rank, source, descriptor],
                    );
                    return true;
                }
                let kind = b
                    .ins()
                    .iconst(types::I32, i64::from(layout.array_slots[&slot]));
                let input = self.cell(module, b, value);
                let from = b
                    .ins()
                    .iconst(types::I32, i64::from(kind_of_operand(layout, value)));
                self.call(
                    module,
                    b,
                    "cvb_paramarray_set",
                    &[self.lower.runtime, array, vector, rank, kind, input, from],
                );
                return true;
            }
            OxInst::Bound {
                dst,
                array,
                dimension,
                which,
            } => {
                let variant = layout.operand_type(array) == Ok(ScalarType::Variant);
                let array = self.cell(module, b, array);
                let dimension = if let Some(dimension) = dimension {
                    self.operand(module, b, dimension, ScalarType::Long)
                } else {
                    b.ins().iconst(types::I32, 1)
                };
                let upper = b
                    .ins()
                    .iconst(types::I32, i64::from(*which == BoundWhich::Upper));
                let out = self.address(b, layout.slot(*dst).unwrap());
                self.call(
                    module,
                    b,
                    if variant {
                        "cvb_variant_array_bound"
                    } else {
                        "cvb_array_bound"
                    },
                    &[array, dimension, upper, out],
                );
                return true;
            }
            OxInst::ArrayRedim {
                dst,
                upper_bounds,
                lower_bounds,
                preserve,
                fixed,
                ..
            } => {
                let slot = layout.slot(*dst).unwrap();
                let array = self.address(b, slot);
                let bounds = b.create_sized_stack_slot(StackSlotData::new(
                    StackSlotKind::ExplicitSlot,
                    (upper_bounds.len() * 8) as u32,
                    3,
                ));
                for (i, upper) in upper_bounds.iter().enumerate() {
                    let lower = if let Some(lower) = lower_bounds.get(i) {
                        self.operand(module, b, lower, ScalarType::Long)
                    } else {
                        b.ins().iconst(types::I32, 0)
                    };
                    let upper = self.operand(module, b, upper, ScalarType::Long);
                    b.ins().stack_store(lower, bounds, (i * 8) as i32);
                    b.ins().stack_store(upper, bounds, (i * 8 + 4) as i32);
                }
                let bounds = b.ins().stack_addr(types::I64, bounds, 0);
                let kind = b
                    .ins()
                    .iconst(types::I32, i64::from(layout.array_slots[&slot]));
                let rank = b.ins().iconst(types::I32, upper_bounds.len() as i64);
                let preserve = b.ins().iconst(types::I32, i64::from(*preserve));
                let fixed = b.ins().iconst(types::I32, i64::from(*fixed));
                if let Some((id, record)) = layout.array_record_slots.get(&slot) {
                    let mut descriptor = record_descriptor(record).expect("admitted array record");
                    let identity = format!("{}:{}:{record:?}", self.program.unit_name, id.0);
                    let mut first = 0xcbf29ce484222325u64;
                    let mut second = 0x84222325cbf29ce4u64;
                    for byte in identity.as_bytes() {
                        first = (first ^ u64::from(*byte)).wrapping_mul(0x100000001b3);
                        second = (second ^ u64::from(*byte)).wrapping_mul(0x100000001b3);
                    }
                    descriptor.extend(first.to_le_bytes());
                    descriptor.extend(second.to_le_bytes());
                    let descriptor = data(module, b, descriptor, 4);
                    self.call(
                        module,
                        b,
                        "cvb_record_array_redim",
                        &[array, rank, bounds, preserve, fixed, descriptor],
                    );
                    return true;
                }
                self.call(
                    module,
                    b,
                    "cvb_array_redim_encoded",
                    &[
                        array,
                        kind,
                        rank,
                        bounds,
                        preserve,
                        fixed,
                        self.lower.runtime,
                    ],
                );
                return true;
            }
            OxInst::NewRecord { dst, .. } => {
                self.clear(module, b, layout.slot(*dst).unwrap());
                return true;
            }
            OxInst::Assign { dst, value }
                if layout.slots[layout.slot(*dst).unwrap()] == ScalarType::Record =>
            {
                let index = layout.slot(*dst).unwrap();
                let out = self.address(b, index);
                let src = self.cell(module, b, value);
                self.copy_record(module, b, index, out, src);
                return true;
            }
            OxInst::RecordGet { dst, record, index }
            | OxInst::RecordArrayGet {
                dst, record, index, ..
            } => {
                let OxOperand::Use(place) = record else {
                    unreachable!("record operand")
                };
                let slot = layout.slot(*place).unwrap();
                let field = &layout.record_slots[&slot].fields()[*index];
                let record = self.cell(module, b, record);
                let source = b.ins().iadd_imm(
                    record,
                    layout.record_slots[&slot]
                        .automation_field_offset(*index)
                        .expect("admitted field") as i64,
                );
                let (source, field) = if let OxInst::RecordArrayGet { indices, .. } = inst {
                    let calcvba_runtime::vba_record::VbaRecordFieldKind::FixedArray {
                        element,
                        bounds,
                    } = field.kind()
                    else {
                        unreachable!("admitted fixed array field")
                    };
                    (
                        self.fixed_element_address(module, b, source, element, bounds, indices),
                        element.as_ref(),
                    )
                } else {
                    (source, field.kind())
                };
                let to = layout.slot(*dst).unwrap();
                let out = self.address(b, to);
                if let calcvba_runtime::vba_record::VbaRecordFieldKind::FixedString { len } = field
                {
                    let len = b.ins().iconst(types::I32, *len as i64);
                    let kind = b.ins().iconst(types::I32, i64::from(self.slot_kind(to)));
                    self.call(module, b, "cvb_fixed_string_get", &[out, source, len, kind]);
                } else if layout.slots[to] == ScalarType::Record {
                    self.copy_record(module, b, to, out, source);
                } else {
                    self.copy(
                        module,
                        b,
                        out,
                        self.slot_kind(to),
                        source,
                        field_kind(field).expect("admitted field"),
                    );
                }
                return true;
            }
            OxInst::RecordSet {
                record,
                index,
                value,
            }
            | OxInst::RecordArraySet {
                record,
                index,
                value,
                ..
            } => {
                let slot = layout.slot(*record).unwrap();
                let field = &layout.record_slots[&slot].fields()[*index];
                let address = self.address(b, slot);
                let out = b.ins().iadd_imm(
                    address,
                    layout.record_slots[&slot]
                        .automation_field_offset(*index)
                        .expect("admitted field") as i64,
                );
                let (out, field) = if let OxInst::RecordArraySet { indices, .. } = inst {
                    let calcvba_runtime::vba_record::VbaRecordFieldKind::FixedArray {
                        element,
                        bounds,
                    } = field.kind()
                    else {
                        unreachable!("admitted fixed array field")
                    };
                    (
                        self.fixed_element_address(module, b, out, element, bounds, indices),
                        element.as_ref(),
                    )
                } else {
                    (out, field.kind())
                };
                let src = self.cell(module, b, value);
                if let calcvba_runtime::vba_record::VbaRecordFieldKind::FixedString { len } = field
                {
                    let len = b.ins().iconst(types::I32, *len as i64);
                    let from = b.ins().iconst(
                        types::I32,
                        i64::from(kind(layout.operand_type(value).unwrap())),
                    );
                    self.call(module, b, "cvb_fixed_string_set", &[out, len, src, from]);
                } else if let calcvba_runtime::vba_record::VbaRecordFieldKind::Record(nested) =
                    field
                {
                    let descriptor = data(
                        module,
                        b,
                        record_descriptor(nested).expect("admitted nested record"),
                        4,
                    );
                    self.call(module, b, "cvb_record_copy", &[out, src, descriptor]);
                } else {
                    self.copy(
                        module,
                        b,
                        out,
                        field_kind(field).expect("admitted field"),
                        src,
                        kind(layout.operand_type(value).unwrap()),
                    );
                }
                return true;
            }
            _ => {}
        }
        if let OxInst::StmtBoundary {
            clear_temps_from, ..
        } = inst
        {
            for index in layout.local_count + clear_temps_from..layout.global_base {
                if self.iterator_slots.contains(&index) {
                    let address = self.address(b, index);
                    self.call(module, b, "cvb_iter_clear", &[self.lower.runtime, address]);
                }
                if owning(layout.slots[index]) {
                    self.clear(module, b, index);
                } else {
                    let value =
                        if matches!(layout.slots[index], ScalarType::Double | ScalarType::Date) {
                            b.ins().f64const(0.0)
                        } else if layout.slots[index] == ScalarType::Single {
                            b.ins().f32const(0.0)
                        } else {
                            b.ins().iconst(machine_type(layout.slots[index]), 0)
                        };
                    self.lower.store_slot(b, index, layout.slots[index], value);
                }
            }
            self.call(module, b, "cvb_drain_terminations", &[self.lower.runtime]);
            return true;
        }
        if let OxInst::CompareObjectIs { dst, lhs, rhs } = inst {
            let a = self.operand(module, b, lhs, ScalarType::HostObject);
            let c = self.operand(module, b, rhs, ScalarType::HostObject);
            let out = self.address(b, layout.slot(*dst).unwrap());
            self.call(module, b, "cvb_object_is", &[a, c, out]);
            return true;
        }
        if let OxInst::ValidateAssignment {
            src,
            target_kind: AssignmentTargetKind::Object,
            target_type_name,
            ..
        } = inst
        {
            self.object_type(module, b, src, target_type_name, true, None);
            return true;
        }
        if matches!(inst, OxInst::ValidateAssignment { .. }) {
            return true;
        }
        let owning_copy = match inst {
            OxInst::Assign { dst, .. } | OxInst::Coerce { dst, .. } => {
                owning(layout.slots[layout.slot(*dst).unwrap()])
                    || extended_numeric(layout.slots[layout.slot(*dst).unwrap()])
                    || matches!(inst,OxInst::Assign {value,..} if layout.operand_type(value).is_ok_and(extended_numeric))
            }
            _ => false,
        };
        if !owning_copy
            && !matches!(inst, OxInst::Logical { .. } | OxInst::Pow { .. })
            && !calcvba_rt_abi::host_entry::value::needs_helper(inst, layout)
        {
            return false;
        }
        if let OxInst::Coerce { dst, src, .. } = inst
            && basic_numeric(layout.slots[layout.slot(*dst).unwrap()])
            && layout.operand_type(src).is_ok_and(basic_numeric)
        {
            return false;
        }
        if let OxInst::Truthy { dst, src } = inst {
            let out = self.address(b, layout.slot(*dst).unwrap());
            let input = self.cell(module, b, src);
            let kind = b
                .ins()
                .iconst(types::I32, i64::from(kind_of_operand(layout, src)));
            let to = b.ins().iconst(types::I32, 11);
            let family = b.ins().iconst(types::I32, 2);
            let mode = b.ins().iconst(types::I32, 0);
            self.call(
                module,
                b,
                "cvb_value_unary",
                &[self.lower.runtime, out, input, to, kind, family, mode],
            );
            return true;
        }
        let (dst, src) = match inst {
            OxInst::Assign { dst, value }
            | OxInst::Coerce {
                dst, src: value, ..
            }
            | OxInst::Box {
                dst, src: value, ..
            }
            | OxInst::Unbox {
                dst, src: value, ..
            }
            | OxInst::Truthy { dst, src: value } => (*dst, value),
            OxInst::Concat { dst, lhs, rhs }
            | OxInst::Logical { dst, lhs, rhs, .. }
            | OxInst::Arith { dst, lhs, rhs, .. }
            | OxInst::Pow { dst, lhs, rhs }
            | OxInst::Div { dst, lhs, rhs }
            | OxInst::Compare { dst, lhs, rhs, .. } => {
                let index = layout.slot(*dst).unwrap();
                let out = self.address(b, index);
                let a = self.cell(module, b, lhs);
                let c = self.cell(module, b, rhs);
                let ak = layout.operand_type(lhs).unwrap();
                let bk = layout.operand_type(rhs).unwrap();
                if matches!(inst, OxInst::Concat { .. })
                    && ak == ScalarType::String
                    && bk == ScalarType::String
                    && layout.slots[index] == ScalarType::String
                {
                    self.call(module, b, "cvb_string_concat", &[out, a, c]);
                    return true;
                }
                let ak = b
                    .ins()
                    .iconst(types::I32, i64::from(kind_of_operand(layout, lhs)));
                let bk = b
                    .ins()
                    .iconst(types::I32, i64::from(kind_of_operand(layout, rhs)));
                let (family, operation, mode) = match inst {
                    OxInst::Arith { op, mode, .. } => (
                        0,
                        raw_arith_op(*op).expect("arithmetic operation"),
                        raw_numeric_mode(*mode).unwrap(),
                    ),
                    OxInst::Div { .. } => (0, RT_ARITH_DIV, RT_NUMERIC_WIDENING),
                    OxInst::Pow { .. } => (0, RT_ARITH_POW, RT_NUMERIC_WIDENING),
                    OxInst::Logical { op, .. } => (1, raw_logical_op(*op), 0),
                    OxInst::Compare {
                        op, mode, coercion, ..
                    } => {
                        let coercion = match coercion {
                            calcvba_bundle::ComparisonCoercion::Strict => 0,
                            calcvba_bundle::ComparisonCoercion::String => 1,
                            calcvba_bundle::ComparisonCoercion::Numeric => 2,
                            calcvba_bundle::ComparisonCoercion::Variant => 3,
                        };
                        (
                            2,
                            raw_compare_op(*op),
                            raw_string_compare_mode(*mode) | (coercion << 8),
                        )
                    }
                    OxInst::Concat { .. } => (3, 0, 0),
                    _ => unreachable!("admitted binary operation"),
                };
                let to = b.ins().iconst(types::I32, i64::from(self.slot_kind(index)));
                let family = b.ins().iconst(types::I32, family);
                let operation = b.ins().iconst(types::I32, i64::from(operation));
                let mode = b.ins().iconst(types::I32, i64::from(mode));
                self.call(
                    module,
                    b,
                    "cvb_value_binary",
                    &[
                        self.lower.runtime,
                        out,
                        a,
                        c,
                        to,
                        ak,
                        bk,
                        family,
                        operation,
                        mode,
                    ],
                );
                return true;
            }
            OxInst::Neg { dst, src, .. } | OxInst::Not { dst, src } => {
                let index = layout.slot(*dst).unwrap();
                let out = self.address(b, index);
                let src_cell = self.cell(module, b, src);
                let to = b.ins().iconst(types::I32, i64::from(self.slot_kind(index)));
                let from = b.ins().iconst(
                    types::I32,
                    i64::from(kind(layout.operand_type(src).unwrap())),
                );
                let (family, mode) = if let OxInst::Neg { mode, .. } = inst {
                    (0, raw_numeric_mode(*mode).unwrap())
                } else {
                    (1, 0)
                };
                let family = b.ins().iconst(types::I32, family);
                let mode = b.ins().iconst(types::I32, i64::from(mode));
                self.call(
                    module,
                    b,
                    "cvb_value_unary",
                    &[self.lower.runtime, out, src_cell, to, from, family, mode],
                );
                return true;
            }
            _ => return false,
        };
        let index = layout.slot(dst).unwrap();
        let out = self.address(b, index);
        let input = self.cell(module, b, src);
        self.copy(
            module,
            b,
            out,
            self.slot_kind(index),
            input,
            kind_of_operand(layout, src),
        );
        true
    }
    fn fixed_element_address(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        base: Value,
        element: &calcvba_runtime::vba_record::VbaRecordFieldKind,
        bounds: &[calcvba_runtime::safe_array::SafeArrayBound],
        indices: &[OxOperand],
    ) -> Value {
        let mut offset = b.ins().iconst(types::I64, 0);
        let (width, align) = element
            .automation_storage_shape()
            .expect("admitted field element");
        let mut stride = width.next_multiple_of(align) as i64;
        for (bound, index) in bounds.iter().zip(indices) {
            let index = self.operand(module, b, index, ScalarType::Long);
            let index = b.ins().sextend(types::I64, index);
            let relative = b.ins().iadd_imm(index, -i64::from(bound.lower));
            let valid = b.ins().icmp_imm(
                ir::condcodes::IntCC::UnsignedLessThan,
                relative,
                i64::from(bound.count),
            );
            let ok = b.ins().iconst(types::I32, 0);
            let error = b.ins().iconst(types::I32, 9);
            let status = b.ins().select(valid, ok, error);
            self.lower.check_status(b, status);
            let term = b.ins().imul_imm(relative, stride);
            offset = b.ins().iadd(offset, term);
            stride *= i64::from(bound.count);
        }
        b.ins().iadd(base, offset)
    }
    fn variant_indices(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        indices: &[OxOperand],
    ) -> (Value, Value) {
        let pointers = b.create_sized_stack_slot(StackSlotData::new(
            StackSlotKind::ExplicitSlot,
            (indices.len().max(1) * 8) as u32,
            3,
        ));
        let mut kinds = Vec::with_capacity(indices.len());
        for (index, operand) in indices.iter().enumerate() {
            let cell = self.cell(module, b, operand);
            b.ins().stack_store(cell, pointers, (index * 8) as i32);
            kinds.push(kind_of_operand(self.lower.layout, operand));
        }
        let kinds = data(
            module,
            b,
            kinds.into_iter().flat_map(u32::to_le_bytes).collect(),
            4,
        );
        let pointers = b.ins().stack_addr(types::I64, pointers, 0);
        (kinds, pointers)
    }
    fn indices(
        &self,
        module: &mut impl Module,
        b: &mut FunctionBuilder<'_>,
        indices: &[OxOperand],
    ) -> Value {
        let slot = b.create_sized_stack_slot(StackSlotData::new(
            StackSlotKind::ExplicitSlot,
            (indices.len().max(1) * 4) as u32,
            2,
        ));
        for (i, index) in indices.iter().enumerate() {
            let value = self.operand(module, b, index, ScalarType::Long);
            b.ins().stack_store(value, slot, (i * 4) as i32);
        }
        b.ins().stack_addr(types::I64, slot, 0)
    }
}

fn declare_kind(kind: &calcvba_bundle::DeclareParamType) -> Option<u32> {
    use calcvba_bundle::DeclareParamType as D;
    Some(match kind {
        D::Byte => 17,
        D::Integer => 2,
        D::Long => 3,
        D::Boolean => 11,
        D::LongLong | D::LongPtr => 20,
        D::Single => 4,
        D::Double => 5,
        D::Currency => 6,
        D::Date => 7,
        D::String => 8,
        D::Variant => 12,
        D::Array => 0x2000,
        D::Any => return None,
    })
}
pub(super) fn validate_declare(
    program: &OxProgram,
    layout: &ScalarLayout,
    id: u32,
    args: &[OxCallArg],
) -> Result<(), String> {
    use calcvba_bundle::DeclareParamType as D;
    let descriptor = program
        .external_calls
        .iter()
        .find(|d| d.descriptor_id == id)
        .ok_or("absent native Declare descriptor")?;
    if descriptor.ordinal_alias || args.len() != descriptor.param_count || args.len() > 32 {
        return Err("native Declare ordinal or argument shape is not yet admitted".into());
    }
    if descriptor.return_type.as_ref().is_some_and(|k| {
        matches!(k, D::Array) || matches!(k, D::Any) && descriptor.return_record_fields.is_none()
    }) {
        return Err("native Declare requires a scalar/pointer return".into());
    }
    for (i, arg) in args.iter().enumerate() {
        let (source, alias) = match arg {
            OxCallArg::ByRef(p) => (OxOperand::Use(*p), true),
            OxCallArg::Operand(op) => (op.clone(), false),
            _ => return Err("native Declare named/omitted argument".into()),
        };
        let ty = layout.operand_type(&source)?;
        let by_ref = descriptor.param_by_ref[i];
        let expected = declare_kind(&descriptor.param_types[i]);
        if by_ref && !alias {
            return Err(
                "native Declare private/forced ByRef argument requires explicit pointer typing"
                    .into(),
            );
        }
        if !by_ref && expected == Some(8) && ty != ScalarType::String {
            return Err("native Declare String input requires a String carrier".into());
        }
        if by_ref && expected == Some(0x2000) && ty != ScalarType::Array {
            return Err("native Declare array input requires an array cell".into());
        }
        if by_ref && expected.is_some_and(|k| k != 0x2000 && k != kind(ty)) {
            return Err("native Declare ByRef carrier mismatch".into());
        }
        if !by_ref
            && (matches!(
                ty,
                ScalarType::Record | ScalarType::Array | ScalarType::HostObject
            ) || expected == Some(12))
        {
            return Err("native Declare ByVal aggregate/object carrier is not yet admitted".into());
        }
    }
    Ok(())
}

pub(super) fn kind_of_operand(layout: &ScalarLayout, operand: &OxOperand) -> u32 {
    if let OxOperand::Use(place) = operand
        && layout.operand_type(operand) == Ok(ScalarType::Array)
    {
        0x2000 | u32::from(layout.array_slots[&layout.slot(*place).unwrap()])
    } else {
        kind(layout.operand_type(operand).unwrap())
    }
}

pub(super) fn field_kind(
    kind: &calcvba_runtime::vba_record::VbaRecordFieldKind,
) -> Result<u32, String> {
    use calcvba_runtime::vba_record::VbaRecordFieldKind as F;
    Ok(match kind {
        F::Integer => 2,
        F::Long => 3,
        F::Single => 4,
        F::Double => 5,
        F::Currency => 6,
        F::Date => 7,
        F::String | F::FixedString { .. } => 8,
        F::Variant => 12,
        F::Boolean => 11,
        F::Byte => 17,
        F::LongLong => 20,
        F::Record(_) => 36,
        _ => return Err(format!("native record field {kind:?} is not yet admitted")),
    })
}

pub(super) fn array_field_kind(
    field: &calcvba_runtime::vba_record::VbaRecordFieldKind,
    rank: usize,
) -> Result<u32, String> {
    if let calcvba_runtime::vba_record::VbaRecordFieldKind::FixedArray { element, bounds } = field
        && rank == bounds.len()
        && rank != 0
    {
        return field_kind(element);
    }
    Err("native record array field requires matching fixed bounds".into())
}

pub(super) fn record_descriptor(
    layout: &calcvba_runtime::vba_record::VbaRecordLayout,
) -> Result<Vec<u8>, String> {
    use calcvba_runtime::vba_record::VbaRecordFieldKind as F;
    fn fields(
        layout: &calcvba_runtime::vba_record::VbaRecordLayout,
        base: usize,
        owners: &mut Vec<(u32, u32)>,
    ) -> Result<(), String> {
        fn field(kind: &F, offset: usize, owners: &mut Vec<(u32, u32)>) -> Result<(), String> {
            match kind {
                F::String => owners.push((offset as u32, 8)),
                F::Variant => owners.push((offset as u32, 12)),
                F::Record(nested) => fields(nested, offset, owners)?,
                F::FixedArray { element, bounds } => {
                    let count = bounds
                        .iter()
                        .try_fold(1usize, |n, b| n.checked_mul(b.count as usize))
                        .ok_or("record array extent overflow")?;
                    let (width, align) = element.automation_storage_shape()?;
                    for index in 0..count {
                        field(
                            element,
                            offset + index * width.next_multiple_of(align),
                            owners,
                        )?;
                    }
                }
                _ => {}
            }
            Ok(())
        }
        for (index, item) in layout.fields().iter().enumerate() {
            field(
                item.kind(),
                base + layout.automation_field_offset(index)?,
                owners,
            )?;
        }
        Ok(())
    }
    let mut owners = Vec::new();
    fields(layout, 0, &mut owners)?;
    let mut bytes = (layout.automation_shape()?.0 as u32).to_le_bytes().to_vec();
    bytes.extend_from_slice(&(owners.len() as u32).to_le_bytes());
    for (offset, kind) in owners {
        bytes.extend_from_slice(&offset.to_le_bytes());
        bytes.extend_from_slice(&kind.to_le_bytes());
    }
    Ok(bytes)
}

pub(super) fn data(
    module: &mut impl Module,
    b: &mut FunctionBuilder<'_>,
    bytes: Vec<u8>,
    align: u64,
) -> Value {
    let id = module
        .declare_anonymous_data(false, false)
        .expect("admitted native constant");
    let mut description = cranelift_module::DataDescription::new();
    description.define(bytes.into_boxed_slice());
    description.set_align(align);
    module
        .define_data(id, &description)
        .expect("native constant data");
    let global = module.declare_data_in_func(id, b.func);
    b.ins().global_value(types::I64, global)
}
