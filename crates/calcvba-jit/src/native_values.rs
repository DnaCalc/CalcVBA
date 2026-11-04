use super::*;

pub(super) const HELPERS: &[(&str, &[ir::Type])] = &[
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
pub(super) fn owning(ty: ScalarType) -> bool {
    !numeric(ty) && ty != ScalarType::Void
}

pub(super) struct Values<'a> {
    pub lower: ScalarLower<'a>,
    pub functions: &'a HashMap<&'static str, ClifFuncId>,
    pub program: &'a OxProgram,
}
impl Values<'_> {
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
        check_status(b, status);
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
            return self.address(b, self.lower.layout.slot(*place).unwrap());
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
        if matches!(operand, OxOperand::Const(OxConst::Str(_))) {
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
            OxInst::ErrFieldGet {
                dst,
                field: ErrField::LastDllError,
            } => {
                let helper =
                    module.declare_func_in_func(self.functions["cvb_last_dll_error"], b.func);
                let call = b.ins().call(helper, &[]);
                let value = b.inst_results(call)[0];
                self.lower
                    .store_slot(b, layout.slot(*dst).unwrap(), ScalarType::Long, value);
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
                let vector = self.indices(module, b, indices);
                let rank = b.ins().iconst(types::I32, indices.len() as i64);
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
                let vector = self.indices(module, b, indices);
                let rank = b.ins().iconst(types::I32, indices.len() as i64);
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
                    "cvb_array_set",
                    &[array, vector, rank, kind, input, from],
                );
                return true;
            }
            OxInst::Bound {
                dst,
                array,
                dimension,
                which,
            } => {
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
                    "cvb_array_bound",
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
            return true;
        }
        if let OxInst::CompareObjectIs { dst, lhs, rhs } = inst {
            let a = self.operand(module, b, lhs, ScalarType::HostObject);
            let c = self.operand(module, b, rhs, ScalarType::HostObject);
            let out = self.address(b, layout.slot(*dst).unwrap());
            self.call(module, b, "cvb_object_is", &[a, c, out]);
            return true;
        }
        if matches!(inst, OxInst::ValidateAssignment { .. }) {
            return true;
        }
        let owning_copy = match inst {
            OxInst::Assign { dst, .. } | OxInst::Coerce { dst, .. } => {
                owning(layout.slots[layout.slot(*dst).unwrap()])
                    || extended_numeric(layout.slots[layout.slot(*dst).unwrap()])
            }
            _ => false,
        };
        if !owning_copy
            && !matches!(inst, OxInst::Logical { .. })
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
            self.call(module, b, "cvb_value_truthy", &[out, input, kind]);
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
                let ak = b.ins().iconst(types::I32, i64::from(kind(ak)));
                let bk = b.ins().iconst(types::I32, i64::from(kind(bk)));
                if let OxInst::Compare { op, mode, .. } = inst {
                    let op = b.ins().iconst(
                        types::I32,
                        match op {
                            CmpOp::Eq => 0,
                            CmpOp::Ne => 1,
                            CmpOp::Lt => 2,
                            CmpOp::Le => 3,
                            CmpOp::Gt => 4,
                            CmpOp::Ge => 5,
                        },
                    );
                    let mode = b
                        .ins()
                        .iconst(types::I32, i64::from(*mode == StringCompareMode::Text));
                    let to = b.ins().iconst(types::I32, i64::from(self.slot_kind(index)));
                    self.call(
                        module,
                        b,
                        "cvb_value_compare",
                        &[out, a, c, ak, bk, op, mode, to],
                    );
                } else {
                    let to = b.ins().iconst(types::I32, i64::from(self.slot_kind(index)));
                    let name = match inst {
                        OxInst::Logical { op, .. } => match op {
                            LogicalOp::And => "cvb_value_and",
                            LogicalOp::Or => "cvb_value_or",
                            LogicalOp::Xor => "cvb_value_xor",
                            LogicalOp::Eqv => "cvb_value_eqv",
                            LogicalOp::Imp => "cvb_value_imp",
                        },
                        OxInst::Concat { .. } => "cvb_value_concat",
                        OxInst::Div { .. } => "cvb_value_div",
                        OxInst::Arith {
                            op: ArithOp::Add, ..
                        } => "cvb_value_add",
                        OxInst::Arith {
                            op: ArithOp::Sub, ..
                        } => "cvb_value_sub",
                        OxInst::Arith {
                            op: ArithOp::Mul, ..
                        } => "cvb_value_mul",
                        _ => unreachable!("admitted value operation"),
                    };
                    self.call(module, b, name, &[out, a, c, to, ak, bk]);
                }
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
                self.call(
                    module,
                    b,
                    if matches!(inst, OxInst::Neg { .. }) {
                        "cvb_value_neg"
                    } else {
                        "cvb_value_not"
                    },
                    &[out, src_cell, to, from],
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
            check_status(b, status);
            let term = b.ins().imul_imm(relative, stride);
            offset = b.ins().iadd(offset, term);
            stride *= i64::from(bound.count);
        }
        b.ins().iadd(base, offset)
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

fn kind_of_operand(layout: &ScalarLayout, operand: &OxOperand) -> u32 {
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
