//! Direct AOT entries using the same ScalarLower as the JIT. The module decides
//! symbol placement. Owning carriers use individual value-operation helpers.

use super::*;
#[path = "native_values.rs"]
mod values;

pub fn value_helper_signatures() -> &'static [(&'static str, &'static [ir::Type])] {
    values::HELPERS
}

/// Dynamic Raise now requires the extended native context's runtime services.
/// Refuse an older template/consumer pairing instead of reading beyond its context.
pub fn value_helper_version(name: &str) -> u16 {
    if matches!(
        name,
        "cvb_raise_dynamic" | "cvb_object_type" | "cvb_paramarray_literal" | "cvb_frame_enter"
    ) {
        2
    } else {
        1
    }
}
use crate::typed::{PhysicalSlots, ScalarLower, machine_type};
use calcvba_oxir::OxArg;
use calcvba_rt_abi::host_entry::{LinkedTarget, ScalarLayout, ScalarType};
use calcvba_rt_abi::native_export::Context as NativeContext;
use cranelift_module::DataId;
use std::collections::{BTreeMap, BTreeSet};

/// Keep native recursion inside a bounded part of the caller's stack. A typed
/// status 28 leaves the DLL through the ordinary zero-on-error wrapper path.
pub const MAX_NATIVE_CALL_DEPTH: i32 = 128;
const DEPTH_OFFSET: i32 = 8;

pub type NativeCallSite = (usize, usize, usize);

fn conversion_target(id: NativeImplId) -> Option<ScalarType> {
    match id {
        NativeImplId::CByte => Some(ScalarType::Byte),
        NativeImplId::CInt => Some(ScalarType::Integer),
        NativeImplId::CLngLng | NativeImplId::CLngPtr => Some(ScalarType::LongLong),
        NativeImplId::CSng => Some(ScalarType::Single),
        NativeImplId::CCur => Some(ScalarType::Currency),
        NativeImplId::CDate => Some(ScalarType::Date),
        NativeImplId::CLng => Some(ScalarType::Long),
        NativeImplId::CDbl => Some(ScalarType::Double),
        NativeImplId::CBool => Some(ScalarType::Boolean),
        _ => None,
    }
}

fn check_status(b: &mut FunctionBuilder<'_>, status: Value) {
    let failed = b.ins().icmp_imm(ir::condcodes::IntCC::NotEqual, status, 0);
    let fail = b.create_block();
    let next = b.create_block();
    b.ins().brif(failed, fail, &[], next, &[]);
    b.switch_to_block(fail);
    b.ins().return_(&[status]);
    b.switch_to_block(next);
}

pub struct NativeScalar {
    pub procedure: FuncId,
    pub function: ClifFuncId,
    pub layout: ScalarLayout,
    offsets: Vec<i32>,
    storage_size: u32,
}
impl NativeScalar {
    fn register_record_result(&self) -> Option<ir::Type> {
        if self.layout.result != ScalarType::Record {
            return None;
        }
        let layout = &self.layout.record_slots[&self.layout.return_slot];
        if values::record_descriptor(layout).ok()?.len() != 8 {
            return None;
        }
        match layout
            .automation_shape()
            .expect("admitted Automation layout")
            .0
        {
            1 => Some(types::I8),
            2 => Some(types::I16),
            4 => Some(types::I32),
            8 => Some(types::I64),
            _ => None,
        }
    }
    fn result_size(&self) -> usize {
        if self.layout.result == ScalarType::Record {
            self.layout.record_slots[&self.layout.return_slot]
                .automation_shape()
                .expect("admitted Automation layout")
                .0
        } else if self.layout.result == ScalarType::Variant {
            24
        } else {
            8
        }
    }
}
fn problem(func: &OxFunc, detail: impl std::fmt::Display) -> JitError {
    JitError::unsupported(format!("native export procedure {}: {detail}", func.name))
}
fn numeric(ty: ScalarType) -> bool {
    basic_numeric(ty) || extended_numeric(ty)
}
fn extended_numeric(ty: ScalarType) -> bool {
    matches!(
        ty,
        ScalarType::Byte
            | ScalarType::Integer
            | ScalarType::LongLong
            | ScalarType::Single
            | ScalarType::Currency
            | ScalarType::Date
    )
}
fn basic_numeric(ty: ScalarType) -> bool {
    matches!(
        ty,
        ScalarType::Long | ScalarType::Double | ScalarType::Boolean
    )
}

/// Compile the complete reachable closure, or return an explicit refusal. The
/// returned functions use (TypedContext*, typed parameters..., result*)->status.
pub fn compile_native_scalars(
    module: &mut impl Module,
    program: &OxProgram,
    roots: &[FuncId],
    library_helpers: &HashMap<NativeImplId, ClifFuncId>,
    long_from_f64: ClifFuncId,
    value_helpers: &HashMap<&'static str, ClifFuncId>,
    foreign_states: &HashMap<NativeCallSite, DataId>,
) -> Result<Vec<NativeScalar>, JitError> {
    calcvba_oxir::verify_program(program)
        .map_err(|e| JitError::unsupported(format!("invalid native OxIR: {e:?}")))?;
    let mut pending = roots.to_vec();
    pending.extend(program.global_initializer);
    let mut reached = BTreeSet::new();
    let mut plans = BTreeMap::new();
    while let Some(id) = pending.pop() {
        if !reached.insert(id.0) {
            continue;
        }
        let func = program
            .funcs
            .get(id.0)
            .ok_or_else(|| JitError::unsupported("native export refers to absent procedure"))?;
        let mut layout =
            ScalarLayout::with_native_program(func, program).map_err(|e| problem(func, e))?;
        if let Some(index) = layout.param_array {
            layout.by_ref_parameters[index] = true;
        }
        if !layout.slots.iter().copied().all(|ty| {
            numeric(ty)
                || matches!(
                    ty,
                    ScalarType::String
                        | ScalarType::Variant
                        | ScalarType::HostObject
                        | ScalarType::Record
                        | ScalarType::Array
                )
        }) {
            return Err(problem(func, "native path does not yet admit this carrier"));
        }
        for record in layout.record_slots.values() {
            values::record_descriptor(record).map_err(|e| problem(func, e))?;
        }
        for (_, proc, _) in &layout.class_calls {
            pending.push(*proc);
        }
        if layout
            .parameters
            .iter()
            .zip(&layout.by_ref_parameters)
            .any(|(ty, by_ref)| *ty == ScalarType::Record && !by_ref)
        {
            return Err(problem(
                func,
                "ByVal UDT parameter requires native aggregate argument lowering",
            ));
        }
        for block in &func.blocks {
            for (index, inst) in block.instrs.iter().enumerate() {
                let operands: Vec<OxOperand> = match inst {
                    OxInst::CallExtern { args, .. } => args
                        .iter()
                        .filter_map(|argument| match argument {
                            OxArg::ByVal(value) => Some(value.clone()),
                            OxArg::ByRef(place) => Some(OxOperand::Use(*place)),
                            OxArg::Omitted => None,
                        })
                        .collect(),
                    OxInst::CallNative {
                        callee: OxNativeCallee::Builtin(_),
                        args,
                        ..
                    } => args
                        .iter()
                        .filter_map(|argument| match argument {
                            OxCallArg::Operand(value) => Some(value.clone()),
                            OxCallArg::ByRef(place) => Some(OxOperand::Use(*place)),
                            _ => None,
                        })
                        .collect(),
                    OxInst::ArrayLiteral { values, .. } => values.clone(),
                    OxInst::ForEachInit { source, .. } => vec![source.clone()],
                    OxInst::Box {
                        src,
                        from: OxTy::Record(_),
                        ..
                    } => vec![src.clone()],
                    _ => Vec::new(),
                };
                for operand in operands {
                    values::validate_library_operand(&layout, &operand)
                        .map_err(|error| problem(func, error))?;
                }
                match inst {
                    OxInst::AsNew {
                        binding: calcvba_oxir::inst::OxAsNew::ProjectClass { class },
                        ..
                    }
                    | OxInst::NewObject { class, .. } => {
                        let class = &program.classes[class.0];
                        if !class.as_new_fields.is_empty()
                            || class.fields.iter().any(|field| {
                                matches!(field.ty, OxTy::Record(_) | OxTy::FixedStr(_))
                            })
                        {
                            return Err(problem(
                                func,
                                format!(
                                    "native class {} has an unsupported field initializer",
                                    class.name
                                ),
                            ));
                        }
                        pending.extend(class.initialize);
                        pending.extend(class.terminate);
                    }
                    OxInst::NewExtern { import, .. } if layout.library_news.contains(import) => {}
                    OxInst::AsNew {
                        binding: calcvba_oxir::inst::OxAsNew::ExternClass { import },
                        ..
                    } if layout.library_news.contains(import) => {}
                    OxInst::AsNew {
                        binding: calcvba_oxir::inst::OxAsNew::ComClass { .. },
                        ..
                    } => {}
                    OxInst::ComCallEarly { method, .. }
                        if program
                            .com_method(*method)
                            .is_some_and(calcvba_com::typed_slots::plan::validate) => {}
                    OxInst::ComCallLate { recv, args, .. }
                    | OxInst::CallByName {
                        object: recv, args, ..
                    } if layout.class_calls.iter().any(|(site, _, _)| site == inst) => {
                        if !matches!(
                            layout.operand_type(recv),
                            Ok(ScalarType::HostObject | ScalarType::Variant)
                        ) {
                            return Err(problem(
                                func,
                                format!(
                                    "native class receiver requires an Object carrier at block {:?} instruction {index}",
                                    block.id
                                ),
                            ));
                        }
                        for (site, procedure, _) in &layout.class_calls {
                            if site != inst {
                                continue;
                            }
                            let callee = &program.funcs[procedure.0];
                            if args.len() + 1 != callee.param_count {
                                return Err(problem(
                                    func,
                                    format!(
                                        "native class call arity at block {:?} instruction {index}",
                                        block.id
                                    ),
                                ));
                            }
                            for (argument, parameter) in
                                args.iter().zip(&callee.locals[1..callee.param_count])
                            {
                                let (operand, alias) = match argument {
                                    OxCallArg::Operand(operand) => (operand.clone(), false),
                                    OxCallArg::ByRef(place) => (
                                        OxOperand::Use(*place),
                                        parameter.param.is_some_and(|p| p.by_ref),
                                    ),
                                    _ => {
                                        return Err(problem(
                                            func,
                                            format!(
                                                "native class call argument at block {:?} instruction {index}",
                                                block.id
                                            ),
                                        ));
                                    }
                                };
                                let from = layout
                                    .operand_type(&operand)
                                    .map_err(|error| problem(func, error))?;
                                let to =
                                    calcvba_rt_abi::host_entry::native_scalar_type(&parameter.ty)
                                        .ok_or_else(|| {
                                        problem(func, "native class parameter carrier")
                                    })?;
                                if from != to
                                    && (alias
                                        || !(to == ScalarType::Variant
                                            || numeric(from) && numeric(to)))
                                {
                                    return Err(problem(
                                        func,
                                        format!(
                                            "native class argument conversion from {from:?} to {to:?} at block {:?} instruction {index}",
                                            block.id
                                        ),
                                    ));
                                }
                            }
                        }
                    }
                    OxInst::CallByName { name, calltype, .. }
                        if !(program.classes.is_empty()
                            || matches!(name, OxOperand::Const(OxConst::Str(_)))
                                && matches!(
                                    calltype,
                                    OxOperand::Const(OxConst::I16(_) | OxConst::I32(_))
                                )) =>
                    {
                        return Err(problem(
                            func,
                            format!(
                                "native dynamic project CallByName is not yet admitted at block {:?} instruction {index}",
                                block.id
                            ),
                        ));
                    }
                    OxInst::ComCallLate { recv, args, .. }
                    | OxInst::CallByName {
                        object: recv, args, ..
                    } => {
                        values::validate_library_operand(&layout, recv)
                            .map_err(|error| problem(func, error))?;
                        for argument in args {
                            match argument {
                                OxCallArg::Operand(operand)
                                | OxCallArg::Named { value: operand, .. } => {
                                    values::validate_library_operand(&layout, operand)
                                        .map_err(|error| problem(func, error))?
                                }
                                OxCallArg::ByRef(place) => values::validate_library_operand(
                                    &layout,
                                    &OxOperand::Use(*place),
                                )
                                .map_err(|error| problem(func, error))?,
                                _ => {}
                            }
                        }
                    }
                    OxInst::FieldGet { .. } | OxInst::FieldSet { .. } | OxInst::TypeOfIs { .. } => {
                    }
                    OxInst::Coerce {
                        dst,
                        src,
                        target:
                            calcvba_oxir::OxCoerceTarget::Numeric(
                                calcvba_bundle::NumericCoerceTarget::Long
                                | calcvba_bundle::NumericCoerceTarget::Double
                                | calcvba_bundle::NumericCoerceTarget::Boolean,
                            ),
                    } if numeric(
                        layout.slots[layout.slot(*dst).map_err(|e| problem(func, e))?],
                    ) && numeric(layout.operand_type(src).map_err(|e| problem(func, e))?) => {}
                    OxInst::CallProc { proc, args, .. } => {
                        if args.iter().any(|arg| {
                            !matches!(arg, OxArg::ByVal(_) | OxArg::ByRef(_) | OxArg::Omitted)
                        }) {
                            return Err(problem(
                                func,
                                format!(
                                    "omitted call argument at block {:?} instruction {index}",
                                    block.id
                                ),
                            ));
                        }
                        pending.push(*proc);
                    }
                    OxInst::CallExtern { import, .. } => {
                        let target = layout
                            .extern_calls
                            .get(import)
                            .ok_or_else(|| problem(func, "unplanned native library call"))?
                            .target;
                        if !matches!(target, LinkedTarget::Builtin(_)) {
                            return Err(problem(
                                func,
                                format!("unsupported native library call {inst:?}"),
                            ));
                        }
                    }
                    OxInst::ArrayErase { array, .. } => {
                        values::validate_library_operand(&layout, &OxOperand::Use(*array))
                            .map_err(|error| problem(func, error))?;
                    }
                    OxInst::ErrFieldGet { .. }
                    | OxInst::ErrFieldSet { .. }
                    | OxInst::ErlGet { .. }
                    | OxInst::SetErrorHandler(_)
                    | OxInst::ClearErr
                    | OxInst::SetLineNumber { .. }
                    | OxInst::NewRecord { .. }
                    | OxInst::CallNative {
                        callee: OxNativeCallee::Builtin(_),
                        ..
                    } => {}
                    OxInst::Ptr { src, kind, .. } => {
                        if layout.operand_type(src) == Ok(ScalarType::Array)
                            || (!matches!(
                                kind,
                                calcvba_oxir::PtrKind::Str | calcvba_oxir::PtrKind::Obj
                            ) && !matches!(src, OxOperand::Use(_)))
                        {
                            return Err(problem(
                                func,
                                "native pointer intrinsic requires an addressable scalar/String/Variant/UDT",
                            ));
                        }
                    }
                    OxInst::CallNative {
                        callee: OxNativeCallee::Declare { descriptor_id, .. },
                        args,
                        ..
                    } => {
                        values::validate_declare(program, &layout, *descriptor_id, args)
                            .map_err(|e| problem(func, e))?;
                    }
                    OxInst::RecordGet { record, index, .. } => {
                        let OxOperand::Use(place) = record else {
                            return Err(problem(func, "record operand"));
                        };
                        let slot = layout.slot(*place).map_err(|e| problem(func, e))?;
                        let field = &layout.record_slots[&slot].fields()[*index];
                        values::field_kind(field.kind()).map_err(|e| problem(func, e))?;
                    }
                    OxInst::RecordSet { record, index, .. } => {
                        let slot = layout.slot(*record).map_err(|e| problem(func, e))?;
                        values::field_kind(layout.record_slots[&slot].fields()[*index].kind())
                            .map_err(|e| problem(func, e))?;
                    }
                    OxInst::RecordArrayGet {
                        record,
                        index,
                        indices,
                        ..
                    } => {
                        let OxOperand::Use(place) = record else {
                            return Err(problem(func, "record array operand"));
                        };
                        let slot = layout.slot(*place).map_err(|e| problem(func, e))?;
                        values::array_field_kind(
                            layout.record_slots[&slot].fields()[*index].kind(),
                            indices.len(),
                        )
                        .map_err(|e| problem(func, e))?;
                    }
                    OxInst::RecordArraySet {
                        record,
                        index,
                        indices,
                        ..
                    } => {
                        let slot = layout.slot(*record).map_err(|e| problem(func, e))?;
                        values::array_field_kind(
                            layout.record_slots[&slot].fields()[*index].kind(),
                            indices.len(),
                        )
                        .map_err(|e| problem(func, e))?;
                    }
                    OxInst::ArrayGet { array, .. } | OxInst::Bound { array, .. } => {
                        if !matches!(
                            layout.operand_type(array),
                            Ok(ScalarType::Array | ScalarType::Variant)
                        ) {
                            return Err(problem(func, format!("Variant array receiver {inst:?}")));
                        }
                    }
                    OxInst::ArraySet { array, .. } => {
                        let slot = layout.slot(*array).map_err(|e| problem(func, e))?;
                        if !matches!(layout.slots[slot], ScalarType::Array | ScalarType::Variant) {
                            return Err(problem(func, format!("Variant array receiver {inst:?}")));
                        }
                    }
                    OxInst::ArrayRedim { dst: array, .. } => {
                        let slot = layout.slot(*array).map_err(|e| problem(func, e))?;
                        if layout.slots[slot] != ScalarType::Array {
                            return Err(problem(func, format!("Variant array receiver {inst:?}")));
                        }
                    }
                    OxInst::ArrayLiteral { .. }
                    | OxInst::ForEachInit { .. }
                    | OxInst::ForEachNext { .. }
                    | OxInst::Box { .. }
                    | OxInst::Unbox { .. }
                    | OxInst::Concat { .. }
                    | OxInst::CompareObjectIs { .. }
                    | OxInst::ValidateAssignment { .. }
                    | OxInst::VariantChanged { .. }
                    | OxInst::Assign { .. }
                    | OxInst::Coerce { .. }
                    | OxInst::Arith { .. }
                    | OxInst::Pow { .. }
                    | OxInst::Div { .. }
                    | OxInst::Neg { .. }
                    | OxInst::Compare { .. }
                    | OxInst::Logical { .. }
                    | OxInst::Not { .. }
                    | OxInst::Truthy { .. }
                    | OxInst::StmtBoundary { .. }
                    | OxInst::DrainTerminations => {}
                    _ => {
                        return Err(problem(
                            func,
                            format!(
                                "unsupported block {:?} instruction {index}: {inst:?}",
                                block.id
                            ),
                        ));
                    }
                }
            }
            if !matches!(
                block.terminator,
                OxTerminator::Jump(_)
                    | OxTerminator::Branch { .. }
                    | OxTerminator::Return
                    | OxTerminator::FaultDispatch { .. }
                    | OxTerminator::Unreachable
                    | OxTerminator::Raise { .. }
                    | OxTerminator::Resume
                    | OxTerminator::ResumeNext
                    | OxTerminator::ResumeLabel(_)
                    | OxTerminator::GoSub { .. }
                    | OxTerminator::GoSubReturn
            ) {
                return Err(problem(
                    func,
                    format!("unsupported terminator {:?}", block.terminator),
                ));
            }
        }
        let mut signature = module.make_signature();
        let ptr = module.target_config().pointer_type();
        signature.params.push(AbiParam::new(ptr));
        signature
            .params
            .extend(layout.parameters.iter().enumerate().map(|(i, t)| {
                AbiParam::new(if layout.by_ref_parameters[i] {
                    ptr
                } else {
                    machine_type(*t)
                })
            }));
        signature.params.push(AbiParam::new(ptr));
        signature.returns.push(AbiParam::new(types::I32));
        let function = module
            .declare_function(&format!("cvb$proc${}", id.0), Linkage::Local, &signature)
            .map_err(module_err)?;
        let mut offsets = Vec::new();
        let mut storage_size = 0;
        for (i, ty) in layout.slots.iter().enumerate() {
            offsets.push(storage_size);
            let size = if *ty == ScalarType::Variant {
                24
            } else if let Some(record) = layout.record_slots.get(&i) {
                record
                    .automation_shape()
                    .expect("admitted Automation layout")
                    .0
                    .next_multiple_of(8)
            } else {
                8
            };
            storage_size += size as i32;
        }
        if storage_size > 65536 {
            return Err(problem(func, "native frame exceeds 65536 bytes"));
        }
        plans.insert(
            id.0,
            NativeScalar {
                procedure: id,
                function,
                layout,
                offsets,
                storage_size: storage_size as u32,
            },
        );
    }
    let output_size = plans
        .values()
        .map(NativeScalar::result_size)
        .max()
        .unwrap_or(8)
        .max(24)
        .next_multiple_of(8) as u32;
    let class_entries: HashMap<_, _> = plans
        .values()
        .map(|plan| (plan.procedure, plan.function))
        .collect();
    for (index, plan) in &plans {
        let func = &program.funcs[*index];
        let mut context = module.make_context();
        context.func.signature = module
            .declarations()
            .get_function_decl(plan.function)
            .signature
            .clone();
        let mut frontend = FunctionBuilderContext::new();
        let depth_guard;
        let runtime;
        let depth;
        let pointers;
        {
            let mut b = FunctionBuilder::new(&mut context.func, &mut frontend);
            let entry = b.create_block();
            b.switch_to_block(entry);
            b.append_block_params_for_function_params(entry);
            let params = b.block_params(entry).to_vec();
            runtime = params[0];
            let out = *params.last().expect("typed result pointer");
            depth = b
                .ins()
                .load(types::I32, ir::MemFlagsData::new(), runtime, DEPTH_OFFSET);
            let exhausted = b.ins().icmp_imm(
                ir::condcodes::IntCC::UnsignedGreaterThanOrEqual,
                depth,
                i64::from(MAX_NATIVE_CALL_DEPTH),
            );
            let refused = b.create_block();
            let accepted = b.create_block();
            b.ins().brif(exhausted, refused, &[], accepted, &[]);
            b.switch_to_block(refused);
            let overflow = b.ins().iconst(types::I32, 28);
            depth_guard = b.ins().return_(&[overflow]);
            b.switch_to_block(accepted);
            let next_depth = b.ins().iadd_imm(depth, 1);
            b.ins()
                .store(ir::MemFlagsData::new(), next_depth, runtime, DEPTH_OFFSET);
            // All calls complete synchronously. Reuse one private result slot
            // instead of growing the frame once for each call instruction.
            let call_output = b.create_sized_stack_slot(StackSlotData::new(
                StackSlotKind::ExplicitSlot,
                output_size,
                3,
            ));
            let call_output_ptr = b.ins().stack_addr(types::I64, call_output, 0);
            let empty = b.ins().iconst(types::I64, 0);
            for offset in (0..output_size).step_by(8) {
                b.ins().stack_store(empty, call_output, offset as i32);
            }
            let slots = b.create_sized_stack_slot(StackSlotData::new(
                StackSlotKind::ExplicitSlot,
                plan.storage_size.max(8),
                3,
            ));
            pointers = b.create_sized_stack_slot(StackSlotData::new(
                StackSlotKind::ExplicitSlot,
                (plan.layout.slots.len().max(1) * 8) as u32,
                3,
            ));
            for offset in (0..plan.storage_size.max(8)).step_by(8) {
                let zero = b.ins().iconst(types::I64, 0);
                b.ins().stack_store(zero, slots, offset as i32);
            }
            let global_offsets = global_offsets(program);
            for (i, ty) in plan.layout.slots.iter().enumerate() {
                let pointer = if i >= plan.layout.global_base {
                    b.ins().iadd_imm(
                        runtime,
                        (std::mem::size_of::<NativeContext>()
                            + global_offsets[i - plan.layout.global_base])
                            as i64,
                    )
                } else if i < plan.layout.parameters.len() && plan.layout.by_ref_parameters[i] {
                    params[i + 1]
                } else {
                    if i < plan.layout.parameters.len() && !values::owning(*ty) {
                        let value = if *ty == ScalarType::Boolean {
                            b.ins().ireduce(types::I16, params[i + 1])
                        } else {
                            params[i + 1]
                        };
                        b.ins().stack_store(value, slots, plan.offsets[i]);
                    }
                    b.ins().stack_addr(types::I64, slots, plan.offsets[i])
                };
                b.ins().stack_store(pointer, pointers, (i * 8) as i32);
            }
            let values = values::Values {
                procedure: plan.procedure,
                lower: ScalarLower {
                    layout: &plan.layout,
                    slots: PhysicalSlots::Native(pointers),
                    runtime,
                    call_conv: module.target_config().default_call_conv,
                    fault_target: std::cell::Cell::new(None),
                },
                functions: value_helpers,
                program,
                class_entries: &class_entries,
                as_new: program
                    .funcs
                    .iter()
                    .flat_map(|function| function.blocks.iter().flat_map(|block| &block.instrs))
                    .filter_map(|instruction| {
                        if let OxInst::AsNew {
                            place: place @ OxPlace::Global(_),
                            binding,
                        } = instruction
                        {
                            Some((*place, binding.clone()))
                        } else {
                            None
                        }
                    })
                    .chain(
                        func.blocks
                            .iter()
                            .flat_map(|block| &block.instrs)
                            .filter_map(|instruction| {
                                if let OxInst::AsNew { place, binding } = instruction {
                                    Some((*place, binding.clone()))
                                } else {
                                    None
                                }
                            }),
                    )
                    .collect(),
                iterator_slots: func
                    .blocks
                    .iter()
                    .flat_map(|block| &block.instrs)
                    .filter_map(|instruction| {
                        if let OxInst::ForEachInit { iter, .. } = instruction {
                            Some(plan.layout.slot(*iter).unwrap())
                        } else {
                            None
                        }
                    })
                    .collect(),
            };
            let lower = &values.lower;
            let source = func
                .name
                .split_once("::")
                .map_or(program.unit_name.as_str(), |(unit, _)| unit);
            let source_text = values::data(module, &mut b, source.as_bytes().to_vec(), 1);
            let source_len = b.ins().iconst(types::I64, source.len() as i64);
            values.call(
                module,
                &mut b,
                "cvb_frame_enter",
                &[runtime, source_text, source_len],
            );
            for (i, ty) in plan.layout.parameters.iter().enumerate() {
                if values::owning(*ty) && !plan.layout.by_ref_parameters[i] {
                    let dst = values.address(&mut b, i);
                    let src = if matches!(ty, ScalarType::Variant | ScalarType::Record) {
                        params[i + 1]
                    } else {
                        let slot = b.create_sized_stack_slot(StackSlotData::new(
                            StackSlotKind::ExplicitSlot,
                            8,
                            3,
                        ));
                        b.ins().stack_store(params[i + 1], slot, 0);
                        b.ins().stack_addr(types::I64, slot, 0)
                    };
                    if *ty == ScalarType::Record {
                        values.copy_record(module, &mut b, i, dst, src);
                    } else {
                        values.copy(
                            module,
                            &mut b,
                            dst,
                            values.slot_kind(i),
                            src,
                            values.slot_kind(i),
                        );
                    }
                }
            }
            let blocks: HashMap<_, _> = func
                .blocks
                .iter()
                .map(|block| (block.id, b.create_block()))
                .collect();
            b.ins().jump(blocks[&func.entry], &[]);
            for block in &func.blocks {
                b.switch_to_block(blocks[&block.id]);
                let fault_bridge = block.fault_target.map(|_| {
                    let bridge = b.create_block();
                    b.append_block_param(bridge, types::I32);
                    bridge
                });
                lower.fault_target.set(fault_bridge);
                for (instruction_index, inst) in block.instrs.iter().enumerate() {
                    let foreign_state = foreign_states
                        .get(&(plan.procedure.0, block.id.0, instruction_index))
                        .copied();
                    let indexed_member = calcvba_rt_abi::host_entry::index_member_call(inst)
                        .filter(|_| {
                            plan.layout
                                .class_calls
                                .iter()
                                .any(|(site, _, _)| site == inst)
                        });
                    if indexed_member.is_none()
                        && values.instruction(module, &mut b, inst, foreign_state)
                    {
                        continue;
                    }
                    if let OxInst::CallExtern { dst, import, args } = inst {
                        let resolved = calcvba_rt_abi::library::resolve_import(program, *import)
                            .map_err(|error| problem(func, error))?;
                        let signature = &plan.layout.extern_calls[import].signature;
                        let specialized = matches!(args.as_slice(), [OxArg::ByVal(operand)] if plan.layout.operand_type(operand).is_ok_and(numeric) && (conversion_target(resolved.implementation).is_some() || library_helpers.contains_key(&resolved.implementation) && signature.parameters == [ScalarType::Double]));
                        if !specialized {
                            values.library_call(
                                module,
                                &mut b,
                                resolved.index,
                                resolved.string_typed_alias,
                                args,
                                *dst,
                            );
                            continue;
                        }
                    }
                    if let OxInst::CallExtern { dst, import, args } = inst
                        && let LinkedTarget::Builtin(native) =
                            plan.layout.extern_calls[import].target
                        && let [OxArg::ByVal(src)] = args.as_slice()
                        && (matches!(
                            native,
                            NativeImplId::Len | NativeImplId::CStr | NativeImplId::VarType
                        ) || (conversion_target(native).is_some()
                            && (!plan.layout.operand_type(src).is_ok_and(basic_numeric)
                                || conversion_target(native).is_some_and(extended_numeric))))
                    {
                        let input = values.cell(module, &mut b, src);
                        let zero = b.ins().iconst(types::I64, 0);
                        for offset in [0, 8, 16] {
                            b.ins().stack_store(zero, call_output, offset);
                        }
                        let to = conversion_target(native).unwrap_or(
                            if matches!(native, NativeImplId::Len | NativeImplId::VarType) {
                                ScalarType::Long
                            } else {
                                ScalarType::String
                            },
                        );
                        if native == NativeImplId::VarType {
                            let kind = b.ins().iconst(
                                types::I32,
                                i64::from(values::kind_of_operand(&plan.layout, src)),
                            );
                            values.call(
                                module,
                                &mut b,
                                "cvb_value_type",
                                &[call_output_ptr, input, kind],
                            );
                        } else if native == NativeImplId::Len {
                            values.call(
                                module,
                                &mut b,
                                "cvb_string_len",
                                &[input, call_output_ptr],
                            );
                        } else {
                            values.copy(
                                module,
                                &mut b,
                                call_output_ptr,
                                values::kind(to),
                                input,
                                values::kind(plan.layout.operand_type(src).unwrap()),
                            );
                        }
                        if let Some(dst) = dst {
                            let index = plan.layout.slot(*dst).unwrap();
                            let address = values.address(&mut b, index);
                            values.copy(
                                module,
                                &mut b,
                                address,
                                values.slot_kind(index),
                                call_output_ptr,
                                values::kind(to),
                            );
                        }
                        if values::owning(to) {
                            let kind = b.ins().iconst(types::I32, i64::from(values::kind(to)));
                            values.call(
                                module,
                                &mut b,
                                "cvb_value_clear",
                                &[call_output_ptr, kind],
                            );
                        }
                        continue;
                    }
                    // Numeric library conversions have polymorphic VBA inputs,
                    // but their bound result type is fixed. Specialize each
                    // numeric call site to unboxed conversion lowering. A
                    // discarded result must still evaluate and report errors.
                    let conversion = if let OxInst::CallExtern { dst, import, args } = inst
                        && let LinkedTarget::Builtin(native) =
                            plan.layout.extern_calls[import].target
                        && let Some(target) = conversion_target(native)
                        && let [OxArg::ByVal(src)] = args.as_slice()
                    {
                        Some((*dst, src, target))
                    } else if let OxInst::Coerce { dst, src, .. } = inst
                        && calcvba_rt_abi::host_entry::value::needs_helper(inst, &plan.layout)
                    {
                        let to =
                            plan.layout.slots[plan.layout.slot(*dst).expect("admitted result")];
                        Some((Some(*dst), src, to))
                    } else {
                        None
                    };
                    if let Some((dst, src, to)) = conversion {
                        let from = plan.layout.operand_type(src).expect("admitted conversion");
                        let input = lower.operand(&mut b, src, from);
                        let value = match (from, to) {
                            (from, to) if from == to => input,
                            (ScalarType::Double, ScalarType::Long) => {
                                let helper = module.declare_func_in_func(long_from_f64, b.func);
                                let call = b.ins().call(helper, &[runtime, input, call_output_ptr]);
                                let status = b.inst_results(call)[0];
                                lower.check_status(&mut b, status);
                                b.ins().stack_load(types::I32, call_output, 0)
                            }
                            (ScalarType::Long | ScalarType::Boolean, ScalarType::Double) => {
                                b.ins().fcvt_from_sint(types::F64, input)
                            }
                            (ScalarType::Boolean, ScalarType::Long) => input,
                            (ScalarType::Double | ScalarType::Long, ScalarType::Boolean) => {
                                let truth = if from == ScalarType::Double {
                                    lower.finite(&mut b, input);
                                    let zero = b.ins().f64const(0.0);
                                    b.ins().fcmp(ir::condcodes::FloatCC::NotEqual, input, zero)
                                } else {
                                    b.ins().icmp_imm(ir::condcodes::IntCC::NotEqual, input, 0)
                                };
                                let truth = b.ins().uextend(types::I32, truth);
                                b.ins().ineg(truth)
                            }
                            _ => unreachable!("admitted numeric conversion"),
                        };
                        if let Some(dst) = dst {
                            lower.store_slot(&mut b, plan.layout.slot(dst).unwrap(), to, value);
                        }
                        continue;
                    }
                    if let OxInst::ComCallLate {
                        dst, recv, args, ..
                    }
                    | OxInst::CallByName {
                        dst,
                        object: recv,
                        args,
                        ..
                    } = indexed_member.as_ref().unwrap_or(inst)
                    {
                        let receiver = values.cell(module, &mut b, recv);
                        let slot = b.create_sized_stack_slot(StackSlotData::new(
                            StackSlotKind::ExplicitSlot,
                            8,
                            3,
                        ));
                        let address = b.ins().stack_addr(types::I64, slot, 0);
                        let kind = b.ins().iconst(
                            types::I32,
                            i64::from(values::kind_of_operand(&plan.layout, recv)),
                        );
                        values.call(
                            module,
                            &mut b,
                            "cvb_member_route",
                            &[runtime, receiver, kind, address],
                        );
                        let route = b.ins().stack_load(types::I32, slot, 0);
                        let end = b.create_block();
                        for (site, procedure, class) in &plan.layout.class_calls {
                            if site != inst {
                                continue;
                            }
                            let selected = b.create_block();
                            let next = b.create_block();
                            let mut matches = b.ins().icmp_imm(
                                ir::condcodes::IntCC::Equal,
                                route,
                                class.0 as i64,
                            );
                            if let OxInst::ArraySet { value, .. } = inst {
                                let input = values.cell(module, &mut b, value);
                                let kind = values::kind_of_operand(&plan.layout, value);
                                let object = if kind == 12 {
                                    let tag =
                                        b.ins().load(types::I16, ir::MemFlagsData::new(), input, 0);
                                    b.ins().icmp_imm(ir::condcodes::IntCC::Equal, tag, 9)
                                } else {
                                    b.ins().iconst(types::I8, i64::from(kind == 9))
                                };
                                let is_set =
                                    program.classes[class.0].methods.iter().any(|method| {
                                        method.proc == *procedure
                                            && method.kind
                                                == calcvba_bundle::ProjectMemberKind::PropertySet
                                    });
                                let kind_matches = if is_set {
                                    object
                                } else {
                                    b.ins().icmp_imm(ir::condcodes::IntCC::Equal, object, 0)
                                };
                                matches = b.ins().band(matches, kind_matches);
                            }
                            b.ins().brif(matches, selected, &[], next, &[]);
                            b.switch_to_block(selected);
                            let callee = &plans[&procedure.0];
                            if args.len() + 1 != callee.layout.parameters.len() {
                                let status = b.ins().iconst(types::I32, 450);
                                lower.check_status(&mut b, status);
                                b.ins().jump(end, &[]);
                                b.switch_to_block(next);
                                continue;
                            }
                            let mut arguments = vec![OxArg::ByVal(recv.clone())];
                            for (index, arg) in args.iter().enumerate() {
                                arguments.push(match arg {
                                    OxCallArg::Operand(operand) => OxArg::ByVal(operand.clone()),
                                    OxCallArg::ByRef(place)
                                        if callee.layout.by_ref_parameters[index + 1] =>
                                    {
                                        OxArg::ByRef(*place)
                                    }
                                    OxCallArg::ByRef(place) => OxArg::ByVal(OxOperand::Use(*place)),
                                    _ => unreachable!("admitted class argument"),
                                });
                            }
                            emit_native_call(
                                module,
                                &mut b,
                                &values,
                                &plans,
                                *dst,
                                &arguments,
                                &callee.layout.parameters,
                                &callee.layout.by_ref_parameters,
                                callee.layout.result,
                                callee.function,
                                output_size,
                                call_output,
                                call_output_ptr,
                            );
                            b.ins().jump(end, &[]);
                            b.switch_to_block(next);
                        }
                        if indexed_member.is_some() {
                            assert!(values.instruction(module, &mut b, inst, foreign_state));
                        } else {
                            values.member_call(module, &mut b, inst);
                        }
                        b.ins().jump(end, &[]);
                        b.switch_to_block(end);
                        continue;
                    }
                    let call_plan = match inst {
                        OxInst::CallProc { dst, proc, args } => {
                            let callee = &plans[&proc.0];
                            Some((
                                *dst,
                                args,
                                &callee.layout.parameters,
                                &callee.layout.by_ref_parameters,
                                callee.layout.result,
                                callee.function,
                            ))
                        }
                        OxInst::CallExtern { dst, import, args } => {
                            let call = &plan.layout.extern_calls[import];
                            let LinkedTarget::Builtin(native) = call.target else {
                                unreachable!("admitted native builtin")
                            };
                            Some((
                                *dst,
                                args,
                                &call.signature.parameters,
                                &call.signature.by_ref,
                                call.signature.result.expect("numeric builtin result"),
                                library_helpers[&native],
                            ))
                        }
                        _ => None,
                    };
                    if let Some((dst, args, parameters, by_ref, result_type, function)) = call_plan
                    {
                        emit_native_call(
                            module,
                            &mut b,
                            &values,
                            &plans,
                            dst,
                            args,
                            parameters,
                            by_ref,
                            result_type,
                            function,
                            output_size,
                            call_output,
                            call_output_ptr,
                        );
                    } else {
                        lower.instruction(&mut b, inst);
                    }
                }
                match &block.terminator {
                    OxTerminator::Jump(next) => {
                        b.ins().jump(blocks[next], &[]);
                    }
                    OxTerminator::Branch {
                        cond,
                        then_blk,
                        else_blk,
                    } => {
                        let condition = lower.operand(&mut b, cond, ScalarType::Boolean);
                        b.ins()
                            .brif(condition, blocks[then_blk], &[], blocks[else_blk], &[]);
                    }
                    OxTerminator::Return => {
                        if plan.layout.has_result {
                            if values::owning(plan.layout.result) {
                                let result = values.address(&mut b, plan.layout.return_slot);
                                if plan.layout.result == ScalarType::Record {
                                    values.copy_record(
                                        module,
                                        &mut b,
                                        plan.layout.return_slot,
                                        out,
                                        result,
                                    );
                                } else {
                                    values.copy(
                                        module,
                                        &mut b,
                                        out,
                                        values.slot_kind(plan.layout.return_slot),
                                        result,
                                        values.slot_kind(plan.layout.return_slot),
                                    );
                                }
                            } else {
                                let result = lower.operand(
                                    &mut b,
                                    &OxOperand::Use(OxPlace::Local(calcvba_oxir::LocalId(
                                        plan.layout.return_slot,
                                    ))),
                                    plan.layout.result,
                                );
                                b.ins().store(ir::MemFlagsData::new(), result, out, 0);
                            }
                        }
                        let zero = b.ins().iconst(types::I32, 0);
                        b.ins().return_(&[zero]);
                    }
                    OxTerminator::Raise {
                        number,
                        source,
                        description,
                        help_file,
                        help_context,
                        inherit,
                    } => {
                        let number_cell = b.create_sized_stack_slot(StackSlotData::new(
                            StackSlotKind::ExplicitSlot,
                            8,
                            3,
                        ));
                        let number_out = b.ins().stack_addr(types::I64, number_cell, 0);
                        let number_src = values.cell(module, &mut b, number);
                        values.copy(
                            module,
                            &mut b,
                            number_out,
                            3,
                            number_src,
                            values::kind(plan.layout.operand_type(number).unwrap()),
                        );
                        let number = b.ins().stack_load(types::I32, number_cell, 0);
                        let mut text = Vec::new();
                        let mut flags = if *inherit { 0x10 } else { 0 };
                        for (index, operand) in [source, description, help_file, help_context]
                            .iter()
                            .enumerate()
                        {
                            if let Some(operand) = operand {
                                let cell = values.cell(module, &mut b, operand);
                                let kind = b.ins().iconst(
                                    types::I32,
                                    i64::from(values::kind(
                                        plan.layout.operand_type(operand).unwrap(),
                                    )),
                                );
                                text.extend([cell, kind]);
                                flags |= 1 << index;
                            } else {
                                let empty = b.ins().iconst(types::I64, 0);
                                let kind = b.ins().iconst(types::I32, 0);
                                text.extend([empty, kind]);
                            }
                        }
                        let flags = b.ins().iconst(types::I32, flags);
                        let helper =
                            module.declare_func_in_func(value_helpers["cvb_raise_dynamic"], b.func);
                        let call = b.ins().call(
                            helper,
                            &[
                                runtime, number, text[0], text[1], text[2], text[3], text[4],
                                text[5], text[6], text[7], flags,
                            ],
                        );
                        let status = b.inst_results(call)[0];
                        lower.return_status(&mut b, status);
                    }
                    OxTerminator::FaultDispatch {
                        resume,
                        resume_next,
                    } => {
                        lower.fault_target.set(None);
                        emit_error_transfer(
                            module,
                            &mut b,
                            &values,
                            &blocks,
                            func,
                            "cvb_fault_route",
                            (resume.0 as u32, resume_next.0 as u32),
                        );
                    }
                    OxTerminator::GoSub { target, ret } => {
                        let push = b.ins().iconst(types::I32, 0);
                        let ret = b.ins().iconst(types::I32, ret.0 as i64);
                        let out = b.ins().iconst(types::I64, 0);
                        values.call(module, &mut b, "cvb_gosub", &[runtime, push, ret, out]);
                        b.ins().jump(blocks[target], &[]);
                    }
                    OxTerminator::GoSubReturn => {
                        emit_error_transfer(
                            module,
                            &mut b,
                            &values,
                            &blocks,
                            func,
                            "cvb_gosub",
                            (1, 0),
                        );
                    }
                    OxTerminator::Resume
                    | OxTerminator::ResumeNext
                    | OxTerminator::ResumeLabel(_) => {
                        let (kind, label) = match block.terminator {
                            OxTerminator::Resume => (0, 0),
                            OxTerminator::ResumeNext => (1, 0),
                            OxTerminator::ResumeLabel(label) => (2, label.0 as u32),
                            _ => unreachable!(),
                        };
                        emit_error_transfer(
                            module,
                            &mut b,
                            &values,
                            &blocks,
                            func,
                            "cvb_resume",
                            (kind, label),
                        );
                    }
                    _ => {
                        let status = b.ins().iconst(types::I32, 5);
                        b.ins().return_(&[status]);
                    }
                }

                if let (Some(bridge), Some(target)) = (fault_bridge, block.fault_target) {
                    lower.fault_target.set(None);
                    b.switch_to_block(bridge);
                    let number = b.block_params(bridge)[0];
                    let action = b.ins().iconst(types::I32, 6);
                    values.call(
                        module,
                        &mut b,
                        "cvb_error_control",
                        &[runtime, action, number],
                    );
                    b.ins().jump(blocks[&target], &[]);
                }
            }
            b.seal_all_blocks();
            b.finalize();
        }
        // ScalarLower emits early status returns too. Restore the depth on
        // every accepted exit, including arithmetic and helper failures.
        let returns: Vec<_> = context
            .func
            .layout
            .blocks()
            .flat_map(|block| context.func.layout.block_insts(block))
            .filter(|inst| {
                *inst != depth_guard && context.func.dfg.insts[*inst].opcode() == ir::Opcode::Return
            })
            .collect();
        use cranelift_codegen::cursor::{Cursor, FuncCursor};
        let cleanup =
            module.declare_func_in_func(value_helpers["cvb_value_clear"], &mut context.func);
        let record_cleanup =
            module.declare_func_in_func(value_helpers["cvb_record_clear"], &mut context.func);
        let frame_exit =
            module.declare_func_in_func(value_helpers["cvb_frame_exit"], &mut context.func);
        let iterator_cleanup =
            module.declare_func_in_func(value_helpers["cvb_iter_clear"], &mut context.func);
        let iterator_slots: BTreeSet<_> = func
            .blocks
            .iter()
            .flat_map(|block| &block.instrs)
            .filter_map(|instruction| {
                if let OxInst::ForEachInit { iter, .. } = instruction {
                    Some(plan.layout.slot(*iter).unwrap())
                } else {
                    None
                }
            })
            .collect();
        let mut record_globals = HashMap::new();
        for (i, layout) in &plan.layout.record_slots {
            let id = module
                .declare_anonymous_data(false, false)
                .map_err(module_err)?;
            let mut description = cranelift_module::DataDescription::new();
            description.define(
                values::record_descriptor(layout)
                    .expect("admitted record")
                    .into_boxed_slice(),
            );
            description.set_align(4);
            module.define_data(id, &description).map_err(module_err)?;
            record_globals.insert(*i, module.declare_data_in_func(id, &mut context.func));
        }
        // One shared cleanup block prevents a copy of every owning local's
        // destructor at every arithmetic/helper error branch.
        let exit = context.func.dfg.make_block();
        let status = context.func.dfg.append_block_param(exit, types::I32);
        context.func.layout.append_block(exit);
        for inst in returns {
            let returned_status = context.func.dfg.inst_args(inst)[0];
            context
                .func
                .replace(inst)
                .jump(exit, &[returned_status.into()]);
        }
        {
            let mut cursor = FuncCursor::new(&mut context.func);
            cursor.goto_bottom(exit);
            for slot in iterator_slots {
                let address = cursor
                    .ins()
                    .stack_load(types::I64, pointers, (slot * 8) as i32);
                cursor.ins().call(iterator_cleanup, &[runtime, address]);
            }
            for (i, ty) in plan.layout.slots[..plan.layout.global_base]
                .iter()
                .enumerate()
            {
                if values::owning(*ty)
                    && !plan
                        .layout
                        .by_ref_parameters
                        .get(i)
                        .copied()
                        .unwrap_or(false)
                {
                    let pointer = cursor
                        .ins()
                        .stack_load(types::I64, pointers, (i * 8) as i32);
                    let kind = cursor
                        .ins()
                        .iconst(types::I32, i64::from(values::kind(*ty)));
                    if let Some(global) = record_globals.get(&i) {
                        let descriptor = cursor.ins().global_value(types::I64, *global);
                        cursor.ins().call(record_cleanup, &[pointer, descriptor]);
                    } else {
                        cursor.ins().call(cleanup, &[pointer, kind]);
                    }
                }
            }
            cursor.ins().call(frame_exit, &[runtime]);
            cursor
                .ins()
                .store(ir::MemFlagsData::new(), depth, runtime, DEPTH_OFFSET);
            cursor.ins().return_(&[status]);
        }
        module
            .define_function(plan.function, &mut context)
            .map_err(|e| JitError::Compile(format!("native procedure {}: {e:?}", func.name)))?;
    }
    Ok(plans.into_values().collect())
}

fn global_offsets(program: &OxProgram) -> Vec<usize> {
    let mut offsets = Vec::new();
    let mut size = 0;
    for global in &program.globals {
        offsets.push(size);
        size += if global.ty == OxTy::Variant {
            24
        } else if let OxTy::Record(id) = global.ty {
            vba_record_layout_for_fields(&program.record_layouts[id.0])
                .expect("admitted global record")
                .automation_shape()
                .expect("admitted Automation layout")
                .0
                .next_multiple_of(8)
        } else {
            8
        };
    }
    offsets.push(size);
    offsets
}

#[allow(clippy::too_many_arguments)]
fn emit_native_call(
    module: &mut impl Module,
    b: &mut FunctionBuilder<'_>,
    values: &values::Values<'_>,
    plans: &BTreeMap<usize, NativeScalar>,
    dst: Option<OxPlace>,
    args: &[OxArg],
    parameters: &[ScalarType],
    by_ref: &[bool],
    result_type: ScalarType,
    function: ClifFuncId,
    output_size: u32,
    call_output: ir::StackSlot,
    call_output_ptr: Value,
) {
    let callee = plans.values().find(|plan| plan.function == function);
    let lower = &values.lower;
    let mut call_args = vec![values.lower.runtime];
    let mut copies = Vec::new();
    for (argument_index, (ty, by_ref)) in parameters.iter().zip(by_ref).enumerate() {
        let arg = args.get(argument_index).unwrap_or(&OxArg::Omitted);
        if matches!(arg, OxArg::Omitted) {
            let slot =
                b.create_sized_stack_slot(StackSlotData::new(StackSlotKind::ExplicitSlot, 24, 3));
            let empty = b.ins().iconst(types::I64, 0);
            for offset in [0, 8, 16] {
                b.ins().stack_store(empty, slot, offset);
            }
            let tag = b.ins().iconst(types::I32, 10);
            let missing = b.ins().iconst(types::I32, i64::from(0x80020004u32 as i32));
            b.ins().stack_store(tag, slot, 0);
            b.ins().stack_store(missing, slot, 8);
            let pointer = b.ins().stack_addr(types::I64, slot, 0);
            call_args.push(pointer);
            copies.push((pointer, ScalarType::Variant, None));
            continue;
        }
        let operand = match arg {
            OxArg::ByVal(operand) => operand.clone(),
            OxArg::ByRef(place) => OxOperand::Use(*place),
            _ => unreachable!("admitted native argument"),
        };
        let value = if callee
            .is_some_and(|callee| callee.layout.param_array == Some(argument_index))
        {
            values.cell(module, b, &operand)
        } else if !by_ref
            && !matches!(ty, ScalarType::Record | ScalarType::Array)
            && values.lower.layout.operand_type(&operand) != Ok(*ty)
        {
            let slot =
                b.create_sized_stack_slot(StackSlotData::new(StackSlotKind::ExplicitSlot, 24, 3));
            let empty = b.ins().iconst(types::I64, 0);
            for offset in [0, 8, 16] {
                b.ins().stack_store(empty, slot, offset);
            }
            let pointer = b.ins().stack_addr(types::I64, slot, 0);
            let source = values.cell(module, b, &operand);
            values.copy(
                module,
                b,
                pointer,
                values::kind(*ty),
                source,
                values::kind_of_operand(values.lower.layout, &operand),
            );
            copies.push((pointer, *ty, None));
            if *ty == ScalarType::Variant {
                pointer
            } else {
                match ty {
                    ScalarType::Byte => {
                        let value = b.ins().load(types::I8, ir::MemFlagsData::new(), pointer, 0);
                        b.ins().uextend(types::I32, value)
                    }
                    ScalarType::Integer | ScalarType::Boolean => {
                        let value = b
                            .ins()
                            .load(types::I16, ir::MemFlagsData::new(), pointer, 0);
                        b.ins().sextend(types::I32, value)
                    }
                    _ => b
                        .ins()
                        .load(machine_type(*ty), ir::MemFlagsData::new(), pointer, 0),
                }
            }
        } else {
            match (arg, by_ref) {
                (OxArg::ByRef(place), true) => {
                    values.address(b, values.lower.layout.slot(*place).unwrap())
                }
                (OxArg::ByVal(op), true) => {
                    let record = if *ty == ScalarType::Record {
                        Some(callee.unwrap().layout.record_slots[&argument_index].as_ref())
                    } else {
                        None
                    };
                    let size = if let Some(record) = record {
                        record
                            .automation_shape()
                            .expect("admitted Automation layout")
                            .0
                            .next_multiple_of(8) as u32
                    } else if *ty == ScalarType::Variant {
                        24
                    } else {
                        8
                    };
                    let slot = b.create_sized_stack_slot(StackSlotData::new(
                        StackSlotKind::ExplicitSlot,
                        size,
                        3,
                    ));
                    let empty = b.ins().iconst(types::I64, 0);
                    for offset in (0..size).step_by(8) {
                        b.ins().stack_store(empty, slot, offset as i32);
                    }
                    let pointer = b.ins().stack_addr(types::I64, slot, 0);
                    let source = values.cell(module, b, op);
                    let descriptor = if let Some(record) = record {
                        Some(values.copy_record_layout(module, b, record, pointer, source))
                    } else {
                        values.copy(
                            module,
                            b,
                            pointer,
                            values::kind(*ty),
                            source,
                            values::kind_of_operand(values.lower.layout, op),
                        );
                        None
                    };
                    if values::owning(*ty) {
                        copies.push((pointer, *ty, descriptor));
                    }
                    pointer
                }
                (OxArg::ByVal(op), false) => values.operand(module, b, op, *ty),
                (OxArg::ByRef(place), false) => {
                    values.operand(module, b, &OxOperand::Use(*place), *ty)
                }
                _ => unreachable!("admitted native call argument"),
            }
        };
        call_args.push(value);
    }
    let empty = b.ins().iconst(types::I64, 0);
    for offset in (0..output_size).step_by(8) {
        b.ins().stack_store(empty, call_output, offset as i32);
    }
    call_args.push(call_output_ptr);
    let target = module.declare_func_in_func(function, b.func);
    let call = b.ins().call(target, &call_args);
    let status = b.inst_results(call)[0];
    for (pointer, ty, descriptor) in copies {
        if let Some(descriptor) = descriptor {
            let clear = module.declare_func_in_func(values.functions["cvb_record_clear"], b.func);
            b.ins().call(clear, &[pointer, descriptor]);
        } else {
            let kind = b.ins().iconst(types::I32, i64::from(values::kind(ty)));
            let clear = module.declare_func_in_func(values.functions["cvb_value_clear"], b.func);
            b.ins().call(clear, &[pointer, kind]);
        }
    }
    lower.check_status(b, status);
    if let Some(dst) = dst {
        let index = values.lower.layout.slot(dst).unwrap();
        if result_type == ScalarType::Record {
            let address = values.address(b, index);
            values.copy_record(module, b, index, address, call_output_ptr);
        } else if values::owning(result_type) || values.lower.layout.slots[index] != result_type {
            let address = values.address(b, index);
            let from = if let Some(callee) = callee {
                values::kind(callee.layout.result)
            } else {
                values::kind(result_type)
            };
            values.copy(
                module,
                b,
                address,
                values.slot_kind(index),
                call_output_ptr,
                from,
            );
        } else {
            let result = b
                .ins()
                .stack_load(machine_type(result_type), call_output, 0);
            lower.store_slot(b, index, result_type, result);
        }
    }
    if result_type == ScalarType::Record {
        let callee = callee.unwrap();
        values.clear_record_layout(
            module,
            b,
            &callee.layout.record_slots[&callee.layout.return_slot],
            call_output_ptr,
        );
    } else if values::owning(result_type) {
        let kind = b
            .ins()
            .iconst(types::I32, i64::from(values::kind(result_type)));
        values.call(module, b, "cvb_value_clear", &[call_output_ptr, kind]);
    }
}

pub fn native_state_size(program: &OxProgram) -> usize {
    std::mem::size_of::<NativeContext>() + global_offsets(program).last().copied().unwrap_or(0)
}

pub struct NativeRuntime {
    pub context: DataId,
    pub enter: ClifFuncId,
    pub record_error: ClifFuncId,
    pub get_last_error: ClifFuncId,
    pub initializer: Option<ClifFuncId>,
    pub values: HashMap<&'static str, ClifFuncId>,
    pub foreign_states: Vec<DataId>,
}

fn abi_type(ty: ScalarType) -> ir::Type {
    if matches!(ty, ScalarType::Boolean | ScalarType::Integer) {
        types::I16
    } else if ty == ScalarType::Byte {
        types::I8
    } else {
        machine_type(ty)
    }
}

fn zero(b: &mut FunctionBuilder<'_>, ty: ScalarType) -> Value {
    if matches!(ty, ScalarType::Double | ScalarType::Date) {
        b.ins().f64const(0.0)
    } else if ty == ScalarType::Single {
        b.ins().f32const(0.0)
    } else {
        b.ins().iconst(abi_type(ty), 0)
    }
}

/// Each ABI wrapper calls its compiled procedure directly, after the owning
/// thread check and the once-only initializer. Refused threads never touch VBA
/// state, ByRef cells or result out pointers.
pub struct NativeWrapperOptions<'a> {
    pub hresult: bool,
    pub string_encoding: u32,
    pub source: &'a str,
    /// The first ABI parameter is an Automation wrapper whose second word
    /// retains the native project instance. Only ComExe emits this wrapper.
    pub com_receiver: bool,
}

pub fn compile_native_wrapper(
    module: &mut impl Module,
    name: &str,
    body: &NativeScalar,
    runtime: &NativeRuntime,
    options: NativeWrapperOptions<'_>,
) -> Result<ClifFuncId, JitError> {
    let NativeWrapperOptions {
        hresult,
        string_encoding,
        source,
        com_receiver,
    } = options;
    let aggregate = body.layout.has_result
        && matches!(body.layout.result, ScalarType::Variant | ScalarType::Record);
    let register_record = body.register_record_result();
    let result_abi = register_record.unwrap_or_else(|| abi_type(body.layout.result));
    let hidden = aggregate && register_record.is_none() && !hresult;
    let mut signature = module.make_signature();
    if hidden {
        signature.params.push(AbiParam::new(types::I64));
    }
    for (ty, by_ref) in body
        .layout
        .parameters
        .iter()
        .zip(&body.layout.by_ref_parameters)
    {
        signature.params.push(AbiParam::new(if *by_ref {
            types::I64
        } else {
            abi_type(*ty)
        }));
    }
    if hresult && body.layout.has_result {
        signature.params.push(AbiParam::new(types::I64));
    }
    if hresult {
        signature.returns.push(AbiParam::new(types::I32));
    } else if hidden {
        signature.returns.push(AbiParam::new(types::I64));
    } else if body.layout.has_result {
        signature.returns.push(AbiParam::new(result_abi));
    }
    let id = module
        .declare_function(name, Linkage::Export, &signature)
        .map_err(module_err)?;
    let source_data = module
        .declare_anonymous_data(false, false)
        .map_err(module_err)?;
    let mut description = cranelift_module::DataDescription::new();
    description.define(source.as_bytes().to_vec().into_boxed_slice());
    module
        .define_data(source_data, &description)
        .map_err(module_err)?;
    let mut context = module.make_context();
    context.func.signature = signature;
    let mut frontend = FunctionBuilderContext::new();
    {
        let mut b = FunctionBuilder::new(&mut context.func, &mut frontend);
        let entry = b.create_block();
        b.switch_to_block(entry);
        b.append_block_params_for_function_params(entry);
        let all_params = b.block_params(entry).to_vec();
        let mut params = all_params
            [usize::from(hidden)..usize::from(hidden) + body.layout.parameters.len()]
            .to_vec();
        let caller_out = if hidden {
            Some(all_params[0])
        } else if hresult && body.layout.has_result {
            all_params.last().copied()
        } else {
            None
        };
        let foreign = b.create_block();
        let accepted = b.create_block();
        let finish = b.create_block();
        b.append_block_param(finish, types::I32);
        let helper = module.declare_func_in_func(runtime.enter, b.func);
        let call = b.ins().call(helper, &[]);
        let status = b.inst_results(call)[0];
        b.ins().brif(status, foreign, &[], accepted, &[]);
        b.switch_to_block(foreign);
        if hresult {
            let error = b.ins().iconst(
                types::I32,
                i64::from(calcvba_rt_abi::native_export::E_WRONG_THREAD),
            );
            b.ins().return_(&[error]);
        } else if body.layout.has_result {
            let value = if register_record.is_some() {
                b.ins().iconst(result_abi, 0)
            } else {
                zero(&mut b, body.layout.result)
            };
            b.ins().return_(&[value]);
        } else {
            b.ins().return_(&[]);
        }
        b.switch_to_block(accepted);
        if com_receiver {
            params[0] = b
                .ins()
                .load(types::I64, ir::MemFlagsData::new(), params[0], 8);
        }
        let state = module.declare_data_in_func(runtime.context, b.func);
        let state = b.ins().global_value(types::I64, state);
        let default_source = module.declare_data_in_func(source_data, b.func);
        let default_source = b.ins().global_value(types::I64, default_source);
        let source_len = b.ins().iconst(types::I64, source.len() as i64);
        b.ins().store(
            ir::MemFlagsData::new(),
            default_source,
            state,
            std::mem::offset_of!(NativeContext, error_source) as i32,
        );
        b.ins().store(
            ir::MemFlagsData::new(),
            source_len,
            state,
            std::mem::offset_of!(NativeContext, error_source_len) as i32,
        );
        let old_array_cells = b.ins().load(
            types::I64,
            ir::MemFlagsData::new(),
            state,
            std::mem::offset_of!(NativeContext, string_array_cells) as i32,
        );
        let old_array_count = b.ins().load(
            types::I64,
            ir::MemFlagsData::new(),
            state,
            std::mem::offset_of!(NativeContext, string_array_count) as i32,
        );
        let old_array_encoding = b.ins().load(
            types::I32,
            ir::MemFlagsData::new(),
            state,
            std::mem::offset_of!(NativeContext, string_array_encoding) as i32,
        );
        let string_arrays: Vec<_> = body
            .layout
            .parameters
            .iter()
            .enumerate()
            .filter(|(i, ty)| {
                **ty == ScalarType::Array && body.layout.array_slots.get(i) == Some(&8)
            })
            .map(|(i, _)| params[i])
            .collect();
        let array_cells = b.create_sized_stack_slot(StackSlotData::new(
            StackSlotKind::ExplicitSlot,
            (string_arrays.len().max(1) * 8) as u32,
            3,
        ));
        for (i, value) in string_arrays.iter().enumerate() {
            b.ins().stack_store(*value, array_cells, (i * 8) as i32);
        }
        let array_cells = b.ins().stack_addr(types::I64, array_cells, 0);
        let array_count = b.ins().iconst(types::I64, string_arrays.len() as i64);
        let array_encoding = b
            .ins()
            .iconst(types::I32, i64::from(options.string_encoding));
        b.ins().store(
            ir::MemFlagsData::new(),
            array_cells,
            state,
            std::mem::offset_of!(NativeContext, string_array_cells) as i32,
        );
        b.ins().store(
            ir::MemFlagsData::new(),
            array_count,
            state,
            std::mem::offset_of!(NativeContext, string_array_count) as i32,
        );
        b.ins().store(
            ir::MemFlagsData::new(),
            array_encoding,
            state,
            std::mem::offset_of!(NativeContext, string_array_encoding) as i32,
        );
        let empty_error = b.ins().iconst(types::I32, 0);
        b.ins().store(
            ir::MemFlagsData::new(),
            empty_error,
            state,
            std::mem::offset_of!(NativeContext, custom_error) as i32,
        );
        let output = b.create_sized_stack_slot(StackSlotData::new(
            StackSlotKind::ExplicitSlot,
            body.result_size().max(24).next_multiple_of(8) as u32,
            3,
        ));
        let empty = b.ins().iconst(types::I64, 0);
        for offset in (0..body.result_size().max(24).next_multiple_of(8)).step_by(8) {
            b.ins().stack_store(empty, output, offset as i32);
        }
        let output_ptr = b.ins().stack_addr(types::I64, output, 0);
        let imports = b.create_sized_stack_slot(StackSlotData::new(
            StackSlotKind::ExplicitSlot,
            (params.len().max(1) * 8) as u32,
            3,
        ));
        for i in 0..params.len() {
            b.ins().stack_store(empty, imports, (i * 8) as i32);
        }
        let fail_if = |b: &mut FunctionBuilder<'_>, bad: Value, code: i64| {
            let failed = b.create_block();
            let next = b.create_block();
            b.ins().brif(bad, failed, &[], next, &[]);
            b.switch_to_block(failed);
            let code = b.ins().iconst(types::I32, code);
            b.ins().jump(finish, &[code.into()]);
            b.switch_to_block(next);
        };
        for (i, value) in params.iter().enumerate() {
            if body.layout.by_ref_parameters[i]
                || matches!(
                    body.layout.parameters[i],
                    ScalarType::Variant | ScalarType::Record
                )
            {
                let null = b.ins().icmp_imm(ir::condcodes::IntCC::Equal, *value, 0);
                fail_if(&mut b, null, 5);
            } else if body.layout.parameters[i] == ScalarType::Double {
                let absolute = b.ins().fabs(*value);
                let max = b.ins().f64const(f64::MAX);
                let bad = b.ins().fcmp(
                    ir::condcodes::FloatCC::UnorderedOrGreaterThan,
                    absolute,
                    max,
                );
                fail_if(&mut b, bad, 6);
            }
        }
        if let Some(out) = caller_out {
            let null = b.ins().icmp_imm(ir::condcodes::IntCC::Equal, out, 0);
            fail_if(&mut b, null, 5);
        }
        let initialized = b.ins().load(types::I32, ir::MemFlagsData::new(), state, 12);
        let ready = b.create_block();
        let initialize = b.create_block();
        let existing = b.create_block();
        b.ins().brif(initialized, existing, &[], initialize, &[]);
        b.switch_to_block(initialize);
        let busy = b.ins().iconst(types::I32, -1);
        b.ins().store(ir::MemFlagsData::new(), busy, state, 12);
        if let Some(initializer) = runtime.initializer {
            let function = module.declare_func_in_func(initializer, b.func);
            let call = b.ins().call(function, &[state, output_ptr]);
            let status = b.inst_results(call)[0];
            let failed = b.create_block();
            let success = b.create_block();
            b.ins().brif(status, failed, &[], success, &[]);
            b.switch_to_block(failed);
            b.ins().store(ir::MemFlagsData::new(), status, state, 12);
            b.ins().jump(finish, &[status.into()]);
            b.switch_to_block(success);
        }
        let done = b.ins().iconst(types::I32, -2);
        b.ins().store(ir::MemFlagsData::new(), done, state, 12);
        b.ins().jump(ready, &[]);
        b.switch_to_block(existing);
        let ok = b
            .ins()
            .icmp_imm(ir::condcodes::IntCC::Equal, initialized, -2);
        let failed = b.create_block();
        b.ins().brif(ok, ready, &[], failed, &[]);
        b.switch_to_block(failed);
        let busy = b
            .ins()
            .icmp_imm(ir::condcodes::IntCC::Equal, initialized, -1);
        let error = b.ins().iconst(types::I32, 5);
        let error = b.ins().select(busy, error, initialized);
        b.ins().jump(finish, &[error.into()]);
        b.switch_to_block(ready);
        let mut args = vec![state];
        for (i, ((value, ty), by_ref)) in params
            .iter()
            .zip(&body.layout.parameters)
            .zip(&body.layout.by_ref_parameters)
            .enumerate()
        {
            let value = if !by_ref && *ty == ScalarType::String && string_encoding != 0 {
                let out = b.ins().stack_addr(types::I64, imports, (i * 8) as i32);
                let encoding = b.ins().iconst(types::I32, i64::from(string_encoding));
                let helper =
                    module.declare_func_in_func(runtime.values["cvb_string_import"], b.func);
                let call = b.ins().call(helper, &[*value, encoding, out]);
                let status = b.inst_results(call)[0];
                let next = b.create_block();
                b.ins().brif(status, finish, &[status.into()], next, &[]);
                b.switch_to_block(next);
                b.ins().stack_load(types::I64, imports, (i * 8) as i32)
            } else if !by_ref && *ty == ScalarType::Byte {
                b.ins().uextend(types::I32, *value)
            } else if !by_ref && *ty == ScalarType::Integer {
                b.ins().sextend(types::I32, *value)
            } else if !by_ref && *ty == ScalarType::Boolean {
                let truth = b.ins().icmp_imm(ir::condcodes::IntCC::NotEqual, *value, 0);
                let truth = b.ins().uextend(types::I32, truth);
                b.ins().ineg(truth)
            } else {
                *value
            };
            args.push(value);
        }
        args.push(output_ptr);
        let target = module.declare_func_in_func(body.function, b.func);
        let call = b.ins().call(target, &args);
        let status = b.inst_results(call)[0];
        b.ins().jump(finish, &[status.into()]);
        b.switch_to_block(finish);
        b.ins().store(
            ir::MemFlagsData::new(),
            old_array_cells,
            state,
            std::mem::offset_of!(NativeContext, string_array_cells) as i32,
        );
        b.ins().store(
            ir::MemFlagsData::new(),
            old_array_count,
            state,
            std::mem::offset_of!(NativeContext, string_array_count) as i32,
        );
        b.ins().store(
            ir::MemFlagsData::new(),
            old_array_encoding,
            state,
            std::mem::offset_of!(NativeContext, string_array_encoding) as i32,
        );

        let status = b.block_params(finish)[0];
        for (i, (ty, by_ref)) in body
            .layout
            .parameters
            .iter()
            .zip(&body.layout.by_ref_parameters)
            .enumerate()
        {
            if !by_ref && *ty == ScalarType::String && string_encoding != 0 {
                let cell = b.ins().stack_addr(types::I64, imports, (i * 8) as i32);
                let kind = b.ins().iconst(types::I32, 8);
                let clear = module.declare_func_in_func(runtime.values["cvb_value_clear"], b.func);
                b.ins().call(clear, &[cell, kind]);
            }
        }
        let source_data = module.declare_data_in_func(source_data, b.func);
        let source_ptr = b.ins().global_value(types::I64, source_data);
        let source_len = b.ins().iconst(types::I64, source.len() as i64);
        let record = module.declare_func_in_func(runtime.record_error, b.func);
        let call = b
            .ins()
            .call(record, &[state, status, source_ptr, source_len]);
        let error = b.inst_results(call)[0];
        if let Some(out) = caller_out {
            let store = b.create_block();
            let returned = b.create_block();
            b.ins().brif(out, store, &[], returned, &[]);
            b.switch_to_block(store);
            if aggregate {
                let mut offset = 0;
                while offset < body.result_size() {
                    let remaining = body.result_size() - offset;
                    let (ty, width) = if remaining >= 8 {
                        (types::I64, 8)
                    } else if remaining >= 4 {
                        (types::I32, 4)
                    } else if remaining >= 2 {
                        (types::I16, 2)
                    } else {
                        (types::I8, 1)
                    };
                    let value = b.ins().stack_load(ty, output, offset as i32);
                    let empty = b.ins().iconst(ty, 0);
                    let value = b.ins().select(status, empty, value);
                    b.ins()
                        .store(ir::MemFlagsData::new(), value, out, offset as i32);
                    offset += width;
                }
            } else {
                let value = b.ins().stack_load(abi_type(body.layout.result), output, 0);
                let empty = zero(&mut b, body.layout.result);
                let value = b.ins().select(status, empty, value);
                b.ins().store(ir::MemFlagsData::new(), value, out, 0);
            }
            b.ins().jump(returned, &[]);
            b.switch_to_block(returned);
        }
        if hresult {
            b.ins().return_(&[error]);
        } else if hidden {
            b.ins().return_(&[caller_out.unwrap()]);
        } else if body.layout.has_result {
            let result = b.ins().stack_load(result_abi, output, 0);
            let empty = if register_record.is_some() {
                b.ins().iconst(result_abi, 0)
            } else {
                zero(&mut b, body.layout.result)
            };
            let result = b.ins().select(status, empty, result);
            b.ins().return_(&[result]);
        } else {
            b.ins().return_(&[]);
        }
        b.seal_all_blocks();
        b.finalize();
    }
    module
        .define_function(id, &mut context)
        .map_err(module_err)?;
    Ok(id)
}

pub fn compile_last_error_wrapper(
    module: &mut impl Module,
    runtime: &NativeRuntime,
) -> Result<ClifFuncId, JitError> {
    compile_error_accessor(
        module,
        runtime,
        "CalcVbaGetLastError",
        runtime.get_last_error,
    )
}

pub fn compile_last_error_help_wrapper(
    module: &mut impl Module,
    runtime: &NativeRuntime,
) -> Result<ClifFuncId, JitError> {
    compile_error_accessor(
        module,
        runtime,
        "CalcVbaGetLastErrorHelp",
        runtime.values["cvb_get_error_help"],
    )
}

fn compile_error_accessor(
    module: &mut impl Module,
    runtime: &NativeRuntime,
    name: &str,
    helper: ClifFuncId,
) -> Result<ClifFuncId, JitError> {
    let mut signature = module.make_signature();
    signature.params = vec![AbiParam::new(types::I64); 2];
    signature.returns.push(AbiParam::new(types::I32));
    let id = module
        .declare_function(name, Linkage::Export, &signature)
        .map_err(module_err)?;
    let mut context = module.make_context();
    context.func.signature = signature;
    let mut frontend = FunctionBuilderContext::new();
    {
        let mut b = FunctionBuilder::new(&mut context.func, &mut frontend);
        let entry = b.create_block();
        b.switch_to_block(entry);
        b.append_block_params_for_function_params(entry);
        let params = b.block_params(entry).to_vec();
        let state = module.declare_data_in_func(runtime.context, b.func);
        let state = b.ins().global_value(types::I64, state);
        let helper = module.declare_func_in_func(helper, b.func);
        let call = b.ins().call(helper, &[state, params[0], params[1]]);
        let value = b.inst_results(call)[0];
        b.ins().return_(&[value]);
        b.seal_all_blocks();
        b.finalize();
    }
    module
        .define_function(id, &mut context)
        .map_err(module_err)?;
    Ok(id)
}

/// Explicitly release module owners on the owning VBA thread before FreeLibrary.
/// The context stays closed afterward, so initialization still runs only once.
pub fn compile_shutdown_wrapper(
    module: &mut impl Module,
    program: &OxProgram,
    runtime: &NativeRuntime,
) -> Result<ClifFuncId, JitError> {
    let mut signature = module.make_signature();
    signature.returns.push(AbiParam::new(types::I32));
    let id = module
        .declare_function("CalcVbaShutdown", Linkage::Export, &signature)
        .map_err(module_err)?;
    let mut context = module.make_context();
    context.func.signature = signature;
    let mut frontend = FunctionBuilderContext::new();
    {
        let mut b = FunctionBuilder::new(&mut context.func, &mut frontend);
        let entry = b.create_block();
        b.switch_to_block(entry);
        let enter = module.declare_func_in_func(runtime.enter, b.func);
        let call = b.ins().call(enter, &[]);
        let status = b.inst_results(call)[0];
        check_status(&mut b, status);
        let data = module.declare_data_in_func(runtime.context, b.func);
        let state = b.ins().global_value(types::I64, data);
        let depth = b
            .ins()
            .load(types::I32, ir::MemFlagsData::new(), state, DEPTH_OFFSET);
        let busy = b.create_block();
        let close = b.create_block();
        b.ins().brif(depth, busy, &[], close, &[]);
        b.switch_to_block(busy);
        let error = b.ins().iconst(types::I32, 5);
        b.ins().return_(&[error]);
        b.switch_to_block(close);
        let closed = b.ins().iconst(types::I32, 5);
        b.ins().store(
            ir::MemFlagsData::new(),
            closed,
            state,
            std::mem::offset_of!(NativeContext, initialization) as i32,
        );
        let one = b.ins().iconst(types::I32, 1);
        b.ins()
            .store(ir::MemFlagsData::new(), one, state, DEPTH_OFFSET);
        let offsets = global_offsets(program);
        let mut status = b.ins().iconst(types::I32, 0);
        for id in &runtime.foreign_states {
            let data = module.declare_data_in_func(*id, b.func);
            let address = b.ins().global_value(types::I64, data);
            let clear = module.declare_func_in_func(runtime.values["cvb_declare_clear"], b.func);
            b.ins().call(clear, &[address]);
        }
        for (i, global) in program.globals.iter().enumerate() {
            let address = b.ins().iadd_imm(
                state,
                (std::mem::size_of::<NativeContext>() + offsets[i]) as i64,
            );
            let (helper, args) = if let OxTy::Record(record) = global.ty {
                let layout = vba_record_layout_for_fields(&program.record_layouts[record.0])
                    .map_err(JitError::unsupported)?;
                let descriptor = values::data(
                    module,
                    &mut b,
                    values::record_descriptor(&layout).map_err(JitError::unsupported)?,
                    4,
                );
                ("cvb_record_clear", vec![address, descriptor])
            } else {
                let kind = match global.ty {
                    OxTy::Str | OxTy::FixedStr(_) => 8,
                    OxTy::Variant => 12,
                    OxTy::Object(_) => 9,
                    OxTy::Array(_, _) => 0x2000,
                    _ => 0,
                };
                let kind = b.ins().iconst(types::I32, kind);
                ("cvb_value_clear", vec![address, kind])
            };
            let helper = module.declare_func_in_func(runtime.values[helper], b.func);
            let call = b.ins().call(helper, &args);
            let next = b.inst_results(call)[0];
            status = b.ins().select(status, status, next);
        }
        let zero = b.ins().iconst(types::I32, 0);
        let drain = module.declare_func_in_func(runtime.values["cvb_drain_terminations"], b.func);
        b.ins().call(drain, &[state]);
        let services = module.declare_func_in_func(runtime.values["cvb_services_clear"], b.func);
        b.ins().call(services, &[state]);
        b.ins()
            .store(ir::MemFlagsData::new(), zero, state, DEPTH_OFFSET);
        let null = b.ins().iconst(types::I64, 0);
        let record = module.declare_func_in_func(runtime.record_error, b.func);
        b.ins().call(record, &[state, status, null, null]);
        b.ins().return_(&[status]);
        b.seal_all_blocks();
        b.finalize();
    }
    module
        .define_function(id, &mut context)
        .map_err(module_err)?;
    Ok(id)
}

fn emit_error_transfer(
    module: &mut impl Module,
    b: &mut FunctionBuilder<'_>,
    values: &values::Values<'_>,
    blocks: &HashMap<calcvba_oxir::BlockId, ir::Block>,
    func: &OxFunc,
    helper: &str,
    operands: (u32, u32),
) {
    let (first, second) = operands;
    let output = b.create_sized_stack_slot(StackSlotData::new(StackSlotKind::ExplicitSlot, 8, 3));
    let pointer = b.ins().stack_addr(types::I64, output, 0);
    let candidates: BTreeSet<_> = if helper == "cvb_fault_route" {
        std::iter::once(second as usize)
            .chain(
                func.blocks
                    .iter()
                    .flat_map(|block| &block.instrs)
                    .filter_map(|instruction| {
                        if let OxInst::SetErrorHandler(calcvba_oxir::ErrorHandler::GotoLabel(
                            label,
                        )) = instruction
                        {
                            Some(label.0)
                        } else {
                            None
                        }
                    }),
            )
            .collect()
    } else if helper == "cvb_gosub" {
        func.blocks
            .iter()
            .filter_map(|block| {
                if let OxTerminator::GoSub { ret, .. } = block.terminator {
                    Some(ret.0)
                } else {
                    None
                }
            })
            .collect()
    } else {
        func.blocks
            .iter()
            .filter_map(|block| match (&block.terminator, first) {
                (OxTerminator::FaultDispatch { resume, .. }, 0) => Some(resume.0),
                (OxTerminator::FaultDispatch { resume_next, .. }, 1) => Some(resume_next.0),
                (_, 2) => Some(second as usize),
                _ => None,
            })
            .collect()
    };
    let first = b.ins().iconst(types::I32, i64::from(first));
    let second = b.ins().iconst(types::I32, i64::from(second));
    values.call(
        module,
        b,
        helper,
        &[values.lower.runtime, first, second, pointer],
    );
    let target = b.ins().stack_load(types::I32, output, 0);
    let invalid = b.create_block();
    let mut switch = cranelift_frontend::Switch::new();
    for id in candidates {
        switch.set_entry(id as u128, blocks[&calcvba_oxir::BlockId(id)]);
    }
    switch.emit(b, target, invalid);
    b.switch_to_block(invalid);
    let status = b.ins().iconst(types::I32, 5);
    values.lower.return_status(b, status);
}
