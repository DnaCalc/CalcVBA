//! Direct scalar AOT entries using the same ScalarLower as the JIT. The module
//! decides where symbols live; neither instruction lowering nor wrappers know
//! any process address. Admission excludes all owning/dynamic operations.

use super::*;
use crate::typed::{PhysicalSlots, ScalarLower, machine_type};
use calcvba_oxir::OxArg;
use calcvba_rt_abi::host_entry::{LinkedTarget, ScalarLayout, ScalarType};
use std::collections::{BTreeMap, BTreeSet};

/// Keep native recursion inside a bounded part of the caller's stack. A typed
/// status 28 leaves the DLL through the ordinary zero-on-error wrapper path.
pub const MAX_NATIVE_CALL_DEPTH: i32 = 128;
const DEPTH_OFFSET: i32 = 8;

fn conversion_target(id: NativeImplId) -> Option<ScalarType> {
    match id {
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
}
fn problem(func: &OxFunc, detail: impl std::fmt::Display) -> JitError {
    JitError::unsupported(format!("native export procedure {}: {detail}", func.name))
}
fn numeric(ty: ScalarType) -> bool {
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
) -> Result<Vec<NativeScalar>, JitError> {
    calcvba_oxir::verify_program(program)
        .map_err(|e| JitError::unsupported(format!("invalid native OxIR: {e:?}")))?;
    if !program.classes.is_empty() || !program.globals.is_empty() {
        return Err(JitError::unsupported(
            "native exports do not yet support module state or classes",
        ));
    }
    let mut pending = roots.to_vec();
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
        let layout =
            ScalarLayout::with_native_program(func, program).map_err(|e| problem(func, e))?;
        if !layout.has_result
            || !layout.slots.iter().copied().all(numeric)
            || layout.by_ref_parameters.iter().any(|p| *p)
            || layout.optional_parameters.iter().any(|p| *p)
            || layout.param_array.is_some()
            || !layout.host_calls.is_empty()
        {
            return Err(problem(
                func,
                "initial native scalar path requires ByVal Long/Double/Boolean Function signatures and numeric locals",
            ));
        }
        for block in &func.blocks {
            for (index, inst) in block.instrs.iter().enumerate() {
                match inst {
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
                        if args.iter().any(|arg| !matches!(arg, OxArg::ByVal(_))) {
                            return Err(problem(
                                func,
                                format!(
                                    "ByRef/omitted call argument at block {:?} instruction {index}",
                                    block.id
                                ),
                            ));
                        }
                        pending.push(*proc);
                    }
                    OxInst::CallExtern { import, args, .. } => {
                        let target = layout
                            .extern_calls
                            .get(import)
                            .ok_or_else(|| problem(func, "unplanned native library call"))?
                            .target;
                        if !matches!(target,LinkedTarget::Builtin(native) if library_helpers.contains_key(&native) || conversion_target(native).is_some())
                            || args.iter().any(|arg| !matches!(arg, OxArg::ByVal(_)))
                        {
                            return Err(problem(
                                func,
                                format!("unsupported native library call {inst:?}"),
                            ));
                        }
                        if matches!(target, LinkedTarget::Builtin(native) if conversion_target(native).is_some())
                            && !matches!(args.as_slice(), [OxArg::ByVal(src)] if layout.operand_type(src).is_ok_and(numeric))
                        {
                            return Err(problem(
                                func,
                                "native numeric conversions require one numeric ByVal argument",
                            ));
                        }
                    }
                    OxInst::Assign { .. }
                    | OxInst::Coerce { .. }
                    | OxInst::Arith { .. }
                    | OxInst::Div { .. }
                    | OxInst::Neg { .. }
                    | OxInst::Compare { .. }
                    | OxInst::Logical { .. }
                    | OxInst::Not { .. }
                    | OxInst::Truthy { .. }
                    | OxInst::StmtBoundary { .. }
                    | OxInst::DrainTerminations => {
                        if calcvba_rt_abi::host_entry::value::needs_helper(inst, &layout) {
                            return Err(problem(
                                func,
                                format!(
                                    "owning helper at block {:?} instruction {index}: {inst:?}",
                                    block.id
                                ),
                            ));
                        }
                    }
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
        signature.params.extend(
            layout
                .parameters
                .iter()
                .map(|t| AbiParam::new(machine_type(*t))),
        );
        signature.params.push(AbiParam::new(ptr));
        signature.returns.push(AbiParam::new(types::I32));
        let function = module
            .declare_function(&format!("cvb$proc${}", id.0), Linkage::Local, &signature)
            .map_err(module_err)?;
        plans.insert(
            id.0,
            NativeScalar {
                procedure: id,
                function,
                layout,
            },
        );
    }
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
            let call_output =
                b.create_sized_stack_slot(StackSlotData::new(StackSlotKind::ExplicitSlot, 8, 3));
            let call_output_ptr = b.ins().stack_addr(types::I64, call_output, 0);
            let slots = b.create_sized_stack_slot(StackSlotData::new(
                StackSlotKind::ExplicitSlot,
                (plan.layout.slots.len().max(1) * 8) as u32,
                3,
            ));
            for (i, ty) in plan.layout.slots.iter().enumerate() {
                let value = if i < plan.layout.parameters.len() {
                    params[i + 1]
                } else if *ty == ScalarType::Double {
                    b.ins().f64const(0.0)
                } else {
                    b.ins().iconst(machine_type(*ty), 0)
                };
                b.ins().stack_store(value, slots, (i * 8) as i32);
            }
            let lower = ScalarLower {
                layout: &plan.layout,
                slots: PhysicalSlots::Stack(slots),
                runtime,
                call_conv: module.target_config().default_call_conv,
            };
            let blocks: HashMap<_, _> = func
                .blocks
                .iter()
                .map(|block| (block.id, b.create_block()))
                .collect();
            b.ins().jump(blocks[&func.entry], &[]);
            for block in &func.blocks {
                b.switch_to_block(blocks[&block.id]);
                for inst in &block.instrs {
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
                            (from, to) if from == to => {
                                if to == ScalarType::Double {
                                    lower.finite(&mut b, input);
                                }
                                input
                            }
                            (ScalarType::Double, ScalarType::Long) => {
                                let helper = module.declare_func_in_func(long_from_f64, b.func);
                                let call = b.ins().call(helper, &[runtime, input, call_output_ptr]);
                                let status = b.inst_results(call)[0];
                                check_status(&mut b, status);
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
                    let call_plan = match inst {
                        OxInst::CallProc { dst, proc, args } => {
                            let callee = &plans[&proc.0];
                            Some((
                                *dst,
                                args,
                                &callee.layout.parameters,
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
                                call.signature.result.expect("numeric builtin result"),
                                library_helpers[&native],
                            ))
                        }
                        _ => None,
                    };
                    if let Some((dst, args, parameters, result_type, function)) = call_plan {
                        let mut values = vec![runtime];
                        for (arg, ty) in args.iter().zip(parameters) {
                            let OxArg::ByVal(op) = arg else {
                                unreachable!("admitted ByVal call")
                            };
                            values.push(lower.operand(&mut b, op, *ty));
                        }
                        values.push(call_output_ptr);
                        let target = module.declare_func_in_func(function, b.func);
                        let call = b.ins().call(target, &values);
                        let status = b.inst_results(call)[0];
                        check_status(&mut b, status);
                        if let Some(dst) = dst {
                            let result =
                                b.ins()
                                    .stack_load(machine_type(result_type), call_output, 0);
                            lower.store_slot(
                                &mut b,
                                plan.layout.slot(dst).expect("admitted call output"),
                                result_type,
                                result,
                            );
                        }
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
                        let result = b.ins().stack_load(
                            machine_type(plan.layout.result),
                            slots,
                            (plan.layout.return_slot * 8) as i32,
                        );
                        b.ins().store(ir::MemFlagsData::new(), result, out, 0);
                        let zero = b.ins().iconst(types::I32, 0);
                        b.ins().return_(&[zero]);
                    }
                    _ => {
                        let status = b.ins().iconst(types::I32, 5);
                        b.ins().return_(&[status]);
                    }
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
        for inst in returns {
            let mut cursor = FuncCursor::new(&mut context.func);
            cursor.goto_inst(inst);
            cursor
                .ins()
                .store(ir::MemFlagsData::new(), depth, runtime, DEPTH_OFFSET);
        }
        module
            .define_function(plan.function, &mut context)
            .map_err(module_err)?;
    }
    Ok(plans.into_values().collect())
}

/// Native ABI wrapper calls this procedure's compiled body directly. enter is
/// a template helper with signature ()->i32, enforcing one owning VBA thread.
pub fn compile_native_wrapper(
    module: &mut impl Module,
    name: &str,
    body: &NativeScalar,
    enter: ClifFuncId,
) -> Result<ClifFuncId, JitError> {
    let abi_type = |ty| {
        if ty == ScalarType::Boolean {
            types::I16
        } else {
            machine_type(ty)
        }
    };
    let mut signature = module.make_signature();
    signature.params.extend(
        body.layout
            .parameters
            .iter()
            .map(|t| AbiParam::new(abi_type(*t))),
    );
    signature
        .returns
        .push(AbiParam::new(abi_type(body.layout.result)));
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
        let refused = b.create_block();
        let accepted = b.create_block();
        let helper = module.declare_func_in_func(enter, b.func);
        let call = b.ins().call(helper, &[]);
        let status = b.inst_results(call)[0];
        b.ins().brif(status, refused, &[], accepted, &[]);
        b.switch_to_block(accepted);
        for (value, ty) in params.iter().zip(&body.layout.parameters) {
            if *ty == ScalarType::Double {
                let absolute = b.ins().fabs(*value);
                let max = b.ins().f64const(f64::MAX);
                let finite = b
                    .ins()
                    .fcmp(ir::condcodes::FloatCC::LessThanOrEqual, absolute, max);
                let next = b.create_block();
                b.ins().brif(finite, next, &[], refused, &[]);
                b.switch_to_block(next);
            }
        }
        let output =
            b.create_sized_stack_slot(StackSlotData::new(StackSlotKind::ExplicitSlot, 8, 3));
        let runtime =
            b.create_sized_stack_slot(StackSlotData::new(StackSlotKind::ExplicitSlot, 16, 3));
        let zero = b.ins().iconst(types::I64, 0);
        b.ins().stack_store(zero, output, 0);
        let budget = b.ins().iconst(types::I64, -1);
        b.ins().stack_store(budget, runtime, 0);
        b.ins().stack_store(zero, runtime, DEPTH_OFFSET);
        let runtime = b.ins().stack_addr(types::I64, runtime, 0);
        let output_ptr = b.ins().stack_addr(types::I64, output, 0);
        let mut args = vec![runtime];
        for (value, ty) in params.iter().zip(&body.layout.parameters) {
            args.push(if *ty == ScalarType::Boolean {
                let truth = b.ins().icmp_imm(ir::condcodes::IntCC::NotEqual, *value, 0);
                let truth = b.ins().uextend(types::I32, truth);
                b.ins().ineg(truth)
            } else {
                *value
            });
        }
        args.push(output_ptr);
        let target = module.declare_func_in_func(body.function, b.func);
        let call = b.ins().call(target, &args);
        let status = b.inst_results(call)[0];
        let returned = b.create_block();
        b.ins().brif(status, refused, &[], returned, &[]);
        b.switch_to_block(returned);
        let result = b.ins().stack_load(abi_type(body.layout.result), output, 0);
        b.ins().return_(&[result]);
        b.switch_to_block(refused);
        let zero = if body.layout.result == ScalarType::Double {
            b.ins().f64const(0.0)
        } else {
            b.ins().iconst(abi_type(body.layout.result), 0)
        };
        b.ins().return_(&[zero]);
        b.seal_all_blocks();
        b.finalize();
    }
    module
        .define_function(id, &mut context)
        .map_err(module_err)?;
    Ok(id)
}
