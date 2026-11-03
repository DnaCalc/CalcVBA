//! Direct scalar AOT entries using the same ScalarLower as the JIT. The module
//! decides where symbols live; neither instruction lowering nor wrappers know
//! any process address. Admission excludes all owning/dynamic operations.

use super::*;
use crate::typed::{PhysicalSlots, ScalarLower, machine_type};
use oxvba_oxir::OxArg;
use oxvba_rt_abi::host_entry::{LinkedTarget, ScalarLayout, ScalarType};
use std::collections::{BTreeMap, BTreeSet};

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
) -> Result<Vec<NativeScalar>, JitError> {
    oxvba_oxir::verify_program(program)
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
                        if !matches!(target,LinkedTarget::Builtin(native) if library_helpers.contains_key(&native))
                            || args.iter().any(|arg| !matches!(arg, OxArg::ByVal(_)))
                        {
                            return Err(problem(
                                func,
                                format!("unsupported native library call {inst:?}"),
                            ));
                        }
                    }
                    OxInst::Assign { .. }
                    | OxInst::Coerce { .. }
                    | OxInst::Arith { .. }
                    | OxInst::Div { .. }
                    | OxInst::Neg { .. }
                    | OxInst::Compare { .. }
                    | OxInst::Truthy { .. }
                    | OxInst::StmtBoundary { .. }
                    | OxInst::DrainTerminations => {
                        if oxvba_rt_abi::host_entry::value::needs_helper(inst, &layout) {
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
        {
            let mut b = FunctionBuilder::new(&mut context.func, &mut frontend);
            let entry = b.create_block();
            b.switch_to_block(entry);
            b.append_block_params_for_function_params(entry);
            let params = b.block_params(entry).to_vec();
            let runtime = params[0];
            let out = *params.last().expect("typed result pointer");
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
                        let output = b.create_sized_stack_slot(StackSlotData::new(
                            StackSlotKind::ExplicitSlot,
                            8,
                            3,
                        ));
                        let output_ptr = b.ins().stack_addr(types::I64, output, 0);
                        let mut values = vec![runtime];
                        for (arg, ty) in args.iter().zip(parameters) {
                            let OxArg::ByVal(op) = arg else {
                                unreachable!("admitted ByVal call")
                            };
                            values.push(lower.operand(&mut b, op, *ty));
                        }
                        values.push(output_ptr);
                        let target = module.declare_func_in_func(function, b.func);
                        let call = b.ins().call(target, &values);
                        let status = b.inst_results(call)[0];
                        let failed = b.ins().icmp_imm(ir::condcodes::IntCC::NotEqual, status, 0);
                        let fail = b.create_block();
                        let next = b.create_block();
                        b.ins().brif(failed, fail, &[], next, &[]);
                        b.switch_to_block(fail);
                        b.ins().return_(&[status]);
                        b.switch_to_block(next);
                        if let Some(dst) = dst {
                            let result = b.ins().stack_load(machine_type(result_type), output, 0);
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
            b.create_sized_stack_slot(StackSlotData::new(StackSlotKind::ExplicitSlot, 8, 3));
        let zero = b.ins().iconst(types::I64, 0);
        b.ins().stack_store(zero, output, 0);
        let budget = b.ins().iconst(types::I64, -1);
        b.ins().stack_store(budget, runtime, 0);
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
