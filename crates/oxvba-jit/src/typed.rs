//! Native typed scalar primary entries. Eligible bodies use physical stack
//! slots/registers and an explicit status/result convention, never JitFrame.
use super::*;
use oxvba_rt_abi::host_entry::{Arguments, Scalar, ScalarLayout, ScalarType, TypedContext};
use std::marker::PhantomData;

pub(crate) struct ScalarEntry {
    pub(crate) pointer: *const u8,
    pub(crate) layout: ScalarLayout,
}

/// A typed resolved address can only execute with its original live image.
/// Keeping a stale copy is harmless: invocation checks image identity first.
pub struct TypedEntry<A, R> {
    pointer: *const u8,
    owner: std::sync::Arc<()>,
    signature: PhantomData<fn(A) -> R>,
}
impl<A, R> Clone for TypedEntry<A, R> {
    fn clone(&self) -> Self {
        Self {
            pointer: self.pointer,
            owner: self.owner.clone(),
            signature: PhantomData,
        }
    }
}
impl<A: Arguments, R: Scalar> TypedEntry<A, R> {
    pub fn invoke(&self, image: &CompiledImage<'_>, args: A, blocks: u64) -> Result<R, i32> {
        if !std::sync::Arc::ptr_eq(&self.owner, &image.typed_identity) {
            return Err(5);
        }
        if !args.is_valid() {
            return Err(6);
        }
        let mut result = R::default();
        let mut context = TypedContext {
            remaining_blocks: blocks,
        };
        // SAFETY: bind checked A/R against the compiler's exact typed signature;
        // image identity proves the original executable owner is still live.
        let status = unsafe { args.invoke(self.pointer, &mut context, &mut result) };
        if status == 0 { Ok(result) } else { Err(status) }
    }
}
impl CompiledImage<'_> {
    pub fn bind_typed<A: Arguments, R: Scalar>(
        &self,
        func: FuncId,
    ) -> Result<TypedEntry<A, R>, JitError> {
        let entry = self
            .typed_functions
            .last()
            .and_then(|p| p.get(func.0))
            .and_then(Option::as_ref)
            .ok_or_else(|| JitError::unsupported("procedure has no typed scalar entry"))?;
        if entry.layout.parameters != A::types() || entry.layout.result != R::TYPE {
            return Err(JitError::unsupported("typed entry signature mismatch"));
        }
        Ok(TypedEntry {
            pointer: entry.pointer,
            owner: self.typed_identity.clone(),
            signature: PhantomData,
        })
    }
}

fn machine_type(ty: ScalarType) -> ir::Type {
    match ty {
        ScalarType::Double => types::F64,
        ScalarType::Long | ScalarType::Boolean => types::I32,
    }
}

pub(crate) type ScalarDefinitions = Vec<Vec<Option<(ClifFuncId, ScalarLayout)>>>;
pub(crate) fn compile_scalar_entries(
    module: &mut JITModule,
    programs: &[&OxProgram],
) -> Result<ScalarDefinitions, JitError> {
    let mut entries = Vec::new();
    let allow_typed = programs
        .iter()
        .all(|p| p.classes.is_empty() && p.record_layouts.is_empty());
    for (p, program) in programs.iter().enumerate() {
        let mut functions = Vec::new();
        for (f, func) in program.funcs.iter().enumerate() {
            if !allow_typed {
                functions.push(None);
                continue;
            }
            let Ok(layout) = ScalarLayout::new(func) else {
                functions.push(None);
                continue;
            };
            let mut signature = module.make_signature();
            let ptr = module.target_config().pointer_type();
            signature.params.push(AbiParam::new(ptr));
            signature.params.extend(
                layout
                    .parameters
                    .iter()
                    .map(|ty| AbiParam::new(machine_type(*ty))),
            );
            signature.params.push(AbiParam::new(ptr));
            signature.returns.push(AbiParam::new(types::I32));
            let id = module
                .declare_function(&format!("ox$typed$p{p}$f{f}"), Linkage::Local, &signature)
                .map_err(module_err)?;
            let mut context = module.make_context();
            context.func.signature = signature;
            context.func.name = UserFuncName::user(p as u32, id.as_u32());
            let mut function_context = FunctionBuilderContext::new();
            {
                let mut b = FunctionBuilder::new(&mut context.func, &mut function_context);
                let entry = b.create_block();
                b.switch_to_block(entry);
                b.append_block_params_for_function_params(entry);
                let params = b.block_params(entry).to_vec();
                let runtime = params[0];
                let output = *params.last().expect("typed result pointer");
                let slots = b.create_sized_stack_slot(StackSlotData::new(
                    StackSlotKind::ExplicitSlot,
                    (layout.slots.len().max(1) * 8) as u32,
                    3,
                ));
                for (index, ty) in layout.slots.iter().enumerate() {
                    let value = if index < layout.parameters.len() {
                        params[index + 1]
                    } else if *ty == ScalarType::Double {
                        b.ins().f64const(0.0)
                    } else {
                        b.ins().iconst(types::I32, 0)
                    };
                    b.ins().stack_store(value, slots, (index * 8) as i32);
                }
                let blocks = func
                    .blocks
                    .iter()
                    .map(|block| (block.id, b.create_block()))
                    .collect::<HashMap<_, _>>();
                b.ins().jump(blocks[&func.entry], &[]);
                let lower = ScalarLower {
                    layout: &layout,
                    slots,
                };
                for block in &func.blocks {
                    b.switch_to_block(blocks[&block.id]);
                    let remaining = b
                        .ins()
                        .load(types::I64, ir::MemFlagsData::new(), runtime, 0);
                    let exhausted = b.ins().icmp_imm(ir::condcodes::IntCC::Equal, remaining, 0);
                    lower.fail_if(&mut b, exhausted, -1);
                    let remaining = b.ins().iadd_imm(remaining, -1);
                    b.ins()
                        .store(ir::MemFlagsData::new(), remaining, runtime, 0);
                    for inst in &block.instrs {
                        lower.instruction(&mut b, inst);
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
                            let value = lower.operand(&mut b, cond, ScalarType::Boolean);
                            b.ins()
                                .brif(value, blocks[then_blk], &[], blocks[else_blk], &[]);
                        }
                        OxTerminator::Return => {
                            let value = b.ins().stack_load(
                                machine_type(layout.result),
                                slots,
                                (layout.return_slot * 8) as i32,
                            );
                            b.ins().store(ir::MemFlagsData::new(), value, output, 0);
                            let status = b.ins().iconst(types::I32, 0);
                            b.ins().return_(&[status]);
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
                .define_function(id, &mut context)
                .map_err(module_err)?;
            functions.push(Some((id, layout)));
        }
        entries.push(functions);
    }
    Ok(entries)
}

struct ScalarLower<'a> {
    layout: &'a ScalarLayout,
    slots: ir::StackSlot,
}
impl ScalarLower<'_> {
    fn fail_if(&self, b: &mut FunctionBuilder<'_>, condition: Value, number: i32) {
        let fail = b.create_block();
        let next = b.create_block();
        b.ins().brif(condition, fail, &[], next, &[]);
        b.switch_to_block(fail);
        let status = b.ins().iconst(types::I32, i64::from(number));
        b.ins().return_(&[status]);
        b.switch_to_block(next);
    }
    fn operand(&self, b: &mut FunctionBuilder<'_>, operand: &OxOperand, to: ScalarType) -> Value {
        let from = self.layout.operand_type(operand).expect("admitted operand");
        let value = match operand {
            OxOperand::Use(place) => b.ins().stack_load(
                machine_type(from),
                self.slots,
                (self.layout.slot(*place).expect("admitted slot") * 8) as i32,
            ),
            OxOperand::Const(OxConst::F64(bits)) => b.ins().f64const(Ieee64::with_bits(*bits)),
            OxOperand::Const(OxConst::I16(value)) => b.ins().iconst(types::I32, i64::from(*value)),
            OxOperand::Const(OxConst::I32(value)) => b.ins().iconst(types::I32, i64::from(*value)),
            OxOperand::Const(OxConst::Bool(value)) => {
                b.ins().iconst(types::I32, if *value { -1 } else { 0 })
            }
            _ => unreachable!("admitted operand"),
        };
        if from == ScalarType::Long && to == ScalarType::Double {
            b.ins().fcvt_from_sint(types::F64, value)
        } else {
            value
        }
    }
    fn finite(&self, b: &mut FunctionBuilder<'_>, value: Value) {
        let abs = b.ins().fabs(value);
        let max = b.ins().f64const(f64::MAX);
        let finite = b
            .ins()
            .fcmp(ir::condcodes::FloatCC::LessThanOrEqual, abs, max);
        let bad = b.ins().icmp_imm(ir::condcodes::IntCC::Equal, finite, 0);
        self.fail_if(b, bad, 6);
    }
    fn instruction(&self, b: &mut FunctionBuilder<'_>, inst: &OxInst) {
        let (dst, value) = match inst {
            OxInst::StmtBoundary {
                clear_temps_from, ..
            } => {
                for i in self.layout.local_count + clear_temps_from..self.layout.slots.len() {
                    let value = if self.layout.slots[i] == ScalarType::Double {
                        b.ins().f64const(0.0)
                    } else {
                        b.ins().iconst(types::I32, 0)
                    };
                    b.ins().stack_store(value, self.slots, (i * 8) as i32);
                }
                return;
            }
            OxInst::DrainTerminations => return,
            OxInst::Assign { dst, value }
            | OxInst::Coerce {
                dst, src: value, ..
            } => {
                let ty = self.layout.slots[self.layout.slot(*dst).expect("admitted slot")];
                let value = self.operand(b, value, ty);
                if matches!(inst, OxInst::Coerce { .. }) && ty == ScalarType::Double {
                    self.finite(b, value);
                }
                (*dst, value)
            }
            OxInst::Arith {
                dst, op, lhs, rhs, ..
            } => {
                let ty = self.layout.slots[self.layout.slot(*dst).expect("admitted slot")];
                let a = self.operand(b, lhs, ty);
                let c = self.operand(b, rhs, ty);
                let value = if ty == ScalarType::Double {
                    let value = match op {
                        ArithOp::Add => b.ins().fadd(a, c),
                        ArithOp::Sub => b.ins().fsub(a, c),
                        ArithOp::Mul => b.ins().fmul(a, c),
                        _ => unreachable!("admitted arithmetic"),
                    };
                    self.finite(b, value);
                    value
                } else {
                    let a = b.ins().sextend(types::I64, a);
                    let c = b.ins().sextend(types::I64, c);
                    let wide = match op {
                        ArithOp::Add => b.ins().iadd(a, c),
                        ArithOp::Sub => b.ins().isub(a, c),
                        ArithOp::Mul => b.ins().imul(a, c),
                        _ => unreachable!("admitted arithmetic"),
                    };
                    let narrow = b.ins().ireduce(types::I32, wide);
                    let back = b.ins().sextend(types::I64, narrow);
                    let overflow = b.ins().icmp(ir::condcodes::IntCC::NotEqual, wide, back);
                    self.fail_if(b, overflow, 6);
                    narrow
                };
                (*dst, value)
            }
            OxInst::Div { dst, lhs, rhs } => {
                let a = self.operand(b, lhs, ScalarType::Double);
                let c = self.operand(b, rhs, ScalarType::Double);
                let zero = b.ins().f64const(0.0);
                let bad = b.ins().fcmp(ir::condcodes::FloatCC::Equal, c, zero);
                self.fail_if(b, bad, 11);
                let value = b.ins().fdiv(a, c);
                self.finite(b, value);
                (*dst, value)
            }
            OxInst::Compare {
                dst, op, lhs, rhs, ..
            } => {
                let ty = self
                    .layout
                    .comparison_type(lhs, rhs)
                    .expect("admitted comparison");
                let a = self.operand(b, lhs, ty);
                let c = self.operand(b, rhs, ty);
                let value = if ty == ScalarType::Double {
                    b.ins().fcmp(float_compare_cc(*op), a, c)
                } else {
                    b.ins().icmp(int_compare_cc(*op), a, c)
                };
                let value = b.ins().uextend(types::I32, value);
                let value = b.ins().ineg(value);
                (*dst, value)
            }
            OxInst::Truthy { dst, src } => {
                let ty = self.layout.operand_type(src).expect("admitted operand");
                let value = self.operand(b, src, ty);
                let truth = if ty == ScalarType::Double {
                    let zero = b.ins().f64const(0.0);
                    b.ins().fcmp(ir::condcodes::FloatCC::NotEqual, value, zero)
                } else {
                    b.ins().icmp_imm(ir::condcodes::IntCC::NotEqual, value, 0)
                };
                let value = b.ins().uextend(types::I32, truth);
                let value = b.ins().ineg(value);
                (*dst, value)
            }
            _ => unreachable!("admitted instruction"),
        };
        let slot = self.layout.slot(dst).expect("admitted destination");
        b.ins().stack_store(value, self.slots, (slot * 8) as i32);
    }
}
