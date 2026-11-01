//! Typed execution of the original OxFunc with unboxed owning physical slots.
use oxvba_oxir::{ArithOp, CmpOp, OxFunc, OxInst, OxOperand, OxTerminator};
use oxvba_rt_abi::host_entry::{
    ScalarBits, ScalarLayout, ScalarType, TypedContext,
    activation::{Activation, CONTROL, FINISHED_PC, SUSPENDED},
};

pub struct TypedInterpreter<'p> {
    func: &'p OxFunc,
    layout: ScalarLayout,
    blocks: Vec<(oxvba_oxir::BlockId, usize, u32)>,
}
impl<'p> TypedInterpreter<'p> {
    pub fn bind(func: &'p OxFunc) -> Result<Self, String> {
        Self::bind_with_interfaces(func, &[])
    }
    pub fn bind_with_interfaces(
        func: &'p OxFunc,
        interfaces: &[oxvba_oxir::ComInterface],
    ) -> Result<Self, String> {
        Self::with_layout(func, ScalarLayout::with_interfaces(func, interfaces)?)
    }
    pub fn bind_with_program(
        func: &'p OxFunc,
        program: &oxvba_oxir::OxProgram,
    ) -> Result<Self, String> {
        Self::with_layout(func, ScalarLayout::with_program(func, program)?)
    }
    pub fn bind_with_programs(
        func: &'p OxFunc,
        program: &oxvba_oxir::OxProgram,
        programs: &[&oxvba_oxir::OxProgram],
    ) -> Result<Self, String> {
        Self::with_layout(func, ScalarLayout::with_programs(func, program, programs)?)
    }
    fn with_layout(func: &'p OxFunc, layout: ScalarLayout) -> Result<Self, String> {
        let mut pc = 0u32;
        let mut blocks = Vec::new();
        for (i, block) in func.blocks.iter().enumerate() {
            blocks.push((block.id, i, pc));
            pc = pc
                .checked_add(
                    u32::try_from(block.instrs.len())
                        .map_err(|_| "typed PC overflow")?
                        .checked_add(2)
                        .ok_or("typed PC overflow")?,
                )
                .ok_or("typed PC overflow")?;
        }
        if pc == FINISHED_PC {
            return Err("typed PC overflow".into());
        }
        Ok(Self {
            func,
            layout,
            blocks,
        })
    }
    pub fn layout(&self) -> &ScalarLayout {
        &self.layout
    }
    pub fn entry_pc(&self) -> u32 {
        self.target(self.func.entry)
    }
    fn target(&self, block: oxvba_oxir::BlockId) -> u32 {
        self.blocks
            .iter()
            .find(|(id, _, _)| *id == block)
            .expect("admitted block")
            .2
    }
    pub fn invoke(
        &self,
        args: &[ScalarBits],
        context: &mut TypedContext,
    ) -> Result<ScalarBits, i32> {
        if args.len() != self.layout.parameters.len() {
            return Err(13);
        }
        let mut slots = [calcvba_host_api::Slot::Long(0); calcvba_host_api::MAX_TYPED_ARGS];
        for (i, (bits, ty)) in args.iter().zip(&self.layout.parameters).enumerate() {
            slots[i] = match ty {
                ScalarType::Long => calcvba_host_api::Slot::Long(bits.as_long()),
                ScalarType::Double => calcvba_host_api::Slot::Double(bits.as_double()),
                _ => return Err(13),
            };
        }
        self.invoke_slots_inner(&slots[..args.len()], context, false)
    }
    pub fn invoke_slots(
        &self,
        args: &[calcvba_host_api::Slot<'_>],
        context: &mut TypedContext,
    ) -> Result<ScalarBits, i32> {
        self.invoke_slots_inner(args, context, true)
    }
    fn invoke_slots_inner(
        &self,
        args: &[calcvba_host_api::Slot<'_>],
        context: &mut TypedContext,
        finite_inputs: bool,
    ) -> Result<ScalarBits, i32> {
        if self.layout.result == ScalarType::HostObject
            || self.layout.slots.len() != self.layout.global_base
        {
            return Err(13);
        }
        // SAFETY: this synchronous local-only frame retains its arguments and
        // borrows the checked layout; no global pointer is dereferenced.
        let mut frame = unsafe {
            if finite_inputs {
                Activation::borrowed(&self.layout, args, std::ptr::null_mut(), 0, self.entry_pc())
            } else {
                Activation::borrowed_exact(&self.layout, args, self.entry_pc())
            }
        }?;
        if self.resume(&mut frame, context, u64::MAX)? != 0 {
            return Err(5);
        }
        Ok(frame.raw().slots[self.layout.return_slot])
    }
    /// Execute at most operations completed physical operations. A paused PC
    /// always names the next operation, so effects before it are never repeated.
    pub fn resume(
        &self,
        frame: &mut Activation<'_>,
        context: &mut TypedContext,
        operations: u64,
    ) -> Result<i32, i32> {
        if !std::ptr::eq(&self.layout, frame.layout()) && self.layout != *frame.layout() {
            return Err(5);
        }
        // SAFETY: this interpreter uniquely owns the admitted frame while running.
        unsafe {
            frame.raw_mut().remaining_operations = operations;
        }
        for _ in 0..operations {
            // SAFETY: unique frame; one physical operation is about to execute.
            unsafe {
                frame.raw_mut().remaining_operations -= 1;
            }
            if frame.pc() == FINISHED_PC {
                return Err(5);
            }
            frame.clear_scratch();
            let (_, index, start) = self
                .blocks
                .iter()
                .rev()
                .find(|(_, _, start)| *start <= frame.pc())
                .ok_or(5)?;
            let block = &self.func.blocks[*index];
            let offset = (frame.pc() - *start) as usize;
            if offset == 0 {
                context.step()?;
                frame.set_pc(frame.pc() + 1);
            } else if let Some(inst) = block.instrs.get(offset - 1) {
                if oxvba_rt_abi::host_entry::stack_instruction(inst) {
                    return Ok(CONTROL);
                }
                self.instruction(frame, context, inst)?;
                frame.set_pc(frame.pc() + 1);
            } else if offset == block.instrs.len() + 1 {
                if oxvba_rt_abi::host_entry::stack_terminator(&block.terminator) {
                    return Ok(CONTROL);
                }
                let next_pc = match &block.terminator {
                    OxTerminator::Jump(next) => self.target(*next),
                    OxTerminator::Branch {
                        cond,
                        then_blk,
                        else_blk,
                    } => self.target(if frame.read(cond, ScalarType::Boolean)?.as_long() != 0 {
                        *then_blk
                    } else {
                        *else_blk
                    }),
                    OxTerminator::Return => {
                        frame.set_pc(FINISHED_PC);
                        frame.clear_scratch();
                        return Ok(0);
                    }
                    _ => return Err(5),
                };
                frame.set_pc(next_pc);
            } else {
                return Err(5);
            }
        }
        frame.clear_scratch();
        Ok(SUSPENDED)
    }
    fn instruction(
        &self,
        frame: &mut Activation<'_>,
        context: &mut TypedContext,
        inst: &OxInst,
    ) -> Result<(), i32> {
        if oxvba_rt_abi::host_entry::value::needs_helper(inst, &self.layout) {
            // SAFETY: frame and original instruction share this checked layout;
            // the helper executes one selected value kernel, never a CFG/call.
            let status = unsafe {
                oxvba_rt_abi::host_entry::value::evaluate(frame.raw_mut(), inst, context)
            };
            return if status == 0 { Ok(()) } else { Err(status) };
        }
        let (dst, value) = match inst {
            OxInst::StmtBoundary {
                clear_temps_from, ..
            } => {
                frame.clear_temps(*clear_temps_from);
                return Ok(());
            }
            OxInst::DrainTerminations => return Ok(()),
            OxInst::ValidateAssignment {
                src,
                target_type_name,
                ..
            } => {
                let pointer = frame.read(src, ScalarType::HostObject)?.as_object();
                if let Some(iid) = self.layout.assignment_iids[target_type_name] {
                    // SAFETY: admitted object reads retain factory tables.
                    if !pointer.is_null() && unsafe { (*pointer).iid } != iid {
                        return Err(13);
                    }
                }
                return Ok(());
            }
            OxInst::Assign { dst, value }
            | OxInst::Coerce {
                dst, src: value, ..
            } => {
                let ty = self.layout.slots[self.layout.slot(*dst).expect("admitted place")];
                let bits = frame.read(value, ty)?;
                if matches!(inst, OxInst::Coerce { .. })
                    && ty == ScalarType::Double
                    && !bits.as_double().is_finite()
                {
                    return Err(6);
                }
                (*dst, bits)
            }
            OxInst::ComCallEarly {
                dst,
                method,
                recv,
                args,
                ..
            } => {
                let plan = &self.layout.host_calls[method];
                let header = frame.read(recv, ScalarType::HostObject)?.as_object();
                let mut values = [ScalarBits::ZERO; calcvba_host_api::MAX_TYPED_ARGS];
                for (i, (arg, ty)) in args.iter().zip(&plan.parameters).enumerate() {
                    values[i] = match arg {
                        oxvba_oxir::OxCallArg::Operand(op) => frame.read(op, ty.carrier())?,
                        oxvba_oxir::OxCallArg::ByRef(place) => {
                            frame.read(&OxOperand::Use(*place), ty.carrier())?
                        }
                        oxvba_oxir::OxCallArg::Const(value) => ScalarBits::long(*value),
                        _ => unreachable!("admitted argument"),
                    };
                }
                // SAFETY: the frame or its scratch owner retains the receiver;
                // no borrow into globals spans the synchronous host callback.
                let value = unsafe {
                    calcvba_host_api::object::value_slots::invoke_value_slot(
                        header,
                        plan.iid,
                        plan.slot,
                        &plan.parameters,
                        plan.result,
                        &values[..args.len()],
                        context,
                    )
                }?;
                if let Some(iid) = plan.result_iid
                    && let calcvba_host_api::Value::Object(Some(object)) = &value
                    && object.header().iid != iid
                {
                    return Err(13);
                }
                if let Some(dst) = dst {
                    frame.write_owned_value(*dst, value)?;
                }
                return Ok(());
            }
            OxInst::Arith {
                dst, op, lhs, rhs, ..
            } => {
                let ty = self.layout.slots[self.layout.slot(*dst).expect("admitted place")];
                let a = frame.read(lhs, ty)?;
                let b = frame.read(rhs, ty)?;
                let value = if ty == ScalarType::Double {
                    let (a, b) = (a.as_double(), b.as_double());
                    let value = match op {
                        ArithOp::Add => a + b,
                        ArithOp::Sub => a - b,
                        ArithOp::Mul => a * b,
                        _ => unreachable!("admitted arithmetic"),
                    };
                    if !value.is_finite() {
                        return Err(6);
                    }
                    ScalarBits::double(value)
                } else {
                    let (a, b) = (a.as_long(), b.as_long());
                    ScalarBits::long(
                        match op {
                            ArithOp::Add => a.checked_add(b),
                            ArithOp::Sub => a.checked_sub(b),
                            ArithOp::Mul => a.checked_mul(b),
                            _ => unreachable!("admitted arithmetic"),
                        }
                        .ok_or(6)?,
                    )
                };
                (*dst, value)
            }
            OxInst::Div { dst, lhs, rhs } => {
                let a = frame.read(lhs, ScalarType::Double)?.as_double();
                let b = frame.read(rhs, ScalarType::Double)?.as_double();
                if b == 0.0 {
                    return Err(11);
                }
                let value = a / b;
                if !value.is_finite() {
                    return Err(6);
                }
                (*dst, ScalarBits::double(value))
            }
            OxInst::Neg { dst, src, .. } => {
                let ty = self.layout.slots[self.layout.slot(*dst).expect("admitted place")];
                let bits = frame.read(src, ty)?;
                (
                    *dst,
                    if ty == ScalarType::Double {
                        ScalarBits::double(-bits.as_double())
                    } else {
                        ScalarBits::long(bits.as_long().checked_neg().ok_or(6)?)
                    },
                )
            }
            OxInst::Compare {
                dst, op, lhs, rhs, ..
            } => {
                let ty = self
                    .layout
                    .comparison_type(lhs, rhs)
                    .expect("admitted comparison");
                let a = frame.read(lhs, ty)?;
                let b = frame.read(rhs, ty)?;
                macro_rules! compare {
                    ($a:expr,$b:expr) => {
                        match op {
                            CmpOp::Eq => $a == $b,
                            CmpOp::Ne => $a != $b,
                            CmpOp::Lt => $a < $b,
                            CmpOp::Le => $a <= $b,
                            CmpOp::Gt => $a > $b,
                            CmpOp::Ge => $a >= $b,
                        }
                    };
                }
                let result = if ty == ScalarType::Double {
                    compare!(a.as_double(), b.as_double())
                } else {
                    compare!(a.as_long(), b.as_long())
                };
                (*dst, ScalarBits::long(if result { -1 } else { 0 }))
            }
            OxInst::Truthy { dst, src } => {
                let ty = self.layout.operand_type(src).expect("admitted operand");
                let bits = frame.read(src, ty)?;
                let truth = if ty == ScalarType::Double {
                    bits.as_double() != 0.0
                } else {
                    bits.as_long() != 0
                };
                (*dst, ScalarBits::long(if truth { -1 } else { 0 }))
            }
            _ => unreachable!("admitted instruction"),
        };
        // SAFETY: the admitted instruction produced the exact carrier and
        // frame/scratch owners retain every object through this store.
        unsafe { frame.write(dst, value) }
    }
}
