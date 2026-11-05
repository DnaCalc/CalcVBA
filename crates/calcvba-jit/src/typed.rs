//! Native typed scalar primary entries. Eligible bodies use physical stack
//! slots/registers and an explicit status/result convention, never JitFrame.
use super::*;
use calcvba_rt_abi::host_entry::activation::{
    self, Activation, CONTROL, FINISHED_PC, RawActivation, SUSPENDED,
};
use calcvba_rt_abi::host_entry::{Arguments, Scalar, ScalarLayout, ScalarType, TypedContext};
use std::marker::PhantomData;

pub(crate) struct ResumableEntry {
    pub(crate) pointer: *const u8,
}

pub(crate) struct ScalarEntry {
    pub(crate) pointer: *const u8,
    pub(crate) slots_pointer: *const u8,
    pub(crate) layout: ScalarLayout,
}

/// A typed resolved address can only execute with its original live image.
/// Keeping a stale copy is harmless: invocation checks image identity first.
pub struct TypedEntry<A, R> {
    pointer: *const u8,
    owner: std::sync::Arc<()>,
    signature: PhantomData<fn() -> (A, R)>,
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
        self.invoke_in_context(
            image,
            args,
            &mut TypedContext {
                remaining_blocks: blocks,
            },
        )
    }
    pub fn invoke_in_context(
        &self,
        image: &CompiledImage<'_>,
        args: A,
        context: &mut TypedContext,
    ) -> Result<R, i32> {
        if !std::sync::Arc::ptr_eq(&self.owner, &image.typed_identity) {
            return Err(5);
        }
        if !args.is_valid() {
            return Err(6);
        }
        let mut result = R::default();
        // SAFETY: bind checked A/R against the compiler's exact typed signature;
        // image identity proves the original executable owner is still live.
        let status = unsafe { args.invoke(self.pointer, context, &mut result) };
        if status == 0 { Ok(result) } else { Err(status) }
    }
}
impl CompiledImage<'_> {
    /// # Safety
    /// frame was admitted and initialized for exactly func in this image; its
    /// canonical global store and factory-object owners remain leased.
    pub unsafe fn resume_activation(
        &self,
        func: FuncId,
        frame: &mut Activation<'_>,
        context: &mut TypedContext,
        operations: u64,
    ) -> Result<i32, i32> {
        // SAFETY: forwards this exact admitted frame to the root unit.
        unsafe {
            self.resume_activation_in(
                self.resumable_functions.len() - 1,
                func,
                frame,
                context,
                operations,
            )
        }
    }
    /// # Safety
    /// Same frame/store contract as resume_activation, for exactly this unit.
    pub unsafe fn resume_activation_in(
        &self,
        unit: usize,
        func: FuncId,
        frame: &mut Activation<'_>,
        context: &mut TypedContext,
        operations: u64,
    ) -> Result<i32, i32> {
        let entry = self
            .resumable_functions
            .get(unit)
            .and_then(|p| p.get(func.0))
            .and_then(Option::as_ref)
            .ok_or(5)?;
        // SAFETY: this exact signature is emitted below; self retains its code.
        let call: unsafe extern "C" fn(*mut TypedContext, *mut RawActivation, u64) -> i32 =
            unsafe { std::mem::transmute(entry.pointer) };
        // SAFETY: caller proves admitted slots, unique frame and live storage.
        let status = unsafe { call(context, frame.raw_mut(), operations) };
        frame.clear_scratch();
        if status == 0 || status == SUSPENDED || status == CONTROL {
            Ok(status)
        } else {
            Err(status)
        }
    }
    /// Slot ABI adapter into the same native primary entry, with no Variant frame.
    pub fn invoke_typed_slots(
        &self,
        func: FuncId,
        args: &[calcvba_host_api::Slot<'_>],
        context: &mut TypedContext,
    ) -> Result<calcvba_host_api::ScalarBits, i32> {
        let entry = self
            .typed_functions
            .last()
            .and_then(|p| p.get(func.0))
            .and_then(Option::as_ref)
            .ok_or(5)?;
        if args.len() != entry.layout.parameters.len()
            || args
                .iter()
                .zip(&entry.layout.parameters)
                .any(|(a, t)| a.ty() != *t)
        {
            return Err(13);
        }
        if args.iter().any(|a| !a.is_valid()) {
            return Err(6);
        }
        let mut bits = [calcvba_host_api::ScalarBits::ZERO; calcvba_host_api::MAX_TYPED_ARGS];
        for (i, arg) in args.iter().enumerate() {
            bits[i] = arg.bits();
        }
        let mut result = calcvba_host_api::ScalarBits::ZERO;
        // SAFETY: generated slots thunk has this exact signature and remains
        // owned by self. Safe Slot borrows retain every referenced host object.
        let call: unsafe extern "C" fn(
            *mut TypedContext,
            *const calcvba_host_api::ScalarBits,
            *mut calcvba_host_api::ScalarBits,
        ) -> i32 = unsafe { std::mem::transmute(entry.slots_pointer) };
        // SAFETY: context, initialized slots and nonaliasing output span entry.
        let status = unsafe { call(context, bits.as_ptr(), &mut result) };
        if status == 0 { Ok(result) } else { Err(status) }
    }
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

pub(crate) type ResumableDefinitions = Vec<Vec<Option<ClifFuncId>>>;
pub(crate) fn compile_resumable_entries(
    module: &mut impl Module,
    programs: &[&OxProgram],
) -> Result<ResumableDefinitions, JitError> {
    let mut entries = Vec::new();
    for (p, program) in programs.iter().enumerate() {
        let mut functions = Vec::new();
        for (f, func) in program.funcs.iter().enumerate() {
            let Ok(layout) = ScalarLayout::with_programs(func, program, programs) else {
                functions.push(None);
                continue;
            };
            let mut signature = module.make_signature();
            signature.params = vec![AbiParam::new(types::I64); 3];
            signature.returns.push(AbiParam::new(types::I32));
            let id = module
                .declare_function(&format!("ox$resume$p{p}$f{f}"), Linkage::Local, &signature)
                .map_err(module_err)?;
            let mut context = module.make_context();
            context.func.signature = signature;
            let mut function_context = FunctionBuilderContext::new();
            {
                let mut b = FunctionBuilder::new(&mut context.func, &mut function_context);
                let entry = b.create_block();
                b.switch_to_block(entry);
                b.append_block_params_for_function_params(entry);
                let params = b.block_params(entry).to_vec();
                let runtime = params[0];
                let frame = params[1];
                b.ins().store(
                    ir::MemFlagsData::new(),
                    params[2],
                    frame,
                    std::mem::offset_of!(RawActivation, remaining_operations) as i32,
                );
                let slice = b.create_sized_stack_slot(StackSlotData::new(
                    StackSlotKind::ExplicitSlot,
                    8,
                    3,
                ));
                b.ins().stack_store(params[2], slice, 0);
                let dispatch = b.create_block();
                b.ins().jump(dispatch, &[]);
                let mut starts = HashMap::new();
                let mut operations = Vec::new();
                let mut pc = 0u32;
                for block in &func.blocks {
                    starts.insert(block.id, pc);
                    let count = u32::try_from(block.instrs.len())
                        .map_err(|_| JitError::unsupported("typed PC overflow"))?
                        .checked_add(2)
                        .ok_or_else(|| JitError::unsupported("typed PC overflow"))?;
                    let end = pc
                        .checked_add(count)
                        .filter(|end| *end != FINISHED_PC)
                        .ok_or_else(|| JitError::unsupported("typed PC overflow"))?;
                    for offset in 0..count {
                        operations.push((pc + offset, block, offset as usize, b.create_block()));
                    }
                    pc = end;
                }
                if pc == FINISHED_PC {
                    return Err(JitError::unsupported("typed PC overflow"));
                }
                let lower = ScalarLower {
                    layout: &layout,
                    slots: PhysicalSlots::Frame(frame),
                    runtime,
                    call_conv: module.target_config().default_call_conv,
                    fault_target: std::cell::Cell::new(None),
                };
                b.switch_to_block(dispatch);
                let current = b.ins().load(
                    types::I32,
                    ir::MemFlagsData::new(),
                    frame,
                    std::mem::offset_of!(RawActivation, pc) as i32,
                );
                let mut switch = cranelift_frontend::Switch::new();
                for (pc, _, _, block) in &operations {
                    switch.set_entry(u128::from(*pc), *block);
                }
                let invalid = b.create_block();
                switch.emit(&mut b, current, invalid);
                b.switch_to_block(invalid);
                let status = b.ins().iconst(types::I32, 5);
                b.ins().return_(&[status]);
                for (pc, block, offset, native) in operations {
                    b.switch_to_block(native);
                    let left = b.ins().stack_load(types::I64, slice, 0);
                    let exhausted = b.ins().icmp_imm(ir::condcodes::IntCC::Equal, left, 0);
                    lower.fail_if(&mut b, exhausted, SUSPENDED);
                    let left = b.ins().iadd_imm(left, -1);
                    b.ins().stack_store(left, slice, 0);
                    b.ins().store(
                        ir::MemFlagsData::new(),
                        left,
                        frame,
                        std::mem::offset_of!(RawActivation, remaining_operations) as i32,
                    );
                    let mut next_pc = None;
                    if offset == 0 {
                        let remaining =
                            b.ins()
                                .load(types::I64, ir::MemFlagsData::new(), runtime, 0);
                        let exhausted = b.ins().icmp_imm(ir::condcodes::IntCC::Equal, remaining, 0);
                        lower.fail_if(&mut b, exhausted, -1);
                        let remaining = b.ins().iadd_imm(remaining, -1);
                        b.ins()
                            .store(ir::MemFlagsData::new(), remaining, runtime, 0);
                    } else if let Some(inst) = block.instrs.get(offset - 1) {
                        if calcvba_rt_abi::host_entry::stack_instruction(inst, &layout)
                            || matches!(inst, OxInst::StmtBoundary { .. })
                        {
                            let status = b.ins().iconst(types::I32, i64::from(CONTROL));
                            b.ins().return_(&[status]);
                            continue;
                        }
                        lower.instruction(&mut b, inst);
                    } else {
                        if calcvba_rt_abi::host_entry::stack_terminator(&block.terminator) {
                            let status = b.ins().iconst(types::I32, i64::from(CONTROL));
                            b.ins().return_(&[status]);
                            continue;
                        }
                        match &block.terminator {
                            OxTerminator::Jump(target) => {
                                next_pc =
                                    Some(b.ins().iconst(types::I32, i64::from(starts[target])));
                            }
                            OxTerminator::Branch {
                                cond,
                                then_blk,
                                else_blk,
                            } => {
                                let test = lower.operand(&mut b, cond, ScalarType::Boolean);
                                let yes = b.ins().iconst(types::I32, i64::from(starts[then_blk]));
                                let no = b.ins().iconst(types::I32, i64::from(starts[else_blk]));
                                next_pc = Some(b.ins().select(test, yes, no));
                            }
                            OxTerminator::Return => {
                                next_pc = Some(b.ins().iconst(types::I32, i64::from(FINISHED_PC)));
                            }
                            _ => {
                                let status = b.ins().iconst(types::I32, 5);
                                b.ins().return_(&[status]);
                                continue;
                            }
                        }
                    }
                    // Release transient global receiver borrows after completing
                    // the instruction, before its continuation can be published.
                    let mut clear = ir::Signature::new(module.target_config().default_call_conv);
                    clear.params.push(AbiParam::new(types::I64));
                    let sig = b.import_signature(clear);
                    let address = b
                        .ins()
                        .iconst(types::I64, activation::scratch_clear as *const u8 as i64);
                    b.ins().call_indirect(sig, address, &[frame]);
                    let next =
                        next_pc.unwrap_or_else(|| b.ins().iconst(types::I32, i64::from(pc + 1)));
                    b.ins().store(
                        ir::MemFlagsData::new(),
                        next,
                        frame,
                        std::mem::offset_of!(RawActivation, pc) as i32,
                    );
                    if offset == block.instrs.len() + 1
                        && matches!(block.terminator, OxTerminator::Return)
                    {
                        let status = b.ins().iconst(types::I32, 0);
                        b.ins().return_(&[status]);
                    } else {
                        b.ins().jump(dispatch, &[]);
                    }
                }
                b.seal_all_blocks();
                b.finalize();
            }
            module
                .define_function(id, &mut context)
                .map_err(module_err)?;
            functions.push(Some(id));
        }
        entries.push(functions);
    }
    Ok(entries)
}

pub(crate) fn machine_type(ty: ScalarType) -> ir::Type {
    match ty {
        ScalarType::Double | ScalarType::Date => types::F64,
        ScalarType::Single => types::F32,
        ScalarType::Long | ScalarType::Boolean | ScalarType::Byte | ScalarType::Integer => {
            types::I32
        }
        ScalarType::LongLong | ScalarType::Currency => types::I64,
        ScalarType::HostObject
        | ScalarType::String
        | ScalarType::Variant
        | ScalarType::Record
        | ScalarType::Array
        | ScalarType::Void
        | ScalarType::Interface(_) => types::I64,
    }
}

pub(crate) type ScalarDefinitions = Vec<Vec<Option<(ClifFuncId, ScalarLayout, ClifFuncId)>>>;
fn compile_slots_thunk(
    module: &mut impl Module,
    p: usize,
    f: usize,
    native: ClifFuncId,
    layout: &ScalarLayout,
) -> Result<ClifFuncId, JitError> {
    let ptr = module.target_config().pointer_type();
    let mut signature = module.make_signature();
    signature.params = vec![AbiParam::new(ptr); 3];
    signature.returns.push(AbiParam::new(types::I32));
    let id = module
        .declare_function(&format!("ox$slots$p{p}$f{f}"), Linkage::Local, &signature)
        .map_err(module_err)?;
    let mut ctx = module.make_context();
    ctx.func.signature = signature;
    let mut function_context = FunctionBuilderContext::new();
    {
        let mut b = FunctionBuilder::new(&mut ctx.func, &mut function_context);
        let block = b.create_block();
        b.switch_to_block(block);
        b.append_block_params_for_function_params(block);
        let params = b.block_params(block).to_vec();
        let mut args = vec![params[0]];
        for (i, ty) in layout.parameters.iter().enumerate() {
            args.push(b.ins().load(
                machine_type(*ty),
                ir::MemFlagsData::new(),
                params[1],
                (i * 8) as i32,
            ));
        }
        args.push(params[2]);
        let native = module.declare_func_in_func(native, b.func);
        let call = b.ins().call(native, &args);
        let status = b.inst_results(call)[0];
        b.ins().return_(&[status]);
        b.seal_all_blocks();
        b.finalize();
    }
    module.define_function(id, &mut ctx).map_err(module_err)?;
    Ok(id)
}
pub(crate) fn compile_scalar_entries(
    module: &mut impl Module,
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
            let Ok(layout) = ScalarLayout::with_interfaces(func, &program.com_interfaces) else {
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
                        b.ins().iconst(machine_type(*ty), 0)
                    };
                    b.ins().stack_store(value, slots, (index * 8) as i32);
                }
                let blocks = func
                    .blocks
                    .iter()
                    .map(|block| (block.id, b.create_block()))
                    .collect::<HashMap<_, _>>();
                let lower = ScalarLower {
                    layout: &layout,
                    slots: PhysicalSlots::Stack(slots),
                    runtime,
                    call_conv: module.target_config().default_call_conv,
                    fault_target: std::cell::Cell::new(None),
                };
                for (index, iid) in layout.parameter_iids.iter().enumerate() {
                    if let Some(iid) = iid {
                        lower.check_interface(&mut b, params[index + 1], *iid);
                    }
                }
                b.ins().jump(blocks[&func.entry], &[]);
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
            let slots_id = compile_slots_thunk(module, p, f, id, &layout)?;
            functions.push(Some((id, layout, slots_id)));
        }
        entries.push(functions);
    }
    Ok(entries)
}

pub(crate) struct ScalarLower<'a> {
    pub(crate) layout: &'a ScalarLayout,
    pub(crate) slots: PhysicalSlots,
    pub(crate) runtime: Value,
    pub(crate) call_conv: cranelift_codegen::isa::CallConv,
    pub(crate) fault_target: std::cell::Cell<Option<ir::Block>>,
}
#[derive(Clone, Copy)]
pub(crate) enum PhysicalSlots {
    Stack(ir::StackSlot),
    Frame(Value),
    /// Each cell contains a native address. ByRef parameters and globals point
    /// at their canonical storage, preserving aliases through direct calls.
    Native(ir::StackSlot),
}
impl ScalarLower<'_> {
    pub(crate) fn return_status(&self, b: &mut FunctionBuilder<'_>, status: Value) {
        if let Some(target) = self.fault_target.get() {
            b.ins().jump(target, &[status.into()]);
        } else {
            b.ins().return_(&[status]);
        }
    }
    pub(crate) fn check_status(&self, b: &mut FunctionBuilder<'_>, status: Value) {
        let failed = b.ins().icmp_imm(ir::condcodes::IntCC::NotEqual, status, 0);
        let fail = b.create_block();
        let next = b.create_block();
        b.ins().brif(failed, fail, &[], next, &[]);
        b.switch_to_block(fail);
        self.return_status(b, status);
        b.switch_to_block(next);
    }

    fn load_slot(&self, b: &mut FunctionBuilder<'_>, index: usize, ty: ScalarType) -> Value {
        match self.slots {
            PhysicalSlots::Native(pointers) => {
                let pointer = b.ins().stack_load(types::I64, pointers, (index * 8) as i32);
                if matches!(ty, ScalarType::Variant | ScalarType::Record) {
                    pointer
                } else if matches!(ty, ScalarType::Boolean | ScalarType::Integer) {
                    let value = b
                        .ins()
                        .load(types::I16, ir::MemFlagsData::new(), pointer, 0);
                    b.ins().sextend(types::I32, value)
                } else if ty == ScalarType::Byte {
                    let value = b.ins().load(types::I8, ir::MemFlagsData::new(), pointer, 0);
                    b.ins().uextend(types::I32, value)
                } else {
                    b.ins()
                        .load(machine_type(ty), ir::MemFlagsData::new(), pointer, 0)
                }
            }
            PhysicalSlots::Stack(slots) => {
                b.ins()
                    .stack_load(machine_type(ty), slots, (index * 8) as i32)
            }
            PhysicalSlots::Frame(frame)
                if index < self.layout.global_base
                    && !self
                        .layout
                        .by_ref_parameters
                        .get(index)
                        .copied()
                        .unwrap_or(false) =>
            {
                b.ins().load(
                    machine_type(ty),
                    ir::MemFlagsData::new(),
                    frame,
                    (std::mem::offset_of!(RawActivation, slots) + index * 8) as i32,
                )
            }
            PhysicalSlots::Frame(frame) => {
                let output = b.create_sized_stack_slot(StackSlotData::new(
                    StackSlotKind::ExplicitSlot,
                    8,
                    3,
                ));
                let ptr = b.ins().stack_addr(types::I64, output, 0);
                let global = b.ins().iconst(types::I64, index as i64);
                let kind = b.ins().iconst(types::I32, i64::from(activation::kind(ty)));
                self.helper(
                    b,
                    activation::slot_read as *const u8,
                    &[frame, global, kind, ptr],
                    &[types::I64, types::I64, types::I32, types::I64],
                );
                b.ins().stack_load(machine_type(ty), output, 0)
            }
        }
    }
    fn check_interface(&self, b: &mut FunctionBuilder<'_>, pointer: Value, iid: [u8; 16]) {
        let low = b.ins().iconst(
            types::I64,
            i64::from_le_bytes(iid[..8].try_into().expect("IID half")),
        );
        let high = b.ins().iconst(
            types::I64,
            i64::from_le_bytes(iid[8..].try_into().expect("IID half")),
        );
        self.helper(
            b,
            activation::check_object_interface as *const u8,
            &[pointer, low, high],
            &[types::I64; 3],
        );
    }

    fn helper(
        &self,
        b: &mut FunctionBuilder<'_>,
        pointer: *const u8,
        args: &[Value],
        types: &[ir::Type],
    ) {
        let mut signature = ir::Signature::new(self.call_conv);
        signature.params = types.iter().copied().map(AbiParam::new).collect();
        signature.returns.push(AbiParam::new(types::I32));
        let sig = b.import_signature(signature);
        let address = b.ins().iconst(ir::types::I64, pointer as i64);
        let call = b.ins().call_indirect(sig, address, args);
        let status = b.inst_results(call)[0];
        let failed = b.ins().icmp_imm(ir::condcodes::IntCC::NotEqual, status, 0);
        let fail = b.create_block();
        let next = b.create_block();
        b.ins().brif(failed, fail, &[], next, &[]);
        b.switch_to_block(fail);
        self.return_status(b, status);
        b.switch_to_block(next);
    }
    pub(crate) fn store_slot(
        &self,
        b: &mut FunctionBuilder<'_>,
        index: usize,
        ty: ScalarType,
        value: Value,
    ) {
        match self.slots {
            PhysicalSlots::Native(pointers) => {
                let pointer = b.ins().stack_load(types::I64, pointers, (index * 8) as i32);
                let value = if matches!(ty, ScalarType::Boolean | ScalarType::Integer) {
                    b.ins().ireduce(types::I16, value)
                } else if ty == ScalarType::Byte {
                    b.ins().ireduce(types::I8, value)
                } else {
                    value
                };
                b.ins().store(ir::MemFlagsData::new(), value, pointer, 0);
            }
            PhysicalSlots::Stack(slots) => {
                b.ins().stack_store(value, slots, (index * 8) as i32);
            }
            PhysicalSlots::Frame(frame)
                if index < self.layout.global_base
                    && matches!(
                        ty,
                        ScalarType::Long | ScalarType::Double | ScalarType::Boolean
                    )
                    && !self
                        .layout
                        .by_ref_parameters
                        .get(index)
                        .copied()
                        .unwrap_or(false) =>
            {
                b.ins().store(
                    ir::MemFlagsData::new(),
                    value,
                    frame,
                    (std::mem::offset_of!(RawActivation, slots) + index * 8) as i32,
                );
            }
            PhysicalSlots::Frame(frame) => {
                let input = b.create_sized_stack_slot(StackSlotData::new(
                    StackSlotKind::ExplicitSlot,
                    8,
                    3,
                ));
                let zero = b.ins().iconst(types::I64, 0);
                b.ins().stack_store(zero, input, 0);
                b.ins().stack_store(value, input, 0);
                let ptr = b.ins().stack_addr(types::I64, input, 0);
                let index = b.ins().iconst(types::I64, index as i64);
                let kind = b.ins().iconst(types::I32, i64::from(activation::kind(ty)));
                self.helper(
                    b,
                    activation::slot_write as *const u8,
                    &[frame, index, kind, ptr],
                    &[types::I64, types::I64, types::I32, types::I64],
                );
            }
        }
    }
    fn fail_if(&self, b: &mut FunctionBuilder<'_>, condition: Value, number: i32) {
        let fail = b.create_block();
        let next = b.create_block();
        b.ins().brif(condition, fail, &[], next, &[]);
        b.switch_to_block(fail);
        let status = b.ins().iconst(types::I32, i64::from(number));
        self.return_status(b, status);
        b.switch_to_block(next);
    }
    pub(crate) fn operand(
        &self,
        b: &mut FunctionBuilder<'_>,
        operand: &OxOperand,
        to: ScalarType,
    ) -> Value {
        let from = self.layout.operand_type(operand).expect("admitted operand");
        let value = match operand {
            OxOperand::Use(place) => {
                self.load_slot(b, self.layout.slot(*place).expect("admitted slot"), from)
            }
            OxOperand::Const(OxConst::F64(bits)) => b.ins().f64const(Ieee64::with_bits(*bits)),
            OxOperand::Const(OxConst::I16(value)) => b.ins().iconst(types::I32, i64::from(*value)),
            OxOperand::Const(OxConst::I32(value)) => b.ins().iconst(types::I32, i64::from(*value)),
            OxOperand::Const(OxConst::I64(value) | OxConst::Currency(value)) => {
                b.ins().iconst(types::I64, *value)
            }
            OxOperand::Const(OxConst::F32(bits)) => b.ins().f32const(Ieee32::with_bits(*bits)),
            OxOperand::Const(OxConst::Date(bits)) => b.ins().f64const(Ieee64::with_bits(*bits)),
            OxOperand::Const(OxConst::Bool(value)) => {
                b.ins().iconst(types::I32, if *value { -1 } else { 0 })
            }
            OxOperand::Const(OxConst::Nothing) => b.ins().iconst(types::I64, 0),
            OxOperand::Const(OxConst::Str(text)) => {
                let PhysicalSlots::Frame(frame) = self.slots else {
                    unreachable!("owning string literal");
                };
                let output = b.create_sized_stack_slot(StackSlotData::new(
                    StackSlotKind::ExplicitSlot,
                    8,
                    3,
                ));
                let out = b.ins().stack_addr(types::I64, output, 0);
                let bytes = b.ins().iconst(types::I64, text.as_ptr() as i64);
                let length = b.ins().iconst(types::I64, text.len() as i64);
                self.helper(
                    b,
                    activation::literal_string as *const u8,
                    &[frame, bytes, length, out],
                    &[types::I64; 4],
                );
                b.ins().stack_load(types::I64, output, 0)
            }
            OxOperand::Const(OxConst::Empty | OxConst::Null) => {
                let PhysicalSlots::Frame(frame) = self.slots else {
                    unreachable!("owning Variant literal");
                };
                let output = b.create_sized_stack_slot(StackSlotData::new(
                    StackSlotKind::ExplicitSlot,
                    8,
                    3,
                ));
                let out = b.ins().stack_addr(types::I64, output, 0);
                let null = b.ins().iconst(
                    types::I32,
                    i64::from(u8::from(matches!(operand, OxOperand::Const(OxConst::Null)))),
                );
                self.helper(
                    b,
                    activation::constant_variant as *const u8,
                    &[frame, null, out],
                    &[types::I64, types::I32, types::I64],
                );
                b.ins().stack_load(types::I64, output, 0)
            }
        };
        if from == ScalarType::Long && to == ScalarType::Double {
            b.ins().fcvt_from_sint(types::F64, value)
        } else {
            value
        }
    }
    pub(crate) fn finite(&self, b: &mut FunctionBuilder<'_>, value: Value) {
        let abs = b.ins().fabs(value);
        let max = b.ins().f64const(f64::MAX);
        let finite = b
            .ins()
            .fcmp(ir::condcodes::FloatCC::LessThanOrEqual, abs, max);
        let bad = b.ins().icmp_imm(ir::condcodes::IntCC::Equal, finite, 0);
        self.fail_if(b, bad, 6);
    }
    fn host_values(
        &self,
        b: &mut FunctionBuilder<'_>,
        plan: &calcvba_rt_abi::host_entry::HostCallPlan,
        dst: Option<OxPlace>,
        recv: &OxOperand,
        args: &[calcvba_oxir::OxCallArg],
    ) {
        use calcvba_host_api::object::{NativeSlot, ObjectHeader};
        let header = self.operand(b, recv, ScalarType::HostObject);
        let null = b.ins().icmp_imm(ir::condcodes::IntCC::Equal, header, 0);
        self.fail_if(b, null, 91);
        for (offset, chunk) in plan.iid.as_chunks::<8>().0.iter().enumerate() {
            let actual = b.ins().load(
                types::I64,
                ir::MemFlagsData::new(),
                header,
                (std::mem::offset_of!(ObjectHeader, iid) + offset * 8) as i32,
            );
            let wrong = b.ins().icmp_imm(
                ir::condcodes::IntCC::NotEqual,
                actual,
                i64::from_le_bytes(*chunk),
            );
            self.fail_if(b, wrong, 438);
        }
        let count = b.ins().load(
            types::I64,
            ir::MemFlagsData::new(),
            header,
            std::mem::offset_of!(ObjectHeader, slot_count) as i32,
        );
        let invalid = b.ins().icmp_imm(
            ir::condcodes::IntCC::UnsignedLessThanOrEqual,
            count,
            plan.slot as i64,
        );
        self.fail_if(b, invalid, 438);
        let table = b.ins().load(
            types::I64,
            ir::MemFlagsData::new(),
            header,
            std::mem::offset_of!(ObjectHeader, slots) as i32,
        );
        let slot = b.ins().iadd_imm(
            table,
            (plan.slot * std::mem::size_of::<NativeSlot>()) as i64,
        );
        let actual = b.ins().load(
            types::I32,
            ir::MemFlagsData::new(),
            slot,
            std::mem::offset_of!(NativeSlot, signature) as i32,
        );
        let expected = calcvba_host_api::value_slot_signature(&plan.parameters, plan.result)
            .expect("admitted host signature");
        let wrong = b
            .ins()
            .icmp_imm(ir::condcodes::IntCC::NotEqual, actual, i64::from(expected));
        self.fail_if(b, wrong, 13);
        let entry = b.ins().load(
            types::I64,
            ir::MemFlagsData::new(),
            slot,
            std::mem::offset_of!(NativeSlot, entry) as i32,
        );
        let data = b.ins().load(
            types::I64,
            ir::MemFlagsData::new(),
            slot,
            std::mem::offset_of!(NativeSlot, data) as i32,
        );
        let input = b.create_sized_stack_slot(StackSlotData::new(
            StackSlotKind::ExplicitSlot,
            (args.len().max(1) * 8) as u32,
            3,
        ));
        for (i, (arg, ty)) in args.iter().zip(&plan.parameters).enumerate() {
            let value = match arg {
                calcvba_oxir::OxCallArg::Operand(op) => self.operand(b, op, ty.carrier()),
                calcvba_oxir::OxCallArg::ByRef(place) => {
                    self.operand(b, &OxOperand::Use(*place), ty.carrier())
                }
                calcvba_oxir::OxCallArg::Const(value) => {
                    b.ins().iconst(types::I32, i64::from(*value))
                }
                _ => unreachable!("admitted host argument"),
            };
            if let Some(iid) = ty.interface() {
                self.check_interface(b, value, iid);
            }
            let zero = b.ins().iconst(types::I64, 0);
            b.ins().stack_store(zero, input, (i * 8) as i32);
            b.ins().stack_store(value, input, (i * 8) as i32);
        }
        let args = b.ins().stack_addr(types::I64, input, 0);
        let output = b.create_sized_stack_slot(StackSlotData::new(
            StackSlotKind::ExplicitSlot,
            std::mem::size_of::<Variant>().max(8) as u32,
            3,
        ));
        let out = b.ins().stack_addr(types::I64, output, 0);
        let mut signature = ir::Signature::new(self.call_conv);
        signature.params = vec![AbiParam::new(types::I64); 4];
        signature.returns.push(AbiParam::new(types::I32));
        let signature = b.import_signature(signature);
        let call = b
            .ins()
            .call_indirect(signature, entry, &[data, self.runtime, args, out]);
        let status = b.inst_results(call)[0];
        let failed = b.ins().icmp_imm(ir::condcodes::IntCC::NotEqual, status, 0);
        let fail = b.create_block();
        let success = b.create_block();
        b.ins().brif(failed, fail, &[], success, &[]);
        b.switch_to_block(fail);
        self.return_status(b, status);
        b.switch_to_block(success);
        let Some(dst) = dst else {
            if matches!(
                plan.result,
                ScalarType::String | ScalarType::Variant | ScalarType::HostObject
            ) {
                let kind = b
                    .ins()
                    .iconst(types::I32, i64::from(activation::kind(plan.result)));
                self.helper(
                    b,
                    activation::result_drop as *const u8,
                    &[kind, out],
                    &[types::I32, types::I64],
                );
            }
            return;
        };
        let index = self.layout.slot(dst).expect("admitted destination");
        match self.slots {
            PhysicalSlots::Stack(_) | PhysicalSlots::Native(_) => {
                let value = b.ins().stack_load(machine_type(plan.result), output, 0);
                self.store_slot(b, index, plan.result, value);
            }
            PhysicalSlots::Frame(frame) => {
                let index = b.ins().iconst(types::I64, index as i64);
                let kind = b
                    .ins()
                    .iconst(types::I32, i64::from(activation::kind(plan.result)));
                let iid = plan.result_iid.unwrap_or([0; 16]);
                let low = b.ins().iconst(
                    types::I64,
                    i64::from_le_bytes(iid[..8].try_into().expect("IID")),
                );
                let high = b.ins().iconst(
                    types::I64,
                    i64::from_le_bytes(iid[8..].try_into().expect("IID")),
                );
                self.helper(
                    b,
                    activation::result_write_checked as *const u8,
                    &[frame, index, kind, out, low, high],
                    &[
                        types::I64,
                        types::I64,
                        types::I32,
                        types::I64,
                        types::I64,
                        types::I64,
                    ],
                );
            }
        }
    }
    pub(crate) fn instruction(&self, b: &mut FunctionBuilder<'_>, inst: &OxInst) {
        if calcvba_rt_abi::host_entry::value::needs_helper(inst, self.layout) {
            let PhysicalSlots::Frame(frame) = self.slots else {
                unreachable!("admitted owning value operation");
            };
            let descriptor = b.ins().iconst(types::I64, inst as *const OxInst as i64);
            self.helper(
                b,
                calcvba_rt_abi::host_entry::value::evaluate as *const u8,
                &[frame, descriptor, self.runtime],
                &[types::I64; 3],
            );
            return;
        }
        if let OxInst::ComCallEarly {
            method,
            dst,
            recv,
            args,
            ..
        } = inst
        {
            let plan = &self.layout.host_calls[method];
            if plan.result != ScalarType::Double
                || !matches!(plan.parameters.as_slice(), [] | [ScalarType::Long])
            {
                self.host_values(b, plan, *dst, recv, args);
                return;
            }
        }
        let (dst, value) = match inst {
            OxInst::ComCallEarly {
                dst: Some(dst),
                method,
                recv,
                args,
                ..
            } => {
                use calcvba_host_api::object::{
                    NativeSlot, ObjectHeader, SLOT_DOUBLE_0, SLOT_DOUBLE_LONG,
                };
                let plan = &self.layout.host_calls[method];
                let header = self.operand(b, recv, ScalarType::HostObject);
                let null = b.ins().icmp_imm(ir::condcodes::IntCC::Equal, header, 0);
                self.fail_if(b, null, 91);
                for (offset, chunk) in plan.iid.as_chunks::<8>().0.iter().enumerate() {
                    let expected = i64::from_le_bytes(*chunk);
                    let actual = b.ins().load(
                        types::I64,
                        ir::MemFlagsData::new(),
                        header,
                        (std::mem::offset_of!(ObjectHeader, iid) + offset * 8) as i32,
                    );
                    let wrong = b
                        .ins()
                        .icmp_imm(ir::condcodes::IntCC::NotEqual, actual, expected);
                    self.fail_if(b, wrong, 438);
                }
                let count = b.ins().load(
                    types::I64,
                    ir::MemFlagsData::new(),
                    header,
                    std::mem::offset_of!(ObjectHeader, slot_count) as i32,
                );
                let out_of_bounds = b.ins().icmp_imm(
                    ir::condcodes::IntCC::UnsignedLessThanOrEqual,
                    count,
                    plan.slot as i64,
                );
                self.fail_if(b, out_of_bounds, 438);
                let table = b.ins().load(
                    types::I64,
                    ir::MemFlagsData::new(),
                    header,
                    std::mem::offset_of!(ObjectHeader, slots) as i32,
                );
                let native = b.ins().iadd_imm(
                    table,
                    (plan.slot * std::mem::size_of::<NativeSlot>()) as i64,
                );
                let signature = b.ins().load(
                    types::I32,
                    ir::MemFlagsData::new(),
                    native,
                    std::mem::offset_of!(NativeSlot, signature) as i32,
                );
                let expected = if plan.long_argument {
                    SLOT_DOUBLE_LONG
                } else {
                    SLOT_DOUBLE_0
                };
                let wrong = b.ins().icmp_imm(
                    ir::condcodes::IntCC::NotEqual,
                    signature,
                    i64::from(expected),
                );
                self.fail_if(b, wrong, 13);
                let entry = b.ins().load(
                    types::I64,
                    ir::MemFlagsData::new(),
                    native,
                    std::mem::offset_of!(NativeSlot, entry) as i32,
                );
                let data = b.ins().load(
                    types::I64,
                    ir::MemFlagsData::new(),
                    native,
                    std::mem::offset_of!(NativeSlot, data) as i32,
                );
                let result = b.create_sized_stack_slot(StackSlotData::new(
                    StackSlotKind::ExplicitSlot,
                    8,
                    3,
                ));
                let result_ptr = b.ins().stack_addr(types::I64, result, 0);
                let mut call_args = vec![data, self.runtime];
                let mut call_sig = ir::Signature::new(self.call_conv);
                call_sig.params = vec![AbiParam::new(types::I64); 2];
                if let Some(arg) = args.first() {
                    let value = match arg {
                        calcvba_oxir::OxCallArg::Operand(op) => {
                            self.operand(b, op, ScalarType::Long)
                        }
                        calcvba_oxir::OxCallArg::ByRef(place) => {
                            self.operand(b, &OxOperand::Use(*place), ScalarType::Long)
                        }
                        calcvba_oxir::OxCallArg::Const(value) => {
                            b.ins().iconst(types::I32, i64::from(*value))
                        }
                        _ => unreachable!("admitted host argument"),
                    };
                    call_args.push(value);
                    call_sig.params.push(AbiParam::new(types::I32));
                }
                call_args.push(result_ptr);
                call_sig.params.push(AbiParam::new(types::I64));
                call_sig.returns.push(AbiParam::new(types::I32));
                let sig = b.import_signature(call_sig);
                let call = b.ins().call_indirect(sig, entry, &call_args);
                let status = b.inst_results(call)[0];
                let failed = b.ins().icmp_imm(ir::condcodes::IntCC::NotEqual, status, 0);
                let fail = b.create_block();
                let success = b.create_block();
                b.ins().brif(failed, fail, &[], success, &[]);
                b.switch_to_block(fail);
                self.return_status(b, status);
                b.switch_to_block(success);
                let value = b.ins().stack_load(types::F64, result, 0);
                self.finite(b, value);
                (*dst, value)
            }
            OxInst::StmtBoundary {
                clear_temps_from, ..
            } => {
                for i in self.layout.local_count + clear_temps_from..self.layout.global_base {
                    if self.layout.slots[i] == ScalarType::Variant {
                        let PhysicalSlots::Frame(frame) = self.slots else {
                            unreachable!("actual Variant requires owning storage");
                        };
                        let index = b.ins().iconst(types::I64, i as i64);
                        self.helper(
                            b,
                            activation::clear_variant as *const u8,
                            &[frame, index],
                            &[types::I64; 2],
                        );
                        continue;
                    }
                    let value = if self.layout.slots[i] == ScalarType::Double {
                        b.ins().f64const(0.0)
                    } else {
                        b.ins().iconst(machine_type(self.layout.slots[i]), 0)
                    };
                    self.store_slot(b, i, self.layout.slots[i], value);
                }
                return;
            }
            OxInst::DrainTerminations => return,
            OxInst::ValidateAssignment {
                src,
                target_type_name,
                ..
            } => {
                if let Some(iid) = self.layout.assignment_iids[target_type_name] {
                    let pointer = self.operand(b, src, ScalarType::HostObject);
                    self.check_interface(b, pointer, iid);
                }
                return;
            }
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
                    if matches!(op, ArithOp::IntDiv | ArithOp::Mod) {
                        let zero = b.ins().icmp_imm(ir::condcodes::IntCC::Equal, c, 0);
                        self.fail_if(b, zero, 11);
                    }
                    let wide = match op {
                        ArithOp::Add => b.ins().iadd(a, c),
                        ArithOp::Sub => b.ins().isub(a, c),
                        ArithOp::Mul => b.ins().imul(a, c),
                        ArithOp::IntDiv => b.ins().sdiv(a, c),
                        ArithOp::Mod => b.ins().srem(a, c),
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
            OxInst::Neg { dst, src, .. } => {
                let ty = self.layout.slots[self.layout.slot(*dst).expect("admitted destination")];
                let value = self.operand(b, src, ty);
                let value = if ty == ScalarType::Double {
                    b.ins().fneg(value)
                } else {
                    let overflow =
                        b.ins()
                            .icmp_imm(ir::condcodes::IntCC::Equal, value, i64::from(i32::MIN));
                    self.fail_if(b, overflow, 6);
                    b.ins().ineg(value)
                };
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
            OxInst::Not { dst, src } => {
                let ty = self.layout.operand_type(src).expect("admitted Not");
                let value = self.operand(b, src, ty);
                (*dst, b.ins().bnot(value))
            }
            OxInst::Logical { dst, op, lhs, rhs } => {
                let left = self
                    .layout
                    .operand_type(lhs)
                    .expect("admitted logical operand");
                let right = self
                    .layout
                    .operand_type(rhs)
                    .expect("admitted logical operand");
                let a = self.operand(b, lhs, left);
                let c = self.operand(b, rhs, right);
                let value = match op {
                    calcvba_oxir::LogicalOp::And => b.ins().band(a, c),
                    calcvba_oxir::LogicalOp::Or => b.ins().bor(a, c),
                    calcvba_oxir::LogicalOp::Xor => b.ins().bxor(a, c),
                    calcvba_oxir::LogicalOp::Eqv => {
                        let xor = b.ins().bxor(a, c);
                        b.ins().bnot(xor)
                    }
                    calcvba_oxir::LogicalOp::Imp => {
                        let not = b.ins().bnot(a);
                        b.ins().bor(not, c)
                    }
                };
                (*dst, value)
            }
            _ => unreachable!("admitted instruction"),
        };
        let slot = self.layout.slot(dst).expect("admitted destination");
        self.store_slot(b, slot, self.layout.slots[slot], value);
    }
}
