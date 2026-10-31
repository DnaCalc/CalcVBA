//! Versioned, allocation-free scalar host-entry convention and admission facts.
//! This is a bounded entry family, not an alternative semantic IR. Backends
//! consume the original OxFunc and the same checked admission facts.
use oxvba_bundle::{NumericCoerceTarget as N, NumericMode};
use oxvba_oxir::{
    ArithOp, OxCoerceTarget, OxConst, OxFunc, OxInst, OxOperand, OxPlace, OxTerminator, OxTy,
};

pub use calcvba_host_api::{
    Arguments, MAX_TYPED_ARGS, Scalar, ScalarBits, ScalarType, TYPED_ENTRY_ABI, TypedContext,
};
pub const MAX_TYPED_SLOTS: usize = 128;

fn scalar_type(ty: &OxTy) -> Option<ScalarType> {
    match ty {
        OxTy::Long => Some(ScalarType::Long),
        OxTy::Double => Some(ScalarType::Double),
        OxTy::Bool => Some(ScalarType::Boolean),
        OxTy::Object(_) => Some(ScalarType::HostObject),
        _ => None,
    }
}

/// Transfer the caller's budget into a stack-owned typed activation.
/// # Safety
/// State follows RawExecState's live unique same-thread contract; context is a
/// nonaliasing, aligned, writable TypedContext for this synchronous call.
pub unsafe extern "C" fn typed_context_start(
    state: *mut crate::RawExecState,
    context: *mut TypedContext,
) -> i32 {
    crate::with_status(|| {
        // SAFETY: the generated adapter retains unique state/context storage.
        let Some(state) = (unsafe { crate::state_from_raw(state) }) else {
            return crate::ST_FAULT;
        };
        if context.is_null() {
            return crate::ST_FAULT;
        }
        // SAFETY: caller supplies aligned writable stack storage.
        unsafe {
            (*context).remaining_blocks = state.remaining_steps.unwrap_or(u64::MAX);
        }
        crate::ST_OK
    })
}

/// Rejoin the dynamic error/budget convention after a typed activation.
/// # Safety
/// State and context follow typed_context_start's contract and remain live;
/// context contains the budget returned by that activation. A nonempty source
/// names one live readable allocation of source_len UTF-8 bytes.
pub unsafe extern "C" fn typed_context_finish(
    state: *mut crate::RawExecState,
    context: *const TypedContext,
    status: i32,
    source: *const u8,
    source_len: usize,
) -> i32 {
    crate::with_status(|| {
        if context.is_null() {
            return crate::ST_FAULT;
        }
        // SAFETY: caller owns both nonaliasing live objects for the call.
        let Some(exec) = (unsafe { crate::state_from_raw(state) }) else {
            return crate::ST_FAULT;
        };
        if exec.remaining_steps.is_some() {
            // SAFETY: the context is live and generated entry has returned.
            exec.remaining_steps = Some(unsafe { (*context).remaining_blocks });
        }
        match status {
            0 => crate::ST_OK,
            -1 => crate::ST_ABORT,
            number => {
                let source = if source_len == 0 {
                    ""
                } else {
                    if source.is_null() {
                        return crate::ST_FAULT;
                    }
                    // SAFETY: generated code retains this UTF-8 program name
                    // for the image lifetime; caller guarantees its full extent.
                    let bytes = unsafe { std::slice::from_raw_parts(source, source_len) };
                    let Ok(source) = std::str::from_utf8(bytes) else {
                        return crate::ST_FAULT;
                    };
                    source
                };
                let mut fault = crate::Fault::new(number, crate::default_error_message(number));
                fault.source = Some(source.into());
                exec.err_engine
                    .raise(fault, exec.default_error_source.clone());
                crate::ST_FAULT
            }
        }
    })
}

/// Inspectable physical layout for a scalar entry. Construction checks every
/// instruction, including unreachable blocks. Unsupported bodies are refused.
#[derive(Debug, Clone)]
pub struct ScalarLayout {
    pub slots: Vec<ScalarType>,
    pub local_count: usize,
    pub parameters: Vec<ScalarType>,
    pub result: ScalarType,
    pub return_slot: usize,
    pub specialized_variant_temps: Vec<usize>,
    pub host_calls: std::collections::HashMap<oxvba_oxir::ComMethodRef, HostCallPlan>,
}
#[derive(Debug, Clone)]
pub struct HostCallPlan {
    pub iid: [u8; 16],
    pub slot: usize,
    pub long_argument: bool,
}
impl ScalarLayout {
    pub fn new(func: &OxFunc) -> Result<Self, String> {
        Self::with_interfaces(func, &[])
    }
    pub fn with_interfaces(
        func: &OxFunc,
        interfaces: &[oxvba_oxir::ComInterface],
    ) -> Result<Self, String> {
        use oxvba_com::{TypeLibParamType as P, TypeLibWireType};
        let mut host_calls = std::collections::HashMap::new();
        for block in &func.blocks {
            for inst in &block.instrs {
                let OxInst::ComCallEarly {
                    method,
                    invoke_kind,
                    ..
                } = inst
                else {
                    continue;
                };
                let descriptor = interfaces
                    .get(method.iface.0)
                    .and_then(|i| i.com_members())
                    .and_then(|m| m.get(method.member))
                    .ok_or("typed host call has no retained interface descriptor")?;
                if descriptor.invoke_kind != *invoke_kind
                    || !matches!(descriptor.parameter_types.as_slice(), [] | [P::Long])
                    || descriptor.return_type != Some(P::Double)
                    || descriptor.parameter_optional.iter().any(|p| *p)
                    || descriptor
                        .parameter_wire_types
                        .iter()
                        .any(|w| *w != TypeLibWireType::Automation(P::Long))
                    || descriptor
                        .return_wire_type
                        .as_ref()
                        .is_some_and(|w| *w != TypeLibWireType::Automation(P::Double))
                {
                    return Err("typed host slot requires Double result and zero or one required ByVal Long".into());
                }
                let iid = descriptor
                    .interface_iid
                    .ok_or("typed host slot requires interface identity")?;
                let mut bytes = [0u8; 16];
                bytes[..4].copy_from_slice(&iid.data1.to_le_bytes());
                bytes[4..6].copy_from_slice(&iid.data2.to_le_bytes());
                bytes[6..8].copy_from_slice(&iid.data3.to_le_bytes());
                bytes[8..].copy_from_slice(&iid.data4);
                let slot = usize::from(descriptor.vtable_slot.ok_or("typed host slot is missing")?);
                if descriptor
                    .vtable_slot_bound
                    .is_none_or(|bound| slot >= usize::from(bound))
                {
                    return Err("typed host slot exceeds declared table".into());
                }
                host_calls.insert(
                    *method,
                    HostCallPlan {
                        iid: bytes,
                        slot,
                        long_argument: !descriptor.parameter_types.is_empty(),
                    },
                );
            }
        }
        let scalar = |ty: &OxTy| {
            scalar_type(ty).ok_or_else(|| format!("typed entry carrier {ty:?} is not supported"))
        };
        if func.param_count > MAX_TYPED_ARGS
            || func.param_count > func.locals.len()
            || func.locals.len().saturating_add(func.temps.len()) > MAX_TYPED_SLOTS
        {
            return Err("typed entry exceeds argument/slot bounds".into());
        }
        let result = func
            .return_local
            .ok_or("typed scalar entries require a Function/Property Get result")?;
        let mut physical = func
            .locals
            .iter()
            .map(|l| scalar(&l.ty).map(Some))
            .collect::<Result<Vec<_>, _>>()?;
        physical.extend(func.temps.iter().map(scalar_type));
        // Widening arithmetic is semantically Variant, but numeric operands
        // with at least one Double prove a Double payload. Specialize only compiler temporaries whose writers
        // prove Double; never a user-declared Variant or Long-only widening.
        loop {
            let mut changed = false;
            for block in &func.blocks {
                for inst in &block.instrs {
                    if let OxInst::ComCallEarly {
                        dst: Some(OxPlace::Temp(id)),
                        method,
                        ..
                    } = inst
                        && host_calls.contains_key(method)
                    {
                        let index = func
                            .locals
                            .len()
                            .checked_add(id.0)
                            .ok_or("invalid temp index")?;
                        let slot = physical.get_mut(index).ok_or("invalid temp index")?;
                        if slot.is_none() {
                            *slot = Some(ScalarType::Double);
                            changed = true;
                        }
                    }
                    if let OxInst::Arith {
                        dst: OxPlace::Temp(id),
                        lhs,
                        rhs,
                        mode: NumericMode::Widening,
                        ..
                    } = inst
                    {
                        let index = func
                            .locals
                            .len()
                            .checked_add(id.0)
                            .ok_or("invalid temp index")?;
                        let Some(slot) = physical.get(index) else {
                            return Err("invalid temp index".into());
                        };
                        if slot.is_some() {
                            continue;
                        }
                        let ty = |o: &OxOperand| match o {
                            OxOperand::Use(OxPlace::Local(i)) if i.0 < func.locals.len() => {
                                physical[i.0]
                            }
                            OxOperand::Use(OxPlace::Temp(i)) => func
                                .locals
                                .len()
                                .checked_add(i.0)
                                .and_then(|i| physical.get(i))
                                .copied()
                                .flatten(),
                            OxOperand::Const(OxConst::F64(_)) => Some(ScalarType::Double),
                            OxOperand::Const(OxConst::I16(_) | OxConst::I32(_)) => {
                                Some(ScalarType::Long)
                            }
                            _ => None,
                        };
                        let (left, right) = (ty(lhs), ty(rhs));
                        let numeric =
                            |ty| matches!(ty, Some(ScalarType::Long | ScalarType::Double));
                        if numeric(left)
                            && numeric(right)
                            && (left == Some(ScalarType::Double)
                                || right == Some(ScalarType::Double))
                        {
                            physical[index] = Some(ScalarType::Double);
                            changed = true;
                        }
                    }
                }
            }
            if !changed {
                break;
            }
        }
        let slots = physical
            .into_iter()
            .collect::<Option<Vec<_>>>()
            .ok_or("typed entry has an unproven carrier")?;
        let parameters = slots[..func.param_count].to_vec();
        if parameters.contains(&ScalarType::Boolean) {
            return Err("typed Boolean parameter ABI remains open".into());
        }
        for local in func.locals.iter().take(func.param_count) {
            let p = local.param.ok_or("missing parameter metadata")?;
            if p.by_ref || p.optional || p.variadic {
                return Err("typed entry requires required ByVal parameters".into());
            }
        }
        let specialized_variant_temps = func
            .temps
            .iter()
            .enumerate()
            .filter(|(_, ty)| **ty == OxTy::Variant)
            .map(|(i, _)| func.locals.len() + i)
            .collect();
        let layout = Self {
            result: *slots.get(result.0).ok_or("invalid return slot")?,
            slots,
            local_count: func.locals.len(),
            parameters,
            return_slot: result.0,
            specialized_variant_temps,
            host_calls,
        };
        if matches!(layout.result, ScalarType::Boolean | ScalarType::HostObject) {
            return Err("typed Boolean/object result ABI remains open".into());
        }
        let ids = func
            .blocks
            .iter()
            .map(|b| b.id)
            .collect::<std::collections::HashSet<_>>();
        if ids.len() != func.blocks.len() || !ids.contains(&func.entry) {
            return Err("invalid typed block identity".into());
        }
        for block in &func.blocks {
            if block.fault_target.is_some_and(|id| !ids.contains(&id))
                || oxvba_oxir::terminator_successors(&block.terminator)
                    .iter()
                    .any(|id| !ids.contains(id))
            {
                return Err("invalid typed block target".into());
            }
            for inst in &block.instrs {
                layout.check(inst)?;
            }
            match &block.terminator {
                OxTerminator::Jump(_)
                | OxTerminator::Return
                | OxTerminator::FaultDispatch { .. }
                | OxTerminator::Unreachable => {}
                OxTerminator::Branch { cond, .. }
                    if layout.operand_type(cond)? == ScalarType::Boolean => {}
                _ => return Err("typed entry terminator requires the dynamic runtime".into()),
            }
        }
        Ok(layout)
    }
    pub fn slot(&self, place: OxPlace) -> Result<usize, String> {
        let index = match place {
            OxPlace::Local(id) if id.0 < self.local_count => id.0,
            OxPlace::Temp(id) => self
                .local_count
                .checked_add(id.0)
                .ok_or("typed slot overflow")?,
            _ => return Err("typed entry cannot access project storage".into()),
        };
        self.slots.get(index).ok_or("typed slot out of bounds")?;
        Ok(index)
    }
    pub fn operand_type(&self, value: &OxOperand) -> Result<ScalarType, String> {
        match value {
            OxOperand::Use(p) => Ok(self.slots[self.slot(*p)?]),
            OxOperand::Const(OxConst::F64(_)) => Ok(ScalarType::Double),
            OxOperand::Const(OxConst::I16(_) | OxConst::I32(_)) => Ok(ScalarType::Long),
            OxOperand::Const(OxConst::Bool(_)) => Ok(ScalarType::Boolean),
            _ => Err("typed operand requires the dynamic runtime".into()),
        }
    }
    pub fn comparison_type(&self, lhs: &OxOperand, rhs: &OxOperand) -> Result<ScalarType, String> {
        let left = self.operand_type(lhs)?;
        let right = self.operand_type(rhs)?;
        if left == ScalarType::HostObject || right == ScalarType::HostObject {
            return Err("typed object value comparison requires default-member semantics".into());
        }
        if left == right {
            return Ok(left);
        }
        if matches!(
            (left, right),
            (ScalarType::Long, ScalarType::Double) | (ScalarType::Double, ScalarType::Long)
        ) {
            return Ok(ScalarType::Double);
        }
        Err("typed comparison requires compatible scalar operands".into())
    }
    fn accepts(&self, dst: OxPlace, src: &OxOperand) -> Result<(), String> {
        let slot = self.slot(dst)?;
        let to = self.slots[slot];
        let from = self.operand_type(src)?;
        if self.specialized_variant_temps.contains(&slot) && from != ScalarType::Double {
            return Err("specialized temporary needs a proven Double writer".into());
        }
        if to == from || (to == ScalarType::Double && from == ScalarType::Long) {
            Ok(())
        } else {
            Err("typed narrowing/coercion requires the dynamic runtime".into())
        }
    }
    fn check(&self, inst: &OxInst) -> Result<(), String> {
        match inst {
            OxInst::ComCallEarly {
                dst: Some(dst),
                method,
                recv,
                args,
                ..
            } => {
                let plan = self.host_calls.get(method).ok_or("unplanned host slot")?;
                if self.slots[self.slot(*dst)?] != ScalarType::Double
                    || self.operand_type(recv)? != ScalarType::HostObject
                    || args.len() != usize::from(plan.long_argument)
                {
                    return Err("typed host call signature mismatch".into());
                }
                if let Some(arg) = args.first() {
                    match arg {
                        oxvba_oxir::OxCallArg::Operand(op)
                            if self.operand_type(op)? == ScalarType::Long => {}
                        oxvba_oxir::OxCallArg::ByRef(place)
                            if self.slots[self.slot(*place)?] == ScalarType::Long => {}
                        oxvba_oxir::OxCallArg::Const(_) => {}
                        _ => return Err("typed host argument requires a Long value".into()),
                    }
                }
                Ok(())
            }
            OxInst::StmtBoundary {
                clear_temps_from, ..
            } if *clear_temps_from <= self.slots.len() - self.local_count => Ok(()),
            OxInst::DrainTerminations => Ok(()),
            OxInst::Assign { dst, value } => self.accepts(*dst, value),
            OxInst::Coerce {
                dst,
                src,
                target: OxCoerceTarget::Numeric(N::Double),
            } if self.slots[self.slot(*dst)?] == ScalarType::Double => self.accepts(*dst, src),
            OxInst::Coerce {
                dst,
                src,
                target: OxCoerceTarget::Numeric(N::Long),
            } if self.slots[self.slot(*dst)?] == ScalarType::Long => self.accepts(*dst, src),
            OxInst::Arith {
                dst,
                lhs,
                rhs,
                mode,
                op,
            } => {
                let ty = self.slots[self.slot(*dst)?];
                if !matches!(
                    (ty, mode),
                    (ScalarType::Long, NumericMode::Checked(N::Long))
                        | (
                            ScalarType::Double,
                            NumericMode::Checked(N::Double) | NumericMode::Widening
                        )
                ) || !matches!(op, ArithOp::Add | ArithOp::Sub | ArithOp::Mul)
                {
                    return Err("typed arithmetic mode requires the dynamic runtime".into());
                }
                let left = self.operand_type(lhs)?;
                let right = self.operand_type(rhs)?;
                if matches!(mode, NumericMode::Widening)
                    && (left != ScalarType::Double && right != ScalarType::Double)
                {
                    return Err("typed widening needs a proven Double operand".into());
                }
                // Arithmetic operands are read after numeric promotion. The
                // writer proof concerns the result, not each operand's tag.
                for from in [left, right] {
                    if from != ty && !(ty == ScalarType::Double && from == ScalarType::Long) {
                        return Err("typed arithmetic needs compatible numeric operands".into());
                    }
                }
                Ok(())
            }
            OxInst::Div { dst, lhs, rhs } if self.slots[self.slot(*dst)?] == ScalarType::Double => {
                self.accepts(*dst, lhs)?;
                self.accepts(*dst, rhs)
            }
            OxInst::Compare { dst, lhs, rhs, .. }
                if self.slots[self.slot(*dst)?] == ScalarType::Boolean =>
            {
                self.comparison_type(lhs, rhs).map(|_| ())
            }
            OxInst::Truthy { dst, src } if self.slots[self.slot(*dst)?] == ScalarType::Boolean => {
                if self.operand_type(src)? == ScalarType::HostObject {
                    Err("typed object condition requires default-member semantics".into())
                } else {
                    Ok(())
                }
            }
            _ => Err(format!(
                "typed instruction {inst:?} requires the dynamic runtime"
            )),
        }
    }
    pub fn read(&self, slots: &[ScalarBits], operand: &OxOperand, to: ScalarType) -> ScalarBits {
        let (bits, from) = match operand {
            OxOperand::Use(place) => {
                let slot = self.slot(*place).expect("admitted slot");
                (slots[slot], self.slots[slot])
            }
            OxOperand::Const(OxConst::F64(bits)) => {
                (ScalarBits::from_bits(*bits), ScalarType::Double)
            }
            OxOperand::Const(OxConst::I16(value)) => {
                (ScalarBits::long(i32::from(*value)), ScalarType::Long)
            }
            OxOperand::Const(OxConst::I32(value)) => (ScalarBits::long(*value), ScalarType::Long),
            OxOperand::Const(OxConst::Bool(value)) => (
                ScalarBits::long(if *value { -1 } else { 0 }),
                ScalarType::Boolean,
            ),
            _ => unreachable!("admitted operand"),
        };
        if to == ScalarType::Double && from == ScalarType::Long {
            ScalarBits::double(f64::from(bits.as_long()))
        } else {
            bits
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn dynamic_adapter_keeps_callee_source_and_budget_abort_does_not_replace_err() {
        let host = oxvba_hal::adapters::builder::HostBuilder::new().build();
        let mut state = crate::ExecState::new(&*host);
        state.default_error_source = "Caller".into();
        state.remaining_steps = Some(8);
        let raw = crate::exec_state_as_raw(&mut state);
        let mut context = TypedContext {
            remaining_blocks: 0,
        };
        // SAFETY: unique live state and initialized nonaliasing stack context.
        let status = unsafe { typed_context_start(raw, &mut context) };
        assert_eq!(status, crate::ST_OK);
        context.step().unwrap();
        let source = b"NumericLibrary";
        // SAFETY: the same unique state/context are live, as is the full UTF-8 source slice.
        let status =
            unsafe { typed_context_finish(raw, &context, 6, source.as_ptr(), source.len()) };
        assert_eq!(status, crate::ST_FAULT);
        assert_eq!(state.remaining_steps, Some(7));
        assert_eq!(state.err_engine.err.source, "NumericLibrary");
        assert_eq!(state.err_engine.err.description, "Overflow");
        let before = state.err_engine.err.clone();
        context.remaining_blocks = 0;
        // SAFETY: no state borrow overlaps the synchronous helper call.
        let status =
            unsafe { typed_context_finish(raw, &context, -1, source.as_ptr(), source.len()) };
        assert_eq!(status, crate::ST_ABORT);
        assert_eq!(state.err_engine.err, before);
    }

    #[test]
    fn malformed_scalar_places_and_blocks_are_refused_before_codegen() {
        use oxvba_oxir::{BlockId, LocalId, OxBlock, OxLocal, TempId};
        let mut func = OxFunc {
            name: "Malformed".into(),
            kind: oxvba_bundle::ProcedureKind::Function,
            locals: vec![OxLocal {
                name: "Malformed".into(),
                ty: OxTy::Long,
                array_element: None,
                param: None,
                escaped: false,
            }],
            temps: vec![OxTy::Long],
            param_count: 0,
            return_local: Some(LocalId(0)),
            entry: BlockId(0),
            blocks: vec![OxBlock {
                id: BlockId(0),
                instrs: vec![OxInst::Assign {
                    dst: OxPlace::Local(LocalId(0)),
                    value: OxOperand::temp(TempId(usize::MAX)),
                }],
                fault_target: None,
                terminator: OxTerminator::Return,
            }],
        };
        assert!(ScalarLayout::new(&func).is_err());
        func.blocks[0].instrs.clear();
        func.entry = BlockId(9);
        assert!(ScalarLayout::new(&func).is_err());
        func.entry = BlockId(0);
        func.param_count = 2;
        assert!(ScalarLayout::new(&func).is_err());

        // Double + integer has a proven Double result; Long-only widening
        // does not. A second writer must not invalidate a successful proof.
        func.param_count = 0;
        func.locals[0].ty = OxTy::Double;
        func.temps[0] = OxTy::Variant;
        func.blocks[0].instrs = vec![OxInst::Arith {
            dst: OxPlace::Temp(TempId(0)),
            op: ArithOp::Add,
            lhs: OxOperand::Const(OxConst::F64(1.0f64.to_bits())),
            rhs: OxOperand::Const(OxConst::I32(2)),
            mode: NumericMode::Widening,
        }];
        assert!(ScalarLayout::new(&func).is_ok());
        if let OxInst::Arith { lhs, .. } = &mut func.blocks[0].instrs[0] {
            *lhs = OxOperand::Const(OxConst::I32(1));
        }
        assert!(ScalarLayout::new(&func).is_err());
        if let OxInst::Arith { lhs, .. } = &mut func.blocks[0].instrs[0] {
            *lhs = OxOperand::Const(OxConst::F64(1.0f64.to_bits()));
        }
        func.blocks[0].instrs.push(OxInst::Assign {
            dst: OxPlace::Temp(TempId(0)),
            value: OxOperand::Const(OxConst::I32(3)),
        });
        assert!(ScalarLayout::new(&func).is_err());

        // Object value comparisons/conditions require default-value dispatch,
        // never integer truth/comparison of the physical object pointer.
        func.locals.push(OxLocal {
            name: "ObjectValue".into(),
            ty: OxTy::Object(oxvba_oxir::ObjClass::Untyped),
            array_element: None,
            param: None,
            escaped: false,
        });
        func.temps = vec![OxTy::Bool];
        let object = OxOperand::Use(OxPlace::Local(LocalId(1)));
        func.blocks[0].instrs = vec![OxInst::Truthy {
            dst: OxPlace::Temp(TempId(0)),
            src: object.clone(),
        }];
        assert!(ScalarLayout::new(&func).is_err());
        func.blocks[0].instrs = vec![OxInst::Compare {
            dst: OxPlace::Temp(TempId(0)),
            op: oxvba_oxir::CmpOp::Eq,
            lhs: object.clone(),
            rhs: object,
            mode: oxvba_bundle::StringCompareMode::Binary,
        }];
        assert!(ScalarLayout::new(&func).is_err());
    }
}
