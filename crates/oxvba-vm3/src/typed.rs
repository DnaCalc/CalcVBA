//! Bounded unboxed scalar interpreter entry. No Variant frame or per-call heap allocation.
use oxvba_oxir::{ArithOp, CmpOp, OxFunc, OxInst, OxTerminator};
use oxvba_rt_abi::host_entry::{
    MAX_TYPED_SLOTS, ScalarBits, ScalarLayout, ScalarType, TypedContext,
};

pub struct TypedInterpreter<'p> {
    func: &'p OxFunc,
    layout: ScalarLayout,
    blocks: Vec<(oxvba_oxir::BlockId, usize)>,
}
impl<'p> TypedInterpreter<'p> {
    pub fn bind(func: &'p OxFunc) -> Result<Self, String> {
        let layout = ScalarLayout::new(func)?;
        let blocks = func
            .blocks
            .iter()
            .enumerate()
            .map(|(i, b)| (b.id, i))
            .collect();
        Ok(Self {
            func,
            layout,
            blocks,
        })
    }
    pub fn layout(&self) -> &ScalarLayout {
        &self.layout
    }
    pub fn invoke(
        &self,
        args: &[ScalarBits],
        context: &mut TypedContext,
    ) -> Result<ScalarBits, i32> {
        if args.len() != self.layout.parameters.len() {
            return Err(13);
        }
        let mut slots = [ScalarBits(0); MAX_TYPED_SLOTS];
        slots[..args.len()].copy_from_slice(args);
        let mut block = self.func.entry;
        loop {
            context.step()?;
            let block_index = self.blocks.iter().find(|(id, _)| *id == block).ok_or(5)?.1;
            let current = &self.func.blocks[block_index];
            for inst in &current.instrs {
                match inst {
                    OxInst::StmtBoundary {
                        clear_temps_from, ..
                    } => {
                        slots[self.layout.local_count + clear_temps_from..self.layout.slots.len()]
                            .fill(ScalarBits(0));
                    }
                    OxInst::DrainTerminations => {}
                    OxInst::Assign { dst, value }
                    | OxInst::Coerce {
                        dst, src: value, ..
                    } => {
                        let slot = self.layout.slot(*dst).expect("admitted destination");
                        slots[slot] = self.layout.read(&slots, value, self.layout.slots[slot]);
                        if matches!(inst, OxInst::Coerce { .. })
                            && self.layout.slots[slot] == ScalarType::Double
                            && !slots[slot].as_double().is_finite()
                        {
                            return Err(6);
                        }
                    }
                    OxInst::Arith {
                        dst, op, lhs, rhs, ..
                    } => {
                        let slot = self.layout.slot(*dst).expect("admitted destination");
                        let ty = self.layout.slots[slot];
                        let lhs = self.layout.read(&slots, lhs, ty);
                        let rhs = self.layout.read(&slots, rhs, ty);
                        slots[slot] = match ty {
                            ScalarType::Double => {
                                let (a, b) = (lhs.as_double(), rhs.as_double());
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
                            }
                            ScalarType::Long => {
                                let (a, b) = (lhs.as_long(), rhs.as_long());
                                let value = match op {
                                    ArithOp::Add => a.checked_add(b),
                                    ArithOp::Sub => a.checked_sub(b),
                                    ArithOp::Mul => a.checked_mul(b),
                                    _ => unreachable!("admitted arithmetic"),
                                }
                                .ok_or(6)?;
                                ScalarBits::long(value)
                            }
                            _ => unreachable!("admitted numeric type"),
                        };
                    }
                    OxInst::Div { dst, lhs, rhs } => {
                        let a = self
                            .layout
                            .read(&slots, lhs, ScalarType::Double)
                            .as_double();
                        let b = self
                            .layout
                            .read(&slots, rhs, ScalarType::Double)
                            .as_double();
                        if b == 0.0 {
                            return Err(11);
                        }
                        let value = a / b;
                        if !value.is_finite() {
                            return Err(6);
                        }
                        slots[self.layout.slot(*dst).expect("admitted destination")] =
                            ScalarBits::double(value);
                    }
                    OxInst::Compare {
                        dst, op, lhs, rhs, ..
                    } => {
                        let ty = self
                            .layout
                            .comparison_type(lhs, rhs)
                            .expect("admitted comparison");
                        let a = self.layout.read(&slots, lhs, ty);
                        let b = self.layout.read(&slots, rhs, ty);
                        macro_rules! compare {
                            ($a:expr, $b:expr) => {
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
                        let value = if ty == ScalarType::Double {
                            compare!(a.as_double(), b.as_double())
                        } else {
                            compare!(a.as_long(), b.as_long())
                        };
                        slots[self.layout.slot(*dst).expect("admitted destination")] =
                            ScalarBits::long(if value { -1 } else { 0 });
                    }
                    OxInst::Truthy { dst, src } => {
                        let ty = self.layout.operand_type(src).expect("admitted operand");
                        let value = self.layout.read(&slots, src, ty);
                        let truth = if ty == ScalarType::Double {
                            value.as_double() != 0.0
                        } else {
                            value.as_long() != 0
                        };
                        slots[self.layout.slot(*dst).expect("admitted destination")] =
                            ScalarBits::long(if truth { -1 } else { 0 });
                    }
                    _ => unreachable!("admitted instruction"),
                }
            }
            block = match &current.terminator {
                OxTerminator::Jump(next) => *next,
                OxTerminator::Branch {
                    cond,
                    then_blk,
                    else_blk,
                } => {
                    if self
                        .layout
                        .read(&slots, cond, ScalarType::Boolean)
                        .as_long()
                        != 0
                    {
                        *then_blk
                    } else {
                        *else_blk
                    }
                }
                OxTerminator::Return => return Ok(slots[self.layout.return_slot]),
                _ => return Err(5),
            };
        }
    }
}
