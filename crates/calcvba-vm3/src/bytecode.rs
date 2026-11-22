//! Compact per-function bytecode compiled lazily from typed OxIR for VM3.
//!
//! This representation stays strictly internal to `calcvba-vm3` (OxIR and `.oxi`
//! images are unchanged). It packs hot instructions into a dense 16-byte flat
//! array, compresses empty jump blocks with exact step accounting, fuses `For`
//! loop headers/steps, `Compare + [Truthy] + Branch`, `ArrayGet + RecordGet`,
//! `ArrayGet + [Assign] + RecordSet + ArraySet`, and single-use temp stores,
//! and falls back to the original `OxInst` / `OxTerminator` for rare operations.

use std::cell::Cell;

use calcvba_bundle::{
    ComparisonCoercion, NativeImplId, NumericCoerceTarget, NumericMode, StringCompareMode,
};
use calcvba_com::TypeLibMemberInvokeKind;
use calcvba_oxir::com::ComMethodRef;
use calcvba_oxir::value::{
    ArithOp, CmpOp, LogicalOp, OxArg, OxCallArg, OxCoerceTarget, OxConst, OxOperand, OxPlace,
};
use calcvba_oxir::{FuncId, ImportId, OxFunc, OxInst, OxProgram, OxTerminator, OxTy};
use calcvba_rt_abi::ox_const_variant;
use calcvba_runtime::Variant;
use calcvba_runtime::object_ref::RuntimeMemberDescriptor;

const TAG_SHIFT: u32 = 30;
const INDEX_MASK: u32 = (1u32 << TAG_SHIFT) - 1;

pub(crate) const TAG_TEMP: u32 = 0;
pub(crate) const TAG_LOCAL: u32 = 1;
pub(crate) const TAG_GLOBAL: u32 = 2;
pub(crate) const TAG_CONST: u32 = 3;

/// Compact 4-byte storage location (`Temp` stores its flat slot index `num_locals + temp_id`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct PackedPlace(u32);

impl PackedPlace {
    #[inline]
    pub(crate) fn from_ox(place: OxPlace, num_locals: usize) -> Option<Self> {
        let (tag, idx) = match place {
            OxPlace::Temp(t) => (TAG_TEMP, num_locals.checked_add(t.0)?),
            OxPlace::Local(l) => (TAG_LOCAL, l.0),
            OxPlace::Global(g) => (TAG_GLOBAL, g.0),
        };
        let idx = u32::try_from(idx).ok()?;
        if idx > INDEX_MASK {
            return None;
        }
        Some(Self((tag << TAG_SHIFT) | idx))
    }

    #[inline]
    pub(crate) const fn tag(self) -> u32 {
        self.0 >> TAG_SHIFT
    }

    #[inline]
    pub(crate) const fn index(self) -> usize {
        (self.0 & INDEX_MASK) as usize
    }

    #[inline]
    pub(crate) fn to_ox(self, num_locals: usize) -> OxPlace {
        match self.tag() {
            TAG_TEMP => OxPlace::Temp(calcvba_oxir::TempId(self.index() - num_locals)),
            TAG_LOCAL => OxPlace::Local(calcvba_oxir::LocalId(self.index())),
            _ => OxPlace::Global(calcvba_oxir::GlobalId(self.index())),
        }
    }
}

/// Compact 4-byte operand (`Temp`, `Local`, `Global`, or `Const` index into `CompiledFunc::consts`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct PackedOp(u32);

impl PackedOp {
    #[inline]
    pub(crate) fn from_place(place: PackedPlace) -> Self {
        Self(place.0)
    }

    #[inline]
    pub(crate) fn from_const_idx(idx: u32) -> Option<Self> {
        if idx > INDEX_MASK {
            return None;
        }
        Some(Self((TAG_CONST << TAG_SHIFT) | idx))
    }

    #[inline]
    pub(crate) const fn tag(self) -> u32 {
        self.0 >> TAG_SHIFT
    }

    #[inline]
    pub(crate) const fn index(self) -> usize {
        (self.0 & INDEX_MASK) as usize
    }

    #[inline]
    pub(crate) const fn as_place(self) -> Option<PackedPlace> {
        if self.tag() == TAG_CONST {
            None
        } else {
            Some(PackedPlace(self.0))
        }
    }
}

/// Reference to an original OxIR instruction inside `func.blocks[block].instrs[inst]`.
#[derive(Clone, Copy, Debug)]
pub(crate) struct InstRef {
    pub(crate) block: u32,
    pub(crate) inst: u32,
}

/// Non-owning constant pool entry. Scalars hold a pre-built non-owning `Variant`
/// (zero heap allocations or live-handle counter effects); strings keep pre-encoded
/// UTF-16 code units so BSTR allocations occur at the exact VBA evaluation points
/// without re-encoding UTF-8 on every execution.
pub(crate) enum FastConst {
    Scalar(Variant),
    Str(Box<[u16]>),
    NullStr,
}

pub(crate) struct CmpBranchData {
    pub(crate) lhs: PackedOp,
    pub(crate) rhs: PackedOp,
    pub(crate) op: CmpOp,
    pub(crate) mode: StringCompareMode,
    pub(crate) coercion: ComparisonCoercion,
    pub(crate) then_pc: u32,
    pub(crate) then_block: u32,
    pub(crate) then_steps: u16,
    pub(crate) else_pc: u32,
    pub(crate) else_block: u32,
    pub(crate) else_steps: u16,
}

pub(crate) struct BranchData {
    pub(crate) cond: PackedOp,
    pub(crate) then_pc: u32,
    pub(crate) then_block: u32,
    pub(crate) then_steps: u16,
    pub(crate) else_pc: u32,
    pub(crate) else_block: u32,
    pub(crate) else_steps: u16,
}

/// Fused `For` header (`head -> asc/desc -> test -> body/after`).
/// Emitted right before `head`'s unfused ops so any slow case simply falls through.
pub(crate) struct ForCheckData {
    pub(crate) var: PackedPlace,
    pub(crate) limit: PackedOp,
    pub(crate) nonneg: PackedOp,
    pub(crate) body_pc: u32,
    pub(crate) body_block: u32,
    pub(crate) body_steps: u16,
    pub(crate) after_pc: u32,
    pub(crate) after_block: u32,
    pub(crate) after_steps: u16,
}

/// Fused `For` step + header (`[body -> empty... ->] step -> head -> asc/desc -> test -> body/after`).
/// When `to_step_steps == 0` it sits at the start of `step_block` and falls through on slow path;
/// when `to_step_steps > 0` it replaces the tail jump at the end of the loop body and jumps to
/// `(step_block, slow_pc)` on overflow or non-fast types.
pub(crate) struct ForStepCheckData {
    pub(crate) var: PackedPlace,
    pub(crate) step: PackedOp,
    pub(crate) limit: PackedOp,
    pub(crate) nonneg: PackedOp,
    pub(crate) body_pc: u32,
    pub(crate) body_block: u32,
    pub(crate) after_pc: u32,
    pub(crate) after_block: u32,
    pub(crate) step_block: u32,
    pub(crate) slow_pc: u32,
    pub(crate) to_step_steps: u16,
    pub(crate) cont_steps: u16,
    pub(crate) exit_steps: u16,
}

pub(crate) struct ArrRecGetData {
    pub(crate) dst: PackedPlace,
    pub(crate) array: PackedPlace,
    pub(crate) index: PackedOp,
    pub(crate) field: u32,
    pub(crate) first_inst: InstRef,
    pub(crate) orig_count: u8,
}

pub(crate) struct ArrRecSetData {
    pub(crate) array: PackedPlace,
    pub(crate) index: PackedOp,
    pub(crate) value: PackedOp,
    pub(crate) field: u32,
    pub(crate) first_inst: InstRef,
    pub(crate) orig_count: u8,
}

pub(crate) enum FastArg {
    ByVal(PackedOp),
    ByRef(PackedPlace),
    Omitted,
}

pub(crate) struct FastProcCall {
    pub(crate) dst: Option<PackedPlace>,
    pub(crate) proc: FuncId,
    pub(crate) args: Vec<FastArg>,
    pub(crate) all_byval: bool,
}

pub(crate) struct FastLibCall {
    pub(crate) dst: Option<PackedPlace>,
    pub(crate) id: NativeImplId,
    pub(crate) string_typed_alias: bool,
    pub(crate) pure_no_events: bool,
    pub(crate) args: Vec<FastArg>,
}

pub(crate) enum FastCallArg {
    Operand(PackedOp),
    ByRef(PackedPlace),
    Omitted,
    Const,
}

pub(crate) struct FastComEarly {
    pub(crate) dst: Option<PackedPlace>,
    pub(crate) method: ComMethodRef,
    pub(crate) invoke_kind: TypeLibMemberInvokeKind,
    pub(crate) recv: PackedOp,
    pub(crate) positional_args: Option<Vec<FastCallArg>>,
    pub(crate) inst: InstRef,
    /// Monomorphic inline cache: `(class_descriptor_ptr, resolved_member)`.
    pub(crate) ic_descriptor: Cell<usize>,
    pub(crate) ic_member: Cell<Option<&'static RuntimeMemberDescriptor>>,
}

pub(crate) struct FastComLate {
    pub(crate) dst: Option<PackedPlace>,
    pub(crate) recv: PackedOp,
    pub(crate) name: String,
    pub(crate) default_member: bool,
    pub(crate) invoke_kind: TypeLibMemberInvokeKind,
    pub(crate) public_only: bool,
    pub(crate) positional_args: Option<Vec<FastCallArg>>,
    pub(crate) inst: InstRef,
    /// Monomorphic inline cache: `(class_descriptor_ptr, resolved_member)`.
    pub(crate) ic_descriptor: Cell<usize>,
    pub(crate) ic_member: Cell<Option<&'static RuntimeMemberDescriptor>>,
}

#[derive(Clone, Copy, Debug)]
pub(crate) enum VmOp {
    Assign {
        dst: PackedPlace,
        src: PackedOp,
    },
    LoadI32 {
        dst: PackedPlace,
        val: i32,
    },
    LoadF64 {
        dst: PackedPlace,
        bits_lo: u32,
        bits_hi: u32,
    },
    LoadBool {
        dst: PackedPlace,
        val: bool,
    },
    ArithLong {
        dst: PackedPlace,
        lhs: PackedOp,
        rhs: PackedOp,
        op: ArithOp,
    },
    ArithLongImm {
        dst: PackedPlace,
        lhs: PackedOp,
        imm: i32,
        op: ArithOp,
    },
    ArithDouble {
        dst: PackedPlace,
        lhs: PackedOp,
        rhs: PackedOp,
        op: ArithOp,
    },
    ArithWidening {
        dst: PackedPlace,
        lhs: PackedOp,
        rhs: PackedOp,
        op: ArithOp,
    },
    DivDouble {
        dst: PackedPlace,
        lhs: PackedOp,
        rhs: PackedOp,
    },
    Neg {
        dst: PackedPlace,
        src: PackedOp,
        mode: NumericMode,
    },
    Concat {
        dst: PackedPlace,
        lhs: PackedOp,
        rhs: PackedOp,
        coerce_to_string: bool,
    },
    Compare {
        dst: PackedPlace,
        lhs: PackedOp,
        rhs: PackedOp,
        op: CmpOp,
        mode: StringCompareMode,
        coercion: ComparisonCoercion,
    },
    CmpBranch(u32),
    ForCheck(u32),
    ForStepCheck(u32),
    Truthy {
        dst: PackedPlace,
        src: PackedOp,
    },
    CoerceNum {
        dst: PackedPlace,
        src: PackedOp,
        target: NumericCoerceTarget,
    },
    CoerceStr {
        dst: PackedPlace,
        src: PackedOp,
    },
    Not {
        dst: PackedPlace,
        src: PackedOp,
    },
    Logical {
        dst: PackedPlace,
        lhs: PackedOp,
        rhs: PackedOp,
        op: LogicalOp,
    },
    ArrayGet1D {
        dst: PackedPlace,
        array: PackedPlace,
        index: PackedOp,
    },
    ArraySet1D {
        array: PackedPlace,
        index: PackedOp,
        value: PackedOp,
    },
    RecordGet {
        dst: PackedPlace,
        record: PackedPlace,
        field: u32,
    },
    RecordSet {
        record: PackedPlace,
        field: u32,
        value: PackedOp,
    },
    ArrayRecordFieldGet1D(u32),
    ArrayRecordFieldSet1D(u32),
    FieldGet {
        dst: PackedPlace,
        object: PackedOp,
        field: i32,
    },
    FieldSet {
        object: PackedOp,
        field: i32,
        value: PackedOp,
    },
    CallProcFast(u32),
    CallExternLib(u32),
    CallComEarly(u32),
    CallComLate(u32),
    StmtBoundary {
        clear_temps_from: u32,
    },
    Jump {
        pc: u32,
        block: u32,
        steps: u16,
    },
    Branch(u32),
    Return,
    SlowInst(InstRef),
    SlowTerm(u32),
}

// Keep VmOp dense (4 instructions per 64-byte cache line).
const _: () = assert!(std::mem::size_of::<VmOp>() == 16);

/// Precomputed non-Empty local slot initializer indices for fast frame setup.
pub(crate) enum LocalSlotInit {
    Object(usize),
    Array(usize),
}

pub(crate) struct CompiledFunc {
    pub(crate) ops: Vec<VmOp>,
    pub(crate) block_start_pc: Vec<u32>,
    pub(crate) consts: Vec<FastConst>,
    pub(crate) cmp_branches: Vec<CmpBranchData>,
    pub(crate) branches: Vec<BranchData>,
    pub(crate) for_checks: Vec<ForCheckData>,
    pub(crate) for_step_checks: Vec<ForStepCheckData>,
    pub(crate) arr_rec_gets: Vec<ArrRecGetData>,
    pub(crate) arr_rec_sets: Vec<ArrRecSetData>,
    pub(crate) proc_calls: Vec<FastProcCall>,
    pub(crate) lib_calls: Vec<FastLibCall>,
    pub(crate) com_early_calls: Vec<FastComEarly>,
    pub(crate) com_late_calls: Vec<FastComLate>,
    /// Cold side-table mapping bytecode `pc` -> `(first_inst, count)` for ops that may
    /// fall back to `exec_slow_inst_range` (`ArrayGet1D`, `ArraySet1D`, `RecordGet`, `RecordSet`).
    pub(crate) fallback_insts: Vec<(u32, InstRef, u8)>,
    pub(crate) num_locals: usize,
    pub(crate) num_slots: usize,
    pub(crate) non_empty_locals: Vec<LocalSlotInit>,
    pub(crate) has_optional_params: bool,
}

impl CompiledFunc {
    #[cold]
    #[inline(never)]
    pub(crate) fn fallback_inst_for_pc(&self, op_pc: u32) -> (InstRef, usize) {
        let idx = self
            .fallback_insts
            .binary_search_by_key(&op_pc, |entry| entry.0)
            .expect("fallback_insts entry for slow-path op");
        let (_, iref, count) = self.fallback_insts[idx];
        (iref, count as usize)
    }
}

/// Whether calling `id` is a pure library computation that cannot fire inbound COM events.
fn is_pure_no_events_builtin(id: NativeImplId) -> bool {
    use NativeImplId::*;
    matches!(
        id,
        Len | LenB
            | Left
            | LeftB
            | Right
            | RightB
            | Mid
            | MidB
            | MidStmt
            | MidBStmt
            | LSetStmt
            | RSetStmt
            | InStr
            | InStrB
            | InStrRev
            | LCase
            | UCase
            | Split
            | Join
            | Replace
            | Trim
            | LTrim
            | RTrim
            | StrComp
            | Like
            | Chr
            | ChrB
            | ChrW
            | Asc
            | AscB
            | AscW
            | Space
            | StringRepeat
            | StrReverse
            | StrConv
            | FormatNumber
            | FormatCurrency
            | FormatPercent
            | FormatDateTime
            | Filter
            | Abs
            | Int
            | Fix
            | Sgn
            | Round
            | Sqr
            | Sin
            | Cos
            | Log
            | Exp
            | Atn
            | Tan
            | TimeSerial
            | TimeValue
            | Weekday
            | Hour
            | Minute
            | Second
            | MonthName
            | WeekdayName
            | CBool
            | CByte
            | CCur
            | CDate
            | CDbl
            | CDec
            | CInt
            | CLng
            | CLngLng
            | CSng
            | CStr
            | CVar
            | CVErr
            | Hex
            | Oct
            | Str
            | Val
            | IsArray
            | IsDate
            | IsEmpty
            | IsError
            | IsMissing
            | IsNull
            | IsNumeric
            | IsObject
            | Rgb
            | QbColor
            | Partition
            | Pmt
            | IPmt
            | PPmt
            | Pv
            | Fv
            | NPer
            | Rate
            | Npv
            | Irr
            | Mirr
            | Sln
            | Syd
            | Ddb
    )
}

struct TempUsage {
    reads: Vec<u32>,
    writes: Vec<u32>,
    escaped: Vec<bool>,
}

impl TempUsage {
    fn analyze(func: &OxFunc) -> Self {
        let n = func.temps.len();
        let mut usage = Self {
            reads: vec![0; n],
            writes: vec![0; n],
            escaped: vec![false; n],
        };
        for blk in &func.blocks {
            for inst in &blk.instrs {
                usage.visit_inst(inst);
            }
            usage.visit_term(&blk.terminator);
        }
        usage
    }

    fn read_op(&mut self, op: &OxOperand) {
        if let OxOperand::Use(OxPlace::Temp(t)) = op
            && let Some(c) = self.reads.get_mut(t.0)
        {
            *c = c.saturating_add(1);
        }
    }

    fn write_place(&mut self, place: &OxPlace) {
        if let OxPlace::Temp(t) = place
            && let Some(c) = self.writes.get_mut(t.0)
        {
            *c = c.saturating_add(1);
        }
    }

    fn escape_place(&mut self, place: &OxPlace) {
        if let OxPlace::Temp(t) = place
            && let Some(e) = self.escaped.get_mut(t.0)
        {
            *e = true;
        }
    }

    fn visit_ox_arg(&mut self, arg: &OxArg) {
        match arg {
            OxArg::ByVal(op) => self.read_op(op),
            OxArg::ByRef(p) => self.escape_place(p),
            OxArg::Omitted => {}
        }
    }

    fn visit_call_arg(&mut self, arg: &OxCallArg) {
        match arg {
            OxCallArg::Operand(op) => self.read_op(op),
            OxCallArg::ByRef(p) => self.escape_place(p),
            OxCallArg::Named { value, by_ref, .. } => {
                self.read_op(value);
                if let Some(p) = by_ref {
                    self.escape_place(p);
                }
            }
            OxCallArg::Omitted | OxCallArg::Const(_) => {}
        }
    }

    fn visit_inst(&mut self, inst: &OxInst) {
        match inst {
            OxInst::AsNew { place, .. } => self.escape_place(place),
            OxInst::Assign { dst, value } => {
                self.read_op(value);
                self.write_place(dst);
            }
            OxInst::Box { dst, src, .. }
            | OxInst::Unbox { dst, src, .. }
            | OxInst::Coerce { dst, src, .. }
            | OxInst::Truthy { dst, src }
            | OxInst::Neg { dst, src, .. }
            | OxInst::Not { dst, src } => {
                self.read_op(src);
                self.write_place(dst);
            }
            OxInst::ValidateAssignment { src, .. }
            | OxInst::PredeclaredSet { value: src, .. }
            | OxInst::PredeclaredExternSet { value: src, .. }
            | OxInst::ErrFieldSet { src, .. }
            | OxInst::AddRef { object: src }
            | OxInst::Release { object: src, .. } => self.read_op(src),
            OxInst::LoadProcRef { dst, .. }
            | OxInst::NewObject { dst, .. }
            | OxInst::NewExtern { dst, .. }
            | OxInst::Predeclared { dst, .. }
            | OxInst::PredeclaredExtern { dst, .. }
            | OxInst::NewRecord { dst, .. }
            | OxInst::ErrFieldGet { dst, .. }
            | OxInst::ErlGet { dst }
            | OxInst::WithEventsNextOwner { dst } => self.write_place(dst),
            OxInst::VariantChanged {
                dst,
                current,
                original,
            } => {
                self.read_op(current);
                self.read_op(original);
                self.write_place(dst);
            }
            OxInst::Arith { dst, lhs, rhs, .. }
            | OxInst::Div { dst, lhs, rhs }
            | OxInst::Pow { dst, lhs, rhs }
            | OxInst::Concat { dst, lhs, rhs }
            | OxInst::Compare { dst, lhs, rhs, .. }
            | OxInst::CompareObjectIs { dst, lhs, rhs }
            | OxInst::Logical { dst, lhs, rhs, .. } => {
                self.read_op(lhs);
                self.read_op(rhs);
                self.write_place(dst);
            }
            OxInst::TypeOfIs { dst, object, .. }
            | OxInst::FieldGet { dst, object, .. }
            | OxInst::WithEventsGet {
                dst, owner: object, ..
            }
            | OxInst::WithEventsClearOwner {
                dst, owner: object, ..
            }
            | OxInst::WithEventsFirstOwner {
                dst,
                source: object,
                ..
            } => {
                self.read_op(object);
                self.write_place(dst);
            }
            OxInst::FieldSet { object, value, .. } => {
                self.read_op(object);
                self.read_op(value);
            }
            OxInst::WithEventsSet {
                dst, owner, value, ..
            } => {
                self.read_op(owner);
                self.read_op(value);
                self.write_place(dst);
            }
            OxInst::RecordGet { dst, record, .. } => {
                self.read_op(record);
                self.write_place(dst);
            }
            OxInst::RecordSet { record, value, .. } | OxInst::RecordLSet { record, value } => {
                self.read_op(value);
                self.write_place(record);
            }
            OxInst::RecordArrayGet {
                dst,
                record,
                indices,
                ..
            } => {
                self.read_op(record);
                for idx in indices {
                    self.read_op(idx);
                }
                self.write_place(dst);
            }
            OxInst::RecordArraySet {
                record,
                indices,
                value,
                ..
            } => {
                for idx in indices {
                    self.read_op(idx);
                }
                self.read_op(value);
                self.write_place(record);
            }
            OxInst::ArrayLiteral {
                dst,
                values,
                aliases,
                ..
            } => {
                for v in values {
                    self.read_op(v);
                }
                for a in aliases.iter().flatten() {
                    self.escape_place(a);
                }
                self.write_place(dst);
            }
            OxInst::ArrayAppend { dst, array, item } => {
                self.read_op(array);
                self.read_op(item);
                self.write_place(dst);
            }
            OxInst::ArrayRedim {
                dst,
                upper_bounds,
                lower_bounds,
                ..
            } => {
                for u in upper_bounds {
                    self.read_op(u);
                }
                for l in lower_bounds {
                    self.read_op(l);
                }
                self.escape_place(dst);
            }
            OxInst::ArrayGet {
                dst,
                array,
                indices,
            } => {
                self.read_op(array);
                for idx in indices {
                    self.read_op(idx);
                }
                self.write_place(dst);
            }
            OxInst::ArraySet {
                array,
                indices,
                value,
            } => {
                for idx in indices {
                    self.read_op(idx);
                }
                self.read_op(value);
                self.write_place(array);
            }
            OxInst::FieldArrayGet {
                dst,
                object,
                indices,
                ..
            } => {
                self.read_op(object);
                for idx in indices {
                    self.read_op(idx);
                }
                self.write_place(dst);
            }
            OxInst::FieldArraySet {
                object,
                indices,
                value,
                ..
            } => {
                self.read_op(object);
                for idx in indices {
                    self.read_op(idx);
                }
                self.read_op(value);
            }
            OxInst::ArrayErase { array, .. } => self.escape_place(array),
            OxInst::Bound {
                dst,
                array,
                dimension,
                ..
            } => {
                self.read_op(array);
                if let Some(dim) = dimension {
                    self.read_op(dim);
                }
                self.write_place(dst);
            }
            OxInst::ForEachInit { iter, source } => {
                self.read_op(source);
                self.escape_place(iter);
            }
            OxInst::ForEachNext {
                iter,
                item,
                has_value,
            } => {
                self.escape_place(iter);
                self.write_place(item);
                self.write_place(has_value);
            }
            OxInst::ForEachEnd { iter } => self.escape_place(iter),
            OxInst::AddrOf {
                dst, base, path, ..
            }
            | OxInst::MakeRef { dst, base, path } => {
                match base {
                    calcvba_oxir::value::OxAddrBase::Place(p) => self.escape_place(p),
                    calcvba_oxir::value::OxAddrBase::ObjectField { object, .. } => {
                        self.read_op(object);
                    }
                }
                for step in path {
                    if let calcvba_oxir::value::OxAddrStep::Index(indices) = step {
                        for idx in indices {
                            self.read_op(idx);
                        }
                    }
                }
                self.escape_place(dst);
            }
            OxInst::ReleaseRef { src } => self.escape_place(src),
            OxInst::CallProc { dst, args, .. } | OxInst::CallExtern { dst, args, .. } => {
                for arg in args {
                    self.visit_ox_arg(arg);
                }
                if let Some(d) = dst {
                    self.write_place(d);
                }
            }
            OxInst::CallProcRef { dst, target, args } => {
                self.read_op(target);
                for arg in args {
                    self.visit_ox_arg(arg);
                }
                if let Some(d) = dst {
                    self.write_place(d);
                }
            }
            OxInst::CallNative { dst, args, .. } => {
                for arg in args {
                    self.visit_call_arg(arg);
                }
                if let Some(d) = dst {
                    self.write_place(d);
                }
            }
            OxInst::CallByName {
                dst,
                object,
                name,
                calltype,
                args,
            } => {
                self.read_op(object);
                self.read_op(name);
                self.read_op(calltype);
                for arg in args {
                    self.visit_call_arg(arg);
                }
                if let Some(d) = dst {
                    self.write_place(d);
                }
            }
            OxInst::ComCallEarly {
                dst, recv, args, ..
            }
            | OxInst::ComCallLate {
                dst, recv, args, ..
            } => {
                self.read_op(recv);
                for arg in args {
                    self.visit_call_arg(arg);
                }
                if let Some(d) = dst {
                    self.write_place(d);
                }
            }
            OxInst::RaiseEvent { source, args, .. } => {
                self.read_op(source);
                for arg in args {
                    self.visit_ox_arg(arg);
                }
            }
            OxInst::SetErrorHandler(_)
            | OxInst::ClearErr
            | OxInst::ClearErrOnExit
            | OxInst::StmtBoundary { .. }
            | OxInst::SetLineNumber { .. }
            | OxInst::DrainTerminations => {}
        }
    }

    fn visit_term(&mut self, term: &OxTerminator) {
        match term {
            OxTerminator::Branch { cond, .. } => self.read_op(cond),
            OxTerminator::Raise {
                number,
                source,
                description,
                help_file,
                help_context,
                ..
            } => {
                self.read_op(number);
                if let Some(op) = source {
                    self.read_op(op);
                }
                if let Some(op) = description {
                    self.read_op(op);
                }
                if let Some(op) = help_file {
                    self.read_op(op);
                }
                if let Some(op) = help_context {
                    self.read_op(op);
                }
            }
            _ => {}
        }
    }

    #[inline]
    fn is_single_use(&self, temp_idx: usize) -> bool {
        temp_idx < self.reads.len()
            && !self.escaped[temp_idx]
            && self.reads[temp_idx] == 1
            && self.writes[temp_idx] == 1
    }

    #[inline]
    fn is_single_update_temp(&self, temp_idx: usize) -> bool {
        temp_idx < self.reads.len()
            && !self.escaped[temp_idx]
            && self.reads[temp_idx] == 1
            && self.writes[temp_idx] == 2
    }
}

fn operand_is_project_typed_receiver(
    program: &OxProgram,
    func: &OxFunc,
    operand: &OxOperand,
) -> bool {
    let ty = match operand {
        OxOperand::Use(OxPlace::Local(id)) => func.locals.get(id.0).map(|local| &local.ty),
        OxOperand::Use(OxPlace::Temp(id)) => func.temps.get(id.0),
        OxOperand::Use(OxPlace::Global(id)) => program.globals.get(id.0).map(|global| &global.ty),
        _ => None,
    };
    matches!(
        ty,
        Some(OxTy::Object(
            calcvba_oxir::ObjClass::Class(_) | calcvba_oxir::ObjClass::Iface(_)
        ))
    )
}

fn is_plain_scalar_or_str_place(program: &OxProgram, func: &OxFunc, place: OxPlace) -> bool {
    let ty = match place {
        OxPlace::Local(id) => func.locals.get(id.0).map(|local| &local.ty),
        OxPlace::Global(id) => program.globals.get(id.0).map(|global| &global.ty),
        OxPlace::Temp(_) => return false,
    };
    matches!(
        ty,
        Some(
            OxTy::Bool
                | OxTy::Byte
                | OxTy::Integer
                | OxTy::Long
                | OxTy::LongLong
                | OxTy::Single
                | OxTy::Double
                | OxTy::Currency
                | OxTy::Date
                | OxTy::Str
        )
    )
}

fn resolve_vba_library_import(
    program: &OxProgram,
    import: ImportId,
) -> Option<(NativeImplId, bool)> {
    let imp = program.imports.get(import.0)?;
    if !imp.unit.eq_ignore_ascii_case("VBA") {
        return None;
    }
    let lib = calcvba_bundle::vba_library_bundle();
    let export = lib.exports.iter().find(|e| e.token.matches(&imp.token))?;
    let calcvba_bundle::ExportTarget::Proc(proc) = export.target else {
        return None;
    };
    match lib.procedures.get(proc).and_then(|p| p.native) {
        Some(calcvba_bundle::NativeBody::Library(id)) => {
            let string_typed_alias = match &imp.token {
                calcvba_bundle::ExportToken::ModuleFunc { member, .. } => {
                    id.is_string_typed_library_alias(member)
                }
                _ => false,
            };
            Some((id, string_typed_alias))
        }
        _ => None,
    }
}

fn const_as_i32(c: &OxConst) -> Option<i32> {
    match *c {
        OxConst::I16(n) => Some(i32::from(n)),
        OxConst::I32(n) => Some(n),
        _ => None,
    }
}

/// Follow empty `Jump`-only blocks, returning `(final_block, total_steps_including_initial_edge)`.
/// Stops without compressing if a cycle is encountered so infinite loops preserve exact budget steps.
fn resolve_jump_chain(func: &OxFunc, start: usize) -> (u32, u16) {
    let mut cur = start;
    let mut steps = 1u16;
    let max_hops = func.blocks.len().min(64);
    for _ in 0..max_hops {
        let Some(blk) = func.blocks.get(cur) else {
            break;
        };
        if !blk.instrs.is_empty() {
            break;
        }
        let OxTerminator::Jump(next) = blk.terminator else {
            break;
        };
        if next.0 == cur || next.0 == start {
            return (start as u32, 1);
        }
        let Some(next_steps) = steps.checked_add(1) else {
            break;
        };
        cur = next.0;
        steps = next_steps;
    }
    (cur as u32, steps)
}

#[derive(Clone)]
struct ForHeaderPattern {
    var: OxPlace,
    limit: OxOperand,
    nonneg: OxOperand,
    body_block: u32,
    body_steps: u16,
    after_block: u32,
    after_steps: u16,
}

#[derive(Clone)]
struct ForStepPattern {
    var: OxPlace,
    step: OxOperand,
    limit: OxOperand,
    nonneg: OxOperand,
    body_block: u32,
    after_block: u32,
    cont_steps_from_step: u16,
    exit_steps_from_step: u16,
}

fn detect_for_headers(func: &OxFunc, usage: &TempUsage) -> Vec<Option<ForHeaderPattern>> {
    let n = func.blocks.len();
    let mut out = vec![None; n];
    for (head_idx, head) in func.blocks.iter().enumerate() {
        if head.instrs.len() != 1 {
            continue;
        }
        let OxInst::Truthy {
            dst: OxPlace::Temp(nonneg_b),
            src: nonneg,
        } = &head.instrs[0]
        else {
            continue;
        };
        let OxTerminator::Branch {
            cond: OxOperand::Use(OxPlace::Temp(cond_b_ref)),
            then_blk: asc_id,
            else_blk: desc_id,
        } = &head.terminator
        else {
            continue;
        };
        if cond_b_ref != nonneg_b || !usage.is_single_use(nonneg_b.0) {
            continue;
        }
        if asc_id.0 == head_idx || desc_id.0 == head_idx || asc_id == desc_id {
            continue;
        }
        let (Some(asc), Some(desc)) = (func.blocks.get(asc_id.0), func.blocks.get(desc_id.0))
        else {
            continue;
        };
        if asc.fault_target != head.fault_target || desc.fault_target != head.fault_target {
            continue;
        }
        if asc.instrs.len() != 1 || desc.instrs.len() != 1 {
            continue;
        }
        let OxInst::Compare {
            dst: OxPlace::Temp(cond_t1),
            op: CmpOp::Le,
            lhs: OxOperand::Use(var1),
            rhs: limit1,
            mode: StringCompareMode::Binary,
            coercion: ComparisonCoercion::Strict,
        } = &asc.instrs[0]
        else {
            continue;
        };
        let OxTerminator::Jump(test1) = asc.terminator else {
            continue;
        };
        let OxInst::Compare {
            dst: OxPlace::Temp(cond_t2),
            op: CmpOp::Ge,
            lhs: OxOperand::Use(var2),
            rhs: limit2,
            mode: StringCompareMode::Binary,
            coercion: ComparisonCoercion::Strict,
        } = &desc.instrs[0]
        else {
            continue;
        };
        let OxTerminator::Jump(test2) = desc.terminator else {
            continue;
        };
        if test1 != test2 || cond_t1 != cond_t2 || var1 != var2 || limit1 != limit2 {
            continue;
        }
        let Some(test) = func.blocks.get(test1.0) else {
            continue;
        };
        if test.fault_target != head.fault_target || test.instrs.len() != 1 {
            continue;
        }
        let OxInst::Truthy {
            dst: OxPlace::Temp(cond_b),
            src: OxOperand::Use(OxPlace::Temp(ct)),
        } = &test.instrs[0]
        else {
            continue;
        };
        if ct != cond_t1 {
            continue;
        }
        let OxTerminator::Branch {
            cond: OxOperand::Use(OxPlace::Temp(cb)),
            then_blk: body_blk,
            else_blk: after_blk,
        } = &test.terminator
        else {
            continue;
        };
        if cb != cond_b || !usage.is_single_use(cond_b.0) {
            continue;
        }
        if cond_t1.0 >= usage.reads.len()
            || usage.escaped[cond_t1.0]
            || usage.reads[cond_t1.0] != 1
            || usage.writes[cond_t1.0] != 2
        {
            continue;
        }
        let (body_target, body_edge_steps) = resolve_jump_chain(func, body_blk.0);
        let (after_target, after_edge_steps) = resolve_jump_chain(func, after_blk.0);
        // 2 steps: head->asc/desc, asc/desc->test, plus test->body/after chain.
        let Some(body_steps) = 2u16.checked_add(body_edge_steps) else {
            continue;
        };
        let Some(after_steps) = 2u16.checked_add(after_edge_steps) else {
            continue;
        };
        out[head_idx] = Some(ForHeaderPattern {
            var: *var1,
            limit: limit1.clone(),
            nonneg: nonneg.clone(),
            body_block: body_target,
            body_steps,
            after_block: after_target,
            after_steps,
        });
    }
    out
}

fn detect_for_steps(
    func: &OxFunc,
    headers: &[Option<ForHeaderPattern>],
) -> Vec<Option<ForStepPattern>> {
    let n = func.blocks.len();
    let mut out = vec![None; n];
    for (step_idx, blk) in func.blocks.iter().enumerate() {
        if blk.instrs.len() != 1 {
            continue;
        }
        let OxInst::Arith {
            dst: var_dst,
            op: ArithOp::Add,
            lhs: OxOperand::Use(var_lhs),
            rhs: step_op,
            mode: NumericMode::Checked(NumericCoerceTarget::Long),
        } = &blk.instrs[0]
        else {
            continue;
        };
        if var_dst != var_lhs {
            continue;
        }
        let OxTerminator::Jump(j_target) = blk.terminator else {
            continue;
        };
        let (head_blk, head_steps) = resolve_jump_chain(func, j_target.0);
        let Some(header) = headers.get(head_blk as usize).and_then(Option::as_ref) else {
            continue;
        };
        if header.var != *var_dst || func.blocks[head_blk as usize].fault_target != blk.fault_target
        {
            continue;
        }
        let (Some(cont_steps), Some(exit_steps)) = (
            head_steps.checked_add(header.body_steps),
            head_steps.checked_add(header.after_steps),
        ) else {
            continue;
        };
        out[step_idx] = Some(ForStepPattern {
            var: *var_dst,
            step: step_op.clone(),
            limit: header.limit.clone(),
            nonneg: header.nonneg.clone(),
            body_block: header.body_block,
            after_block: header.after_block,
            cont_steps_from_step: cont_steps,
            exit_steps_from_step: exit_steps,
        });
    }
    out
}

impl CompiledFunc {
    pub(crate) fn compile(program: &OxProgram, func: &OxFunc) -> Self {
        let num_locals = func.locals.len();
        let num_slots = num_locals + func.temps.len();
        let usage = TempUsage::analyze(func);
        let for_headers = detect_for_headers(func, &usage);
        let for_steps = detect_for_steps(func, &for_headers);

        let mut non_empty_locals = Vec::new();
        let mut has_optional_params = false;
        for (idx, local) in func.locals.iter().enumerate() {
            if idx < func.param_count
                && local
                    .param
                    .as_ref()
                    .is_some_and(|p| p.optional && !p.variadic)
            {
                has_optional_params = true;
            }
            match &local.ty {
                OxTy::Object(_) => non_empty_locals.push(LocalSlotInit::Object(idx)),
                OxTy::Array(_, _) => non_empty_locals.push(LocalSlotInit::Array(idx)),
                _ => {}
            }
        }

        let mut out = Self {
            ops: Vec::new(),
            block_start_pc: vec![0; func.blocks.len()],
            consts: Vec::new(),
            cmp_branches: Vec::new(),
            branches: Vec::new(),
            for_checks: Vec::new(),
            for_step_checks: Vec::new(),
            arr_rec_gets: Vec::new(),
            arr_rec_sets: Vec::new(),
            proc_calls: Vec::new(),
            lib_calls: Vec::new(),
            com_early_calls: Vec::new(),
            com_late_calls: Vec::new(),
            fallback_insts: Vec::new(),
            num_locals,
            num_slots,
            non_empty_locals,
            has_optional_params,
        };

        let mut step_block_slow_pc: Vec<u32> = vec![0; func.blocks.len()];
        let mut jump_patches: Vec<(usize, u32)> = Vec::new();
        let mut branch_patches: Vec<usize> = Vec::new();
        let mut cmp_branch_patches: Vec<usize> = Vec::new();
        let mut for_check_patches: Vec<usize> = Vec::new();
        let mut for_step_patches: Vec<usize> = Vec::new();

        for (blk_idx, blk) in func.blocks.iter().enumerate() {
            out.block_start_pc[blk_idx] = out.ops.len() as u32;

            // Prefix fused ForCheck before the original head ops (falls through on slow path).
            if let Some(hdr) = &for_headers[blk_idx]
                && let (Some(p_var), Some(p_lim), Some(p_nn)) = (
                    PackedPlace::from_ox(hdr.var, num_locals),
                    out.pack_op(&hdr.limit),
                    out.pack_op(&hdr.nonneg),
                )
            {
                let idx = out.for_checks.len();
                out.for_checks.push(ForCheckData {
                    var: p_var,
                    limit: p_lim,
                    nonneg: p_nn,
                    body_pc: 0,
                    body_block: hdr.body_block,
                    body_steps: hdr.body_steps,
                    after_pc: 0,
                    after_block: hdr.after_block,
                    after_steps: hdr.after_steps,
                });
                for_check_patches.push(idx);
                out.ops.push(VmOp::ForCheck(idx as u32));
            }

            // Prefix fused ForStepCheck before step_blk's original ops (falls through on slow path).
            if let Some(stp) = &for_steps[blk_idx]
                && let (Some(p_var), Some(p_step), Some(p_lim), Some(p_nn)) = (
                    PackedPlace::from_ox(stp.var, num_locals),
                    out.pack_op(&stp.step),
                    out.pack_op(&stp.limit),
                    out.pack_op(&stp.nonneg),
                )
            {
                let idx = out.for_step_checks.len();
                out.for_step_checks.push(ForStepCheckData {
                    var: p_var,
                    step: p_step,
                    limit: p_lim,
                    nonneg: p_nn,
                    body_pc: 0,
                    body_block: stp.body_block,
                    after_pc: 0,
                    after_block: stp.after_block,
                    step_block: blk_idx as u32,
                    slow_pc: 0,
                    to_step_steps: 0,
                    cont_steps: stp.cont_steps_from_step,
                    exit_steps: stp.exit_steps_from_step,
                });
                for_step_patches.push(idx);
                out.ops.push(VmOp::ForStepCheck(idx as u32));
            }
            step_block_slow_pc[blk_idx] = out.ops.len() as u32;

            let instrs = &blk.instrs;
            let len = instrs.len();

            // Check if the tail of this block fuses into `CmpBranch`.
            let mut fused_tail_len = 0usize;
            if let OxTerminator::Branch {
                cond: OxOperand::Use(OxPlace::Temp(cond_temp)),
                then_blk,
                else_blk,
            } = &blk.terminator
            {
                if len >= 2
                    && let OxInst::Compare {
                        dst: OxPlace::Temp(cmp_temp),
                        op,
                        lhs,
                        rhs,
                        mode,
                        coercion,
                    } = &instrs[len - 2]
                    && let OxInst::Truthy {
                        dst: OxPlace::Temp(truth_temp),
                        src: OxOperand::Use(OxPlace::Temp(truth_src)),
                    } = &instrs[len - 1]
                    && truth_temp == cond_temp
                    && truth_src == cmp_temp
                    && usage.is_single_use(cmp_temp.0)
                    && usage.is_single_use(truth_temp.0)
                    && let (Some(l), Some(r)) = (out.pack_op(lhs), out.pack_op(rhs))
                {
                    let (then_block, then_steps) = resolve_jump_chain(func, then_blk.0);
                    let (else_block, else_steps) = resolve_jump_chain(func, else_blk.0);
                    fused_tail_len = 2;
                    let idx = out.cmp_branches.len();
                    out.cmp_branches.push(CmpBranchData {
                        lhs: l,
                        rhs: r,
                        op: *op,
                        mode: *mode,
                        coercion: *coercion,
                        then_pc: 0,
                        then_block,
                        then_steps,
                        else_pc: 0,
                        else_block,
                        else_steps,
                    });
                    cmp_branch_patches.push(idx);
                } else if len >= 1
                    && let OxInst::Compare {
                        dst: OxPlace::Temp(cmp_temp),
                        op,
                        lhs,
                        rhs,
                        mode,
                        coercion,
                    } = &instrs[len - 1]
                    && cmp_temp == cond_temp
                    && usage.is_single_use(cmp_temp.0)
                    && let (Some(l), Some(r)) = (out.pack_op(lhs), out.pack_op(rhs))
                {
                    let (then_block, then_steps) = resolve_jump_chain(func, then_blk.0);
                    let (else_block, else_steps) = resolve_jump_chain(func, else_blk.0);
                    fused_tail_len = 1;
                    let idx = out.cmp_branches.len();
                    out.cmp_branches.push(CmpBranchData {
                        lhs: l,
                        rhs: r,
                        op: *op,
                        mode: *mode,
                        coercion: *coercion,
                        then_pc: 0,
                        then_block,
                        then_steps,
                        else_pc: 0,
                        else_block,
                        else_steps,
                    });
                    cmp_branch_patches.push(idx);
                }
            }

            let emit_end = len - fused_tail_len;
            let mut i = 0usize;
            while i < emit_end {
                let iref = InstRef {
                    block: blk_idx as u32,
                    inst: i as u32,
                };
                // Pattern 1a: Fused 1D UDT array element field write (mutable_base_place form):
                //   ArrayGet  { dst: Temp(t1), array: Use(arr), indices: [idx1] }
                //   Assign    { dst: Temp(t2), value: Use(Temp(t1)) }
                //   RecordSet { record: Temp(t2), index: field, value: val }
                //   ArraySet  { array: arr2, indices: [idx2], value: Use(Temp(t2)) }
                if i + 3 < emit_end
                    && let OxInst::ArrayGet {
                        dst: OxPlace::Temp(t1),
                        array: OxOperand::Use(arr),
                        indices: idx1,
                    } = &instrs[i]
                    && idx1.len() == 1
                    && let OxInst::Assign {
                        dst: OxPlace::Temp(t2),
                        value: OxOperand::Use(OxPlace::Temp(t1_ref)),
                    } = &instrs[i + 1]
                    && t1_ref == t1
                    && t2 != t1
                    && let OxInst::RecordSet {
                        record: OxPlace::Temp(t_rec),
                        index: field,
                        value: val,
                    } = &instrs[i + 2]
                    && t_rec == t2
                    && !matches!(val, OxOperand::Use(OxPlace::Temp(tv)) if tv == t1 || tv == t2)
                    && let OxInst::ArraySet {
                        array: arr2,
                        indices: idx2,
                        value: OxOperand::Use(OxPlace::Temp(t_set)),
                    } = &instrs[i + 3]
                    && t_set == t2
                    && arr2 == arr
                    && idx2.len() == 1
                    && idx2[0] == idx1[0]
                    && usage.is_single_use(t1.0)
                    && usage.is_single_update_temp(t2.0)
                    && let (Some(p_arr), Some(p_idx), Some(p_val), Ok(field_u32)) = (
                        PackedPlace::from_ox(*arr, num_locals),
                        out.pack_op(&idx1[0]),
                        out.pack_op(val),
                        u32::try_from(*field),
                    )
                {
                    let data_idx = out.arr_rec_sets.len() as u32;
                    out.arr_rec_sets.push(ArrRecSetData {
                        array: p_arr,
                        index: p_idx,
                        value: p_val,
                        field: field_u32,
                        first_inst: iref,
                        orig_count: 4,
                    });
                    out.ops.push(VmOp::ArrayRecordFieldSet1D(data_idx));
                    i += 4;
                    continue;
                }

                // Pattern 1b: Fused 1D UDT array element field write (direct 3-inst form):
                //   ArrayGet  { dst: Temp(t), array: Use(arr), indices: [idx1] }
                //   RecordSet { record: Temp(t), index: field, value: val }
                //   ArraySet  { array: arr2, indices: [idx2], value: Use(Temp(t)) }
                if i + 2 < emit_end
                    && let OxInst::ArrayGet {
                        dst: OxPlace::Temp(t),
                        array: OxOperand::Use(arr),
                        indices: idx1,
                    } = &instrs[i]
                    && idx1.len() == 1
                    && let OxInst::RecordSet {
                        record: OxPlace::Temp(t_rec),
                        index: field,
                        value: val,
                    } = &instrs[i + 1]
                    && t_rec == t
                    && !matches!(val, OxOperand::Use(OxPlace::Temp(tv)) if tv == t)
                    && let OxInst::ArraySet {
                        array: arr2,
                        indices: idx2,
                        value: OxOperand::Use(OxPlace::Temp(t_set)),
                    } = &instrs[i + 2]
                    && t_set == t
                    && arr2 == arr
                    && idx2.len() == 1
                    && idx2[0] == idx1[0]
                    && usage.is_single_update_temp(t.0)
                    && let (Some(p_arr), Some(p_idx), Some(p_val), Ok(field_u32)) = (
                        PackedPlace::from_ox(*arr, num_locals),
                        out.pack_op(&idx1[0]),
                        out.pack_op(val),
                        u32::try_from(*field),
                    )
                {
                    let data_idx = out.arr_rec_sets.len() as u32;
                    out.arr_rec_sets.push(ArrRecSetData {
                        array: p_arr,
                        index: p_idx,
                        value: p_val,
                        field: field_u32,
                        first_inst: iref,
                        orig_count: 3,
                    });
                    out.ops.push(VmOp::ArrayRecordFieldSet1D(data_idx));
                    i += 3;
                    continue;
                }

                // Pattern 2: Fused 1D UDT array element field read (+ optional trailing Assign):
                //   ArrayGet  { dst: Temp(t), array: Use(arr), indices: [idx] }
                //   RecordGet { dst: dst2, record: Use(Temp(t)), index: field }
                if i + 1 < emit_end
                    && let OxInst::ArrayGet {
                        dst: OxPlace::Temp(t),
                        array: OxOperand::Use(arr),
                        indices,
                    } = &instrs[i]
                    && indices.len() == 1
                    && let OxInst::RecordGet {
                        dst: dst2,
                        record: OxOperand::Use(OxPlace::Temp(t_rec)),
                        index: field,
                    } = &instrs[i + 1]
                    && t_rec == t
                    && usage.is_single_use(t.0)
                {
                    let (final_dst, orig_count) = if i + 2 < emit_end
                        && let OxPlace::Temp(t_out) = *dst2
                        && let OxInst::Assign {
                            dst: assign_dst,
                            value: OxOperand::Use(OxPlace::Temp(t_src)),
                        } = &instrs[i + 2]
                        && *t_src == t_out
                        && usage.is_single_use(t_out.0)
                        && is_plain_scalar_or_str_place(program, func, *assign_dst)
                    {
                        (*assign_dst, 3u8)
                    } else {
                        (*dst2, 2u8)
                    };
                    if let (Some(p_dst), Some(p_arr), Some(p_idx), Ok(field_u32)) = (
                        PackedPlace::from_ox(final_dst, num_locals),
                        PackedPlace::from_ox(*arr, num_locals),
                        out.pack_op(&indices[0]),
                        u32::try_from(*field),
                    ) {
                        let data_idx = out.arr_rec_gets.len() as u32;
                        out.arr_rec_gets.push(ArrRecGetData {
                            dst: p_dst,
                            array: p_arr,
                            index: p_idx,
                            field: field_u32,
                            first_inst: iref,
                            orig_count,
                        });
                        out.ops.push(VmOp::ArrayRecordFieldGet1D(data_idx));
                        i += orig_count as usize;
                        continue;
                    }
                }

                // A typed String assignment coerces the result of `&` before copying it.
                // Carry the coercion on the concat and write the owned BSTR once.
                if i + 2 < emit_end
                    && let OxInst::Concat {
                        dst: OxPlace::Temp(joined),
                        lhs,
                        rhs,
                    } = &instrs[i]
                    && let OxInst::Coerce {
                        dst: OxPlace::Temp(converted),
                        src: OxOperand::Use(OxPlace::Temp(input)),
                        target: OxCoerceTarget::Str,
                    } = &instrs[i + 1]
                    && input == joined
                    && usage.is_single_use(joined.0)
                    && let OxInst::Assign {
                        dst,
                        value: OxOperand::Use(OxPlace::Temp(output)),
                    } = &instrs[i + 2]
                    && output == converted
                    && usage.is_single_use(converted.0)
                    && is_plain_scalar_or_str_place(program, func, *dst)
                    && let (Some(dst), Some(lhs), Some(rhs)) = (
                        PackedPlace::from_ox(*dst, num_locals),
                        out.pack_op(lhs),
                        out.pack_op(rhs),
                    )
                {
                    out.ops.push(VmOp::Concat {
                        dst,
                        lhs,
                        rhs,
                        coerce_to_string: true,
                    });
                    i += 3;
                    continue;
                }

                // Pattern 3: Single-use temporary destination forwarding:
                //   op(dst = Temp(t), ...); Assign(dst = local/global, src = Temp(t))
                //   -> op(dst = local/global, ...)
                // Only applied to ops that never write `dst` on error.
                let forward_target: Option<OxPlace> = if i + 1 < emit_end
                    && let OxInst::Assign {
                        dst: final_dst,
                        value: OxOperand::Use(OxPlace::Temp(t_src)),
                    } = &instrs[i + 1]
                    && usage.is_single_use(t_src.0)
                    && is_plain_scalar_or_str_place(program, func, *final_dst)
                {
                    let inst_writes_t = match &instrs[i] {
                        OxInst::Arith {
                            dst,
                            mode: NumericMode::Checked(NumericCoerceTarget::Long),
                            ..
                        }
                        | OxInst::Concat { dst, .. }
                        | OxInst::Compare { dst, .. }
                        | OxInst::Truthy { dst, .. }
                        | OxInst::Not { dst, .. }
                        | OxInst::Logical { dst, .. }
                        | OxInst::FieldGet { dst, .. } => *dst == OxPlace::Temp(*t_src),
                        OxInst::ArrayGet {
                            dst,
                            array: OxOperand::Use(_),
                            indices,
                        } if indices.len() == 1 => *dst == OxPlace::Temp(*t_src),
                        OxInst::CallProc { dst: Some(dst), .. }
                        | OxInst::CallExtern { dst: Some(dst), .. } => {
                            *dst == OxPlace::Temp(*t_src)
                        }
                        _ => false,
                    };
                    inst_writes_t.then_some(*final_dst)
                } else {
                    None
                };

                let op_pc = out.ops.len() as u32;
                let (op, consumed) =
                    out.compile_inst(program, func, &instrs[i], iref, op_pc, forward_target);
                out.ops.push(op);
                i += consumed;
            }

            if fused_tail_len > 0 {
                let idx = (out.cmp_branches.len() - 1) as u32;
                out.ops.push(VmOp::CmpBranch(idx));
            } else {
                match &blk.terminator {
                    OxTerminator::Jump(target) => {
                        let (final_blk, steps) = resolve_jump_chain(func, target.0);
                        // Fuse loop body's tail jump straight into `ForStepCheck`!
                        if let Some(stp) = &for_steps[final_blk as usize]
                            && let (
                                Some(p_var),
                                Some(p_step),
                                Some(p_lim),
                                Some(p_nn),
                                Some(cont_steps),
                                Some(exit_steps),
                            ) = (
                                PackedPlace::from_ox(stp.var, num_locals),
                                out.pack_op(&stp.step),
                                out.pack_op(&stp.limit),
                                out.pack_op(&stp.nonneg),
                                steps.checked_add(stp.cont_steps_from_step),
                                steps.checked_add(stp.exit_steps_from_step),
                            )
                        {
                            let idx = out.for_step_checks.len();
                            out.for_step_checks.push(ForStepCheckData {
                                var: p_var,
                                step: p_step,
                                limit: p_lim,
                                nonneg: p_nn,
                                body_pc: 0,
                                body_block: stp.body_block,
                                after_pc: 0,
                                after_block: stp.after_block,
                                step_block: final_blk,
                                slow_pc: 0,
                                to_step_steps: steps,
                                cont_steps,
                                exit_steps,
                            });
                            for_step_patches.push(idx);
                            out.ops.push(VmOp::ForStepCheck(idx as u32));
                        } else {
                            let op_pos = out.ops.len();
                            out.ops.push(VmOp::Jump {
                                pc: 0,
                                block: final_blk,
                                steps,
                            });
                            jump_patches.push((op_pos, final_blk));
                        }
                    }
                    OxTerminator::Branch {
                        cond,
                        then_blk,
                        else_blk,
                    } => {
                        if let Some(p_cond) = out.pack_op(cond) {
                            let (then_block, then_steps) = resolve_jump_chain(func, then_blk.0);
                            let (else_block, else_steps) = resolve_jump_chain(func, else_blk.0);
                            let idx = out.branches.len();
                            out.branches.push(BranchData {
                                cond: p_cond,
                                then_pc: 0,
                                then_block,
                                then_steps,
                                else_pc: 0,
                                else_block,
                                else_steps,
                            });
                            branch_patches.push(idx);
                            out.ops.push(VmOp::Branch(idx as u32));
                        } else {
                            out.ops.push(VmOp::SlowTerm(blk_idx as u32));
                        }
                    }
                    OxTerminator::Return => out.ops.push(VmOp::Return),
                    _ => out.ops.push(VmOp::SlowTerm(blk_idx as u32)),
                }
            }
        }

        // Resolve all block start PCs into Jump / Branch / CmpBranch / ForCheck / ForStepCheck.
        for (op_pos, target_blk) in jump_patches {
            let target_pc = out.block_start_pc[target_blk as usize];
            if let VmOp::Jump { pc, .. } = &mut out.ops[op_pos] {
                *pc = target_pc;
            }
        }
        for idx in branch_patches {
            let b = &mut out.branches[idx];
            b.then_pc = out.block_start_pc[b.then_block as usize];
            b.else_pc = out.block_start_pc[b.else_block as usize];
        }
        for idx in cmp_branch_patches {
            let b = &mut out.cmp_branches[idx];
            b.then_pc = out.block_start_pc[b.then_block as usize];
            b.else_pc = out.block_start_pc[b.else_block as usize];
        }
        for idx in for_check_patches {
            let f = &mut out.for_checks[idx];
            f.body_pc = out.block_start_pc[f.body_block as usize];
            f.after_pc = out.block_start_pc[f.after_block as usize];
        }
        for idx in for_step_patches {
            let f = &mut out.for_step_checks[idx];
            f.body_pc = out.block_start_pc[f.body_block as usize];
            f.after_pc = out.block_start_pc[f.after_block as usize];
            f.slow_pc = step_block_slow_pc[f.step_block as usize];
        }

        out
    }

    fn pack_const(&mut self, c: &OxConst) -> Option<PackedOp> {
        let entry = match c {
            OxConst::Str(s) => {
                FastConst::Str(s.encode_utf16().collect::<Vec<u16>>().into_boxed_slice())
            }
            OxConst::NullStr => FastConst::NullStr,
            other => FastConst::Scalar(ox_const_variant(other)),
        };
        let idx = u32::try_from(self.consts.len()).ok()?;
        let op = PackedOp::from_const_idx(idx)?;
        self.consts.push(entry);
        Some(op)
    }

    fn pack_op(&mut self, op: &OxOperand) -> Option<PackedOp> {
        match op {
            OxOperand::Use(place) => {
                PackedPlace::from_ox(*place, self.num_locals).map(PackedOp::from_place)
            }
            OxOperand::Const(c) => self.pack_const(c),
        }
    }

    fn pack_ox_args(&mut self, args: &[OxArg]) -> Option<(Vec<FastArg>, bool)> {
        let mut packed = Vec::with_capacity(args.len());
        let mut all_byval = true;
        for arg in args {
            match arg {
                OxArg::ByVal(op) => packed.push(FastArg::ByVal(self.pack_op(op)?)),
                OxArg::ByRef(place) => {
                    all_byval = false;
                    packed.push(FastArg::ByRef(PackedPlace::from_ox(
                        *place,
                        self.num_locals,
                    )?));
                }
                OxArg::Omitted => {
                    all_byval = false;
                    packed.push(FastArg::Omitted);
                }
            }
        }
        Some((packed, all_byval))
    }

    fn pack_positional_call_args(&mut self, args: &[OxCallArg]) -> Option<Vec<FastCallArg>> {
        let mut packed = Vec::with_capacity(args.len());
        for arg in args {
            match arg {
                OxCallArg::Operand(op) => packed.push(FastCallArg::Operand(self.pack_op(op)?)),
                OxCallArg::ByRef(place) => {
                    packed.push(FastCallArg::ByRef(PackedPlace::from_ox(
                        *place,
                        self.num_locals,
                    )?));
                }
                OxCallArg::Omitted => packed.push(FastCallArg::Omitted),
                OxCallArg::Const(_) => packed.push(FastCallArg::Const),
                OxCallArg::Named { .. } => return None,
            }
        }
        Some(packed)
    }

    fn compile_inst(
        &mut self,
        program: &OxProgram,
        func: &OxFunc,
        inst: &OxInst,
        iref: InstRef,
        op_pc: u32,
        forward_dst: Option<OxPlace>,
    ) -> (VmOp, usize) {
        let num_locals = self.num_locals;
        let consumed_if_fwd = if forward_dst.is_some() { 2 } else { 1 };
        match inst {
            OxInst::Assign { dst, value }
            | OxInst::Box {
                dst, src: value, ..
            } => {
                let Some(p_dst) = PackedPlace::from_ox(*dst, num_locals) else {
                    return (VmOp::SlowInst(iref), 1);
                };
                let op = match value {
                    OxOperand::Const(OxConst::I32(v)) => VmOp::LoadI32 {
                        dst: p_dst,
                        val: *v,
                    },
                    OxOperand::Const(OxConst::F64(bits)) => VmOp::LoadF64 {
                        dst: p_dst,
                        bits_lo: *bits as u32,
                        bits_hi: (*bits >> 32) as u32,
                    },
                    OxOperand::Const(OxConst::Bool(v)) => VmOp::LoadBool {
                        dst: p_dst,
                        val: *v,
                    },
                    _ => match self.pack_op(value) {
                        Some(src) => VmOp::Assign { dst: p_dst, src },
                        None => VmOp::SlowInst(iref),
                    },
                };
                (op, 1)
            }
            OxInst::Arith {
                dst,
                op,
                lhs,
                rhs,
                mode,
            } => {
                let effective_dst = match mode {
                    NumericMode::Checked(NumericCoerceTarget::Long) => forward_dst.unwrap_or(*dst),
                    _ => *dst,
                };
                let used_fwd = matches!(mode, NumericMode::Checked(NumericCoerceTarget::Long))
                    && forward_dst.is_some();
                let Some(p_dst) = PackedPlace::from_ox(effective_dst, num_locals) else {
                    return (VmOp::SlowInst(iref), 1);
                };
                match mode {
                    NumericMode::Checked(NumericCoerceTarget::Long) => {
                        if let (Some(l), OxOperand::Const(c)) = (self.pack_op(lhs), rhs)
                            && let Some(imm) = const_as_i32(c)
                        {
                            return (
                                VmOp::ArithLongImm {
                                    dst: p_dst,
                                    lhs: l,
                                    imm,
                                    op: *op,
                                },
                                if used_fwd { 2 } else { 1 },
                            );
                        }
                        if let (Some(l), Some(r)) = (self.pack_op(lhs), self.pack_op(rhs)) {
                            return (
                                VmOp::ArithLong {
                                    dst: p_dst,
                                    lhs: l,
                                    rhs: r,
                                    op: *op,
                                },
                                if used_fwd { 2 } else { 1 },
                            );
                        }
                        (VmOp::SlowInst(iref), 1)
                    }
                    NumericMode::Checked(NumericCoerceTarget::Double) => {
                        if let (Some(l), Some(r)) = (self.pack_op(lhs), self.pack_op(rhs)) {
                            (
                                VmOp::ArithDouble {
                                    dst: p_dst,
                                    lhs: l,
                                    rhs: r,
                                    op: *op,
                                },
                                1,
                            )
                        } else {
                            (VmOp::SlowInst(iref), 1)
                        }
                    }
                    NumericMode::Widening => {
                        if let (Some(l), Some(r)) = (self.pack_op(lhs), self.pack_op(rhs)) {
                            (
                                VmOp::ArithWidening {
                                    dst: p_dst,
                                    lhs: l,
                                    rhs: r,
                                    op: *op,
                                },
                                1,
                            )
                        } else {
                            (VmOp::SlowInst(iref), 1)
                        }
                    }
                    _ => (VmOp::SlowInst(iref), 1),
                }
            }
            OxInst::Div { dst, lhs, rhs } => {
                if let (Some(p_dst), Some(l), Some(r)) = (
                    PackedPlace::from_ox(*dst, num_locals),
                    self.pack_op(lhs),
                    self.pack_op(rhs),
                ) {
                    (
                        VmOp::DivDouble {
                            dst: p_dst,
                            lhs: l,
                            rhs: r,
                        },
                        1,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::Neg { dst, src, mode } => {
                if let (Some(p_dst), Some(p_src)) =
                    (PackedPlace::from_ox(*dst, num_locals), self.pack_op(src))
                {
                    (
                        VmOp::Neg {
                            dst: p_dst,
                            src: p_src,
                            mode: *mode,
                        },
                        1,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::Concat { dst, lhs, rhs } => {
                if let (Some(p_dst), Some(l), Some(r)) = (
                    PackedPlace::from_ox(forward_dst.unwrap_or(*dst), num_locals),
                    self.pack_op(lhs),
                    self.pack_op(rhs),
                ) {
                    (
                        VmOp::Concat {
                            dst: p_dst,
                            lhs: l,
                            rhs: r,
                            coerce_to_string: false,
                        },
                        consumed_if_fwd,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::Compare {
                dst,
                op,
                lhs,
                rhs,
                mode,
                coercion,
            } => {
                if let (Some(p_dst), Some(l), Some(r)) = (
                    PackedPlace::from_ox(forward_dst.unwrap_or(*dst), num_locals),
                    self.pack_op(lhs),
                    self.pack_op(rhs),
                ) {
                    (
                        VmOp::Compare {
                            dst: p_dst,
                            lhs: l,
                            rhs: r,
                            op: *op,
                            mode: *mode,
                            coercion: *coercion,
                        },
                        consumed_if_fwd,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::Truthy { dst, src } => {
                if let (Some(p_dst), Some(p_src)) = (
                    PackedPlace::from_ox(forward_dst.unwrap_or(*dst), num_locals),
                    self.pack_op(src),
                ) {
                    (
                        VmOp::Truthy {
                            dst: p_dst,
                            src: p_src,
                        },
                        consumed_if_fwd,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::Coerce { dst, src, target } => {
                let Some(p_dst) = PackedPlace::from_ox(*dst, num_locals) else {
                    return (VmOp::SlowInst(iref), 1);
                };
                let Some(p_src) = self.pack_op(src) else {
                    return (VmOp::SlowInst(iref), 1);
                };
                let op = match target {
                    OxCoerceTarget::Numeric(t) => VmOp::CoerceNum {
                        dst: p_dst,
                        src: p_src,
                        target: *t,
                    },
                    OxCoerceTarget::Str => VmOp::CoerceStr {
                        dst: p_dst,
                        src: p_src,
                    },
                    OxCoerceTarget::ImplicitVariant => VmOp::Assign {
                        dst: p_dst,
                        src: p_src,
                    },
                    OxCoerceTarget::FixedStr(_) => VmOp::SlowInst(iref),
                };
                (op, 1)
            }
            OxInst::Not { dst, src } => {
                if let (Some(p_dst), Some(p_src)) = (
                    PackedPlace::from_ox(forward_dst.unwrap_or(*dst), num_locals),
                    self.pack_op(src),
                ) {
                    (
                        VmOp::Not {
                            dst: p_dst,
                            src: p_src,
                        },
                        consumed_if_fwd,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::Logical { dst, op, lhs, rhs } => {
                if let (Some(p_dst), Some(l), Some(r)) = (
                    PackedPlace::from_ox(forward_dst.unwrap_or(*dst), num_locals),
                    self.pack_op(lhs),
                    self.pack_op(rhs),
                ) {
                    (
                        VmOp::Logical {
                            dst: p_dst,
                            lhs: l,
                            rhs: r,
                            op: *op,
                        },
                        consumed_if_fwd,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::ArrayGet {
                dst,
                array: OxOperand::Use(arr),
                indices,
            } if indices.len() == 1 => {
                if let (Some(p_dst), Some(p_arr), Some(p_idx)) = (
                    PackedPlace::from_ox(forward_dst.unwrap_or(*dst), num_locals),
                    PackedPlace::from_ox(*arr, num_locals),
                    self.pack_op(&indices[0]),
                ) {
                    self.fallback_insts
                        .push((op_pc, iref, consumed_if_fwd as u8));
                    (
                        VmOp::ArrayGet1D {
                            dst: p_dst,
                            array: p_arr,
                            index: p_idx,
                        },
                        consumed_if_fwd,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::ArraySet {
                array,
                indices,
                value,
            } if indices.len() == 1 => {
                if let (Some(p_arr), Some(p_idx), Some(p_val)) = (
                    PackedPlace::from_ox(*array, num_locals),
                    self.pack_op(&indices[0]),
                    self.pack_op(value),
                ) {
                    self.fallback_insts.push((op_pc, iref, 1));
                    (
                        VmOp::ArraySet1D {
                            array: p_arr,
                            index: p_idx,
                            value: p_val,
                        },
                        1,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::RecordGet {
                dst,
                record: OxOperand::Use(rec),
                index,
            } => {
                if let (Some(p_dst), Some(p_rec), Ok(field)) = (
                    PackedPlace::from_ox(*dst, num_locals),
                    PackedPlace::from_ox(*rec, num_locals),
                    u32::try_from(*index),
                ) {
                    self.fallback_insts.push((op_pc, iref, 1));
                    (
                        VmOp::RecordGet {
                            dst: p_dst,
                            record: p_rec,
                            field,
                        },
                        1,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::RecordSet {
                record,
                index,
                value,
            } => {
                if let (Some(p_rec), Some(p_val), Ok(field)) = (
                    PackedPlace::from_ox(*record, num_locals),
                    self.pack_op(value),
                    u32::try_from(*index),
                ) {
                    self.fallback_insts.push((op_pc, iref, 1));
                    (
                        VmOp::RecordSet {
                            record: p_rec,
                            field,
                            value: p_val,
                        },
                        1,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::FieldGet { dst, object, field } => {
                if let (Some(p_dst), Some(p_obj)) = (
                    PackedPlace::from_ox(forward_dst.unwrap_or(*dst), num_locals),
                    self.pack_op(object),
                ) {
                    (
                        VmOp::FieldGet {
                            dst: p_dst,
                            object: p_obj,
                            field: *field,
                        },
                        consumed_if_fwd,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::FieldSet {
                object,
                field,
                value,
            } => {
                if let (Some(p_obj), Some(p_val)) = (self.pack_op(object), self.pack_op(value)) {
                    (
                        VmOp::FieldSet {
                            object: p_obj,
                            field: *field,
                            value: p_val,
                        },
                        1,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::CallProc { dst, proc, args } => {
                let effective_dst = match (dst, forward_dst) {
                    (Some(_), Some(fwd)) => Some(fwd),
                    (Some(d), None) => Some(*d),
                    (None, _) => None,
                };
                let used_fwd = dst.is_some() && forward_dst.is_some();
                let p_dst = match effective_dst {
                    Some(d) => match PackedPlace::from_ox(d, num_locals) {
                        Some(p) => Some(p),
                        None => return (VmOp::SlowInst(iref), 1),
                    },
                    None => None,
                };
                if let Some((packed_args, all_byval)) = self.pack_ox_args(args) {
                    let idx = self.proc_calls.len() as u32;
                    self.proc_calls.push(FastProcCall {
                        dst: p_dst,
                        proc: *proc,
                        args: packed_args,
                        all_byval,
                    });
                    (VmOp::CallProcFast(idx), if used_fwd { 2 } else { 1 })
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::CallExtern { dst, import, args } => {
                let Some((id, string_typed_alias)) = resolve_vba_library_import(program, *import)
                else {
                    return (VmOp::SlowInst(iref), 1);
                };
                if id == NativeImplId::StrConv {
                    return (VmOp::SlowInst(iref), 1);
                }
                let effective_dst = match (dst, forward_dst) {
                    (Some(_), Some(fwd)) => Some(fwd),
                    (Some(d), None) => Some(*d),
                    (None, _) => None,
                };
                let used_fwd = dst.is_some() && forward_dst.is_some();
                let p_dst = match effective_dst {
                    Some(d) => match PackedPlace::from_ox(d, num_locals) {
                        Some(p) => Some(p),
                        None => return (VmOp::SlowInst(iref), 1),
                    },
                    None => None,
                };
                if let Some((packed_args, _)) = self.pack_ox_args(args) {
                    let idx = self.lib_calls.len() as u32;
                    self.lib_calls.push(FastLibCall {
                        dst: p_dst,
                        id,
                        string_typed_alias,
                        pure_no_events: is_pure_no_events_builtin(id),
                        args: packed_args,
                    });
                    (VmOp::CallExternLib(idx), if used_fwd { 2 } else { 1 })
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            OxInst::ComCallEarly {
                dst,
                method,
                invoke_kind,
                recv,
                args,
            } => {
                let p_dst = match dst {
                    Some(d) => match PackedPlace::from_ox(*d, num_locals) {
                        Some(p) => Some(p),
                        None => return (VmOp::SlowInst(iref), 1),
                    },
                    None => None,
                };
                let Some(p_recv) = self.pack_op(recv) else {
                    return (VmOp::SlowInst(iref), 1);
                };
                let positional_args = self.pack_positional_call_args(args);
                let idx = self.com_early_calls.len() as u32;
                self.com_early_calls.push(FastComEarly {
                    dst: p_dst,
                    method: *method,
                    invoke_kind: *invoke_kind,
                    recv: p_recv,
                    positional_args,
                    inst: iref,
                    ic_descriptor: Cell::new(0),
                    ic_member: Cell::new(None),
                });
                (VmOp::CallComEarly(idx), 1)
            }
            OxInst::ComCallLate {
                dst,
                recv,
                name,
                default_member,
                invoke_kind,
                args,
            } => {
                let p_dst = match dst {
                    Some(d) => match PackedPlace::from_ox(*d, num_locals) {
                        Some(p) => Some(p),
                        None => return (VmOp::SlowInst(iref), 1),
                    },
                    None => None,
                };
                let Some(p_recv) = self.pack_op(recv) else {
                    return (VmOp::SlowInst(iref), 1);
                };
                let positional_args = self.pack_positional_call_args(args);
                let public_only = !operand_is_project_typed_receiver(program, func, recv);
                let idx = self.com_late_calls.len() as u32;
                self.com_late_calls.push(FastComLate {
                    dst: p_dst,
                    recv: p_recv,
                    name: name.clone(),
                    default_member: *default_member,
                    invoke_kind: *invoke_kind,
                    public_only,
                    positional_args,
                    inst: iref,
                    ic_descriptor: Cell::new(0),
                    ic_member: Cell::new(None),
                });
                (VmOp::CallComLate(idx), 1)
            }
            OxInst::StmtBoundary {
                clear_temps_from, ..
            } => {
                if let Ok(from) = u32::try_from(*clear_temps_from) {
                    (
                        VmOp::StmtBoundary {
                            clear_temps_from: from,
                        },
                        1,
                    )
                } else {
                    (VmOp::SlowInst(iref), 1)
                }
            }
            _ => (VmOp::SlowInst(iref), 1),
        }
    }
}
