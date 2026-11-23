//! Operands, places, constants, and the small operator enums.
//!
//! OxIR keeps a slot-like value model — simple **places** (locals/globals/temps)
//! that ops read and write — exactly like the proven `calcvba_bundle::isa` machine,
//! but typed. Complex/fallible accesses (a COM field, an array element, a record
//! field) are *not* hidden inside an operand; they are explicit instructions
//! ([`crate::inst::OxInst`]) that read into / write out of a simple place. This
//! mirrors the bundle ISA and keeps effects and faults visible.

use serde::{Deserialize, Serialize};

use crate::ids::{GlobalId, LocalId, TempId};

/// A typed literal. Floats/date/currency are stored as bit patterns so the IR stays
/// `Eq`/`Hash` and round-trips losslessly (mirrors `calcvba_bundle::coreir::CoreConst`).
#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum OxConst {
    Empty,
    Null,
    Nothing,
    Bool(bool),
    I16(i16),
    I32(i32),
    I64(i64),
    /// IEEE-754 `f32` bit pattern.
    F32(u32),
    /// IEEE-754 `f64` bit pattern.
    F64(u64),
    /// Currency scaled ×10_000.
    Currency(i64),
    /// `Date` serial as an `f64` bit pattern.
    Date(u64),
    Str(String),
    /// `vbNullString`: a String whose BSTR is null (`StrPtr` 0, `NULL` to a `Declare`);
    /// everywhere else it is `""`.
    NullStr,
    /// Missing: the Variant an omitted argument carries (`VT_ERROR` holding
    /// `DISP_E_PARAMNOTFOUND`, 0x80020004), so `IsMissing` is True.
    Missing,
}

/// A simple, directly-addressable place: a local, a global, or a temporary. These
/// are the read/write targets of ops; everything else (fields, array elements,
/// record fields, WithEvents sinks) is reached through an explicit instruction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum OxPlace {
    Local(LocalId),
    Global(GlobalId),
    Temp(TempId),
}

/// An operand: a constant, or a pure read of a simple place.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum OxOperand {
    Const(OxConst),
    Use(OxPlace),
}

impl OxOperand {
    /// Convenience: a `Use` of a local.
    pub fn local(id: LocalId) -> Self {
        OxOperand::Use(OxPlace::Local(id))
    }
    /// Convenience: a `Use` of a temporary.
    pub fn temp(id: TempId) -> Self {
        OxOperand::Use(OxPlace::Temp(id))
    }
}

/// Binary arithmetic operators that carry a `NumericMode` (the `Checked`/`Widening`
/// regime). `Div`/`Pow` are separate instructions because they are always `f64`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum ArithOp {
    Add,
    Sub,
    Mul,
    IntDiv,
    Mod,
}

/// Ordered/equality comparison operators (carry a `StringCompareMode`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum CmpOp {
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
}

/// Bitwise/logical operators (VBA `And`/`Or`/… are bitwise on numbers, logical on
/// Booleans — the runtime decides by operand type).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum LogicalOp {
    And,
    Or,
    Xor,
    Eqv,
    Imp,
}

/// Where an explicit value-changing coercion lands. Pure representation changes
/// (typed ↔ Variant) are [`crate::inst::OxInst::Box`]/[`crate::inst::OxInst::Unbox`],
/// not coercions. `ImplicitVariant` widens to the [`crate::ty::OxTy`] `Variant` top
/// type without changing the value.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum OxCoerceTarget {
    Numeric(calcvba_bundle::NumericCoerceTarget),
    Str,
    FixedStr(u32),
    ImplicitVariant,
}

/// An argument to a compiled VBA procedure call (`CallProc`/`CallProcRef`/
/// `CallExtern`). VBA passes `ByRef` by default — a true alias of the caller's
/// place; `ByVal` copies the value in; `Omitted` leaves an optional parameter unset.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum OxArg {
    ByVal(OxOperand),
    ByRef(OxPlace),
    Omitted,
}

/// An argument to a `CallNative` / `CallByName` (the late/native dispatch path).
/// `ByRef` carries a place for marshaled copy-out; `Named` carries a `name := value`
/// argument; `Const` is a compile-time integer threaded as a trailing argument
/// (e.g. an `Option Compare` mode).
#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum OxCallArg {
    Operand(OxOperand),
    /// Explicit call-site `ByVal` overriding a `ByRef` `Declare` parameter.
    ByValOverride(OxOperand),
    ByRef(OxPlace),
    Omitted,
    /// `by_ref` is the place a late-bound named argument aliases, passed ByRef like
    /// [`OxCallArg::ByRef`]; `value` is then `Use` of that place.
    Named {
        name: String,
        value: OxOperand,
        by_ref: Option<OxPlace>,
    },
    Const(i32),
}

/// What a `CallNative` targets. COM dispatch is **not** here — it becomes a
/// first-class, typed instruction added with the typed COM interface +
/// method-descriptor tables (next sub-section); `CallNative` covers only the
/// base-library built-ins and `Declare Lib` external calls.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum OxNativeCallee {
    /// A base-library built-in with a native body.
    Builtin(calcvba_bundle::NativeImplId),
    /// A `Declare Lib` external call; the marshalling descriptor lives in the
    /// program's `external_calls` table, keyed by `descriptor_id`.
    Declare { descriptor_id: u32 },
}

/// Where an [`OxInst::AddrOf`](crate::OxInst::AddrOf) path starts: a variable or
/// temporary's cell, or a field of a class instance.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum OxAddrBase {
    Place(OxPlace),
    ObjectField { object: OxOperand, field: i32 },
}

/// One step of an [`OxInst::AddrOf`](crate::OxInst::AddrOf) path.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum OxAddrStep {
    /// An element of the array held here, by declared subscripts.
    Index(Vec<OxOperand>),
    /// Positional field of the UDT record held here.
    Field(usize),
}

// ── Small leaf enums — reused from `calcvba_bundle::coreir` (the one truth) ──────
//
// `BoundWhich` (`LBound`/`UBound`), `PtrKind` (`VarPtr`/`StrPtr`/…) and `ErrField`
// (`Err.Number`/…) are defined once in `coreir` and reused here verbatim — OxIR does not
// re-model them. They carry `serde` (added on the canonical definitions), so they
// serialize into the `.oxi` image directly.
pub use calcvba_bundle::coreir::{BoundWhich, ErrField, PtrKind};
