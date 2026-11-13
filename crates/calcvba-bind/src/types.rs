//! VBA type lattice + coercion (binder side). Pure functions mapping the symbol
//! model's `VarTypeRef` into the coreir coercion / array / target-kind vocabulary,
//! plus the operator result lattice and the `coerce`-node insertion the
//! expression binder uses.
//!
//! Type inference here is deliberately a *safe over-approximation*: when in doubt
//! the result is `Variant`. The VM computes the real value; the binder only uses
//! the inferred type to decide which coercion node (if any) to insert. Inferring
//! `Variant` therefore never produces a wrong result — at worst a redundant
//! `Coerce` on assignment to a typed variable.

use calcvba_bundle::coreir::{CoerceTarget, CoreArg, CoreBinOp, CoreCallee, CoreConst, CoreValue};
use calcvba_bundle::native::NativeImplId;
use calcvba_bundle::{ArrayElementType, AssignmentTargetKind, NumericCoerceTarget, NumericMode};
use calcvba_symbol::cond_comp::{ConditionalCompilationPointerWidth, ConditionalCompilationTarget};
use calcvba_symbol::signature::{BuiltinType, VarTypeRef};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct TypeContext {
    long_ptr_width: ConditionalCompilationPointerWidth,
}

impl TypeContext {
    pub(crate) fn from_target(target: ConditionalCompilationTarget) -> Self {
        Self {
            long_ptr_width: target.pointer_width,
        }
    }

    pub(crate) fn longptr_is_64(self) -> bool {
        self.long_ptr_width == ConditionalCompilationPointerWidth::Bits64
    }
}

impl Default for TypeContext {
    fn default() -> Self {
        Self::from_target(ConditionalCompilationTarget::default())
    }
}

// ── Type → coreir mappings (frame builder + ReDim) ──────────────────────────

/// The coreir array element type for a declared element type.
pub(crate) fn array_element_of_with(ctx: TypeContext, elem: &VarTypeRef) -> ArrayElementType {
    match elem {
        VarTypeRef::Builtin(b) => builtin_element(ctx, *b),
        VarTypeRef::FixedString(_) => ArrayElementType::String,
        VarTypeRef::FixedArray { element, bounds } => ArrayElementType::FixedArray {
            element: Box::new(array_element_of_with(ctx, element)),
            bounds: bounds.clone(),
        },
        VarTypeRef::Object(_) | VarTypeRef::Udt(_) | VarTypeRef::Variant | VarTypeRef::Array(_) => {
            ArrayElementType::Variant
        }
    }
}

fn builtin_element(ctx: TypeContext, b: BuiltinType) -> ArrayElementType {
    match b {
        BuiltinType::Boolean => ArrayElementType::Boolean,
        BuiltinType::Byte => ArrayElementType::Byte,
        BuiltinType::Integer => ArrayElementType::Integer,
        BuiltinType::Long => ArrayElementType::Long,
        BuiltinType::LongLong => ArrayElementType::LongLong,
        BuiltinType::LongPtr if ctx.longptr_is_64() => ArrayElementType::LongLong,
        BuiltinType::LongPtr => ArrayElementType::Long,
        BuiltinType::Single => ArrayElementType::Single,
        BuiltinType::Double => ArrayElementType::Double,
        BuiltinType::Currency => ArrayElementType::Currency,
        BuiltinType::Date => ArrayElementType::Date,
        BuiltinType::String => ArrayElementType::String,
    }
}

/// `Some(element)` when the declared type is an array, for `CoreLocal`/`CoreGlobal`.
pub(crate) fn array_element_with(ctx: TypeContext, ty: &VarTypeRef) -> Option<ArrayElementType> {
    match ty {
        VarTypeRef::Array(inner) => Some(array_element_of_with(ctx, inner)),
        VarTypeRef::FixedArray { element, .. } => Some(array_element_of_with(ctx, element)),
        _ => None,
    }
}

/// The static assignment-target kind (drives Let/Set legality + `ValidateAssignment`).
pub fn assignment_target_kind(ty: &VarTypeRef) -> AssignmentTargetKind {
    match ty {
        VarTypeRef::Object(_) => AssignmentTargetKind::Object,
        VarTypeRef::Variant => AssignmentTargetKind::Variant,
        // A UDT is a value aggregate: `q = p` is a `Let` (value copy of the record),
        // never `Set`.
        VarTypeRef::Builtin(_)
        | VarTypeRef::Udt(_)
        | VarTypeRef::Array(_)
        | VarTypeRef::FixedArray { .. }
        | VarTypeRef::FixedString(_) => AssignmentTargetKind::Scalar,
    }
}

/// A display name for a declared type (for `ValidateAssignment` diagnostics).
pub fn type_name(ty: &VarTypeRef) -> String {
    match ty {
        VarTypeRef::Builtin(b) => format!("{b:?}"),
        VarTypeRef::Object(name) | VarTypeRef::Udt(name) => name.clone(),
        VarTypeRef::Variant => "Variant".into(),
        VarTypeRef::Array(inner) => format!("{}()", type_name(inner)),
        VarTypeRef::FixedArray { element, bounds } => {
            format!(
                "{}({})",
                type_name(element),
                fixed_array_bounds_name(bounds)
            )
        }
        VarTypeRef::FixedString(_) => "String".into(),
    }
}

fn fixed_array_bounds_name(bounds: &[calcvba_bundle::FixedArrayBound]) -> String {
    bounds
        .iter()
        .map(|bound| {
            let upper = i64::from(bound.lower)
                + i64::try_from(bound.len.saturating_sub(1)).unwrap_or(i64::MAX);
            format!("{} To {}", bound.lower, upper)
        })
        .collect::<Vec<_>>()
        .join(", ")
}

// ── Coercion insertion ──────────────────────────────────────────────────────

/// The coercion target when storing into a declared scalar type, or `None` when
/// the VM's store coercion suffices (`Variant`/`Object`/array).
pub(crate) fn coerce_target_with(ctx: TypeContext, to: &VarTypeRef) -> Option<CoerceTarget> {
    match to {
        VarTypeRef::Builtin(BuiltinType::String) => Some(CoerceTarget::String),
        VarTypeRef::Builtin(b) => numeric_target(ctx, *b).map(CoerceTarget::Numeric),
        // Assigning to a fixed-length string pads/truncates to its length.
        VarTypeRef::FixedString(len) => Some(CoerceTarget::FixedString(*len as usize)),
        _ => None,
    }
}

fn numeric_target(ctx: TypeContext, b: BuiltinType) -> Option<NumericCoerceTarget> {
    Some(match b {
        // Storing into a `Boolean` is `CBool`: non-zero → `True`, zero → `False`.
        BuiltinType::Boolean => NumericCoerceTarget::Boolean,
        BuiltinType::Byte => NumericCoerceTarget::Byte,
        BuiltinType::Integer => NumericCoerceTarget::Integer,
        BuiltinType::Long => NumericCoerceTarget::Long,
        BuiltinType::LongPtr if !ctx.longptr_is_64() => NumericCoerceTarget::Long,
        BuiltinType::LongPtr | BuiltinType::LongLong => NumericCoerceTarget::LongLong,
        BuiltinType::Single => NumericCoerceTarget::Single,
        BuiltinType::Double => NumericCoerceTarget::Double,
        BuiltinType::Currency => NumericCoerceTarget::Currency,
        BuiltinType::Date => NumericCoerceTarget::Date,
        BuiltinType::String => return None,
    })
}

/// The [`NumericMode`] for an arithmetic op (`+`/`-`/`*`/`\`/`Mod`/unary `-`) whose
/// static result type is `ty`: `Checked(target)` when the operands promote to a fixed
/// numeric type (so the VM computes in it and raises Overflow on out-of-range), else
/// `Widening` (a `Variant` operand, or the `Double` ceiling — the VBA widening regime).
/// VBA's overflow-vs-widen choice is *static* (identical run-time tags can mean either),
/// so the binder picks it here and stamps it on the op; the VM and JIT read it back.
/// `Boolean`/`String`/`Date`/`Variant` results map to `Widening` (non-arithmetic or
/// already-widest), as does `Double`.
pub(crate) fn numeric_mode_with(ctx: TypeContext, ty: &VarTypeRef) -> NumericMode {
    let VarTypeRef::Builtin(b) = ty else {
        return NumericMode::Widening;
    };
    match b {
        BuiltinType::Double | BuiltinType::Boolean | BuiltinType::Date | BuiltinType::String => {
            NumericMode::Widening
        }
        other => numeric_target(ctx, *other).map_or(NumericMode::Widening, NumericMode::Checked),
    }
}

/// Coerce a value being **stored** into a declared scalar target to that target's
/// type — unconditionally (unlike [`coerce`], which trusts the source's static type
/// and skips identity conversions). A declared variable must hold its declared type
/// at run time regardless of the value's runtime tag (for example, a small
/// unsuffixed literal carries `Integer`, while a declared `Long` target stores
/// `Long`). No-op for
/// `Object`/`Variant`/array targets (their store needs no scalar coercion). Skips
/// re-wrapping a value already coerced to the same target.
pub(crate) fn coerce_store_with(ctx: TypeContext, value: CoreValue, to: &VarTypeRef) -> CoreValue {
    let Some(target) = coerce_target_with(ctx, to) else {
        return value;
    };
    if let CoreValue::Coerce { to: existing, .. } = &value
        && *existing == target
    {
        return value;
    }
    CoreValue::Coerce {
        value: Box::new(value),
        to: target,
    }
}

/// Coerce `value` (of type `from`) to type `to`, wrapping in a `Coerce` node only
/// when a scalar conversion is actually needed (skips identity conversions).
fn is_string_into_byte_array(from: &VarTypeRef, to: &VarTypeRef) -> bool {
    matches!(from, VarTypeRef::Builtin(BuiltinType::String))
        && matches!(to, VarTypeRef::Array(inner) if matches!(**inner, VarTypeRef::Builtin(BuiltinType::Byte)))
}

/// `bytes = "text"` (a String stored into a Byte array) copies the string's UTF-16LE bytes, as
/// VBA does; lowered to StrConv with the internal UTF-16 mode. Any other pairing is unchanged.
pub(crate) fn string_into_byte_array(
    value: CoreValue,
    from: &VarTypeRef,
    to: &VarTypeRef,
) -> CoreValue {
    if !is_string_into_byte_array(from, to) {
        return value;
    }
    CoreValue::Call {
        callee: CoreCallee::Native(NativeImplId::StrConv),
        args: vec![
            CoreArg::ByVal(value),
            CoreArg::ByVal(CoreValue::Const(CoreConst::I32(
                calcvba_bundle::STRCONV_UTF16_BYTES,
            ))),
        ],
    }
}

pub(crate) fn coerce_with(
    ctx: TypeContext,
    value: CoreValue,
    from: &VarTypeRef,
    to: &VarTypeRef,
) -> CoreValue {
    if from == to {
        return value;
    }
    if is_string_into_byte_array(from, to) {
        return string_into_byte_array(value, from, to);
    }
    // An array value is never numerically coercible — passing an array argument to a
    // parameter (`f(arr)`) copies the array, it does not coerce its elements. Without
    // this guard a `Dim a() As Byte` passed by value (or to a scalar-typed slot) would
    // emit a `CoerceNumeric` on the whole array and fault at run time.
    if matches!(from, VarTypeRef::Array(_)) {
        return value;
    }
    match coerce_target_with(ctx, to) {
        Some(target) => CoreValue::Coerce {
            value: Box::new(value),
            to: target,
        },
        None => value,
    }
}

pub(crate) fn is_longlong_with(ctx: TypeContext, ty: &VarTypeRef) -> bool {
    matches!(ty, VarTypeRef::Builtin(BuiltinType::LongLong))
        || (matches!(ty, VarTypeRef::Builtin(BuiltinType::LongPtr)) && ctx.longptr_is_64())
}

/// Office rejects an implicit `LongLong` -> `Byte`/`Integer`/`Long`/`Currency` conversion in a
/// store, argument or return at compile time (type mismatch); `Single` and `Double` accept it.
pub(crate) fn is_rejected_longlong_narrowing(
    ctx: TypeContext,
    from: &VarTypeRef,
    to: &VarTypeRef,
) -> bool {
    matches!(from, VarTypeRef::Builtin(BuiltinType::LongLong))
        && matches!(
            typed_numeric(ctx, to),
            Some(
                BuiltinType::Byte
                    | BuiltinType::Integer
                    | BuiltinType::Long
                    | BuiltinType::Currency
            )
        )
}

fn is_fixed_integer(ty: &VarTypeRef) -> bool {
    matches!(
        ty,
        VarTypeRef::Builtin(
            BuiltinType::Byte
                | BuiltinType::Integer
                | BuiltinType::Long
                | BuiltinType::LongLong
                | BuiltinType::LongPtr
        )
    )
}

// ── Operator result lattice ─────────────────────────────────────────────────

/// The (over-approximated) result type of a binary operator.
pub(crate) fn result_type_with(
    ctx: TypeContext,
    op: CoreBinOp,
    lhs: &VarTypeRef,
    rhs: &VarTypeRef,
) -> VarTypeRef {
    use CoreBinOp::*;
    match op {
        Eq | Ne | Lt | Le | Gt | Ge | Is | Like => builtin(BuiltinType::Boolean),
        Concat => builtin(BuiltinType::String),
        And | Or | Xor | Eqv | Imp => {
            if is_boolean(lhs) && is_boolean(rhs) {
                builtin(BuiltinType::Boolean)
            } else if is_variant(lhs) || is_variant(rhs) {
                VarTypeRef::Variant
            } else if (is_longlong_with(ctx, lhs) || is_longlong_with(ctx, rhs))
                && is_fixed_integer(lhs)
                && is_fixed_integer(rhs)
            {
                builtin(BuiltinType::LongLong)
            } else {
                builtin(BuiltinType::Long)
            }
        }
        IntDiv | Mod => int_div_result(ctx, lhs, rhs),
        Div => div_result(ctx, lhs, rhs),
        Pow => {
            if is_variant(lhs) || is_variant(rhs) {
                VarTypeRef::Variant
            } else {
                builtin(BuiltinType::Double)
            }
        }
        Add | Sub | Mul => arith_result(ctx, op, lhs, rhs),
    }
}

/// The numeric builtin behind a typed operand, if it is one of the seven VBA numeric types.
fn typed_numeric(ctx: TypeContext, ty: &VarTypeRef) -> Option<BuiltinType> {
    match ty {
        VarTypeRef::Builtin(b) => match b {
            BuiltinType::Byte
            | BuiltinType::Integer
            | BuiltinType::Long
            | BuiltinType::LongLong
            | BuiltinType::Single
            | BuiltinType::Double
            | BuiltinType::Currency => Some(*b),
            BuiltinType::LongPtr => Some(if ctx.longptr_is_64() {
                BuiltinType::LongLong
            } else {
                BuiltinType::Long
            }),
            _ => None,
        },
        _ => None,
    }
}

fn is_integer_builtin(b: BuiltinType) -> bool {
    matches!(
        b,
        BuiltinType::Byte | BuiltinType::Integer | BuiltinType::Long | BuiltinType::LongLong
    )
}

/// Result type of typed `+`, `-` and `*` as measured against Excel VBA (16.0 build 20527,
/// conformance/excel-measurements/numeric-mode-matrix.excel.txt): the wider integer type; a
/// `Double` operand gives `Double` (but `Currency` for `+`/`-` with `Currency`); `Currency` wins
/// over the other types except that `Currency * Single` is `Double`; `Single` with `Long` or
/// `LongLong` is `Double`, otherwise `Single`. A non-numeric operand makes the result `Variant`.
fn arith_result(ctx: TypeContext, op: CoreBinOp, lhs: &VarTypeRef, rhs: &VarTypeRef) -> VarTypeRef {
    use BuiltinType::*;
    let (Some(l), Some(r)) = (typed_numeric(ctx, lhs), typed_numeric(ctx, rhs)) else {
        return VarTypeRef::Variant;
    };
    let either = |t: BuiltinType| l == t || r == t;
    let additive = matches!(op, CoreBinOp::Add | CoreBinOp::Sub);
    let result = if either(Double) {
        if either(Currency) && additive {
            Currency
        } else {
            Double
        }
    } else if either(Currency) {
        if either(Single) && !additive {
            Double
        } else {
            Currency
        }
    } else if either(Single) {
        if either(Long) || either(LongLong) {
            Double
        } else {
            Single
        }
    } else {
        rank_to_builtin(rank_of_integer(l).max(rank_of_integer(r)))
    };
    builtin(result)
}

fn rank_of_integer(b: BuiltinType) -> u8 {
    match b {
        BuiltinType::Byte => 0,
        BuiltinType::Integer => 1,
        BuiltinType::Long => 2,
        _ => 3,
    }
}

/// `/`: `Double`, except `Single` when both operands are `Byte`, `Integer` or `Single` and at
/// least one is `Single`.
fn div_result(ctx: TypeContext, lhs: &VarTypeRef, rhs: &VarTypeRef) -> VarTypeRef {
    use BuiltinType::*;
    let (Some(l), Some(r)) = (typed_numeric(ctx, lhs), typed_numeric(ctx, rhs)) else {
        return VarTypeRef::Variant;
    };
    let small = |t: BuiltinType| matches!(t, Byte | Integer | Single);
    if small(l) && small(r) && (l == Single || r == Single) {
        builtin(Single)
    } else {
        builtin(Double)
    }
}

/// `\` and `Mod`: the wider of the two integer operand types, where a `Single`, `Double` or
/// `Currency` operand counts as `Long`, so `Byte \ Byte` is `Byte` and `Byte \ Integer` is
/// `Integer`. Non-numeric operands keep the engine's long-standing `Long` answer.
fn int_div_result(ctx: TypeContext, lhs: &VarTypeRef, rhs: &VarTypeRef) -> VarTypeRef {
    if is_longlong_with(ctx, lhs) || is_longlong_with(ctx, rhs) {
        return builtin(BuiltinType::LongLong);
    }
    // A Variant operand decides the result subtype at run time.
    if is_variant(lhs) || is_variant(rhs) {
        return VarTypeRef::Variant;
    }
    let (Some(l), Some(r)) = (typed_numeric(ctx, lhs), typed_numeric(ctx, rhs)) else {
        return builtin(BuiltinType::Long);
    };
    let rank = |t: BuiltinType| {
        if is_integer_builtin(t) {
            rank_of_integer(t)
        } else {
            2
        }
    };
    builtin(rank_to_builtin(rank(l).max(rank(r))))
}

fn rank_to_builtin(rank: u8) -> BuiltinType {
    match rank {
        0 => BuiltinType::Byte,
        1 => BuiltinType::Integer,
        2 => BuiltinType::Long,
        _ => BuiltinType::LongLong,
    }
}

fn builtin(b: BuiltinType) -> VarTypeRef {
    VarTypeRef::Builtin(b)
}

pub fn is_variant(ty: &VarTypeRef) -> bool {
    matches!(ty, VarTypeRef::Variant)
}

pub fn is_boolean(ty: &VarTypeRef) -> bool {
    matches!(ty, VarTypeRef::Builtin(BuiltinType::Boolean))
}

/// Consumed by the objects/COM phase (early/late dispatch by receiver type).
#[allow(dead_code)]
pub fn is_object(ty: &VarTypeRef) -> bool {
    matches!(ty, VarTypeRef::Object(_))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn long() -> VarTypeRef {
        builtin(BuiltinType::Long)
    }

    fn ctx() -> TypeContext {
        TypeContext::default()
    }

    #[test]
    fn widening_picks_the_wider_type() {
        assert_eq!(
            result_type_with(
                ctx(),
                CoreBinOp::Add,
                &builtin(BuiltinType::Integer),
                &builtin(BuiltinType::Double)
            ),
            builtin(BuiltinType::Double)
        );
        assert_eq!(
            result_type_with(ctx(), CoreBinOp::Mul, &long(), &long()),
            long()
        );
    }

    #[test]
    fn comparisons_and_concat_have_fixed_result() {
        assert_eq!(
            result_type_with(ctx(), CoreBinOp::Lt, &long(), &long()),
            builtin(BuiltinType::Boolean)
        );
        assert_eq!(
            result_type_with(
                ctx(),
                CoreBinOp::Concat,
                &long(),
                &builtin(BuiltinType::String)
            ),
            builtin(BuiltinType::String)
        );
    }

    #[test]
    fn variant_operand_propagates() {
        assert_eq!(
            result_type_with(ctx(), CoreBinOp::Add, &VarTypeRef::Variant, &long()),
            VarTypeRef::Variant
        );
    }

    #[test]
    fn coerce_skips_identity_and_wraps_narrowing() {
        let v = CoreValue::Const(calcvba_bundle::coreir::CoreConst::I32(5));
        // Long → Long: no node.
        assert!(matches!(
            coerce_with(ctx(), v.clone(), &long(), &long()),
            CoreValue::Const(_)
        ));
        // Long → Integer: a Coerce node.
        assert!(matches!(
            coerce_with(ctx(), v, &long(), &builtin(BuiltinType::Integer)),
            CoreValue::Coerce { .. }
        ));
        assert!(matches!(
            coerce_target_with(ctx(), &builtin(BuiltinType::String)),
            Some(CoerceTarget::String)
        ));
        // Currency target: now expressible.
        assert!(coerce_target_with(ctx(), &builtin(BuiltinType::Currency)).is_some());
    }
}
