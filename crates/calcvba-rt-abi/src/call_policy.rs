use crate::{Fault, coerce_declared_value, hidden_me_receiver_param_count};
use calcvba_eval::arith::{self, ArithError};
use calcvba_oxir::{OxFunc, OxLocal, OxTy};
use calcvba_runtime::{VarType, Variant};

/// Bind late-bound actuals by name or position while preserving the engine's argument modes.
/// Empty slots become omitted arguments at the engine's call boundary.
pub fn order_project_arguments<'a, A>(
    callee: &OxFunc,
    actuals: impl IntoIterator<Item = (A, Option<&'a str>)>,
) -> Result<Vec<Option<A>>, Fault> {
    let mut ordered = Vec::new();
    let mut next_positional = 0;
    for (arg, name) in actuals {
        let index = match name {
            Some(name) => callee
                .locals
                .iter()
                .take(callee.param_count)
                .skip(hidden_me_receiver_param_count(callee))
                .position(|param| param.name.eq_ignore_ascii_case(name))
                .ok_or_else(|| Fault::standard(448))?,
            None => {
                while ordered.get(next_positional).is_some_and(Option::is_some) {
                    next_positional += 1;
                }
                next_positional
            }
        };
        if ordered.len() <= index {
            ordered.resize_with(index + 1, || None);
        }
        if ordered[index].is_some() {
            return Err(Fault::standard(448));
        }
        ordered[index] = Some(arg);
    }
    Ok(ordered)
}

/// Initialize a private argument cell for a ByVal formal or a ByRef expression.
pub fn coerce_project_argument(value: &Variant, formal: &OxLocal) -> Result<Variant, ArithError> {
    validate_project_argument(value, formal)?;
    coerce_declared_value(value, &formal.ty)
}

/// A Null supplied to a late-bound ByVal scalar parameter raises type mismatch.
pub fn validate_project_argument(value: &Variant, formal: &OxLocal) -> Result<(), ArithError> {
    if value.vtype() == VarType::Null
        && formal
            .param
            .as_ref()
            .is_some_and(|p| !p.by_ref && !p.variadic)
        && (crate::declared_numeric_target(&formal.ty).is_some()
            || matches!(formal.ty, OxTy::Str | OxTy::FixedStr(_)))
    {
        return Err(ArithError::type_mismatch());
    }
    Ok(())
}

/// A slot already carrying a reference is forwarded through its existing engine alias.
/// For a typed slot, build a `VT_BYREF` view of its payload for a Variant formal.
///
/// # Safety
/// The caller must keep the cell's storage live and stable with its declared type
/// for the entire use of the returned reference.
pub unsafe fn byref_variant_payload(
    formal: &OxTy,
    declared: &OxTy,
    cell: &mut Variant,
) -> Option<Variant> {
    if *formal != OxTy::Variant
        || !declared.is_byref_variant_reference()
        || cell.byref_target().is_some()
    {
        return None;
    }
    let vt = cell.core().raw_vt();
    // SAFETY: the caller preserves the typed cell's payload for the reference lifetime.
    Some(unsafe { Variant::by_ref(vt, Variant::payload_ptr(cell)) })
}

/// Prepare a ParamArray write before touching either the actual slot or the pack.
/// The pack sees a variable-length String even when the actual has a fixed width.
pub fn param_array_write_values(
    value: &Variant,
    actual_ty: &OxTy,
) -> Result<(Variant, Variant), ArithError> {
    if let OxTy::FixedStr(_) = actual_ty {
        let element = arith::coerce_string(value)?;
        let actual = coerce_declared_value(&element, actual_ty)?;
        return Ok((actual, element));
    }
    let actual = coerce_declared_value(value, actual_ty)?;
    Ok((actual.clone(), actual))
}

/// Every element that aliases the actual must observe the same write.
pub fn write_param_array_aliases<A: PartialEq>(
    array: &mut Variant,
    aliases: &[Option<A>],
    target: A,
    element: &Variant,
) -> Result<(), Fault> {
    for (index, alias) in aliases.iter().enumerate() {
        if alias.as_ref() == Some(&target) {
            array
                .set_safearray_element(index, element)
                .map_err(|message| Fault::new(13, message))?;
        }
    }
    Ok(())
}
