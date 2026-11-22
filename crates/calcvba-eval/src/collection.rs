//! Built-in `VBA.Collection`.
//!
//! The data model lives in `calcvba-runtime` (carried on the object box's `native_state`
//! slot) so vm3 and the JIT share ONE implementation; this module re-exports it so
//! the long-standing `calcvba_eval::collection::*` import path keeps resolving. The keyed
//! method *dispatch* shim ([`dispatch_collection`]) lives here in `calcvba-eval`, beside the
//! rest of the shared value/builtin kernel — it is the single place that turns a member +
//! `Variant` args into a `CollectionData` operation, so no VM re-implements it.

pub use calcvba_runtime::collection::{CollectionData, CollectionError, Selector, fold_key};

use calcvba_runtime::{VarType, Variant};

/// The omitted-argument sentinel (`DISP_E_PARAMNOTFOUND`) the VMs marshal an omitted optional
/// into; the shared dispatcher recognises it to tell "omitted" from a real value. (The VMs
/// carry their own copies of this stable HRESULT for marshalling.)
const MISSING_ARG: i32 = 0x8002_0004u32 as i32;

/// A built-in `Collection` member, resolved from the call-site member name/dispid by the VM.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CollectionMethod {
    Add,
    Item,
    Count,
    Remove,
}

/// Put a call's arguments in `method`'s parameter order (`Item, Key, Before, After` /
/// `Index`): `names[i]` is the name a named argument was written with (`Before:=1`). Gaps are
/// the omitted sentinel. Errors are VBA run-time error numbers: 448 for an unknown or repeated
/// name, 450 for too many arguments.
pub fn order_collection_args(
    method: CollectionMethod,
    values: &[Variant],
    names: &[Option<String>],
) -> Result<Vec<Variant>, i32> {
    if names.iter().all(Option::is_none) {
        return Ok(values.to_vec());
    }
    let parameters: &[&str] = match method {
        CollectionMethod::Add => &["Item", "Key", "Before", "After"],
        CollectionMethod::Item | CollectionMethod::Remove => &["Index"],
        CollectionMethod::Count => &[],
    };
    let mut ordered = vec![Variant::from_error_code(MISSING_ARG); parameters.len()];
    let mut occupied = vec![false; parameters.len()];
    for (position, value) in values.iter().enumerate() {
        let index = match names.get(position).and_then(Option::as_deref) {
            Some(name) => parameters
                .iter()
                .position(|parameter| parameter.eq_ignore_ascii_case(name))
                .ok_or(448)?,
            None => position,
        };
        if index >= parameters.len() {
            return Err(450);
        }
        if occupied[index] {
            return Err(448);
        }
        occupied[index] = true;
        ordered[index] = value.clone();
    }
    Ok(ordered)
}

/// The shared keyed-`Collection` dispatch: run `method` against `data` (the receiver's
/// box-owned [`CollectionData`], reached via `ObjectRef::with_native_collection`) with the
/// positional method `args` (omitted optionals arrive as the `MISSING_ARG` sentinel). Returns
/// the method result, or a [`CollectionError`] the caller maps to a VBA run-time error number.
/// This is the single implementation VM3 and the JIT both call.
pub fn dispatch_collection(
    method: CollectionMethod,
    data: &mut CollectionData,
    args: &[Variant],
) -> Result<Variant, CollectionError> {
    match method {
        CollectionMethod::Count => Ok(Variant::from_i32(data.count())),
        CollectionMethod::Add => {
            let value = args
                .first()
                .filter(|value| !is_missing_arg(value))
                .ok_or(CollectionError::WrongArgumentCount)?;
            let key = optional_key(args.get(1))?;
            let before = optional_selector(args.get(2));
            let after = optional_selector(args.get(3));
            data.add(value.clone(), key, before, after)?;
            Ok(Variant::empty())
        }
        CollectionMethod::Item => {
            let sel = required_selector(data, args.first())?;
            data.item(&sel)
        }
        CollectionMethod::Remove => {
            let sel = required_selector(data, args.first())?;
            data.remove(&sel)?;
            Ok(Variant::empty())
        }
    }
}

fn is_missing_arg(v: &Variant) -> bool {
    v.as_error_code() == Some(MISSING_ARG)
}

fn fold_string_variant(v: &Variant) -> Option<String> {
    let units = v.borrowed_string_units()?;
    if units.iter().all(|&u| u < 0x80) {
        let mut folded = String::with_capacity(units.len());
        for &u in units {
            folded.push((u as u8).to_ascii_lowercase() as char);
        }
        Some(folded)
    } else {
        Some(fold_key(&String::from_utf16_lossy(units)))
    }
}

/// A selector from a Variant: a string is a (folded) key, anything else a 1-based index
/// (lenient numeric coercion; a non-numeric, non-string value yields index 0 → out of range).
fn variant_selector(v: &Variant) -> Selector {
    if let Some(folded) = fold_string_variant(v) {
        Selector::Key(folded)
    } else {
        // Any numeric subtype is a 1-based Long index. Coerce through the shared
        // integer conversion (VBA banker's rounding) rather than a hand-rolled
        // as_i16/i32/i64/f64 chain that let Byte/SignedByte/Single/unsigned
        // subtypes fall through to index 0 (so `c.Item(CByte(1))` raised error 9).
        let index = crate::arith::int(v)
            .ok()
            .and_then(|n| i32::try_from(n).ok())
            .unwrap_or(0);
        Selector::Index(index)
    }
}

/// The folded string key from an `Add` key argument, `None` if omitted; a key that is not a
/// String is a type mismatch.
fn optional_key(arg: Option<&Variant>) -> Result<Option<String>, CollectionError> {
    let Some(v) = arg.filter(|v| !is_missing_arg(v)) else {
        return Ok(None);
    };
    match fold_string_variant(v) {
        Some(folded) => Ok(Some(folded)),
        None => Err(CollectionError::TypeMismatch),
    }
}

/// A `Selector` from an optional `before`/`after` argument, or `None` if omitted.
fn optional_selector(arg: Option<&Variant>) -> Option<Selector> {
    let v = arg?;
    (!is_missing_arg(v)).then(|| variant_selector(v))
}

/// A required `Item`/`Remove` selector. Omitted → [`CollectionError::ArgNotOptional`] (449).
fn required_selector(
    data: &CollectionData,
    arg: Option<&Variant>,
) -> Result<Selector, CollectionError> {
    let value = arg
        .filter(|value| !is_missing_arg(value))
        .ok_or(CollectionError::ArgNotOptional)?;
    if data.count() == 0 {
        return Err(CollectionError::BadArgument);
    }
    if matches!(
        value.vtype(),
        VarType::Null | VarType::Error | VarType::ArrayVariant | VarType::Object
    ) {
        return Err(CollectionError::TypeMismatch);
    }
    Ok(variant_selector(value))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn count_add_item_remove_through_dispatch() {
        let mut c = CollectionData::default();
        dispatch_collection(CollectionMethod::Add, &mut c, &[Variant::from_i32(10)]).unwrap();
        dispatch_collection(CollectionMethod::Add, &mut c, &[Variant::from_i32(20)]).unwrap();
        assert_eq!(
            dispatch_collection(CollectionMethod::Count, &mut c, &[])
                .unwrap()
                .as_i32(),
            Some(2)
        );
        assert_eq!(
            dispatch_collection(CollectionMethod::Item, &mut c, &[Variant::from_i32(1)])
                .unwrap()
                .as_i32(),
            Some(10)
        );
        dispatch_collection(CollectionMethod::Remove, &mut c, &[Variant::from_i32(1)]).unwrap();
        assert_eq!(
            dispatch_collection(CollectionMethod::Item, &mut c, &[Variant::from_i32(1)])
                .unwrap()
                .as_i32(),
            Some(20)
        );
    }

    #[test]
    fn keyed_add_dup_missing_and_omitted_map_to_error_codes() {
        let mut c = CollectionData::default();
        dispatch_collection(
            CollectionMethod::Add,
            &mut c,
            &[Variant::from_i32(1), Variant::from_string("A")],
        )
        .unwrap();
        // Case-insensitive keyed Item → 1.
        assert_eq!(
            dispatch_collection(CollectionMethod::Item, &mut c, &[Variant::from_string("a")])
                .unwrap()
                .as_i32(),
            Some(1)
        );
        // Duplicate key → 457.
        assert_eq!(
            dispatch_collection(
                CollectionMethod::Add,
                &mut c,
                &[Variant::from_i32(2), Variant::from_string("a")]
            )
            .unwrap_err(),
            CollectionError::DuplicateKey
        );
        // Missing key → 9.
        assert_eq!(
            dispatch_collection(
                CollectionMethod::Item,
                &mut c,
                &[Variant::from_string("zz")]
            )
            .unwrap_err(),
            CollectionError::BadArgument
        );
        // Omitted required selector → 449.
        assert_eq!(
            dispatch_collection(CollectionMethod::Item, &mut c, &[]).unwrap_err(),
            CollectionError::ArgNotOptional
        );
        // An explicit MISSING_ARG sentinel also reads as omitted → 449.
        assert_eq!(
            dispatch_collection(
                CollectionMethod::Item,
                &mut c,
                &[Variant::from_error_code(MISSING_ARG)]
            )
            .unwrap_err(),
            CollectionError::ArgNotOptional
        );
    }

    #[test]
    fn add_before_positions_the_insertion_via_dispatch() {
        let mut c = CollectionData::default();
        dispatch_collection(CollectionMethod::Add, &mut c, &[Variant::from_i32(10)]).unwrap();
        dispatch_collection(CollectionMethod::Add, &mut c, &[Variant::from_i32(30)]).unwrap();
        // Add 20 with key omitted (MISSING_ARG) and before=index 2 → [10, 20, 30].
        dispatch_collection(
            CollectionMethod::Add,
            &mut c,
            &[
                Variant::from_i32(20),
                Variant::from_error_code(MISSING_ARG),
                Variant::from_i32(2),
            ],
        )
        .unwrap();
        let vals: Vec<i32> = c.values().iter().filter_map(|v| v.as_i32()).collect();
        assert_eq!(vals, vec![10, 20, 30]);
    }

    #[test]
    fn integer_literal_selector_indexes_collection() {
        let mut c = CollectionData::default();
        dispatch_collection(CollectionMethod::Add, &mut c, &[Variant::from_i32(10)]).unwrap();
        dispatch_collection(CollectionMethod::Add, &mut c, &[Variant::from_i32(20)]).unwrap();
        assert_eq!(
            dispatch_collection(CollectionMethod::Item, &mut c, &[Variant::from_i16(2)])
                .unwrap()
                .as_i32(),
            Some(20)
        );
    }
}
