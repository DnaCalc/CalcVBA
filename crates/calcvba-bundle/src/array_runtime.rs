//! Pure `ArrayElementType` → runtime-value/layout helpers, shared by every executor.
//!
//! `ReDim`, array literals, and UDT construction need to (a) pick the SAFEARRAY element
//! vartype for a declared element type, (b) seed a freshly-`ReDim`-ed element with its typed
//! default, (c) build a [`SafeArray`] with the right element storage, and (d) project a UDT
//! field list into a native [`VbaRecordLayout`]. These are pure functions of
//! [`ArrayElementType`] (no VM state), so vm3 and the Cranelift backend share one
//! definition rather than mirroring it.

use std::sync::Arc;

use calcvba_runtime::Variant;
use calcvba_runtime::safe_array::{
    SafeArray, SafeArrayBound, VT_BOOL_VALUE, VT_BSTR_VALUE, VT_CY_VALUE, VT_DATE_VALUE,
    VT_I2_VALUE, VT_I4_VALUE, VT_I8_VALUE, VT_R4_VALUE, VT_R8_VALUE, VT_RECORD_VALUE, VT_UI1_VALUE,
    VT_VARIANT_VALUE,
};
use calcvba_runtime::vba_record::{
    VbaRecord, VbaRecordFieldKind, VbaRecordFieldSpec, VbaRecordLayout,
};

use crate::ArrayElementType;

/// An unallocated UDT array still needs its declared record prototype when a
/// native callee allocates its first SAFEARRAY. Its engine descriptor has no
/// bounds or payload; the native boundary projects it as a null SAFEARRAY*.
pub fn unallocated_typed_array(element: &ArrayElementType) -> Result<Variant, String> {
    if let ArrayElementType::Record(fields) = element {
        // Declaration itself does not allocate elements. Layouts outside the
        // runtime's record limits must still raise at ReDim, inside VBA's error
        // handling, rather than aborting activation before any statement runs.
        if let Ok(layout) = vba_record_layout_for_fields(fields) {
            return SafeArray::from_vba_records_nd(Vec::new(), layout, Vec::new())
                .map(Variant::from_safearray);
        }
    }
    Ok(Variant::unallocated_array(safearray_vartype_for_element(
        element,
    )))
}

#[test]
fn unallocated_record_array_retains_layout_but_projects_a_null_native_pointer() {
    let element = ArrayElementType::Record(vec![
        ArrayElementType::Variant,
        ArrayElementType::FixedString(3),
    ]);
    let value = unallocated_typed_array(&element).unwrap();
    for value in [value.clone(), value] {
        assert!(value.safearray_bounds_len().is_none());
        let array = value.try_as_safearray().unwrap().unwrap();
        assert!(array.bounds().is_none());
        assert_eq!(array.len(), 0);
        assert_eq!(
            array.vba_record_layout().unwrap(),
            vba_record_layout_for_fields(match &element {
                ArrayElementType::Record(fields) => fields,
                _ => unreachable!(),
            })
            .unwrap()
        );
        let wire = value.to_wire_bytes();
        assert_eq!(&wire[..2], &0x2024u16.to_le_bytes());
        assert_eq!(&wire[2..], &[0; 14]);
    }
}
/// Choose the actual element type for ReDim. A Variant keeps its contained
/// array's type unless the source explicitly supplies an As clause.
pub fn resolve_redim_element(
    current: &Variant,
    declared: &ArrayElementType,
    inherit_current: bool,
    preserve: bool,
) -> Result<ArrayElementType, i32> {
    let old = current.array_element_vartype();
    let element = if inherit_current {
        if let Some(vt) = old {
            let element = array_element_type_for_vartype(vt);
            if safearray_vartype_for_element(&element) != vt {
                return Err(13);
            }
            element
        } else {
            declared.clone()
        }
    } else {
        declared.clone()
    };
    if preserve && old.is_some_and(|vt| vt != safearray_vartype_for_element(&element)) {
        return Err(13);
    }
    Ok(element)
}

/// Build a [`SafeArray`] of `elems` shaped by `bounds`, using the SAFEARRAY storage that
/// matches the declared `element_type`: a UDT element stores native VBA records; every other
/// element type stores a typed scalar / Variant element.
///
/// `fixed` carries the VBA fixed-vs-dynamic distinction onto the descriptor's
/// `FADF_FIXEDSIZE` bit: a fixed-size declarator (`Dim a(1 To 3)`) and a UDT
/// fixed-array field build `fixed = true` arrays (whose `Erase` resets elements),
/// while a user `ReDim` builds `fixed = false` (whose `Erase` deallocates).
pub fn redim_safearray_from_elements(
    bounds: Vec<SafeArrayBound>,
    element_type: &ArrayElementType,
    elems: Vec<Variant>,
    fixed: bool,
) -> Result<SafeArray, String> {
    if let ArrayElementType::Record(fields) = element_type {
        let layout = vba_record_layout_for_fields(fields)?;
        let records = elems
            .into_iter()
            .map(|value| {
                value
                    .as_vba_record()
                    .ok_or_else(|| "UDT array element default was not a VBA record".to_string())
            })
            .collect::<Result<Vec<_>, String>>()?;
        return Ok(SafeArray::from_vba_records_nd(bounds, layout, records)?.with_fixed_size(fixed));
    }
    Ok(SafeArray::from_typed_variants_nd(
        bounds,
        safearray_vartype_for_element(element_type),
        elems,
    )?
    .with_fixed_size(fixed))
}

/// The SAFEARRAY element vartype (`VT_*`) for a declared array element type.
pub fn safearray_vartype_for_element(element_type: &ArrayElementType) -> u16 {
    match element_type {
        ArrayElementType::Variant => VT_VARIANT_VALUE,
        ArrayElementType::Record(_) => VT_RECORD_VALUE,
        ArrayElementType::FixedArray { .. } => VT_VARIANT_VALUE,
        ArrayElementType::Integer => VT_I2_VALUE,
        ArrayElementType::Long => VT_I4_VALUE,
        ArrayElementType::LongLong => VT_I8_VALUE,
        ArrayElementType::Byte => VT_UI1_VALUE,
        ArrayElementType::Single => VT_R4_VALUE,
        ArrayElementType::Double => VT_R8_VALUE,
        ArrayElementType::Currency => VT_CY_VALUE,
        ArrayElementType::Date => VT_DATE_VALUE,
        ArrayElementType::String | ArrayElementType::FixedString(_) => VT_BSTR_VALUE,
        ArrayElementType::Boolean => VT_BOOL_VALUE,
    }
}

/// The declared array element type for a SAFEARRAY's runtime element vartype (`VT_*`) — the
/// inverse of [`safearray_vartype_for_element`] for the intrinsic scalar/string/Variant carriers.
/// Used by `Erase` of a fixed-size array to re-default each element to the array's OWN element
/// type (runtime truth) rather than a possibly-erased bind-site type (a fixed array held in a
/// `Variant` slot keeps its real element type). `VT_RECORD` and any object/unknown carrier fall
/// back to `Variant` — record arrays are handled separately via their stored layout, and a typed
/// object SAFEARRAY is stored as `VT_VARIANT` by this runtime, so neither reaches the fallback.
pub fn array_element_type_for_vartype(element_vt: u16) -> ArrayElementType {
    match element_vt {
        VT_I2_VALUE => ArrayElementType::Integer,
        VT_I4_VALUE => ArrayElementType::Long,
        VT_I8_VALUE => ArrayElementType::LongLong,
        VT_UI1_VALUE => ArrayElementType::Byte,
        VT_R4_VALUE => ArrayElementType::Single,
        VT_R8_VALUE => ArrayElementType::Double,
        VT_CY_VALUE => ArrayElementType::Currency,
        VT_DATE_VALUE => ArrayElementType::Date,
        VT_BOOL_VALUE => ArrayElementType::Boolean,
        VT_BSTR_VALUE => ArrayElementType::String,
        _ => ArrayElementType::Variant,
    }
}

/// The default value for a freshly-`ReDim`-ed array element. A UDT-record element is a native
/// VBA record with descriptor-backed field offsets, built recursively so nested scalar-UDT
/// subfields are themselves default records — mirroring the binder's `emit_udt_record_init` so
/// `Lines(i).Words(j).Rect.X` reads back. Scalar defaults are typed zero values so typed
/// SAFEARRAY storage can encode them directly instead of relying on a boundary-time Variant
/// projection.
pub fn default_array_element(element_type: &ArrayElementType) -> Result<Variant, String> {
    let value = match element_type {
        ArrayElementType::Variant => Variant::empty(),
        ArrayElementType::Integer => Variant::from_i16(0),
        ArrayElementType::Long => Variant::from_i32(0),
        ArrayElementType::LongLong => Variant::from_i64(0),
        ArrayElementType::Byte => Variant::from_u8(0),
        ArrayElementType::Single => Variant::from_f32(0.0),
        ArrayElementType::Double => Variant::from_f64(0.0),
        ArrayElementType::Currency => Variant::from_currency_scaled_i64(0),
        ArrayElementType::Date => Variant::from_date_f64(0.0),
        ArrayElementType::String => Variant::from_string(""),
        ArrayElementType::FixedString(len) => Variant::from_utf16_units(&vec![0; *len]),
        ArrayElementType::Boolean => Variant::from_bool(false),
        // A UDT-record element can fail to lay out or allocate for guest-legal
        // types (a record above the 64 KiB size limit, a zero-sized fixed-array
        // or string field, excessive nesting depth). These used to `expect()`
        // and abort the host on an ordinary `ReDim arr(..) As BigUdt`; return the
        // error so callers can seat a recoverable type mismatch (13), mirroring
        // the scalar `NewRecord` path.
        ArrayElementType::Record(fields) => {
            let layout = vba_record_layout_for_fields(fields)?;
            let record = VbaRecord::new_default(layout)?;
            Variant::from_vba_record(record)
        }
        ArrayElementType::FixedArray { .. } => Variant::empty(),
    };
    Ok(value)
}

/// Project a UDT field list into a native [`VbaRecordLayout`] (anonymous fields, in order).
pub fn vba_record_layout_for_fields(
    fields: &[ArrayElementType],
) -> Result<Arc<VbaRecordLayout>, String> {
    let specs = fields
        .iter()
        .map(|field| Ok(VbaRecordFieldSpec::anonymous(vba_record_field_kind(field)?)))
        .collect::<Result<Vec<_>, String>>()?;
    Ok(Arc::new(VbaRecordLayout::new(specs)?))
}

/// Map one UDT field's declared element type to its native record field kind (recursive for
/// nested records and fixed arrays).
pub fn vba_record_field_kind(
    element_type: &ArrayElementType,
) -> Result<VbaRecordFieldKind, String> {
    let kind = match element_type {
        ArrayElementType::Variant => VbaRecordFieldKind::Variant,
        ArrayElementType::Integer => VbaRecordFieldKind::Integer,
        ArrayElementType::Long => VbaRecordFieldKind::Long,
        ArrayElementType::LongLong => VbaRecordFieldKind::LongLong,
        ArrayElementType::Byte => VbaRecordFieldKind::Byte,
        ArrayElementType::Single => VbaRecordFieldKind::Single,
        ArrayElementType::Double => VbaRecordFieldKind::Double,
        ArrayElementType::Currency => VbaRecordFieldKind::Currency,
        ArrayElementType::Date => VbaRecordFieldKind::Date,
        ArrayElementType::String => VbaRecordFieldKind::String,
        ArrayElementType::FixedString(len) => VbaRecordFieldKind::FixedString { len: *len },
        ArrayElementType::Boolean => VbaRecordFieldKind::Boolean,
        ArrayElementType::Record(fields) => {
            VbaRecordFieldKind::Record(vba_record_layout_for_fields(fields)?)
        }
        ArrayElementType::FixedArray { element, bounds } => {
            let bounds = bounds
                .iter()
                .map(|bound| {
                    Ok(SafeArrayBound {
                        count: u32::try_from(bound.len).map_err(|_| {
                            "fixed-array record field length exceeds SAFEARRAY capacity".to_string()
                        })?,
                        lower: bound.lower,
                    })
                })
                .collect::<Result<Vec<_>, String>>()?;
            VbaRecordFieldKind::FixedArray {
                element: Box::new(vba_record_field_kind(element)?),
                bounds,
            }
        }
    };
    Ok(kind)
}
