//! Explicit COM compatibility adapters for scalar slot-token projections.

use calcvba_runtime::{
    ObjectRefExt, VarType, Variant,
    safe_array::{array_tag_from_safe_array, marshal_dispatch_argument},
};

/// The scalar i32 token of a cell, for the legacy compat-object dispatch lane.
pub fn variant_to_runtime_token(value: &Variant) -> Result<i32, String> {
    let invalid = || format!("invalid {:?} payload", value.vtype());
    match value.vtype() {
        VarType::ByRef => variant_to_runtime_token(&value.clone()),
        VarType::Empty | VarType::Null => Ok(0),
        VarType::Error => value.as_error_code().ok_or_else(invalid),
        VarType::Boolean => Ok(i32::from(value.as_bool().ok_or_else(invalid)?)),
        VarType::Integer => Ok(i32::from(value.as_i16().ok_or_else(invalid)?)),
        VarType::Long => value.as_i32().ok_or_else(invalid),
        VarType::Byte => Ok(i32::from(value.as_u8().ok_or_else(invalid)?)),
        VarType::SignedByte => Ok(i32::from(value.as_i8().ok_or_else(invalid)?)),
        VarType::UnsignedInteger => Ok(i32::from(value.as_u16().ok_or_else(invalid)?)),
        VarType::UnsignedLong | VarType::UnsignedInt => {
            let raw = value.as_u32().ok_or_else(invalid)?;
            i32::try_from(raw).map_err(|_| format!("COM value {raw} is outside i32 token range"))
        }
        VarType::LongLong => {
            let raw = value.as_i64().ok_or_else(invalid)?;
            i32::try_from(raw)
                .map_err(|_| format!("COM i64 value {raw} is outside i32 token range"))
        }
        VarType::UnsignedLongLong => {
            let raw = value.as_u64().ok_or_else(invalid)?;
            i32::try_from(raw)
                .map_err(|_| format!("COM u64 value {raw} is outside i32 token range"))
        }
        VarType::Single => Ok(value.as_f32().ok_or_else(invalid)? as i32),
        VarType::Double => Ok(value.as_f64().ok_or_else(invalid)? as i32),
        VarType::Date => Ok(value.as_date_f64().ok_or_else(invalid)? as i32),
        VarType::Decimal => value
            .as_decimal96()
            .ok_or_else(invalid)?
            .to_string()
            .parse::<i32>()
            .map_err(|err| format!("COM decimal value cannot be represented as i32: {err}")),
        VarType::Currency => {
            Ok((value.as_currency_scaled_i64().ok_or_else(invalid)? / 10_000) as i32)
        }
        VarType::String => value
            .as_bstr()
            .ok_or_else(invalid)?
            .as_str()
            .trim()
            .parse::<i32>()
            .map_err(|err| format!("COM string value cannot be represented as i32: {err}")),
        VarType::ArrayVariant => Ok(value
            .as_safearray()
            .map_or(0, |array| array.len().min(i32::MAX as usize) as i32)),
        VarType::Object => Ok(value.as_object_ref().map_or(0, |object| object.raw())),
        VarType::Record => Err("COM record value cannot be represented as i32".to_string()),
        VarType::ProcRef => Err("procedure references cannot be passed to COM".to_string()),
    }
}

/// As [`variant_to_runtime_token`], with arrays as legacy dispatch array tags.
pub fn variant_to_legacy_dispatch_token(value: &Variant) -> Result<i32, String> {
    if let Some(array) = value.as_safearray() {
        let token = array_tag_from_safe_array(&array)
            .ok_or_else(|| "unsupported array intent for legacy dispatch transport".to_string())?;
        return Ok(marshal_dispatch_argument(token));
    }
    variant_to_runtime_token(value)
}
