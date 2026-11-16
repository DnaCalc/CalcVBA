//! `Declare` calls over the runtime cells (design doc, step 9): the callee gets the caller's
//! storage itself wherever VBA passes it, and copies are made only where VBA makes them.
//!
//! - A ByRef scalar, Variant, typed array or array element is the address of the caller's
//!   storage (the `VT_BYREF` reference's target), so the callee writes the variable directly and
//!   two arguments naming one variable are one address. A ByRef argument of another type (an
//!   expression, or a variable of a different type) gets a temporary of the parameter's type,
//!   written back through the reference after the call.
//! - A String (ByVal or ByRef, also inside `As Any`) is an ANSI BSTR temporary, as VBA converts
//!   it (best-fit `CP_ACP`); a null BSTR is a null pointer. After the call the temporary is
//!   converted back at its full byte length into the variable when there is one (a ByVal String
//!   variable binds as a reference for that copy-back, the pre-sized-buffer idiom).
//! - String arrays and records holding Strings are ANSI copies converted back after the call;
//!   other arrays pass the `SAFEARRAY*` slot itself.
//! - ByVal scalars take VBA's argument coercion (rounding, overflow), never truncation; a ByVal
//!   object is its interface pointer; a ByVal Variant is a pointer to a copy of the cell (x64).
//! - A Variant result moves into the runtime as it is.
//!
//! The library, the entry point and a libffi CIF are prepared once per `Declare` and kept.
use super::StandardHostServices;
use crate::error::{HalError, HalResult};
use crate::model::CapabilityId;
use crate::traits::{DynLinkDescriptorView, write_through_reference};
use calcvba_com::call::{ComArg, ComArgValue};
use calcvba_com::windows_ffi_bridge::{
    FfiReturnType, OwnedLibrary, PreparedNativeCall, get_proc_address, get_proc_address_ordinal,
};
use calcvba_runtime::{VarType, Variant};
use std::rc::Rc;
use std::sync::Arc;
use windows_sys::Win32::Foundation::{SysAllocStringByteLen, SysFreeString, SysStringByteLen};
use windows_sys::Win32::Globalization::{CP_ACP, MultiByteToWideChar, WideCharToMultiByte};

const VT_I2: u16 = 0x0002;
const VT_I4: u16 = 0x0003;
const VT_R4: u16 = 0x0004;
const VT_R8: u16 = 0x0005;
const VT_CY: u16 = 0x0006;
const VT_DATE: u16 = 0x0007;
const VT_BSTR: u16 = 0x0008;
const VT_BOOL: u16 = 0x000B;
const VT_VARIANT: u16 = 0x000C;
const VT_UI1: u16 = 0x0011;
const VT_I8: u16 = 0x0014;
const VT_ARRAY: u16 = 0x2000;

/// A `Declare` prepared once: its library (kept loaded), entry point and CIF.
pub(crate) struct DeclarePlan {
    /// The loaded DLL, kept for the plan; `None` for one of the runtime's own exports.
    library: Option<Arc<OwnedLibrary>>,
    call: PreparedNativeCall,
    result: ResultShape,
}

#[derive(Clone)]
enum ResultShape {
    /// A scalar, String (an ANSI BSTR) or nothing, decoded by type name.
    Scalar(Option<String>),
    /// A VARIANT returned through the hidden result pointer.
    Variant,
    /// A record returned by value, in its declared (ANSI) shape.
    Record(Arc<calcvba_runtime::vba_record::VbaRecordLayout>),
}

/// The prepared-plan cache key: everything the signature depends on.
pub(crate) type PlanKey = (u32, String, String, Vec<String>, Vec<bool>, Option<String>);

fn fault(host: &StandardHostServices, message: impl Into<String>) -> HalError {
    HalError::adapter_fault(
        host.profile,
        CapabilityId::DynamicLinking,
        "invoke_symbol",
        message,
    )
}

fn vba_fault(host: &StandardHostServices, number: i32, message: impl Into<String>) -> HalError {
    fault(host, message).with_host_error_code(number)
}

/// The native type of one argument slot.
fn slot_type(param_type: &str, by_ref: bool) -> FfiReturnType {
    if by_ref {
        return FfiReturnType::LongPtr;
    }
    match param_type {
        "Long" => FfiReturnType::Long,
        "Integer" => FfiReturnType::Integer,
        "Byte" => FfiReturnType::Byte,
        "Boolean" => FfiReturnType::Boolean,
        "Double" | "Date" => FfiReturnType::Double,
        "Single" => FfiReturnType::Single,
        "LongLong" | "Currency" => FfiReturnType::LongLong,
        // LongPtr, String (an LPSTR), Variant (a pointer to a copy), Any, Array.
        _ => FfiReturnType::LongPtr,
    }
}

fn scalar_result(return_type: Option<&str>) -> FfiReturnType {
    match return_type {
        None => FfiReturnType::Void,
        Some("Long") => FfiReturnType::Long,
        Some("Integer") => FfiReturnType::Integer,
        Some("Byte") => FfiReturnType::Byte,
        Some("Boolean") => FfiReturnType::Boolean,
        Some("Double" | "Date") => FfiReturnType::Double,
        Some("Single") => FfiReturnType::Single,
        Some("LongLong" | "Currency") => FfiReturnType::LongLong,
        Some(_) => FfiReturnType::LongPtr,
    }
}

/// The `VARTYPE` a typed parameter's storage has.
fn param_vt(param_type: &str) -> Option<u16> {
    Some(match param_type {
        "Long" => VT_I4,
        "Integer" => VT_I2,
        "Byte" => VT_UI1,
        "Boolean" => VT_BOOL,
        "Double" => VT_R8,
        "Single" => VT_R4,
        "Currency" => VT_CY,
        "Date" => VT_DATE,
        "LongLong" | "LongPtr" => VT_I8,
        "String" => VT_BSTR,
        "Variant" => VT_VARIANT,
        _ => return None,
    })
}

/// Loads (or reuses) the plan for `descriptor`.
pub(super) fn plan(
    host: &StandardHostServices,
    descriptor: &DynLinkDescriptorView<'_>,
) -> HalResult<Rc<DeclarePlan>> {
    let key: PlanKey = (
        descriptor.descriptor_id,
        descriptor.library.to_ascii_lowercase(),
        descriptor.alias.to_string(),
        descriptor.param_types.to_vec(),
        descriptor.param_by_ref.to_vec(),
        descriptor.return_type.as_deref().map(str::to_string),
    );
    let mut state = host
        .dynlink_state
        .try_borrow_mut()
        .map_err(|_| fault(host, "dynlink binding table already borrowed"))?;
    if let Some(plan) = state.declare_plans.get(&key) {
        return Ok(Rc::clone(plan));
    }
    // The VBA runtime's own exports are CalcVBA's (it is the runtime); the rest load the DLL.
    if let Some(address) = runtime_export(descriptor.library, descriptor.alias) {
        let plan = Rc::new(prepare(host, descriptor, None, address)?);
        state.declare_plans.insert(key, Rc::clone(&plan));
        return Ok(plan);
    }
    let library_key = descriptor.library.to_ascii_lowercase();
    let library = match state.native_modules.get(&library_key) {
        Some(library) => Arc::clone(library),
        None => {
            let library = Arc::new(
                OwnedLibrary::load(descriptor.library)
                    .map_err(|message| vba_fault(host, 48, message))?,
            );
            state
                .native_modules
                .insert(library_key, Arc::clone(&library));
            library
        }
    };
    let address = if descriptor.ordinal_alias {
        let digits = descriptor
            .alias
            .strip_prefix('#')
            .unwrap_or(descriptor.alias);
        let ordinal: u16 = digits
            .parse()
            .map_err(|_| fault(host, format!("invalid ordinal `{}`", descriptor.alias)))?;
        get_proc_address_ordinal(library.handle(), ordinal)
    } else {
        get_proc_address(library.handle(), descriptor.alias)
    }
    .map_err(|_| {
        let entry = descriptor
            .alias
            .strip_prefix('#')
            .unwrap_or(descriptor.alias);
        vba_fault(
            host,
            if descriptor.ordinal_alias { 452 } else { 453 },
            format!(
                "Can't find DLL entry point {entry} in {}",
                descriptor.library
            ),
        )
    })?;
    let plan = Rc::new(prepare(host, descriptor, Some(library), address)?);
    state.declare_plans.insert(key, Rc::clone(&plan));
    Ok(plan)
}

/// The VBA runtime's own exports a `Declare` may name, which CalcVBA, being the runtime,
/// provides itself: `Lib "VBE7" Alias "VarPtr"` (also `vbe6`, `msvbvm60`) returns its
/// argument, so `ByRef a() As Any` gives the address of the array's `SAFEARRAY*`.
fn runtime_export(library: &str, entry: &str) -> Option<usize> {
    extern "system" fn var_ptr(pointer: usize) -> usize {
        pointer
    }
    let file = library
        .trim()
        .rsplit(['\\', '/'])
        .next()?
        .to_ascii_lowercase();
    let stem = file.strip_suffix(".dll").unwrap_or(&file);
    let runtime = matches!(stem, "vbe7" | "vbe6" | "msvbvm60" | "msvbvm50");
    (runtime && entry.eq_ignore_ascii_case("VarPtr"))
        .then_some((var_ptr as *const ()).expose_provenance())
}

/// The address a runtime export (`VarPtr`) gets for a ByRef UDT array: the variable's own
/// `SAFEARRAY*` slot, as VBA passes it, Strings and all (probe variant-layout: Excel reports the
/// original array's `fFeatures` through `VarPtrArray(fa)`). A DLL gets an Automation-layout copy
/// instead, which lives only for the call; the identity export needs no conversion, and its
/// result must stay valid while the variable does. `None` for anything else (marshalled as for
/// any `Declare`).
fn runtime_own_storage(arg: ComArgValue<'_>) -> Option<u64> {
    let ComArgValue::Reference(reference) = arg else {
        return None;
    };
    // Cloning a reference reads its target: the array value.
    let value = reference.clone();
    value.record_array_layout()?;
    match reference.byref_target()? {
        // A typed array variable: its own SAFEARRAY* slot.
        (vt, pointer) if vt & VT_ARRAY != 0 => Some(pointer.expose_provenance() as u64),
        // A cell holding the array: the parray field of that cell.
        (VT_VARIANT, pointer) => {
            Some(Variant::payload_ptr(pointer.cast()).expose_provenance() as u64)
        }
        _ => None,
    }
}

/// The prepared call (CIF and result shape) for `address`.
fn prepare(
    host: &StandardHostServices,
    descriptor: &DynLinkDescriptorView<'_>,
    library: Option<Arc<OwnedLibrary>>,
    address: usize,
) -> HalResult<DeclarePlan> {
    let slots: Vec<FfiReturnType> = (0..descriptor.param_count)
        .map(|index| {
            slot_type(
                descriptor
                    .param_types
                    .get(index)
                    .map_or("Long", String::as_str),
                descriptor.param_by_ref.get(index).copied().unwrap_or(false),
            )
        })
        .collect();
    let (call, result) = if let Some(layout) = &descriptor.return_record_layout {
        let bytes = layout
            .declare_shape()
            .map_err(|message| fault(host, message))?
            .0;
        (
            PreparedNativeCall::aggregate_return(address, &slots, bytes),
            ResultShape::Record(layout.clone()),
        )
    } else if descriptor.return_type.as_deref() == Some("Variant") {
        (
            PreparedNativeCall::aggregate_return(address, &slots, calcvba_runtime::WIRE_SIZE),
            ResultShape::Variant,
        )
    } else {
        let return_type = descriptor.return_type.as_deref();
        (
            PreparedNativeCall::new(address, &slots, scalar_result(return_type)),
            ResultShape::Scalar(return_type.map(str::to_string)),
        )
    };
    Ok(DeclarePlan {
        library,
        call,
        result,
    })
}

/// An ANSI BSTR of `units` (`CP_ACP`, best fit, as VBA converts a `Declare` String).
fn ansi_bstr(units: &[u16]) -> Result<*mut u16, String> {
    let bytes = if units.is_empty() {
        Vec::new()
    } else {
        let len = i32::try_from(units.len()).map_err(|_| "String too long for ANSI")?;
        // SAFETY: a size query over a live slice.
        let required = unsafe {
            WideCharToMultiByte(
                CP_ACP,
                0,
                units.as_ptr(),
                len,
                core::ptr::null_mut(),
                0,
                core::ptr::null(),
                core::ptr::null_mut(),
            )
        };
        let mut bytes = vec![0u8; required.max(0) as usize];
        // SAFETY: the output buffer has the queried size.
        let written = unsafe {
            WideCharToMultiByte(
                CP_ACP,
                0,
                units.as_ptr(),
                len,
                bytes.as_mut_ptr(),
                required,
                core::ptr::null(),
                core::ptr::null_mut(),
            )
        };
        bytes.truncate(written.max(0) as usize);
        bytes
    };
    let len = u32::try_from(bytes.len()).map_err(|_| "ANSI String too long")?;
    // SAFETY: copies `len` initialized bytes into a new BSTR.
    let bstr = unsafe { SysAllocStringByteLen(bytes.as_ptr(), len) }.cast_mut();
    if bstr.is_null() {
        return Err("ANSI BSTR allocation failed".into());
    }
    Ok(bstr)
}

/// The Unicode String an ANSI BSTR converts back to, at its full byte length.
///
/// # Safety
/// `bstr` is null or a live BSTR.
unsafe fn unicode_from_ansi(bstr: *mut u16) -> Variant {
    if bstr.is_null() {
        return Variant::from_string(calcvba_runtime::bstr::BStr::empty());
    }
    // SAFETY: a live BSTR's byte length and payload.
    let bytes = unsafe {
        let len = SysStringByteLen(bstr) as usize;
        core::slice::from_raw_parts(bstr.cast::<u8>(), len)
    };
    if bytes.is_empty() {
        return Variant::from_string(calcvba_runtime::bstr::BStr::empty());
    }
    let len = i32::try_from(bytes.len()).unwrap_or(i32::MAX);
    // SAFETY: a size query over a live slice.
    let required =
        unsafe { MultiByteToWideChar(CP_ACP, 0, bytes.as_ptr(), len, core::ptr::null_mut(), 0) };
    let mut units = vec![0u16; required.max(0) as usize];
    // SAFETY: the output buffer has the queried size.
    let written = unsafe {
        MultiByteToWideChar(CP_ACP, 0, bytes.as_ptr(), len, units.as_mut_ptr(), required)
    };
    units.truncate(written.max(0) as usize);
    Variant::from_utf16_units(&units)
}

/// The UTF-16 units of a String value (empty for a non-String) and whether it is a null BSTR.
fn string_units(value: &Variant) -> (Vec<u16>, bool) {
    if value.vtype() == VarType::String && value.raw_bstr_ptr().is_null() {
        return (Vec::new(), true);
    }
    (
        value
            .as_bstr()
            .map(|text| text.payload_units().to_vec())
            .unwrap_or_default(),
        false,
    )
}

/// What happens to one argument after the call.
enum After<'a> {
    /// An ANSI BSTR temporary (boxed, so `&*cell` is the `BSTR*` a ByRef String passes),
    /// converted back into `target` when there is one, then freed.
    Ansi {
        cell: Box<*mut u16>,
        target: Option<&'a Variant>,
    },
    /// A temporary cell of the parameter's type, written back through `target`; `before` is
    /// what it owned before the call (the callee may take that over).
    Temp {
        cell: Box<Variant>,
        target: Option<&'a Variant>,
        before: calcvba_runtime::live_counters::OwnedPayload,
    },
    /// An ANSI record copy converted back through `target`.
    Record {
        record: calcvba_runtime::pointer_helpers::AutomationRecord,
        target: Option<&'a Variant>,
    },
    /// A record-array copy converted back through `target`.
    RecordArray {
        array: calcvba_runtime::pointer_helpers::AutomationRecordArray,
        target: Option<&'a Variant>,
    },
    /// A String array whose elements were made ANSI, converted back through `target`.
    AnsiArray {
        cell: Box<Variant>,
        target: Option<&'a Variant>,
    },
    /// A cell kept alive for the call (a ByVal Variant copy, a project object's IDispatch).
    Keep(Box<Variant>),
    /// Caller storage passed directly (a Variant cell or an array's `SAFEARRAY*` slot), which the
    /// callee may take over: what it owned before the call, for the live-handle accounting.
    Watch {
        storage: *const u8,
        slot_vt: Option<u16>,
        before: calcvba_runtime::live_counters::OwnedPayload,
    },
}

/// What a temporary cell owns, for the live-handle accounting.
fn temp_payload(cell: &Variant) -> calcvba_runtime::live_counters::OwnedPayload {
    // SAFETY: a live cell.
    unsafe {
        calcvba_runtime::live_counters::owned_payload((cell as *const Variant).cast::<u8>(), None)
    }
}

/// Passes caller storage directly, watching what it owns.
fn direct<'a>(pointer: *mut u8, slot_vt: Option<u16>, after: &mut Vec<After<'a>>) -> u64 {
    // SAFETY: the reference's target is live storage of that shape for the call.
    let before = unsafe { calcvba_runtime::live_counters::owned_payload(pointer, slot_vt) };
    after.push(After::Watch {
        storage: pointer,
        slot_vt,
        before,
    });
    pointer.expose_provenance() as u64
}

/// Writes `value` into the storage `reference` names, with VBA's assignment coercion.
fn write_back(host: &StandardHostServices, reference: &Variant, value: Variant) -> HalResult<()> {
    write_through_reference(reference, value).map_err(|error| {
        let number = error.host_error_code.unwrap_or(13);
        vba_fault(host, number, error.message)
    })
}

/// VBA's argument coercion of `value` to a typed parameter's `VARTYPE`.
fn coerce(host: &StandardHostServices, value: &Variant, vt: u16) -> HalResult<Variant> {
    if value.core().raw_vt() == vt {
        return Ok(value.clone());
    }
    calcvba_eval::arith::coerce_for_byref_target(value, vt)
        .map_err(|error| vba_fault(host, error.code, error.message))
}

/// The 8 payload bytes of a cell as a slot value.
fn payload_word(value: &Variant) -> u64 {
    let bytes = value.to_wire_bytes();
    u64::from_le_bytes(bytes[8..16].try_into().expect("8 payload bytes"))
}

/// An array's element `VARTYPE` (`VT_BSTR` for String arrays), from a cell holding or referring
/// to it.
fn array_element_vt(value: &Variant) -> Option<u16> {
    let vt = match value.byref_target() {
        Some((vt, _)) => vt,
        None => value.core().raw_vt(),
    };
    (vt & VT_ARRAY != 0).then_some(vt & 0x0FFF)
}

/// One argument's slot word plus what follows the call.
fn marshal<'a>(
    host: &StandardHostServices,
    param_type: &str,
    by_ref: bool,
    arg: ComArgValue<'a>,
    after: &mut Vec<After<'a>>,
) -> HalResult<u64> {
    let (value, reference): (Variant, Option<&'a Variant>) = match arg {
        ComArgValue::Value(value) => (value.clone(), None),
        // Cloning a reference reads its target.
        ComArgValue::Reference(reference) => (reference.clone(), Some(reference)),
        // An omitted optional parameter (VBA's `Optional` in a Declare) is the type's zero.
        ComArgValue::Missing => (Variant::empty(), None),
    };
    let target = reference.and_then(Variant::byref_target);
    let address = |pointer: *mut u8| pointer.expose_provenance() as u64;

    // Strings: an ANSI BSTR temporary, ByVal or ByRef.
    if param_type == "String" || (param_type == "Any" && value.vtype() == VarType::String) {
        let (units, null) = string_units(&value);
        let bstr = if null {
            core::ptr::null_mut()
        } else {
            ansi_bstr(&units).map_err(|message| fault(host, message))?
        };
        let mut cell = Box::new(bstr);
        let word = if by_ref {
            address((&mut *cell as *mut *mut u16).cast())
        } else {
            bstr.expose_provenance() as u64
        };
        after.push(After::Ansi {
            cell,
            target: reference,
        });
        return Ok(word);
    }

    if !by_ref {
        return match param_type {
            "Variant" => {
                // A pointer to a copy of the cell (an x64 by-value aggregate).
                let mut cell = Box::new(value);
                let word = address((&mut *cell as *mut Variant).cast());
                after.push(After::Keep(cell));
                Ok(word)
            }
            "Any" | "Array" => byval_any(host, value, after),
            _ => {
                let vt = param_vt(param_type).unwrap_or(VT_I4);
                Ok(payload_word(&coerce(host, &value, vt)?))
            }
        };
    }

    // ByRef arrays: the SAFEARRAY* slot, except String and record arrays (copies).
    if param_type == "Array" || (param_type == "Any" && value.holds_array()) {
        if let Some(layout) = value.record_array_layout() {
            let array = value
                .try_as_safearray()
                .map_err(|message| fault(host, message))?;
            let mut array = calcvba_runtime::pointer_helpers::AutomationRecordArray::from_array(
                layout,
                array.as_ref(),
            )
            .map_err(|message| fault(host, message))?;
            let word = array.as_cell_ptr().expose_provenance() as u64;
            after.push(After::RecordArray {
                array,
                target: reference,
            });
            return Ok(word);
        }
        if array_element_vt(&value) == Some(VT_BSTR) {
            let mut cell = Box::new(ansi_array(host, value)?);
            let word = address(Variant::payload_ptr(&mut *cell));
            after.push(After::AnsiArray {
                cell,
                target: reference,
            });
            return Ok(word);
        }
        return match target {
            // A typed array variable: its own SAFEARRAY* slot.
            Some((vt, pointer)) if vt & VT_ARRAY != 0 => Ok(direct(pointer, Some(vt), after)),
            // A Variant holding the array: the parray field of its cell.
            Some((VT_VARIANT, pointer)) if value.holds_array() => {
                direct(pointer, None, after);
                Ok(address(Variant::payload_ptr(pointer.cast())))
            }
            _ => {
                let mut cell = Box::new(value);
                let word = address(Variant::payload_ptr(&mut *cell));
                // The callee may take over what the temporary owns (VariantClear): watched.
                let before = temp_payload(&cell);
                after.push(After::Temp {
                    cell,
                    target: reference,
                    before,
                });
                Ok(word)
            }
        };
    }

    // Records (a UDT `As Any`): the record's own storage when it holds no Strings, otherwise an
    // ANSI copy converted back after the call.
    if value.vtype() == VarType::Record {
        // A record variable arrives as a reference to the cell holding it.
        if let Some((VT_VARIANT, pointer)) = target
            && value
                .vba_record_layout()
                .is_some_and(|layout| !layout.holds_text())
            // SAFETY: the reference's target is the live record variable for the call.
            && let Some(data) = unsafe { Variant::vba_record_data_at(pointer.cast()) }
        {
            return Ok(address(data));
        }
        let record = value
            .as_vba_record()
            .ok_or_else(|| fault(host, "native ByRef record requires a VBA record value"))?;
        let mut record =
            calcvba_runtime::pointer_helpers::AutomationRecord::from_declared_record(&record)
                .map_err(|message| fault(host, message))?;
        let word = record.as_mut_ptr().expose_provenance() as u64;
        after.push(After::Record {
            record,
            target: reference,
        });
        return Ok(word);
    }

    // The caller's storage itself when it has the parameter's type (or any type for As Any).
    let wanted = match param_type {
        "Any" => None,
        other => param_vt(other),
    };
    if let Some((vt, pointer)) = target {
        let matches = match wanted {
            Some(wanted) => vt == wanted,
            None => true,
        };
        if matches {
            return Ok(if vt == VT_VARIANT {
                direct(pointer, None, after)
            } else {
                address(pointer)
            });
        }
    }
    // A temporary of the parameter's type.
    let mut cell = Box::new(match wanted {
        Some(VT_VARIANT) => value,
        Some(vt) => coerce(host, &value, vt)?,
        None => value,
    });
    let pointer = if wanted == Some(VT_VARIANT) {
        (&mut *cell as *mut Variant).cast::<u8>()
    } else {
        Variant::payload_ptr(&mut *cell)
    };
    let word = address(pointer);
    // The callee may take over what the temporary owns (VariantClear): watched.
    let before = temp_payload(&cell);
    after.push(After::Temp {
        cell,
        target: reference,
        before,
    });
    Ok(word)
}

/// A ByVal `As Any` (or ByVal array) argument: a pointer value, an object's interface pointer,
/// or a scalar's payload.
fn byval_any<'a>(
    host: &StandardHostServices,
    value: Variant,
    after: &mut Vec<After<'a>>,
) -> HalResult<u64> {
    match value.vtype() {
        VarType::Object => {
            let Some(object) = value.as_object_ref() else {
                return Ok(0);
            };
            if object.is_nothing_placeholder() {
                return Ok(0);
            }
            if object.is_compat_object() {
                // A project object crosses as its IDispatch, kept alive for the call.
                let dispatch = calcvba_com::automation::project_dispatch(&object)
                    .map_err(|fault| vba_fault(host, fault.number, "project object"))?;
                let word = payload_word(&dispatch);
                after.push(After::Keep(Box::new(dispatch)));
                return Ok(word);
            }
            Ok(payload_word(&value))
        }
        VarType::ArrayVariant => Ok(value.raw_safearray_ptr().expose_provenance() as u64),
        _ => Ok(payload_word(&value)),
    }
}

/// A copy of a String array whose elements are ANSI BSTRs.
fn ansi_array(host: &StandardHostServices, value: Variant) -> HalResult<Variant> {
    let Ok(Some(array)) = value.try_as_safearray() else {
        return Ok(value);
    };
    let copy = Variant::from_safearray(array);
    let pointer = copy.raw_safearray_ptr();
    for index in 0..raw_len(pointer) {
        // SAFETY: element `index` of a live BSTR array copy this frame owns.
        unsafe {
            let slot = raw_element(pointer, index).cast::<*mut u16>();
            let unicode = *slot;
            let units = if unicode.is_null() {
                Vec::new()
            } else {
                let len = SysStringByteLen(unicode) as usize / 2;
                core::slice::from_raw_parts(unicode, len).to_vec()
            };
            let ansi = if unicode.is_null() {
                core::ptr::null_mut()
            } else {
                ansi_bstr(&units).map_err(|message| fault(host, message))?
            };
            // The copy's own element, freed as the runtime allocated it.
            drop(calcvba_runtime::bstr::BStr::from_raw_bstr(unicode));
            *slot = ansi;
        }
    }
    Ok(copy)
}

/// Converts an ANSI String array (what the callee left, possibly redimensioned) back to Unicode.
fn unicode_array(cell: &mut Variant) {
    let pointer = cell.raw_safearray_ptr();
    for index in 0..raw_len(pointer) {
        // SAFETY: element `index` of the live BSTR array the cell owns.
        unsafe {
            let slot = raw_element(pointer, index).cast::<*mut u16>();
            let ansi = *slot;
            let unicode = unicode_from_ansi(ansi);
            SysFreeString(ansi);
            *slot = if ansi.is_null() {
                core::ptr::null_mut()
            } else {
                unicode
                    .as_bstr()
                    .and_then(|text| text.clone_raw_bstr().ok())
                    .unwrap_or(core::ptr::null_mut())
            };
        }
    }
}

/// The element count of a raw SAFEARRAY (0 for null).
fn raw_len(pointer: *mut core::ffi::c_void) -> usize {
    if pointer.is_null() {
        return 0;
    }
    // SAFETY: a live descriptor: cDims at 0, rgsabound at 24 (x64), 8 bytes each.
    unsafe {
        let base = pointer.cast::<u8>();
        let dims = base.cast::<u16>().read_unaligned() as usize;
        (0..dims)
            .map(|d| base.add(24 + d * 8).cast::<u32>().read_unaligned() as usize)
            .product()
    }
}

/// The address of flat element `index` of a raw SAFEARRAY.
///
/// # Safety
/// `pointer` is a live descriptor and `index` is in bounds.
unsafe fn raw_element(pointer: *mut core::ffi::c_void, index: usize) -> *mut u8 {
    // SAFETY: cbElements at 4, pvData at 16 (x64).
    unsafe {
        let base = pointer.cast::<u8>();
        let size = base.add(4).cast::<u32>().read_unaligned() as usize;
        let data = base.add(16).cast::<*mut u8>().read_unaligned();
        data.add(index * size)
    }
}

/// Calls the `Declare` with `args` (design doc, step 9); returns the result.
pub(super) fn invoke(
    host: &StandardHostServices,
    descriptor: &DynLinkDescriptorView<'_>,
    args: &[ComArg<'_>],
) -> HalResult<Variant> {
    if args.len() != descriptor.param_count {
        return Err(vba_fault(host, 450, "Wrong number of arguments"));
    }
    let plan = plan(host, descriptor)?;
    let mut after: Vec<After<'_>> = Vec::with_capacity(args.len());
    let mut words: Vec<u64> = Vec::with_capacity(args.len());
    for (index, arg) in args.iter().enumerate() {
        let param_type = descriptor
            .param_types
            .get(index)
            .map_or("Long", String::as_str);
        let by_ref = descriptor.param_by_ref.get(index).copied().unwrap_or(false);
        if plan.library.is_none()
            && by_ref
            && let Some(word) = runtime_own_storage(arg.value)
        {
            words.push(word);
            continue;
        }
        words.push(marshal(host, param_type, by_ref, arg.value, &mut after)?);
    }
    let mut cells: Vec<*mut core::ffi::c_void> = words
        .iter_mut()
        .map(|word| (word as *mut u64).cast())
        .collect();
    // The result: up to a VARIANT, or a record's declared shape.
    let result_bytes = match &plan.result {
        ResultShape::Record(layout) => layout
            .declare_shape()
            .map_err(|message| fault(host, message))?
            .0
            .max(8),
        _ => calcvba_runtime::WIRE_SIZE,
    };
    let mut result = vec![0u64; result_bytes.div_ceil(8)];
    let mut record = match &plan.result {
        ResultShape::Record(layout) => Some(
            calcvba_runtime::pointer_helpers::AutomationRecord::empty_declared(layout.clone())
                .map_err(|message| fault(host, message))?,
        ),
        _ => None,
    };
    let result_pointer = match &mut record {
        Some(record) => record.as_mut_ptr().cast(),
        None => result.as_mut_ptr().cast(),
    };
    // SAFETY: the plan's CIF is the descriptor's signature; every slot holds a value of its
    // native type or the address of live storage (the caller's, or a temporary in `after`),
    // all of which outlive the call; the result storage fits the declared result.
    let last_error = unsafe { plan.call.invoke_cells(&mut cells, result_pointer) };
    host.last_dll_error.set(last_error as i32);

    let value = match &plan.result {
        ResultShape::Scalar(return_type) => {
            super::dynlink::unmarshal_ffi_to_variant(result[0] as i64, return_type.as_deref())
        }
        ResultShape::Variant => {
            let mut bytes = [0u8; calcvba_runtime::WIRE_SIZE];
            for (chunk, word) in bytes.chunks_mut(8).zip(&result) {
                chunk.copy_from_slice(&word.to_le_bytes()[..chunk.len()]);
            }
            // SAFETY: the callee's VARIANT result is ours to take.
            match unsafe { Variant::adopt_automation(bytes) } {
                Ok(value) => value,
                Err(_) => return Err(vba_fault(host, 458, "unsupported Variant result type")),
            }
        }
        ResultShape::Record(_) => record
            .as_ref()
            .expect("record result")
            .read()
            .map_err(|message| fault(host, message))?,
    };

    // Copy-backs, in argument order.
    let mut first_fault = None;
    // Storage passed twice (one variable for two parameters) is accounted once.
    let mut watched = std::collections::HashSet::new();
    for step in after {
        let outcome = match step {
            After::Ansi { cell, target } => {
                let bstr = *cell;
                // SAFETY: the temporary (possibly replaced by the callee) is ours.
                let unicode = unsafe { unicode_from_ansi(bstr) };
                // SAFETY: as above; freed once.
                unsafe { SysFreeString(bstr) };
                match target {
                    // A null temporary stays what it was.
                    Some(target) if !bstr.is_null() => write_back(host, target, unicode),
                    _ => Ok(()),
                }
            }
            After::Temp {
                cell,
                target,
                before,
            } => {
                calcvba_runtime::live_counters::account_foreign_rewrite(
                    before,
                    temp_payload(&cell),
                );
                match target {
                    Some(target) => write_back(host, target, (*cell).clone()),
                    None => Ok(()),
                }
            }
            After::Record { record, target } => match target {
                Some(target) => record
                    .read()
                    .map_err(|message| fault(host, message))
                    .and_then(|value| write_back(host, target, value)),
                None => Ok(()),
            },
            After::RecordArray { array, target } => match target {
                Some(target) => array
                    .read()
                    .map_err(|message| fault(host, message))
                    .and_then(|value| write_back(host, target, value)),
                None => Ok(()),
            },
            After::AnsiArray { mut cell, target } => {
                unicode_array(&mut cell);
                match target {
                    Some(target) => write_back(host, target, *cell),
                    None => Ok(()),
                }
            }
            After::Watch {
                storage,
                slot_vt,
                before,
            } => {
                if !watched.insert(storage as usize) {
                    continue;
                }
                // SAFETY: the caller's storage, still live after the call.
                let now =
                    unsafe { calcvba_runtime::live_counters::owned_payload(storage, slot_vt) };
                calcvba_runtime::live_counters::account_foreign_rewrite(before, now);
                Ok(())
            }
            After::Keep(cell) => {
                // Kept alive until the call returned; released now.
                drop(cell);
                Ok(())
            }
        };
        if let Err(error) = outcome {
            first_fault.get_or_insert(error);
        }
    }
    if let Some(error) = first_fault {
        return Err(error);
    }
    Ok(value)
}

impl std::fmt::Debug for DeclarePlan {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("DeclarePlan")
    }
}
