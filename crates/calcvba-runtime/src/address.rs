//! Real addresses of VBA storage, for `VarPtr`, `StrPtr` and `ObjPtr`.
//!
//! VBA hands out the address of the storage itself (Excel x64, `temp/variant-layout-probe`):
//! a Variant's cell, the payload of a typed variable, `pvData` + offset for an array element,
//! the field's offset inside a UDT. Writes through the address are seen by later reads. VM3
//! and the JIT both locate storage with [`locate`], starting from the engine cell that holds
//! the variable (a local, global, temporary or class field), so they give identical
//! addresses.

use std::sync::Arc;

use crate::safe_array::{
    SafeArray, SafeArrayBound, VT_BSTR_VALUE, VT_DISPATCH_VALUE, VT_UNKNOWN_VALUE,
    VT_VARIANT_VALUE, element_index,
};
use crate::vba_record::{VbaRecordFieldKind, VbaRecordLayout};
use crate::{VarType, Variant};

/// One step from a storage location into the storage it holds.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AddressStep {
    /// An element of the array held here, by declared subscripts (dimension 1 first).
    Index(Vec<i64>),
    /// Positional field of the UDT record held here.
    Field(usize),
}

/// What the pointer helper asks for.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AddressKind {
    /// `VarPtr` of a non-Variant: the storage of the value (a typed cell's payload).
    Value,
    /// `VarPtr` of a Variant: the `VARIANT` cell itself.
    Cell,
    /// `StrPtr`: the BSTR the storage holds.
    Bstr,
}

/// Storage inside a record, an array's payload or a cell's payload.
#[derive(Debug, Clone)]
pub enum Storage {
    /// A scalar, Decimal or other plain value.
    Plain,
    /// A BSTR pointer.
    Bstr,
    /// An interface pointer (null for `Nothing`).
    Object,
    /// A record of this layout, in place.
    Record(Arc<VbaRecordLayout>),
    /// An inline fixed-size array field.
    FixedArray {
        element: Box<VbaRecordFieldKind>,
        bounds: Vec<SafeArrayBound>,
    },
    /// A `SAFEARRAY*` (a dynamic-array field), null while unallocated.
    DynamicArray { element: Box<VbaRecordFieldKind> },
}

/// A located piece of storage.
#[derive(Debug, Clone)]
pub enum Location {
    /// An engine cell (`VARIANT`): a variable, a Variant element or a Variant field.
    Cell(*mut Variant),
    /// Raw storage of the given kind.
    Memory(*mut u8, Storage),
}

/// VBA's "Subscript out of range".
const SUBSCRIPT_OUT_OF_RANGE: i32 = 9;
/// VBA's "Type mismatch".
const TYPE_MISMATCH: i32 = 13;

/// Walk `path` from the engine cell `root`.
///
/// # Safety
/// `root` is a live engine cell, and nothing else accesses the storage it reaches while the
/// returned location is in use.
pub unsafe fn locate(root: *mut Variant, path: &[AddressStep]) -> Result<Location, i32> {
    let mut at = Location::Cell(root);
    for step in path {
        // SAFETY: `at` addresses live storage reached from `root`, per the caller.
        at = unsafe { step_into(at, step)? };
    }
    Ok(at)
}

/// # Safety
/// As [`locate`].
unsafe fn step_into(at: Location, step: &AddressStep) -> Result<Location, i32> {
    match (at, step) {
        (Location::Cell(cell), AddressStep::Index(subscripts)) => {
            // SAFETY: `cell` is a live engine cell.
            let cell = unsafe { &*cell };
            if cell.vtype() != VarType::ArrayVariant {
                return Err(TYPE_MISMATCH);
            }
            let raw = cell.raw_safearray_ptr();
            // SAFETY: an array cell owns its descriptor (or null); the view only reads it.
            let Some(array) = (unsafe { SafeArray::borrow_raw(raw, cell.record_array_layout()) })
            else {
                return Err(SUBSCRIPT_OUT_OF_RANGE);
            };
            array_element(&array, subscripts)
        }
        (Location::Cell(cell), AddressStep::Field(index)) => {
            // SAFETY: `cell` is a live engine cell.
            let Some((data, layout)) = (unsafe { (*cell).engine_record_memory() }) else {
                return Err(TYPE_MISMATCH);
            };
            record_field(data, &layout, *index)
        }
        (Location::Memory(data, Storage::Record(layout)), AddressStep::Field(index)) => {
            record_field(data, &layout, *index)
        }
        (
            Location::Memory(data, Storage::FixedArray { element, bounds }),
            AddressStep::Index(s),
        ) => {
            let flat = element_index(&bounds, s).ok_or(SUBSCRIPT_OUT_OF_RANGE)?;
            let (size, align) = element.storage_shape().map_err(|_| TYPE_MISMATCH)?;
            let stride = size.next_multiple_of(align);
            // SAFETY: `flat` is within the inline array's elements.
            Ok(field_location(unsafe { data.add(flat * stride) }, &element))
        }
        (Location::Memory(slot, Storage::DynamicArray { element }), AddressStep::Index(s)) => {
            // SAFETY: the slot holds the field's `SAFEARRAY*` (null while unallocated).
            let raw = unsafe { slot.cast::<*mut core::ffi::c_void>().read() };
            // SAFETY: a non-null field descriptor is live; the view only reads it.
            let Some(array) =
                (unsafe { SafeArray::borrow_raw(raw, element.array_record_layout()) })
            else {
                return Err(SUBSCRIPT_OUT_OF_RANGE);
            };
            array_element(&array, s)
        }
        _ => Err(TYPE_MISMATCH),
    }
}

fn array_element(array: &SafeArray, subscripts: &[i64]) -> Result<Location, i32> {
    let bounds = array.bounds().ok_or(SUBSCRIPT_OUT_OF_RANGE)?;
    if bounds.len() != subscripts.len() {
        return Err(SUBSCRIPT_OUT_OF_RANGE);
    }
    let flat = element_index(&bounds, subscripts).ok_or(SUBSCRIPT_OUT_OF_RANGE)?;
    let element = array.element_address(flat).ok_or(SUBSCRIPT_OUT_OF_RANGE)?;
    Ok(match array.vba_record_layout() {
        Some(layout) => Location::Memory(element, Storage::Record(layout)),
        None => match array.element_vartype() {
            VT_VARIANT_VALUE => Location::Cell(element.cast()),
            VT_BSTR_VALUE => Location::Memory(element, Storage::Bstr),
            VT_DISPATCH_VALUE | VT_UNKNOWN_VALUE => Location::Memory(element, Storage::Object),
            _ => Location::Memory(element, Storage::Plain),
        },
    })
}

fn record_field(data: *mut u8, layout: &VbaRecordLayout, index: usize) -> Result<Location, i32> {
    let field = layout.fields().get(index).ok_or(TYPE_MISMATCH)?;
    // SAFETY: the field offset lies inside the record's memory.
    Ok(field_location(
        unsafe { data.add(field.offset()) },
        field.kind(),
    ))
}

fn field_location(ptr: *mut u8, kind: &VbaRecordFieldKind) -> Location {
    match kind {
        VbaRecordFieldKind::Variant => Location::Cell(ptr.cast()),
        VbaRecordFieldKind::String => Location::Memory(ptr, Storage::Bstr),
        VbaRecordFieldKind::Object => Location::Memory(ptr, Storage::Object),
        VbaRecordFieldKind::Record(layout) => {
            Location::Memory(ptr, Storage::Record(layout.clone()))
        }
        VbaRecordFieldKind::FixedArray { element, bounds } => Location::Memory(
            ptr,
            Storage::FixedArray {
                element: element.clone(),
                bounds: bounds.clone(),
            },
        ),
        VbaRecordFieldKind::DynamicArray { element } => Location::Memory(
            ptr,
            Storage::DynamicArray {
                element: element.clone(),
            },
        ),
        _ => Location::Memory(ptr, Storage::Plain),
    }
}

impl Location {
    /// The address (or BSTR) the pointer helper returns, or `None` for `StrPtr` of storage
    /// that holds no String (VBA converts such an operand to a temporary String first).
    ///
    /// # Safety
    /// The location is live.
    pub unsafe fn address(&self, kind: AddressKind) -> Option<usize> {
        match (self, kind) {
            (Self::Cell(cell), AddressKind::Cell) => Some(cell.expose_provenance()),
            (Self::Cell(cell), AddressKind::Value) => {
                // SAFETY: a live engine cell.
                let cell_ref = unsafe { &**cell };
                // A UDT variable's address is its record memory.
                if let Some((data, _)) = cell_ref.engine_record_memory() {
                    return Some(data.expose_provenance());
                }
                Some(cell_ref.payload_address())
            }
            (Self::Cell(cell), AddressKind::Bstr) => {
                // SAFETY: a live engine cell.
                let cell_ref = unsafe { &**cell };
                (cell_ref.vtype() == VarType::String)
                    .then(|| cell_ref.raw_bstr_ptr().expose_provenance())
            }
            (Self::Memory(ptr, Storage::Bstr), AddressKind::Bstr) => {
                // SAFETY: BSTR storage holds a pointer.
                Some(unsafe { ptr.cast::<*mut u16>().read() }.expose_provenance())
            }
            (Self::Memory(_, _), AddressKind::Bstr) => None,
            (Self::Memory(ptr, _), _) => Some(ptr.expose_provenance()),
        }
    }

    /// The interface pointer `ObjPtr` returns for raw object storage (`None` for a cell,
    /// whose object the caller resolves through its host).
    ///
    /// # Safety
    /// The location is live.
    pub unsafe fn object_pointer(&self) -> Option<usize> {
        match self {
            // SAFETY: object storage holds an interface pointer or null.
            Self::Memory(ptr, Storage::Object) => {
                Some(unsafe { ptr.cast::<*mut u8>().read() }.expose_provenance())
            }
            _ => None,
        }
    }
}
