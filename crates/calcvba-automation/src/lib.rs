//! Automation value layouts shared by CalcVBA's runtime, its JIT and native lanes, and other
//! in-process libraries (CalcDb, DNA Calc) that exchange values with it without conversion.
//!
//! The platform's Automation ABI is the specification: a `VARIANT` is 24 bytes on 64-bit
//! targets and 16 bytes where pointers are 32-bit (wasm32 included), because its union holds
//! `BRECORD { pvRecord, pRecInfo }`. See `docs/design/runtime-variant-automation-layout-2026-11-13.md`.
//!
//! This crate holds the values themselves: the cell ([`Variant`]), `BSTR` ([`bstr`]),
//! `SAFEARRAY` ([`safe_array`]), `BRECORD` ([`com_record`], [`vba_record`]), interface
//! references ([`object_ref`]) and the Decimal and Currency payloads. On Windows they are
//! allocated by oleaut32; elsewhere this crate lays out the same bytes. `calcvba-runtime`
//! re-exports all of it and adds VBA's semantics.

pub mod abi;
pub mod automation_record;
pub mod bstr;
pub mod com;
pub mod com_record;
pub mod decimal;
pub mod live_counters;
pub mod object_ref;
pub mod oleaut;
pub mod safe_array;
pub mod value_types;
pub mod variant;
pub mod vba_record;

pub use bstr::Utf16String;
pub use com_record::{ComRecord, ComRecordRef};
pub use decimal::Decimal96;
#[cfg(any(test, feature = "live-counters"))]
pub use live_counters::current_thread_live_handle_counts;
pub use live_counters::{HandleBalance, LiveHandleCounts, live_handle_counts};
pub use object_ref::{ObjectRef, RawRuntimeIUnknown};
pub use value_types::{BindingHandle, CurrencyValue, DynLinkSymbol, F64Subtype, F64Value};
pub use variant::{VarType, Variant, VariantCore, VariantData, WIRE_SIZE};
pub use vba_record::{
    MAX_VBA_RECORD_FIXED_ARRAY_RANK, MAX_VBA_RECORD_LAYOUT_GRAPH_DEPTH,
    MAX_VBA_RECORD_LAYOUT_GRAPH_VISITS, MAX_VBA_RECORD_SIZE, VbaRecord, VbaRecordFieldHandle,
    VbaRecordFieldKind, VbaRecordFieldSpec, VbaRecordLayout,
};
