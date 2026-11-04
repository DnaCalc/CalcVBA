//! `Declare Lib` external procedures — the complete source-declared shape. The
//! runtime-resolved fields of `calcvba_bundle::ExternalCallDescriptor` (the
//! `DynLinkSymbol` handle, marshal lane, selection policy, writebacks) are
//! intentionally *not* here; they are not source-declared and the binder/loader
//! fills them. Everything declarable from source is complete here.

use calcvba_bundle::DeclareParamType;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CallConv {
    Stdcall,
    Cdecl,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DeclareSymbol {
    pub descriptor_id: u32,
    pub declared_name: String,
    pub library: String,
    /// The `Alias` export name, or the declared name when no alias is given.
    pub alias: String,
    /// True when `alias` is an ordinal (`Alias "#123"`).
    pub ordinal_alias: bool,
    /// Statically checked ordinal value, retaining the original spelling in `alias`.
    pub ordinal_value: Option<u16>,
    pub calling_convention: CallConv,
    /// `Declare Function` (has a return) vs `Declare Sub`.
    pub is_function: bool,
    pub param_names: Vec<String>,
    pub param_types: Vec<DeclareParamType>,
    pub param_by_ref: Vec<bool>,
    pub param_optional: Vec<bool>,
    pub param_param_array: bool,
    pub return_type: Option<DeclareParamType>,
    pub return_record_type: Option<String>,
}
