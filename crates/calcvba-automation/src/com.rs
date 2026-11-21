//! COM interface layouts with the Windows ABI, usable on every platform.
//!
//! Off Windows these are CalcVBA's portable stand-ins: the same vtable shapes and calling
//! convention (`extern "system"` is the C convention there), so a CalcVBA object can sit in a
//! `VARIANT`'s `punkVal`/`pdispVal` and a CalcVBA record type in `pRecInfo` on every platform.
//! Interface identifiers travel as `REFIID` (`*const Guid`), never by value.

use core::ffi::c_void;

/// `GUID` / `IID` / `CLSID`.
#[repr(C)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct Guid {
    pub data1: u32,
    pub data2: u16,
    pub data3: u16,
    pub data4: [u8; 8],
}

impl Guid {
    pub const fn new(data1: u32, data2: u16, data3: u16, data4: [u8; 8]) -> Self {
        Self {
            data1,
            data2,
            data3,
            data4,
        }
    }

    /// Canonical 16-byte COM GUID memory representation (little-endian `data1`..`data3`,
    /// followed by `data4`).
    pub const fn to_le_bytes(self) -> [u8; 16] {
        let d1 = self.data1.to_le_bytes();
        let d2 = self.data2.to_le_bytes();
        let d3 = self.data3.to_le_bytes();
        let d4 = self.data4;
        [
            d1[0], d1[1], d1[2], d1[3], d2[0], d2[1], d3[0], d3[1], d4[0], d4[1], d4[2], d4[3],
            d4[4], d4[5], d4[6], d4[7],
        ]
    }

    /// Reconstruct a [`Guid`] from its 16-byte COM memory layout.
    pub const fn from_le_bytes(bytes: [u8; 16]) -> Self {
        Self {
            data1: u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]),
            data2: u16::from_le_bytes([bytes[4], bytes[5]]),
            data3: u16::from_le_bytes([bytes[6], bytes[7]]),
            data4: [
                bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14],
                bytes[15],
            ],
        }
    }

    /// Deterministic UUID-shaped COM identity for a project-defined type (matches
    /// `calcvba_build::deterministic_uuid`).
    pub fn deterministic(namespace: &str, name: &str) -> Self {
        const NAMESPACE_SEED: u128 = 0x4f58_5642_4100_0000_0000_0000_0000_0001;

        let mut hash: u128 = NAMESPACE_SEED;
        for byte in namespace
            .bytes()
            .chain(core::iter::once(0u8))
            .chain(name.bytes())
        {
            hash ^= u128::from(byte);
            hash = hash.wrapping_mul(0x0100_0000_01b3_0000_0000_0000_0000_001d);
        }

        let mut bytes = hash.to_be_bytes();
        bytes[6] = (bytes[6] & 0x0f) | 0x50;
        bytes[8] = (bytes[8] & 0x3f) | 0x80;

        Self {
            data1: u32::from_be_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]),
            data2: u16::from_be_bytes([bytes[4], bytes[5]]),
            data3: u16::from_be_bytes([bytes[6], bytes[7]]),
            data4: [
                bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14],
                bytes[15],
            ],
        }
    }

    /// Parse `XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX`, with or without braces.
    pub fn parse(input: &str) -> Option<Self> {
        fn hex<T: TryFrom<u32>>(text: &str, len: usize) -> Option<T> {
            if text.len() != len {
                return None;
            }
            u32::from_str_radix(text, 16).ok()?.try_into().ok()
        }
        let trimmed = input.trim();
        let body = trimmed
            .strip_prefix('{')
            .and_then(|inner| inner.strip_suffix('}'))
            .unwrap_or(trimmed);
        let mut parts = body.split('-');
        let data1 = hex::<u32>(parts.next()?, 8)?;
        let data2 = hex::<u16>(parts.next()?, 4)?;
        let data3 = hex::<u16>(parts.next()?, 4)?;
        let data4_hi = parts.next()?;
        let data4_lo = parts.next()?;
        if parts.next().is_some() || data4_hi.len() != 4 || data4_lo.len() != 12 {
            return None;
        }
        let mut data4 = [0u8; 8];
        data4[0] = hex::<u8>(data4_hi.get(0..2)?, 2)?;
        data4[1] = hex::<u8>(data4_hi.get(2..4)?, 2)?;
        for (index, byte) in data4[2..].iter_mut().enumerate() {
            *byte = hex::<u8>(data4_lo.get(index * 2..index * 2 + 2)?, 2)?;
        }
        Some(Self::new(data1, data2, data3, data4))
    }
}

/// Upper-case canonical form without braces.
impl core::fmt::Display for Guid {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        let d = &self.data4;
        write!(
            f,
            "{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}",
            self.data1, self.data2, self.data3, d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7]
        )
    }
}

/// An interface pointer: its first word is the vtable.
#[repr(C)]
pub struct IUnknown {
    pub vtbl: *const IUnknownVtbl,
}

/// `HRESULT`.
pub type HResult = i32;

pub const S_OK: HResult = 0;
pub const S_FALSE: HResult = 1;
pub const E_NOTIMPL: HResult = 0x8000_4001u32 as i32;
pub const E_NOINTERFACE: HResult = 0x8000_4002u32 as i32;
pub const E_POINTER: HResult = 0x8000_4003u32 as i32;
pub const E_FAIL: HResult = 0x8000_4005u32 as i32;
pub const E_INVALIDARG: HResult = 0x8007_0057u32 as i32;
pub const RPC_E_WRONG_THREAD: HResult = 0x8001_010Eu32 as i32;

pub const CONNECT_E_NOCONNECTION: HResult = 0x8004_0200u32 as i32;
pub const CONNECT_E_ADVISELIMIT: HResult = 0x8004_0201u32 as i32;
pub const CONNECT_E_CANNOTCONNECT: HResult = 0x8004_0202u32 as i32;

pub const DISP_E_UNKNOWNINTERFACE: HResult = 0x8002_0001u32 as i32;
pub const DISP_E_MEMBERNOTFOUND: HResult = 0x8002_0003u32 as i32;
pub const DISP_E_PARAMNOTFOUND: HResult = 0x8002_0004u32 as i32;
pub const DISP_E_TYPEMISMATCH: HResult = 0x8002_0005u32 as i32;
pub const DISP_E_UNKNOWNNAME: HResult = 0x8002_0006u32 as i32;
pub const DISP_E_BADVARTYPE: HResult = 0x8002_0008u32 as i32;
pub const DISP_E_EXCEPTION: HResult = 0x8002_0009u32 as i32;
pub const DISP_E_OVERFLOW: HResult = 0x8002_000Au32 as i32;
pub const DISP_E_BADINDEX: HResult = 0x8002_000Bu32 as i32;
pub const DISP_E_BADPARAMCOUNT: HResult = 0x8002_000Eu32 as i32;
pub const DISP_E_PARAMNOTOPTIONAL: HResult = 0x8002_000Fu32 as i32;
pub const TYPE_E_ELEMENTNOTFOUND: HResult = 0x8002_802Bu32 as i32;
pub const TYPE_E_BADMODULEKIND: HResult = 0x8002_88BDu32 as i32;

pub const DISPID_UNKNOWN: i32 = -1;
pub const DISPID_VALUE: i32 = 0;
pub const DISPID_PROPERTYPUT: i32 = -3;
pub const DISPID_NEWENUM: i32 = -4;
pub const MEMBERID_NIL: i32 = -1;

pub const DISPATCH_METHOD: u16 = 0x1;
pub const DISPATCH_PROPERTYGET: u16 = 0x2;
pub const DISPATCH_PROPERTYPUT: u16 = 0x4;
pub const DISPATCH_PROPERTYPUTREF: u16 = 0x8;

pub const TKIND_INTERFACE: i32 = 3;
pub const TKIND_DISPATCH: i32 = 4;

pub const FUNC_PUREVIRTUAL: i32 = 1;
pub const FUNC_DISPATCH: i32 = 4;

pub const INVOKE_FUNC: i32 = 1;
pub const INVOKE_PROPERTYGET: i32 = 2;
pub const INVOKE_PROPERTYPUT: i32 = 4;
pub const INVOKE_PROPERTYPUTREF: i32 = 8;

pub const CC_STDCALL: i32 = 4;

pub const TYPEFLAG_FCANCREATE: u16 = 0x0002;
pub const TYPEFLAG_FDUAL: u16 = 0x0040;
pub const TYPEFLAG_FNONEXTENSIBLE: u16 = 0x0080;
pub const TYPEFLAG_FOLEAUTOMATION: u16 = 0x0100;
/// Exact `TYPEATTR.wTypeFlags` Excel VBA reports for a class module's `ITypeInfo` (`0x1C2`).
pub const VBA_CLASS_TYPEFLAGS: u16 =
    TYPEFLAG_FCANCREATE | TYPEFLAG_FDUAL | TYPEFLAG_FNONEXTENSIBLE | TYPEFLAG_FOLEAUTOMATION;

pub const FUNCFLAG_FRESTRICTED: u16 = 0x0001;

pub const PARAMFLAG_NONE: u16 = 0x0000;
pub const PARAMFLAG_FIN: u16 = 0x0001;
pub const PARAMFLAG_FOUT: u16 = 0x0002;
pub const PARAMFLAG_FRETVAL: u16 = 0x0008;
pub const PARAMFLAG_FOPT: u16 = 0x0010;

pub const IDLFLAG_NONE: u16 = 0x0000;
pub const IDLFLAG_FIN: u16 = 0x0001;
pub const IDLFLAG_FOUT: u16 = 0x0002;
pub const IDLFLAG_FRETVAL: u16 = 0x0008;

pub const VT_VOID: u16 = 24;
pub const VT_HRESULT: u16 = 25;
pub const VT_PTR: u16 = 26;
pub const VT_SAFEARRAY: u16 = 27;

/// `IID_NULL` {00000000-0000-0000-0000-000000000000}.
pub const IID_NULL: Guid = Guid::new(0, 0, 0, [0; 8]);
/// `IID_IUnknown` {00000000-0000-0000-C000-000000000046}.
pub const IID_IUNKNOWN: Guid = Guid::new(0, 0, 0, [0xC0, 0, 0, 0, 0, 0, 0, 0x46]);
/// `IID_IDispatch` {00020400-0000-0000-C000-000000000046}.
pub const IID_IDISPATCH: Guid = Guid::new(0x0002_0400, 0, 0, [0xC0, 0, 0, 0, 0, 0, 0, 0x46]);
/// `IID_ITypeInfo` {00020401-0000-0000-C000-000000000046}.
pub const IID_ITYPEINFO: Guid = Guid::new(0x0002_0401, 0, 0, [0xC0, 0, 0, 0, 0, 0, 0, 0x46]);
/// `IID_IRecordInfo` {0000002F-0000-0000-C000-000000000046}.
pub const IID_IRECORDINFO: Guid = Guid::new(0x0000_002F, 0, 0, [0xC0, 0, 0, 0, 0, 0, 0, 0x46]);
/// `IID_IErrorInfo` {1CF2B120-547D-101B-8E65-08002B2BD119}.
pub const IID_IERRORINFO: Guid = Guid::new(
    0x1CF2_B120,
    0x547D,
    0x101B,
    [0x8E, 0x65, 0x08, 0x00, 0x2B, 0x2B, 0xD1, 0x19],
);
/// `IID_ISupportErrorInfo` {DF0B3D60-548F-101B-8E65-08002B2BD119}.
pub const IID_ISUPPORTERRORINFO: Guid = Guid::new(
    0xDF0B_3D60,
    0x548F,
    0x101B,
    [0x8E, 0x65, 0x08, 0x00, 0x2B, 0x2B, 0xD1, 0x19],
);
/// `IID_IConnectionPointContainer` {B196B284-BAB4-101A-B69C-00AA00341D07}.
pub const IID_ICONNECTIONPOINTCONTAINER: Guid = Guid::new(
    0xB196_B284,
    0xBAB4,
    0x101A,
    [0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07],
);
/// `IID_IEnumConnectionPoints` {B196B285-BAB4-101A-B69C-00AA00341D07}.
pub const IID_IENUMCONNECTIONPOINTS: Guid = Guid::new(
    0xB196_B285,
    0xBAB4,
    0x101A,
    [0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07],
);
/// `IID_IConnectionPoint` {B196B286-BAB4-101A-B69C-00AA00341D07}.
pub const IID_ICONNECTIONPOINT: Guid = Guid::new(
    0xB196_B286,
    0xBAB4,
    0x101A,
    [0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07],
);
/// `IID_IEnumConnections` {B196B287-BAB4-101A-B69C-00AA00341D07}.
pub const IID_IENUMCONNECTIONS: Guid = Guid::new(
    0xB196_B287,
    0xBAB4,
    0x101A,
    [0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07],
);
/// `IID__Collection` (`VBA.Collection` default interface) {A4C46780-499F-101B-BB78-00AA00383CBB}.
pub const IID_VBA_COLLECTION: Guid = Guid::new(
    0xA4C4_6780,
    0x499F,
    0x101B,
    [0xBB, 0x78, 0x00, 0xAA, 0x00, 0x38, 0x3C, 0xBB],
);

/// `IUnknown`'s vtable.
#[repr(C)]
pub struct IUnknownVtbl {
    pub query_interface: unsafe extern "system" fn(
        this: *mut c_void,
        riid: *const Guid,
        object: *mut *mut c_void,
    ) -> HResult,
    pub add_ref: unsafe extern "system" fn(this: *mut c_void) -> u32,
    pub release: unsafe extern "system" fn(this: *mut c_void) -> u32,
}

/// `IDispatch`'s vtable. Names and type info pointers are left untyped here; the marshalling
/// layer owns their meaning.
#[repr(C)]
pub struct IDispatchVtbl {
    pub unknown: IUnknownVtbl,
    pub get_type_info_count:
        unsafe extern "system" fn(this: *mut c_void, count: *mut u32) -> HResult,
    pub get_type_info: unsafe extern "system" fn(
        this: *mut c_void,
        index: u32,
        lcid: u32,
        type_info: *mut *mut c_void,
    ) -> HResult,
    pub get_ids_of_names: unsafe extern "system" fn(
        this: *mut c_void,
        riid: *const Guid,
        names: *const *const u16,
        count: u32,
        lcid: u32,
        dispids: *mut i32,
    ) -> HResult,
    #[allow(clippy::type_complexity)]
    pub invoke: unsafe extern "system" fn(
        this: *mut c_void,
        dispid: i32,
        riid: *const Guid,
        lcid: u32,
        flags: u16,
        params: *mut c_void,
        result: *mut c_void,
        exception: *mut c_void,
        arg_error: *mut u32,
    ) -> HResult,
}

/// Portable `DISPPARAMS` with the Windows Automation layout.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct DispParams {
    pub rgvarg: *mut crate::VariantCore,
    pub rgdispid_named_args: *mut i32,
    pub c_args: u32,
    pub c_named_args: u32,
}

impl Default for DispParams {
    fn default() -> Self {
        Self {
            rgvarg: core::ptr::null_mut(),
            rgdispid_named_args: core::ptr::null_mut(),
            c_args: 0,
            c_named_args: 0,
        }
    }
}

/// Portable `EXCEPINFO` with the Windows Automation layout.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct ExcepInfo {
    pub w_code: u16,
    pub w_reserved: u16,
    pub bstr_source: *mut u16,
    pub bstr_description: *mut u16,
    pub bstr_help_file: *mut u16,
    pub dw_help_context: u32,
    pub pv_reserved: *mut c_void,
    pub pfn_deferred_fill_in: Option<unsafe extern "system" fn(*mut ExcepInfo) -> HResult>,
    pub scode: HResult,
}

impl Default for ExcepInfo {
    fn default() -> Self {
        Self {
            w_code: 0,
            w_reserved: 0,
            bstr_source: core::ptr::null_mut(),
            bstr_description: core::ptr::null_mut(),
            bstr_help_file: core::ptr::null_mut(),
            dw_help_context: 0,
            pv_reserved: core::ptr::null_mut(),
            pfn_deferred_fill_in: None,
            scode: 0,
        }
    }
}

/// Portable `TYPEDESC` payload union with the Windows Automation layout.
#[repr(C)]
#[derive(Clone, Copy)]
pub union TypeDescPayload {
    pub lptdesc: *mut TypeDesc,
    pub lpadesc: *mut c_void,
    pub hreftype: u32,
}

/// Portable `TYPEDESC` with the Windows Automation layout.
#[repr(C)]
#[derive(Clone, Copy)]
pub struct TypeDesc {
    pub payload: TypeDescPayload,
    pub vt: u16,
}

impl Default for TypeDesc {
    fn default() -> Self {
        Self {
            payload: TypeDescPayload {
                lptdesc: core::ptr::null_mut(),
            },
            vt: 0,
        }
    }
}

/// Portable `IDLDESC` with the Windows Automation layout.
#[repr(C)]
#[derive(Clone, Copy, Debug, Default)]
pub struct IdlDesc {
    pub dw_reserved: usize,
    pub w_idl_flags: u16,
}

/// Portable `PARAMDESCEX` with the Windows Automation layout.
#[repr(C)]
#[derive(Clone, Copy)]
pub struct ParamDescEx {
    pub c_bytes: u32,
    pub var_default_value: crate::VariantCore,
}

/// Portable `PARAMDESC` with the Windows Automation layout.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct ParamDesc {
    pub pparamdescex: *mut ParamDescEx,
    pub w_param_flags: u16,
}

impl Default for ParamDesc {
    fn default() -> Self {
        Self {
            pparamdescex: core::ptr::null_mut(),
            w_param_flags: 0,
        }
    }
}

/// Portable `ELEMDESC` payload union with the Windows Automation layout.
#[repr(C)]
#[derive(Clone, Copy)]
pub union ElemDescPayload {
    pub idldesc: IdlDesc,
    pub paramdesc: ParamDesc,
}

/// Portable `ELEMDESC` with the Windows Automation layout.
#[repr(C)]
#[derive(Clone, Copy)]
pub struct ElemDesc {
    pub tdesc: TypeDesc,
    pub payload: ElemDescPayload,
}

impl Default for ElemDesc {
    fn default() -> Self {
        Self {
            tdesc: TypeDesc::default(),
            payload: ElemDescPayload {
                paramdesc: ParamDesc::default(),
            },
        }
    }
}

/// Portable `TYPEATTR` with the Windows Automation layout.
#[repr(C)]
#[derive(Clone, Copy)]
pub struct TypeAttr {
    pub guid: Guid,
    pub lcid: u32,
    pub dw_reserved: u32,
    pub memid_constructor: i32,
    pub memid_destructor: i32,
    pub lpstr_schema: *mut u16,
    pub cb_size_instance: u32,
    pub typekind: i32,
    pub c_funcs: u16,
    pub c_vars: u16,
    pub c_impl_types: u16,
    pub cb_size_vft: u16,
    pub cb_alignment: u16,
    pub w_type_flags: u16,
    pub w_major_ver_num: u16,
    pub w_minor_ver_num: u16,
    pub tdesc_alias: TypeDesc,
    pub idldesc_type: IdlDesc,
}

/// Portable `FUNCDESC` with the Windows Automation layout.
#[repr(C)]
#[derive(Clone, Copy)]
pub struct FuncDesc {
    pub memid: i32,
    pub lprgscode: *mut i32,
    pub lprgelemdesc_param: *mut ElemDesc,
    pub funckind: i32,
    pub invkind: i32,
    pub callconv: i32,
    pub c_params: i16,
    pub c_params_opt: i16,
    pub o_vft: i16,
    pub c_scodes: i16,
    pub elemdesc_func: ElemDesc,
    pub w_func_flags: u16,
}

/// Portable `ITypeInfo` interface pointer.
#[repr(C)]
pub struct ITypeInfo {
    pub vtbl: *const ITypeInfoVtbl,
}

/// `ITypeInfo`'s vtable in the Windows `oaidl.h` slot order (3 `IUnknown` + 19 `ITypeInfo` slots).
#[repr(C)]
pub struct ITypeInfoVtbl {
    pub unknown: IUnknownVtbl,
    pub get_type_attr:
        unsafe extern "system" fn(this: *mut c_void, type_attr: *mut *mut TypeAttr) -> HResult,
    pub get_type_comp:
        unsafe extern "system" fn(this: *mut c_void, type_comp: *mut *mut c_void) -> HResult,
    pub get_func_desc: unsafe extern "system" fn(
        this: *mut c_void,
        index: u32,
        func_desc: *mut *mut FuncDesc,
    ) -> HResult,
    pub get_var_desc: unsafe extern "system" fn(
        this: *mut c_void,
        index: u32,
        var_desc: *mut *mut c_void,
    ) -> HResult,
    pub get_names: unsafe extern "system" fn(
        this: *mut c_void,
        memid: i32,
        names: *mut *mut u16,
        max_names: u32,
        count: *mut u32,
    ) -> HResult,
    pub get_ref_type_of_impl_type:
        unsafe extern "system" fn(this: *mut c_void, index: u32, ref_type: *mut u32) -> HResult,
    pub get_impl_type_flags: unsafe extern "system" fn(
        this: *mut c_void,
        index: u32,
        impl_type_flags: *mut i32,
    ) -> HResult,
    pub get_ids_of_names: unsafe extern "system" fn(
        this: *mut c_void,
        names: *const *const u16,
        count: u32,
        memids: *mut i32,
    ) -> HResult,
    #[allow(clippy::type_complexity)]
    pub invoke: unsafe extern "system" fn(
        this: *mut c_void,
        instance: *mut c_void,
        memid: i32,
        flags: u16,
        params: *mut DispParams,
        result: *mut c_void,
        excep_info: *mut ExcepInfo,
        arg_err: *mut u32,
    ) -> HResult,
    pub get_documentation: unsafe extern "system" fn(
        this: *mut c_void,
        memid: i32,
        name: *mut *mut u16,
        doc_string: *mut *mut u16,
        help_context: *mut u32,
        help_file: *mut *mut u16,
    ) -> HResult,
    pub get_dll_entry: unsafe extern "system" fn(
        this: *mut c_void,
        memid: i32,
        invkind: i32,
        dll_name: *mut *mut u16,
        name: *mut *mut u16,
        ordinal: *mut u16,
    ) -> HResult,
    pub get_ref_type_info: unsafe extern "system" fn(
        this: *mut c_void,
        ref_type: u32,
        type_info: *mut *mut c_void,
    ) -> HResult,
    pub address_of_member: unsafe extern "system" fn(
        this: *mut c_void,
        memid: i32,
        invkind: i32,
        address: *mut *mut c_void,
    ) -> HResult,
    pub create_instance: unsafe extern "system" fn(
        this: *mut c_void,
        outer: *mut c_void,
        riid: *const Guid,
        object: *mut *mut c_void,
    ) -> HResult,
    pub get_mops:
        unsafe extern "system" fn(this: *mut c_void, memid: i32, mops: *mut *mut u16) -> HResult,
    pub get_containing_type_lib: unsafe extern "system" fn(
        this: *mut c_void,
        type_lib: *mut *mut c_void,
        index: *mut u32,
    ) -> HResult,
    pub release_type_attr: unsafe extern "system" fn(this: *mut c_void, type_attr: *mut TypeAttr),
    pub release_func_desc: unsafe extern "system" fn(this: *mut c_void, func_desc: *mut FuncDesc),
    pub release_var_desc: unsafe extern "system" fn(this: *mut c_void, var_desc: *mut c_void),
}

/// `IRecordInfo`'s vtable, in the header's slot order. Record data pointers are untyped; field
/// values are `VARIANT*` (`*mut c_void` here, typed by the value crate that owns `Variant`).
#[repr(C)]
pub struct IRecordInfoVtbl {
    pub unknown: IUnknownVtbl,
    pub record_init: unsafe extern "system" fn(this: *mut c_void, record: *mut c_void) -> HResult,
    pub record_clear: unsafe extern "system" fn(this: *mut c_void, record: *mut c_void) -> HResult,
    pub record_copy: unsafe extern "system" fn(
        this: *mut c_void,
        existing: *mut c_void,
        new: *mut c_void,
    ) -> HResult,
    pub get_guid: unsafe extern "system" fn(this: *mut c_void, guid: *mut Guid) -> HResult,
    pub get_name: unsafe extern "system" fn(this: *mut c_void, name: *mut *mut u16) -> HResult,
    pub get_size: unsafe extern "system" fn(this: *mut c_void, size: *mut u32) -> HResult,
    pub get_type_info:
        unsafe extern "system" fn(this: *mut c_void, type_info: *mut *mut c_void) -> HResult,
    pub get_field: unsafe extern "system" fn(
        this: *mut c_void,
        data: *mut c_void,
        name: *const u16,
        field: *mut c_void,
    ) -> HResult,
    pub get_field_no_copy: unsafe extern "system" fn(
        this: *mut c_void,
        data: *mut c_void,
        name: *const u16,
        field: *mut c_void,
        data_carray: *mut *mut c_void,
    ) -> HResult,
    pub put_field: unsafe extern "system" fn(
        this: *mut c_void,
        flags: u32,
        data: *mut c_void,
        name: *const u16,
        field: *mut c_void,
    ) -> HResult,
    pub put_field_no_copy: unsafe extern "system" fn(
        this: *mut c_void,
        flags: u32,
        data: *mut c_void,
        name: *const u16,
        field: *mut c_void,
    ) -> HResult,
    pub get_field_names: unsafe extern "system" fn(
        this: *mut c_void,
        count: *mut u32,
        names: *mut *mut u16,
    ) -> HResult,
    pub is_matching_type: unsafe extern "system" fn(this: *mut c_void, other: *mut c_void) -> i32,
    pub record_create: unsafe extern "system" fn(this: *mut c_void) -> *mut c_void,
    pub record_create_copy: unsafe extern "system" fn(
        this: *mut c_void,
        source: *mut c_void,
        dest: *mut *mut c_void,
    ) -> HResult,
    pub record_destroy:
        unsafe extern "system" fn(this: *mut c_void, record: *mut c_void) -> HResult,
}

/// Portable `IErrorInfo` interface pointer.
#[repr(C)]
pub struct IErrorInfo {
    pub vtbl: *const IErrorInfoVtbl,
}

/// `IErrorInfo`'s vtable in Windows `oaidl.h` slot order (3 `IUnknown` + 5 slots).
#[repr(C)]
pub struct IErrorInfoVtbl {
    pub unknown: IUnknownVtbl,
    pub get_guid: unsafe extern "system" fn(this: *mut c_void, guid: *mut Guid) -> HResult,
    pub get_source: unsafe extern "system" fn(this: *mut c_void, source: *mut *mut u16) -> HResult,
    pub get_description:
        unsafe extern "system" fn(this: *mut c_void, description: *mut *mut u16) -> HResult,
    pub get_help_file:
        unsafe extern "system" fn(this: *mut c_void, help_file: *mut *mut u16) -> HResult,
    pub get_help_context:
        unsafe extern "system" fn(this: *mut c_void, help_context: *mut u32) -> HResult,
}

/// Portable `ISupportErrorInfo` interface pointer.
#[repr(C)]
pub struct ISupportErrorInfo {
    pub vtbl: *const ISupportErrorInfoVtbl,
}

/// `ISupportErrorInfo`'s vtable in Windows `oaidl.h` slot order (3 `IUnknown` + 1 slot).
#[repr(C)]
pub struct ISupportErrorInfoVtbl {
    pub unknown: IUnknownVtbl,
    pub interface_supports_error_info:
        unsafe extern "system" fn(this: *mut c_void, riid: *const Guid) -> HResult,
}

/// Portable `IConnectionPointContainer` interface pointer.
#[repr(C)]
pub struct IConnectionPointContainer {
    pub vtbl: *const IConnectionPointContainerVtbl,
}

/// `IConnectionPointContainer`'s vtable in Windows `ocidl.h` slot order (3 `IUnknown` + 2 slots).
#[repr(C)]
pub struct IConnectionPointContainerVtbl {
    pub unknown: IUnknownVtbl,
    pub enum_connection_points:
        unsafe extern "system" fn(this: *mut c_void, pp_enum: *mut *mut c_void) -> HResult,
    pub find_connection_point: unsafe extern "system" fn(
        this: *mut c_void,
        riid: *const Guid,
        pp_cp: *mut *mut c_void,
    ) -> HResult,
}

/// Portable `IConnectionPoint` interface pointer.
#[repr(C)]
pub struct IConnectionPoint {
    pub vtbl: *const IConnectionPointVtbl,
}

/// `IConnectionPoint`'s vtable in Windows `ocidl.h` slot order (3 `IUnknown` + 5 slots).
#[repr(C)]
pub struct IConnectionPointVtbl {
    pub unknown: IUnknownVtbl,
    pub get_connection_interface:
        unsafe extern "system" fn(this: *mut c_void, p_iid: *mut Guid) -> HResult,
    pub get_connection_point_container:
        unsafe extern "system" fn(this: *mut c_void, pp_cpc: *mut *mut c_void) -> HResult,
    pub advise: unsafe extern "system" fn(
        this: *mut c_void,
        punk_sink: *mut c_void,
        pdw_cookie: *mut u32,
    ) -> HResult,
    pub unadvise: unsafe extern "system" fn(this: *mut c_void, dw_cookie: u32) -> HResult,
    pub enum_connections:
        unsafe extern "system" fn(this: *mut c_void, pp_enum: *mut *mut c_void) -> HResult,
}

/// Portable `IEnumConnectionPoints` interface pointer.
#[repr(C)]
pub struct IEnumConnectionPoints {
    pub vtbl: *const IEnumConnectionPointsVtbl,
}

/// `IEnumConnectionPoints`'s vtable in Windows `ocidl.h` slot order (3 `IUnknown` + 4 slots).
#[repr(C)]
pub struct IEnumConnectionPointsVtbl {
    pub unknown: IUnknownVtbl,
    pub next: unsafe extern "system" fn(
        this: *mut c_void,
        c_connections: u32,
        pp_cp: *mut *mut c_void,
        pc_fetched: *mut u32,
    ) -> HResult,
    pub skip: unsafe extern "system" fn(this: *mut c_void, c_connections: u32) -> HResult,
    pub reset: unsafe extern "system" fn(this: *mut c_void) -> HResult,
    pub clone: unsafe extern "system" fn(this: *mut c_void, pp_enum: *mut *mut c_void) -> HResult,
}

/// Portable `CONNECTDATA` with the Windows `ocidl.h` layout.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct ConnectData {
    pub punk: *mut c_void,
    pub dw_cookie: u32,
}

impl Default for ConnectData {
    fn default() -> Self {
        Self {
            punk: core::ptr::null_mut(),
            dw_cookie: 0,
        }
    }
}

/// Portable `IEnumConnections` interface pointer.
#[repr(C)]
pub struct IEnumConnections {
    pub vtbl: *const IEnumConnectionsVtbl,
}

/// `IEnumConnections`'s vtable in Windows `ocidl.h` slot order (3 `IUnknown` + 4 slots).
#[repr(C)]
pub struct IEnumConnectionsVtbl {
    pub unknown: IUnknownVtbl,
    pub next: unsafe extern "system" fn(
        this: *mut c_void,
        c_connections: u32,
        rgcd: *mut ConnectData,
        pc_fetched: *mut u32,
    ) -> HResult,
    pub skip: unsafe extern "system" fn(this: *mut c_void, c_connections: u32) -> HResult,
    pub reset: unsafe extern "system" fn(this: *mut c_void) -> HResult,
    pub clone: unsafe extern "system" fn(this: *mut c_void, pp_enum: *mut *mut c_void) -> HResult,
}

#[cfg(test)]
mod tests {
    use super::*;
    use core::mem::size_of;

    const P: usize = size_of::<*const ()>();

    #[test]
    fn vtable_slot_counts() {
        assert_eq!(size_of::<IUnknownVtbl>(), 3 * P);
        assert_eq!(size_of::<IDispatchVtbl>(), 7 * P);
        // IUnknown (3) + GetTypeAttr .. ReleaseVarDesc (19).
        assert_eq!(size_of::<ITypeInfoVtbl>(), 22 * P);
        // IUnknown (3) + RecordInit .. RecordDestroy (16).
        assert_eq!(size_of::<IRecordInfoVtbl>(), 19 * P);
        // IUnknown (3) + GetGUID .. GetHelpContext (5).
        assert_eq!(size_of::<IErrorInfoVtbl>(), 8 * P);
        // IUnknown (3) + InterfaceSupportsErrorInfo (1).
        assert_eq!(size_of::<ISupportErrorInfoVtbl>(), 4 * P);
        // IUnknown (3) + EnumConnectionPoints, FindConnectionPoint (2).
        assert_eq!(size_of::<IConnectionPointContainerVtbl>(), 5 * P);
        // IUnknown (3) + GetConnectionInterface .. EnumConnections (5).
        assert_eq!(size_of::<IConnectionPointVtbl>(), 8 * P);
        // IUnknown (3) + Next, Skip, Reset, Clone (4).
        assert_eq!(size_of::<IEnumConnectionPointsVtbl>(), 7 * P);
        assert_eq!(size_of::<IEnumConnectionsVtbl>(), 7 * P);
    }

    #[test]
    fn guid_parse_and_display_round_trip() {
        let parsed = Guid::parse("{00020400-0000-0000-c000-000000000046}").unwrap();
        assert_eq!(parsed, IID_IDISPATCH);
        assert_eq!(
            Guid::parse("00000000-0000-0000-C000-000000000046"),
            Some(IID_IUNKNOWN)
        );
        assert!(Guid::parse("00020400-0000-0000-C000").is_none());
        assert!(Guid::parse("not-a-guid").is_none());
        assert!(Guid::parse("00020400-0000-0000-C0é0-000000000046").is_none());
        struct Buf([u8; 36], usize);
        impl core::fmt::Write for Buf {
            fn write_str(&mut self, s: &str) -> core::fmt::Result {
                self.0[self.1..self.1 + s.len()].copy_from_slice(s.as_bytes());
                self.1 += s.len();
                Ok(())
            }
        }
        let mut buf = Buf([0; 36], 0);
        core::fmt::write(&mut buf, format_args!("{IID_IDISPATCH}")).unwrap();
        assert_eq!(&buf.0, b"00020400-0000-0000-C000-000000000046");
        assert_eq!(
            Guid::from_le_bytes(IID_IDISPATCH.to_le_bytes()),
            IID_IDISPATCH
        );
        assert_eq!(
            Guid::from_le_bytes(IID_VBA_COLLECTION.to_le_bytes()),
            IID_VBA_COLLECTION
        );
        assert_eq!(
            Guid::parse("A4C46780-499F-101B-BB78-00AA00383CBB"),
            Some(IID_VBA_COLLECTION)
        );
        let det = Guid::deterministic("Demo", "IWidget");
        assert_eq!(det, Guid::deterministic("Demo", "IWidget"));
        assert_ne!(det, Guid::deterministic("Demo", "IOther"));
        assert_ne!(det, IID_IDISPATCH);
    }

    #[test]
    #[cfg(windows)]
    fn guid_matches_windows() {
        use core::mem::{align_of, offset_of};
        use windows_sys::Win32::System::Com::{DISPPARAMS, EXCEPINFO};
        use windows_sys::core::GUID;
        assert_eq!(size_of::<Guid>(), size_of::<GUID>());
        assert_eq!(align_of::<Guid>(), align_of::<GUID>());
        assert_eq!(offset_of!(Guid, data2), offset_of!(GUID, data2));
        assert_eq!(offset_of!(Guid, data4), offset_of!(GUID, data4));

        assert_eq!(size_of::<DispParams>(), size_of::<DISPPARAMS>());
        assert_eq!(align_of::<DispParams>(), align_of::<DISPPARAMS>());
        assert_eq!(
            offset_of!(DispParams, rgvarg),
            offset_of!(DISPPARAMS, rgvarg)
        );
        assert_eq!(
            offset_of!(DispParams, rgdispid_named_args),
            offset_of!(DISPPARAMS, rgdispidNamedArgs)
        );
        assert_eq!(
            offset_of!(DispParams, c_args),
            offset_of!(DISPPARAMS, cArgs)
        );
        assert_eq!(
            offset_of!(DispParams, c_named_args),
            offset_of!(DISPPARAMS, cNamedArgs)
        );

        assert_eq!(size_of::<ExcepInfo>(), size_of::<EXCEPINFO>());
        assert_eq!(align_of::<ExcepInfo>(), align_of::<EXCEPINFO>());
        assert_eq!(offset_of!(ExcepInfo, w_code), offset_of!(EXCEPINFO, wCode));
        assert_eq!(
            offset_of!(ExcepInfo, bstr_source),
            offset_of!(EXCEPINFO, bstrSource)
        );
        assert_eq!(
            offset_of!(ExcepInfo, bstr_description),
            offset_of!(EXCEPINFO, bstrDescription)
        );
        assert_eq!(
            offset_of!(ExcepInfo, bstr_help_file),
            offset_of!(EXCEPINFO, bstrHelpFile)
        );
        assert_eq!(
            offset_of!(ExcepInfo, dw_help_context),
            offset_of!(EXCEPINFO, dwHelpContext)
        );
        assert_eq!(
            offset_of!(ExcepInfo, pv_reserved),
            offset_of!(EXCEPINFO, pvReserved)
        );
        assert_eq!(
            offset_of!(ExcepInfo, pfn_deferred_fill_in),
            offset_of!(EXCEPINFO, pfnDeferredFillIn)
        );
        assert_eq!(offset_of!(ExcepInfo, scode), offset_of!(EXCEPINFO, scode));

        use windows_sys::Win32::System::Com::{ELEMDESC, FUNCDESC, IDLDESC, TYPEATTR, TYPEDESC};
        use windows_sys::Win32::System::Ole::{PARAMDESC, PARAMDESCEX};

        assert_eq!(size_of::<TypeDesc>(), size_of::<TYPEDESC>());
        assert_eq!(align_of::<TypeDesc>(), align_of::<TYPEDESC>());
        assert_eq!(offset_of!(TypeDesc, vt), offset_of!(TYPEDESC, vt));

        assert_eq!(size_of::<IdlDesc>(), size_of::<IDLDESC>());
        assert_eq!(align_of::<IdlDesc>(), align_of::<IDLDESC>());
        assert_eq!(
            offset_of!(IdlDesc, w_idl_flags),
            offset_of!(IDLDESC, wIDLFlags)
        );

        assert_eq!(size_of::<ParamDesc>(), size_of::<PARAMDESC>());
        assert_eq!(align_of::<ParamDesc>(), align_of::<PARAMDESC>());
        assert_eq!(
            offset_of!(ParamDesc, w_param_flags),
            offset_of!(PARAMDESC, wParamFlags)
        );

        assert_eq!(size_of::<ParamDescEx>(), size_of::<PARAMDESCEX>());
        assert_eq!(align_of::<ParamDescEx>(), align_of::<PARAMDESCEX>());
        assert_eq!(
            offset_of!(ParamDescEx, var_default_value),
            offset_of!(PARAMDESCEX, varDefaultValue)
        );

        assert_eq!(size_of::<ElemDesc>(), size_of::<ELEMDESC>());
        assert_eq!(align_of::<ElemDesc>(), align_of::<ELEMDESC>());
        assert_eq!(
            offset_of!(ElemDesc, payload),
            offset_of!(ELEMDESC, Anonymous)
        );

        assert_eq!(size_of::<TypeAttr>(), size_of::<TYPEATTR>());
        assert_eq!(align_of::<TypeAttr>(), align_of::<TYPEATTR>());
        assert_eq!(offset_of!(TypeAttr, lcid), offset_of!(TYPEATTR, lcid));
        assert_eq!(
            offset_of!(TypeAttr, memid_constructor),
            offset_of!(TYPEATTR, memidConstructor)
        );
        assert_eq!(
            offset_of!(TypeAttr, lpstr_schema),
            offset_of!(TYPEATTR, lpstrSchema)
        );
        assert_eq!(
            offset_of!(TypeAttr, cb_size_instance),
            offset_of!(TYPEATTR, cbSizeInstance)
        );
        assert_eq!(
            offset_of!(TypeAttr, typekind),
            offset_of!(TYPEATTR, typekind)
        );
        assert_eq!(offset_of!(TypeAttr, c_funcs), offset_of!(TYPEATTR, cFuncs));
        assert_eq!(
            offset_of!(TypeAttr, cb_size_vft),
            offset_of!(TYPEATTR, cbSizeVft)
        );
        assert_eq!(
            offset_of!(TypeAttr, w_type_flags),
            offset_of!(TYPEATTR, wTypeFlags)
        );
        assert_eq!(
            offset_of!(TypeAttr, tdesc_alias),
            offset_of!(TYPEATTR, tdescAlias)
        );
        assert_eq!(
            offset_of!(TypeAttr, idldesc_type),
            offset_of!(TYPEATTR, idldescType)
        );

        assert_eq!(size_of::<FuncDesc>(), size_of::<FUNCDESC>());
        assert_eq!(align_of::<FuncDesc>(), align_of::<FUNCDESC>());
        assert_eq!(
            offset_of!(FuncDesc, lprgscode),
            offset_of!(FUNCDESC, lprgscode)
        );
        assert_eq!(
            offset_of!(FuncDesc, lprgelemdesc_param),
            offset_of!(FUNCDESC, lprgelemdescParam)
        );
        assert_eq!(
            offset_of!(FuncDesc, funckind),
            offset_of!(FUNCDESC, funckind)
        );
        assert_eq!(offset_of!(FuncDesc, invkind), offset_of!(FUNCDESC, invkind));
        assert_eq!(
            offset_of!(FuncDesc, callconv),
            offset_of!(FUNCDESC, callconv)
        );
        assert_eq!(
            offset_of!(FuncDesc, c_params),
            offset_of!(FUNCDESC, cParams)
        );
        assert_eq!(
            offset_of!(FuncDesc, c_params_opt),
            offset_of!(FUNCDESC, cParamsOpt)
        );
        assert_eq!(offset_of!(FuncDesc, o_vft), offset_of!(FUNCDESC, oVft));
        assert_eq!(
            offset_of!(FuncDesc, elemdesc_func),
            offset_of!(FUNCDESC, elemdescFunc)
        );
        assert_eq!(
            offset_of!(FuncDesc, w_func_flags),
            offset_of!(FUNCDESC, wFuncFlags)
        );

        use windows_sys::Win32::System::Com::CONNECTDATA;
        assert_eq!(size_of::<ConnectData>(), size_of::<CONNECTDATA>());
        assert_eq!(align_of::<ConnectData>(), align_of::<CONNECTDATA>());
        assert_eq!(offset_of!(ConnectData, punk), offset_of!(CONNECTDATA, pUnk));
        assert_eq!(
            offset_of!(ConnectData, dw_cookie),
            offset_of!(CONNECTDATA, dwCookie)
        );
    }
}
