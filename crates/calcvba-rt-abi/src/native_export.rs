//! Header of the per-image, single-owner native-library state. Generated code
//! addresses globals immediately after this header; no procedure lookup occurs.
#[repr(C)]
pub struct Context {
    pub remaining_blocks: u64,
    pub depth: u32,
    /// 0: uninitialised, -1: initialising, -2: ready, positive: failed VBA code.
    pub initialization: i32,
    pub error_source: *const u8,
    pub error_source_len: usize,
    pub error_description: *const u8,
    pub error_description_len: usize,
    pub error_number: i32,
    pub custom_error: u32,
    /// Owned error text retained past the activation that raised it.
    pub error_owner: *mut core::ffi::c_void,
    /// Borrowed String() boundary cells of the active public wrapper.
    pub string_array_cells: *const usize,
    pub string_array_count: usize,
    pub string_array_encoding: u32,
    /// Runtime library/host services, created on demand outside the loader lock.
    pub services: *mut core::ffi::c_void,
}

pub const WRONG_THREAD: i32 = 1444;
pub const E_WRONG_THREAD: i32 = 0x8001010eu32 as i32;

pub fn hresult(code: i32) -> i32 {
    match code {
        0 => 0,
        WRONG_THREAD => E_WRONG_THREAD,
        code if code < 0 => code,
        code => (0x800a0000 | (code as u32 & 0xffff)) as i32,
    }
}

/// Retained member-call metadata for native cells. This contains imported COM
/// ABI descriptors and argument ownership facts, never executable source/IR.
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct NativeMemberCall {
    pub receiver_kind: u32,
    pub result_kind: u32,
    pub binding: NativeMemberBinding,
    pub arguments: Vec<NativeMemberArgument>,
}
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub enum NativeMemberBinding {
    Early {
        member: Box<calcvba_com::TypeLibMemberMetadata>,
        interfaces: Vec<calcvba_com::TypeLibInterfaceMetadata>,
    },
    Late {
        name: Option<String>,
        kind: calcvba_com::TypeLibMemberInvokeKind,
    },
    CallByName,
}
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub struct NativeMemberArgument {
    pub kind: u32,
    pub by_ref: bool,
    pub omitted: bool,
    pub name: Option<String>,
}
