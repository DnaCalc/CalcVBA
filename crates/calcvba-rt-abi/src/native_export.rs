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
