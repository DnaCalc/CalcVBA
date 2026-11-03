#[unsafe(no_mangle)]
pub unsafe extern "system" fn MutateThenCall(value: *mut i32, callback: extern "system" fn() -> i32) -> i32 {
    // SAFETY: VBA Declare supplies a live exact ByRef Long and AddressOf callback.
    unsafe { *value += 10; }
    callback()
}
