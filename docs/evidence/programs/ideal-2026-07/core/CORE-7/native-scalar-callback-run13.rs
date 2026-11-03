#[unsafe(no_mangle)]
pub unsafe extern "system" fn MutateDoubleThenCall(value: *mut f64, same: *mut f64, callback: extern "system" fn() -> f64) -> f64 {
    if value != same { return -999.0; }
    // SAFETY: VBA Declare supplies one live exact ByRef Double and AddressOf callback.
    unsafe { *value += 10.0; }
    callback()
}
#[unsafe(no_mangle)]
pub unsafe extern "system" fn MutateBooleanThenCall(value: *mut i16, same: *mut i16, callback: extern "system" fn() -> i16) -> i16 {
    if value != same { return -1; }
    // SAFETY: VBA Boolean and Automation VARIANT_BOOL are exact signed 16-bit cells.
    unsafe { *value = -1; }
    callback()
}
