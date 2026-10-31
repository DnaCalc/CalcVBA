//! Independent x64 Win64 ABI fixture, compiled as a DLL by the owning test.
#[link(name = "kernel32")]
unsafe extern "system" {
    fn SetLastError(error: u32);
}

#[unsafe(no_mangle)]
pub extern "system" fn ScalarAdd(a: i32, b: i32) -> i32 {
    // SAFETY: SetLastError only changes this caller thread's Win32 error state.
    unsafe { SetLastError(0) };
    a + b
}

#[unsafe(no_mangle)]
pub unsafe extern "system" fn ScalarWrite(value: *mut i32) {
    // SAFETY: the fixture contract requires one writable i32 from the caller.
    unsafe { *value = 73; SetLastError(1234); }
}
