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
    unsafe {
        *value = 73;
        SetLastError(1234);
    }
}

static ALIAS_CALLS: std::sync::atomic::AtomicI32 = std::sync::atomic::AtomicI32::new(0);

macro_rules! alias_scalar {
    ($name:ident, $ty:ty, $first:expr, $increment:expr) => {
        #[unsafe(no_mangle)]
        pub unsafe extern "system" fn $name(first: *mut $ty, second: *mut $ty) -> $ty {
            // SAFETY: the fixture requires writable scalar pointers. Raw reads
            // and writes deliberately allow the two pointers to be identical.
            unsafe {
                first.write($first);
                second.write(second.read() + $increment);
                ALIAS_CALLS.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
                SetLastError(4321);
                first.read()
            }
        }
    };
}
alias_scalar!(AliasLong, i32, 41, 1);
alias_scalar!(AliasInteger, i16, -32767, 1);
alias_scalar!(AliasByte, u8, 250, 1);
alias_scalar!(AliasBoolean, i16, -1, 1);
alias_scalar!(AliasI64, i64, 4_294_967_337, 1);
alias_scalar!(AliasSingle, f32, 123.5, 0.25);
alias_scalar!(AliasDouble, f64, 123456789.5, 0.25);
alias_scalar!(AliasCurrency, i64, 1_234_567_890_000, 10_000);
alias_scalar!(AliasDate, f64, 45123.5, 0.25);

#[unsafe(no_mangle)]
pub extern "system" fn AliasCallCount() -> i32 {
    // SAFETY: only this thread's OS error state is changed.
    unsafe { SetLastError(4321) };
    ALIAS_CALLS.load(std::sync::atomic::Ordering::Relaxed)
}

#[unsafe(no_mangle)]
pub extern "system" fn PointerProbe(
    _first: *mut std::ffi::c_void,
    _second: *mut std::ffi::c_void,
) -> i32 {
    ALIAS_CALLS.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    42
}

#[unsafe(no_mangle)]
pub extern "system" fn CurrencyValue() -> i64 {
    1_234_567_890_123
}
#[unsafe(no_mangle)]
pub extern "system" fn DateValue() -> f64 {
    45123.75
}
#[unsafe(no_mangle)]
pub extern "system" fn ByteValue() -> u8 {
    251
}
// Retained for the explicitly open foreign-Boolean-payload observation.
#[unsafe(no_mangle)]
pub extern "system" fn BoolOne() -> i16 {
    1
}

#[unsafe(no_mangle)]
pub extern "system" fn Mixed(
    a: i32,
    b: f64,
    c: f32,
    d: i64,
    e: i16,
    f: f64,
    g: i64,
    h: i16,
    i: u8,
    j: i64,
) -> i32 {
    // SAFETY: only this thread's OS error state is changed.
    unsafe { SetLastError(4321) };
    if a != 19 {
        return -1001;
    }
    if b != 23.5 {
        return -1002;
    }
    if c != 12.5 {
        return -1003;
    }
    if d != 12_345_000 {
        return -1004;
    }
    if e != -2 {
        return -1005;
    }
    if f != 45123.75 {
        return -1006;
    }
    if g != 4_294_967_338 {
        return -1007;
    }
    if h != -1 {
        return -1008;
    }
    if i != 251 {
        return -1009;
    }
    if j != -9_876_543_210_987 {
        return -1010;
    }
    42
}
