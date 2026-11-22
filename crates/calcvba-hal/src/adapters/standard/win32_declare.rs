use calcvba_runtime::win32_declare::Win32Declare;
use std::time::Duration;

pub(super) fn address(entry: Win32Declare) -> usize {
    let address = match entry {
        Win32Declare::Sleep => sleep as *const (),
        Win32Declare::GetCurrentProcessId => process_id as *const (),
        Win32Declare::GetTickCount => tick_count as *const (),
        Win32Declare::GetTickCount64 => tick_count64 as *const (),
        Win32Declare::QueryPerformanceCounter => performance_counter as *const (),
        Win32Declare::QueryPerformanceFrequency => performance_frequency as *const (),
        Win32Declare::StringLengthWide => string_length_wide as *const (),
        Win32Declare::MoveMemory => move_memory as *const (),
        Win32Declare::MoveMemoryLongCount => move_memory_long_count as *const (),
    };
    address.expose_provenance()
}

fn uptime() -> Option<Duration> {
    #[cfg(target_os = "linux")]
    let clock = libc::CLOCK_BOOTTIME;
    #[cfg(target_os = "macos")]
    let clock = libc::CLOCK_MONOTONIC;
    let mut value = libc::timespec {
        tv_sec: 0,
        tv_nsec: 0,
    };
    // SAFETY: the selected clock is supported by this OS and value is valid output storage.
    (unsafe { libc::clock_gettime(clock, &mut value) } == 0)
        .then(|| Duration::new(value.tv_sec as u64, value.tv_nsec as u32))
}

extern "C" fn sleep(milliseconds: u32) {
    match milliseconds {
        0 => std::thread::yield_now(),
        u32::MAX => loop {
            std::thread::park();
        },
        ms => std::thread::sleep(Duration::from_millis(u64::from(ms))),
    }
}

extern "C" fn process_id() -> u32 {
    std::process::id()
}

extern "C" fn tick_count() -> u32 {
    tick_count64() as u32
}

extern "C" fn tick_count64() -> u64 {
    uptime().map_or(0, |value| value.as_millis() as u64)
}

unsafe extern "C" fn performance_counter(output: *mut i64) -> i32 {
    let Some(value) = uptime() else { return 0 };
    // SAFETY: the Declare caller supplies a writable LARGE_INTEGER.
    unsafe { output.write_unaligned(value.as_nanos() as i64) };
    1
}

unsafe extern "C" fn performance_frequency(output: *mut i64) -> i32 {
    // SAFETY: the Declare caller supplies a writable LARGE_INTEGER.
    unsafe { output.write_unaligned(1_000_000_000) };
    1
}

unsafe extern "C" fn string_length_wide(text: *const u16) -> i32 {
    if text.is_null() {
        return 0;
    }
    let mut len = 0;
    // SAFETY: the Declare caller supplies a readable, NUL-terminated UTF-16 string.
    while unsafe { text.add(len).read_unaligned() } != 0 {
        len += 1;
    }
    len as i32
}

unsafe extern "C" fn move_memory(dest: *mut u8, src: *const u8, bytes: usize) {
    if bytes != 0 {
        // SAFETY: the Declare caller supplies readable and writable spans of bytes.
        unsafe { std::ptr::copy(src, dest, bytes) };
    }
}

unsafe extern "C" fn move_memory_long_count(dest: *mut u8, src: *const u8, bytes: u32) {
    // SAFETY: the Declare caller supplies readable and writable spans of bytes.
    unsafe { move_memory(dest, src, bytes as usize) };
}
