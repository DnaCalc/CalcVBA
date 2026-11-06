//! Developer-built executable startup. User builds fill three RVA slots and
//! append compiled VBA; this startup cannot compile or interpret a procedure.
#[cfg(windows)]
#[repr(C)]
struct Entries {
    marker: [u8; 16],
    main: u32,
    shutdown: u32,
    error: u32,
}
#[cfg(windows)]
#[used]
static ENTRIES: Entries = Entries {
    marker: *b"CVBEXE_ENTRY_V1!",
    main: 0,
    shutdown: 0,
    error: 0,
};

#[cfg(windows)]
fn main() -> std::process::ExitCode {
    use windows_sys::Win32::{
        Foundation::{SysFreeString, SysStringLen},
        System::LibraryLoader::GetModuleHandleW,
    };
    calcvba_native_template::retain_native_helpers();
    // SAFETY: these immutable image cells are filled by the checked PE patcher
    // before loading. Volatile reads prevent folding the template's zero slots.
    let entries = unsafe { std::ptr::addr_of!(ENTRIES).read_volatile() };
    if entries.main == 0 || entries.shutdown == 0 || entries.error == 0 {
        eprintln!("CalcVBA executable template has no compiled entry");
        return std::process::ExitCode::FAILURE;
    }
    // SAFETY: the process image is loaded throughout startup and shutdown; all
    // three patched entries have the fixed native ABI emitted by the compiler.
    let (status, shutdown) = unsafe {
        let base = GetModuleHandleW(std::ptr::null()).cast::<u8>();
        let run: unsafe extern "C" fn() -> i32 =
            std::mem::transmute(base.add(entries.main as usize));
        let stop: unsafe extern "C" fn() -> i32 =
            std::mem::transmute(base.add(entries.shutdown as usize));
        let error: unsafe extern "C" fn(*mut *mut u16, *mut *mut u16) -> i32 =
            std::mem::transmute(base.add(entries.error as usize));
        let status = run();
        if status < 0 {
            let mut source = std::ptr::null_mut();
            let mut description = std::ptr::null_mut();
            let number = error(&mut source, &mut description);
            let take = |pointer: *mut u16| {
                let text = if pointer.is_null() {
                    String::new()
                } else {
                    String::from_utf16_lossy(std::slice::from_raw_parts(
                        pointer,
                        SysStringLen(pointer) as usize,
                    ))
                };
                SysFreeString(pointer);
                text
            };
            eprintln!(
                "VBA error {number} ({}): {}",
                take(source),
                take(description)
            );
        }
        (status, stop())
    };
    if status < 0 || shutdown != 0 {
        std::process::ExitCode::FAILURE
    } else {
        std::process::ExitCode::SUCCESS
    }
}

#[cfg(not(windows))]
fn main() {
    eprintln!("CalcVBA native executable templates require Windows");
    std::process::exit(1);
}
