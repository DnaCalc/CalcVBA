//! Developer-built executable startup. User builds fill three RVA slots and
//! append compiled VBA; this startup cannot compile or interpret a procedure.
#[cfg(any(
    windows,
    all(target_os = "linux", target_arch = "x86_64", target_env = "gnu")
))]
#[repr(C)]
struct Entries {
    marker: [u8; 16],
    main: u32,
    shutdown: u32,
    error: u32,
}
#[cfg(any(
    windows,
    all(target_os = "linux", target_arch = "x86_64", target_env = "gnu")
))]
#[used]
#[cfg_attr(target_os = "linux", unsafe(link_section = ".cvbentry"))]
static ENTRIES: Entries = Entries {
    marker: *b"CVBEXE_ENTRY_V1!",
    main: 0,
    shutdown: 0,
    error: 0,
};

#[cfg(any(
    windows,
    all(target_os = "linux", target_arch = "x86_64", target_env = "gnu")
))]
fn main() -> std::process::ExitCode {
    calcvba_native_template::retain_native_helpers();
    // SAFETY: these immutable image cells are filled by the checked template patcher
    // before loading. Volatile reads prevent folding the template's zero slots.
    let entries = unsafe { std::ptr::addr_of!(ENTRIES).read_volatile() };
    if entries.main == 0 || entries.shutdown == 0 || entries.error == 0 {
        eprintln!("CalcVBA executable template has no compiled entry");
        return std::process::ExitCode::FAILURE;
    }
    // SAFETY: the process image is loaded throughout startup and shutdown; all
    // three patched entries have the fixed native ABI emitted by the compiler.
    let (status, shutdown) = unsafe {
        #[cfg(windows)]
        let entry = |rva: u32| {
            windows_sys::Win32::System::LibraryLoader::GetModuleHandleW(std::ptr::null())
                .cast::<u8>()
                .wrapping_add(rva as usize)
                .cast_const()
        };
        #[cfg(target_os = "linux")]
        let entry = |delta: u32| {
            std::ptr::addr_of!(ENTRIES)
                .cast::<u8>()
                .wrapping_offset(delta as i32 as isize)
        };
        let run: unsafe extern "C" fn() -> i32 = std::mem::transmute(entry(entries.main));
        let stop: unsafe extern "C" fn() -> i32 = std::mem::transmute(entry(entries.shutdown));
        let error: unsafe extern "C" fn(*mut *mut u16, *mut *mut u16) -> i32 =
            std::mem::transmute(entry(entries.error));
        let status = run();
        if status < 0 {
            let mut source = std::ptr::null_mut();
            let mut description = std::ptr::null_mut();
            let number = error(&mut source, &mut description);
            let take = |pointer| calcvba_automation::bstr::BStr::from_raw_bstr(pointer).to_string();
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

#[cfg(not(any(
    windows,
    all(target_os = "linux", target_arch = "x86_64", target_env = "gnu")
)))]
fn main() {
    eprintln!("CalcVBA native executable templates require Windows x64 or Linux x86-64");
    std::process::exit(1);
}
