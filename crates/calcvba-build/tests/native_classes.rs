#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_build::{NativeLibraryBuildOptions, build_native_library};
use calcvba_engine::{Engine, HostConfig};
use std::{
    path::PathBuf,
    sync::atomic::{AtomicUsize, Ordering},
};

#[test]
fn project_construction_members_as_new_and_lifetimes_agree() {
    let project = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../examples/native/classes/NativeClasses.basproj");
    let closure = calcvba_project::load_project_closure(&project).unwrap();
    // Captured by VBE Compile and execution in one owned, hidden Excel.
    let expected = "10:4:True:True:Counter:2:3:3";
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let engine = Engine::new(config);
        let values = engine
            .execute_manifest_with_variant_snapshot(closure.last().unwrap())
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|v| v.as_bstr())
                .any(|s| s.to_string() == expected),
            "{values:?}"
        );
    }
    static NEXT: AtomicUsize = AtomicUsize::new(0);
    let out_dir = std::env::temp_dir().join(format!(
        "calcvba-classes-{}-{}",
        std::process::id(),
        NEXT.fetch_add(1, Ordering::Relaxed)
    ));
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: out_dir.clone(),
        template_path: None,
    })
    .unwrap();
    // SAFETY: generated export signatures and same-thread shutdown match this
    // caller; the library remains loaded until all returned BSTR owners drop.
    unsafe {
        use windows_sys::Win32::{
            Foundation::{SysFreeString, SysStringLen},
            System::LibraryLoader::{GetProcAddress, LoadLibraryW},
        };
        let path: Vec<_> = output
            .dll_path
            .as_os_str()
            .to_string_lossy()
            .encode_utf16()
            .chain(Some(0))
            .collect();
        let library = LoadLibraryW(path.as_ptr());
        assert!(!library.is_null());
        let call: unsafe extern "C" fn() -> *mut u16 =
            std::mem::transmute(GetProcAddress(library, c"ClassSuite".as_ptr().cast()).unwrap());
        let shutdown: unsafe extern "C" fn() -> i32 = std::mem::transmute(
            GetProcAddress(library, c"CalcVbaShutdown".as_ptr().cast()).unwrap(),
        );
        let error: unsafe extern "C" fn(*mut usize, *mut usize) -> i32 = std::mem::transmute(
            GetProcAddress(library, c"CalcVbaGetLastError".as_ptr().cast()).unwrap(),
        );
        for _ in 0..2 {
            let text = call();
            let mut source = 0;
            let mut description = 0;
            let number = error(&mut source, &mut description);
            let take = |pointer: *mut u16| {
                let text = String::from_utf16_lossy(std::slice::from_raw_parts(
                    if pointer.is_null() {
                        std::ptr::NonNull::dangling().as_ptr()
                    } else {
                        pointer
                    },
                    SysStringLen(pointer) as usize,
                ));
                SysFreeString(pointer);
                text
            };
            assert_eq!(
                number,
                0,
                "{}:{}",
                take(source as *mut u16),
                take(description as *mut u16)
            );
            assert_eq!(take(text), expected);
        }
        assert_eq!(shutdown(), 0);
        windows_sys::Win32::Foundation::FreeLibrary(library);
    }
    let refused = out_dir.join("refused");
    std::fs::create_dir(&refused).unwrap();
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../examples/native/classes");
    for name in ["Counter.cls", "NativeClasses.basproj"] {
        std::fs::copy(fixture.join(name), refused.join(name)).unwrap();
    }
    let source = std::fs::read_to_string(fixture.join("Counts.bas"))
        .unwrap()
        .replace(
            "70 Set same = first",
            "70 Set same = first\nIf False Then answer = same.Bump(first)",
        );
    std::fs::write(refused.join("Counts.bas"), source).unwrap();
    let error = build_native_library(&NativeLibraryBuildOptions {
        project_path: refused.join("NativeClasses.basproj"),
        out_dir: refused.join("out"),
        template_path: None,
    })
    .unwrap_err()
    .to_string();
    assert!(
        error.contains("Counts:") && error.contains("native class argument conversion"),
        "{error}"
    );
    assert!(!refused.join("out").exists());
    std::fs::remove_dir_all(out_dir).unwrap();
}
