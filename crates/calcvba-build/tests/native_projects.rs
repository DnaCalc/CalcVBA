#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_build::{NativeLibraryBuildOptions, build_native_library};
use calcvba_engine::{Engine, HostConfig};
use std::path::PathBuf;
use windows_sys::Win32::{
    Foundation::{FreeLibrary, SysFreeString, SysStringLen},
    System::LibraryLoader::{GetProcAddress, LoadLibraryW},
};

#[test]
fn source_project_diamond_calls_classes_globals_and_error_sources_agree() {
    let project = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/support/native_projects/Root.basproj");
    let closure = calcvba_project::load_project_closure(&project).unwrap();
    assert_eq!(
        closure.len(),
        4,
        "one shared leaf state for both reference branches"
    );
    let expected = "1:2:0:9|3:2:11:Token:False:True:1:6:Leaf#7:4:11:Token:False:True:2:6:Leaf";
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let values = Engine::new(config)
            .execute_project_closure_with_variant_snapshot(&closure)
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|v| v.as_bstr())
                .any(|v| v.as_str() == expected),
            "{values:?}"
        );
    }
    let directory =
        std::env::temp_dir().join(format!("calcvba-native-projects-{}", std::process::id()));
    assert!(!directory.exists());
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: directory.clone(),
        template_path: std::env::var_os("CALCVBA_NATIVE_TEMPLATE_TEST").map(PathBuf::from),
    })
    .unwrap();
    // SAFETY: the generated entries have these exact signatures; all owners and
    // calls stay on this test thread until the module is shut down and unloaded.
    unsafe {
        let path: Vec<_> = output
            .dll_path
            .to_string_lossy()
            .encode_utf16()
            .chain(Some(0))
            .collect();
        let dll = LoadLibraryW(path.as_ptr());
        assert!(!dll.is_null());
        let call: unsafe extern "C" fn() -> *mut u16 =
            std::mem::transmute(GetProcAddress(dll, c"Probe".as_ptr().cast()).unwrap());
        let error: unsafe extern "C" fn(*mut usize, *mut usize) -> i32 = std::mem::transmute(
            GetProcAddress(dll, c"CalcVbaGetLastError".as_ptr().cast()).unwrap(),
        );
        let shutdown: unsafe extern "C" fn() -> i32 =
            std::mem::transmute(GetProcAddress(dll, c"CalcVbaShutdown".as_ptr().cast()).unwrap());
        for _ in 0..2 {
            let text = call();
            let number = error(std::ptr::null_mut(), std::ptr::null_mut());
            assert_eq!(number, 0);
            assert!(!text.is_null());
            let result = String::from_utf16_lossy(std::slice::from_raw_parts(
                text,
                SysStringLen(text) as usize,
            ));
            SysFreeString(text);
            assert_eq!(result, expected);
        }
        assert_eq!(shutdown(), 0);
        FreeLibrary(dll);
    }
    std::fs::remove_dir_all(directory).unwrap();
}
