#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_build::{NativeLibraryBuildOptions, build_native_library};
use calcvba_engine::{Engine, HostConfig};
use std::{
    path::PathBuf,
    sync::atomic::{AtomicUsize, Ordering},
};
use windows_sys::Win32::{
    Foundation::{FreeLibrary, HMODULE, SysFreeString, SysStringLen},
    System::LibraryLoader::{GetProcAddress, LoadLibraryW},
};

struct Directory(PathBuf);
impl Directory {
    fn new() -> Self {
        static NEXT: AtomicUsize = AtomicUsize::new(0);
        let path = std::env::temp_dir().join(format!(
            "calcvba-native-members-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        std::fs::create_dir_all(&path).unwrap();
        assert_eq!(
            path.canonicalize().unwrap().parent(),
            Some(std::env::temp_dir().canonicalize().unwrap().as_path())
        );
        Self(path)
    }
}
impl Drop for Directory {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}
struct Dll(HMODULE);
impl Dll {
    fn load(path: &std::path::Path) -> Self {
        let name: Vec<_> = path
            .to_string_lossy()
            .encode_utf16()
            .chain(Some(0))
            .collect();
        // SAFETY: terminated absolute path; this test owns the DLL until Drop.
        let module = unsafe { LoadLibraryW(name.as_ptr()) };
        assert!(!module.is_null());
        Self(module)
    }
    fn call(&self, name: &str) -> String {
        let name = std::ffi::CString::new(name).unwrap();
        // SAFETY: test exports are zero-argument BSTR Functions; owning exports
        // and the error accessor are called on this DLL's first calling thread.
        unsafe {
            let entry: unsafe extern "C" fn() -> *mut u16 =
                std::mem::transmute(GetProcAddress(self.0, name.as_ptr().cast()).unwrap());
            let error: unsafe extern "C" fn(*mut usize, *mut usize) -> i32 = std::mem::transmute(
                GetProcAddress(self.0, c"CalcVbaGetLastError".as_ptr().cast()).unwrap(),
            );
            let result = entry();
            let mut source = 0;
            let mut description = 0;
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
            let result = take(result);
            let source = take(source as *mut u16);
            let description = take(description as *mut u16);
            assert_eq!(number, 0, "{source}: {description}");
            result
        }
    }
}
impl Drop for Dll {
    fn drop(&mut self) {
        // SAFETY: same-thread cleanup runs before the owned image is unloaded.
        unsafe {
            let shutdown: unsafe extern "C" fn() -> i32 = std::mem::transmute(
                GetProcAddress(self.0, c"CalcVbaShutdown".as_ptr().cast()).unwrap(),
            );
            let _ = shutdown();
            FreeLibrary(self.0);
        }
    }
}

#[test]
fn collection_interfaces_late_calls_and_failed_construction_agree() {
    let directory = Directory::new();
    let support = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/support/native_members");
    for entry in std::fs::read_dir(support).unwrap() {
        let entry = entry.unwrap();
        std::fs::copy(entry.path(), directory.0.join(entry.file_name())).unwrap();
    }
    let project = directory.0.join("Members.basproj");
    std::fs::write(&project,r#"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>NativeMembers</ProjectName></PropertyGroup><ItemGroup><Module Include="Members.bas"/><ClassModule Include="ICounter.cls"/><ClassModule Include="Worker.cls"/><ClassModule Include="OtherWorker.cls"/><ClassModule Include="BadCtor.cls"/><NativeExport Include="MemberSuite"><Module>Members</Module><Procedure>MemberSuite</Procedure></NativeExport><NativeExport Include="ConstructorFaultSuite"><Module>Members</Module><Procedure>ConstructorFaultSuite</Procedure></NativeExport><NativeExport Include="NothingSuite"><Module>Members</Module><Procedure>NothingSuite</Procedure></NativeExport></ItemGroup></Project>"#).unwrap();
    let closure = calcvba_project::load_project_closure(&project).unwrap();
    // Captured by VBE Compile and execution in the reused owned Excel process.
    let expected = "15:5|35:7:True|Collection:3:a:a23:2:0|z:3|1:9|438:3#513:Nothing:1:0#91";
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(calcvba_hal::HostPolicy::native());
        let values = engine
            .execute_manifest_with_variant_snapshot(closure.last().unwrap())
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|value| value.as_bstr())
                .any(|value| value.to_string() == expected),
            "{values:?}"
        );
    }
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: directory.0.join("out"),
        template_path: None,
    })
    .unwrap();
    let dll = Dll::load(&output.dll_path);
    assert_eq!(
        format!(
            "{}#{}#{}",
            dll.call("MemberSuite"),
            dll.call("ConstructorFaultSuite"),
            dll.call("NothingSuite")
        ),
        expected
    );
}

#[test]
fn imported_early_dictionary_members_agree() {
    let directory = Directory::new();
    let support = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/support/native_members/NativeCom.bas");
    std::fs::copy(support, directory.0.join("NativeCom.bas")).unwrap();
    let project = directory.0.join("Early.basproj");
    std::fs::write(&project,r#"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>NativeEarly</ProjectName></PropertyGroup><ItemGroup><Module Include="NativeCom.bas"/><COMReference Include="Scripting"><Guid>{420B2830-E718-11CF-893D-00A0C9054228}</Guid><VersionMajor>1</VersionMajor><VersionMinor>0</VersionMinor><Lcid>0</Lcid><ImportLib>scrrun.dll</ImportLib></COMReference><NativeExport Include="EarlyOwnDictionarySuite"><Module>NativeCom</Module><Procedure>EarlyOwnDictionarySuite</Procedure></NativeExport></ItemGroup></Project>"#).unwrap();
    let closure = calcvba_project::load_project_closure(&project).unwrap();
    let expected = "2:7:11:True:True:False|5:Invalid procedure call or argument";
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(calcvba_hal::HostPolicy::native());
        let values = engine
            .execute_manifest_with_variant_snapshot(closure.last().unwrap())
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|value| value.as_bstr())
                .any(|value| value.to_string() == expected),
            "{values:?}"
        );
    }
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: directory.0.join("out"),
        template_path: None,
    })
    .unwrap();
    let dll = Dll::load(&output.dll_path);
    assert_eq!(dll.call("EarlyOwnDictionarySuite"), expected);
}
