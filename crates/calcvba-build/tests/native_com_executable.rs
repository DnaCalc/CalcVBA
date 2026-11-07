#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_build::{NativeLibraryBuildOptions, build_native_com_executable, build_native_library};
use calcvba_engine::{Engine, HostConfig};
use std::{
    path::PathBuf,
    sync::atomic::{AtomicUsize, Ordering},
};

struct Directory(PathBuf);
impl Directory {
    fn new() -> Self {
        static NEXT: AtomicUsize = AtomicUsize::new(0);
        let path = std::env::temp_dir().join(format!(
            "calcvba-com-exe-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        std::fs::create_dir_all(&path).unwrap();
        Self(path)
    }
}
impl Drop for Directory {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

#[test]
fn com_executable_has_compiled_factories_dual_slots_and_matching_class_bodies() {
    let directory = Directory::new();
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../examples/native/comexe");
    let built = build_native_com_executable(&NativeLibraryBuildOptions {
        project_path: fixture.join("NativeCounter.basproj"),
        out_dir: directory.0.join("server"),
        template_path: None,
    })
    .unwrap();
    assert!(built.exe_path.is_file());
    let descriptor: calcvba_rt_abi::native_export::NativeComServer = serde_json::from_slice(
        &std::fs::read(built.exe_path.with_extension("native-com.json")).unwrap(),
    )
    .unwrap();
    assert_eq!(descriptor.classes.len(), 1);
    assert_eq!(descriptor.classes[0].entries.len(), 7);
    assert!(!descriptor.classes[0].single_use);
    // Loading a local TLB does not register it or activate a COM server.
    let library = calcvba_com::windows_typelib_loader::load_typelib_from_path(
        &built.exe_path.with_extension("tlb").display().to_string(),
    )
    .unwrap();
    let members = calcvba_com::windows_typelib_loader::enumerate_typelib_members(library).unwrap();
    // SAFETY: loading above gave this scope one retained type-library reference.
    unsafe {
        calcvba_com::windows_typelib_loader::release_typelib(library);
    }
    assert!(
        members
            .iter()
            .any(|member| member.name == "Add" && member.is_dual && member.vtable_slot == Some(7))
    );
    assert!(
        members
            .iter()
            .any(|member| member.name.eq_ignore_ascii_case("Value")
                && member.vtable_slot == Some(11))
    );

    std::fs::copy(fixture.join("Counter.cls"), directory.0.join("Counter.cls")).unwrap();
    let client = std::fs::read_to_string(fixture.join("Client.bas"))
        .unwrap()
        .replace("CalcVbaNativeCounter27Lib.", "");
    let source = format!(
        "{client}\nPublic observed As String\nPublic Sub Main()\nobserved = Probe()\nEnd Sub\n"
    );
    // Module declarations must precede procedures, as in VBE source.
    let source = source.replace("Public observed As String\n", "").replace(
        "Option Explicit",
        "Option Explicit\nPublic observed As String",
    );
    std::fs::write(directory.0.join("Client.bas"), source).unwrap();
    let project = directory.0.join("Bodies.basproj");
    std::fs::write(&project, r#"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>CounterBodies</ProjectName></PropertyGroup><ItemGroup><Module Include="Client.bas"/><ClassModule Include="Counter.cls"/><NativeExport Include="Probe"><Module>ComExeClient</Module><Procedure>Probe</Procedure></NativeExport></ItemGroup></Project>"#).unwrap();
    let closure = calcvba_project::load_project_closure(&project).unwrap();
    let expected = "15:10:7:41:hello:42:True:513:late:42";
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(calcvba_hal::HostPolicy::native());
        let values = engine
            .execute_manifest_with_variant_snapshot(closure.last().unwrap())
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|v| v.as_bstr())
                .any(|v| v.to_string() == expected),
            "{values:?}"
        );
    }
    let native = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: directory.0.join("bodies"),
        template_path: None,
    })
    .unwrap();
    let path: Vec<_> = native
        .dll_path
        .to_string_lossy()
        .encode_utf16()
        .chain(Some(0))
        .collect();
    // SAFETY: this test owns its generated DLL and exact zero-argument BSTR ABI;
    // calls, owning result disposal, shutdown and unload use the same thread.
    unsafe {
        use windows_sys::Win32::{
            Foundation::{FreeLibrary, SysFreeString, SysStringLen},
            System::LibraryLoader::{GetProcAddress, LoadLibraryW},
        };
        let library = LoadLibraryW(path.as_ptr());
        assert!(!library.is_null());
        let call: unsafe extern "system" fn() -> *mut u16 =
            std::mem::transmute(GetProcAddress(library, c"Probe".as_ptr().cast()).unwrap());
        let stop: unsafe extern "system" fn() -> i32 = std::mem::transmute(
            GetProcAddress(library, c"CalcVbaShutdown".as_ptr().cast()).unwrap(),
        );
        let value = call();
        assert!(!value.is_null());
        let text = String::from_utf16_lossy(std::slice::from_raw_parts(
            value,
            SysStringLen(value) as usize,
        ));
        SysFreeString(value);
        assert_eq!(stop(), 0);
        FreeLibrary(library);
        assert_eq!(text, expected);
    }
    let invalid = std::fs::read_to_string(fixture.join("Counter.cls"))
        .unwrap()
        .replace(
            "Echo(ByVal text As String) As String",
            "Echo(ByVal text As String, Optional ByVal extra As Long) As String",
        );
    std::fs::write(directory.0.join("Counter.cls"), invalid).unwrap();
    std::fs::copy(
        fixture.join("NativeCounter.basproj"),
        directory.0.join("NativeCounter.basproj"),
    )
    .unwrap();
    let refused = build_native_com_executable(&NativeLibraryBuildOptions {
        project_path: directory.0.join("NativeCounter.basproj"),
        out_dir: directory.0.join("refused"),
        template_path: None,
    })
    .unwrap_err()
    .to_string();
    assert!(
        refused.contains("Optional parameter 1 is not a Variant"),
        "{refused}"
    );
    assert!(!directory.0.join("refused").exists());
}
