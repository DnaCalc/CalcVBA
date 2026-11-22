#![cfg(all(target_os = "linux", target_arch = "x86_64", target_env = "gnu"))]

use calcvba_build::{NativeLibraryBuildOptions, build_native_executable, build_native_library};
use std::{path::PathBuf, process::Command};

struct Directory(PathBuf);

impl Directory {
    fn new(case: &str) -> Self {
        let path =
            std::env::temp_dir().join(format!("calcvba-linux-{}-{case}", std::process::id()));
        std::fs::create_dir(&path).unwrap();
        Self(path)
    }

    fn project(&self, output_type: &str, exports: &str, source: &str) -> NativeLibraryBuildOptions {
        let project_path = self.0.join("Project.basproj");
        std::fs::write(
            &project_path,
            format!(
                "<Project><PropertyGroup><OutputType>{output_type}</OutputType>                 <ProjectName>LinuxFixture</ProjectName></PropertyGroup>                 <ItemGroup><Module Include=\"Code.bas\"/>{exports}</ItemGroup></Project>"
            ),
        )
        .unwrap();
        std::fs::write(self.0.join("Code.bas"), source).unwrap();
        NativeLibraryBuildOptions {
            project_path,
            out_dir: self.0.join("out"),
            template_path: None,
        }
    }
}

impl Drop for Directory {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

#[test]
fn linux_library_exports_a_system_v_procedure_and_its_project_image() {
    use object::{Object, ObjectSection};

    let directory = Directory::new("library");
    let output = build_native_library(&directory.project(
        "NativeLibrary",
        "<NativeExport Include=\"Scale\"><Module>Code</Module><Procedure>Scale</Procedure>         <Ordinal>7</Ordinal></NativeExport>",
        r#"Private Declare PtrSafe Function ForeignPid Lib "libc" Alias "getpid" () As Long
Private Function Unused() As Long
    Unused = ForeignPid()
End Function
Public Function Scale(ByVal value As Double, ByRef factor As Long) As Double
    factor = factor + 1
    Scale = Sqr(value) * CDbl(factor)
End Function"#,
    ))
    .unwrap();
    assert_eq!(output.dll_path.extension().unwrap(), "so");
    let bytes = std::fs::read(&output.dll_path).unwrap();
    let file = object::File::parse(bytes.as_slice()).unwrap();
    assert_eq!(file.format(), object::BinaryFormat::Elf);
    let mut names: Vec<_> = file
        .exports()
        .unwrap()
        .into_iter()
        .map(|export| std::str::from_utf8(export.name()).unwrap())
        .collect();
    names.sort_unstable();
    assert_eq!(
        names,
        [
            "CalcVbaFreeBstr",
            "CalcVbaGetLastError",
            "CalcVbaGetLastErrorHelp",
            "CalcVbaShutdown",
            "Scale"
        ]
    );
    assert!(file.section_by_name(".cvbframe").is_some());
    assert!(file.section_by_name(".eh_frame_hdr").is_some());
    let image = file.section_by_name(".cvbimg").unwrap().data().unwrap();
    assert_eq!(
        calcvba_oxir::OxImage::from_bytes(image)
            .unwrap()
            .programs
            .len(),
        1
    );
    let map = std::fs::read_to_string(&output.map_path).unwrap();
    assert!(map.contains("Scale\t7\tCode.Scale\t"), "{map}");

    use calcvba_com::windows_ffi_bridge::{OwnedLibrary, get_proc_address};
    let library = OwnedLibrary::load(output.dll_path.to_str().unwrap()).unwrap();
    let handle = library.handle();
    let scale = get_proc_address(handle, "Scale").unwrap();
    let error = get_proc_address(handle, "CalcVbaGetLastError").unwrap();
    let free = get_proc_address(handle, "CalcVbaFreeBstr").unwrap();
    let shutdown = get_proc_address(handle, "CalcVbaShutdown").unwrap();
    for helper in [
        "cvb_sqr_f64",
        "cvb_template_abi",
        "rt_add_i16",
        "cvb_free_bstr",
    ] {
        assert!(get_proc_address(handle, helper).is_err(), "{helper}");
    }
    // SAFETY: the live library supplies C ABI functions with the declared signatures.
    unsafe {
        let scale = std::mem::transmute::<usize, extern "C" fn(f64, &mut i32) -> f64>(scale);
        let error =
            std::mem::transmute::<usize, extern "C" fn(*mut *mut u16, *mut *mut u16) -> i32>(error);
        let free = std::mem::transmute::<usize, extern "C" fn(*mut u16)>(free);
        let shutdown = std::mem::transmute::<usize, extern "C" fn() -> i32>(shutdown);
        let mut factor = 13;
        assert_eq!(scale(9.0, &mut factor).to_bits(), 42.0f64.to_bits());
        assert_eq!(factor, 14);
        std::thread::spawn(move || {
            let mut factor = 13;
            assert_eq!(scale(9.0, &mut factor).to_bits(), 0.0f64.to_bits());
            assert_eq!(factor, 13);
            assert_eq!(shutdown(), 1444);
        })
        .join()
        .unwrap();
        assert_eq!(scale(-1.0, &mut factor).to_bits(), 0.0f64.to_bits());
        let mut source = std::ptr::null_mut();
        assert_eq!(error(&mut source, std::ptr::null_mut()), 5);
        let name = "LinuxFixture".encode_utf16().collect::<Vec<_>>();
        assert_eq!(std::slice::from_raw_parts(source, name.len()), name);
        free(source);
        assert_eq!(shutdown(), 0);
    }
}

#[test]
fn linux_executable_runs_main_and_reports_unhandled_errors() {
    let directory = Directory::new("executable");
    let source = "Public observed As String\nPublic Sub Main()\n                  observed = \"Linux native\"\nDebug.Print observed\nEnd Sub\n";
    for config in [
        calcvba_engine::HostConfig::vm3(),
        calcvba_engine::HostConfig::jit(),
    ] {
        let values = calcvba_engine::Engine::new(config)
            .execute_source_with_variant_snapshot_clean(source)
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|v| v.as_bstr())
                .any(|v| v.to_string() == "Linux native")
        );
    }
    let options = directory.project("Exe", "", source);
    let output = build_native_executable(&options).unwrap();
    assert!(output.exe_path.extension().is_none());
    let result = Command::new(&output.exe_path).output().unwrap();
    assert!(result.status.success(), "{result:?}");
    assert_eq!(
        String::from_utf8_lossy(&result.stderr).trim(),
        "Linux native"
    );

    std::fs::write(
        directory.0.join("Code.bas"),
        "Public Sub Main()\nErr.Raise 13\nEnd Sub\n",
    )
    .unwrap();
    let output = build_native_executable(&options).unwrap();
    let result = Command::new(output.exe_path).output().unwrap();
    assert!(!result.status.success(), "{result:?}");
    assert!(
        String::from_utf8_lossy(&result.stderr).contains("VBA error 13 (LinuxFixture)"),
        "{result:?}"
    );
}

#[test]
fn linux_scalar_exports_match_win64_fixture_through_both_declare_backends() {
    let directory = Directory::new("parity");
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/native/scalars/NativeScalars.basproj"),
        out_dir: directory.0.clone(),
        template_path: None,
    })
    .unwrap();
    declare_parity(
        &output.dll_path,
        include_str!("support/native_scalar_client.bas"),
        include_str!("support/native_scalar_expected.txt"),
    );
}

#[test]
fn linux_numeric_boundary_widths_match_win64_fixture_through_both_declare_backends() {
    let directory = Directory::new("widths");
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/native/widths/NativeWide.basproj"),
        out_dir: directory.0.clone(),
        template_path: None,
    })
    .unwrap();
    // Keep the width cases and their expected values. Replace the Win32
    // allocation and copy calls used by the fixture's returned-BSTR reader.
    let source = include_str!("support/native_wide_client.bas")
        .replace(
            "Private Declare PtrSafe Function BStrLength Lib \"oleaut32\" Alias \"SysStringLen\" (ByVal text As LongPtr) As Long",
            "Private Function BStrLength(ByVal text As LongPtr) As Long\n             Dim byteCount As Long\n             If text <> 0 Then CopyBytes VarPtr(byteCount), text - 4, 4\n             BStrLength = byteCount \\ 2\nEnd Function",
        )
        .replace(
            "Lib \"oleaut32\" Alias \"SysFreeString\"",
            "Lib \"__DLL__\" Alias \"CalcVbaFreeBstr\"",
        )
        .replace("Lib \"kernel32\" Alias \"RtlMoveMemory\"", "Lib \"libc\" Alias \"memmove\"");
    declare_parity(
        &output.dll_path,
        &source,
        include_str!("support/native_wide_expected.txt"),
    );
}

fn declare_parity(library: &std::path::Path, source: &str, expected: &str) {
    let source = source.replace("__DLL__", library.to_str().unwrap());
    let expected = expected.trim();
    for config in [
        calcvba_engine::HostConfig::vm3(),
        calcvba_engine::HostConfig::jit(),
    ] {
        let mut engine = calcvba_engine::Engine::new(config);
        engine.set_host_policy(calcvba_hal::model::HostPolicy::native());
        let values = engine
            .execute_source_with_variant_snapshot_clean(&source)
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|v| v.as_bstr())
                .any(|v| v.to_string() == expected),
            "Declare results: {values:?}"
        );
    }
}

#[test]
fn linux_utf16_and_hresult_descriptor_controls_the_c_abi() {
    use calcvba_com::windows_ffi_bridge::{OwnedLibrary, get_proc_address};

    let directory = Directory::new("descriptor");
    let output = build_native_library(&directory.project(
        "NativeLibrary",
        "<NativeExport Include=\"ReadLength\"><Module>Code</Module><Procedure>ReadLength</Procedure>         <CallingConvention>Cdecl</CallingConvention><StringEncoding>Utf16</StringEncoding>         <ErrorPolicy>HResult</ErrorPolicy></NativeExport>",
        "Public Function ReadLength(ByVal text As String) As Long\n         If Len(text) = 0 Then Err.Raise 5\n         ReadLength = Len(text)\nEnd Function\n",
    ))
    .unwrap();
    let library = OwnedLibrary::load(output.dll_path.to_str().unwrap()).unwrap();
    let read = get_proc_address(library.handle(), "ReadLength").unwrap();
    let stop = get_proc_address(library.handle(), "CalcVbaShutdown").unwrap();
    // SAFETY: the tests use both fixed signatures, terminated UTF-16 and a live out pointer.
    unsafe {
        let read = std::mem::transmute::<usize, extern "C" fn(*const u16, &mut i32) -> i32>(read);
        let stop = std::mem::transmute::<usize, extern "C" fn() -> i32>(stop);
        let mut result = -1;
        let text = [0x03bb, 0xd835, 0xdd18, 0];
        assert_eq!(read(text.as_ptr(), &mut result), 0);
        assert_eq!(result, 3);
        assert_eq!(read(std::ptr::null(), &mut result) as u32, 0x800a0005);
        assert_eq!(result, 0);
        assert_eq!(stop(), 0);
    }
}
