#![cfg(all(windows, target_arch = "x86_64"))]

use oxvba_build::{NativeLibraryBuildOptions, build_native_library};
use std::{
    path::PathBuf,
    sync::atomic::{AtomicUsize, Ordering},
};

struct Directory(PathBuf);
impl Directory {
    fn new() -> Self {
        static NEXT: AtomicUsize = AtomicUsize::new(0);
        let path = std::env::temp_dir().join(format!(
            "calcvba-native-{}-{}",
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
fn patched_numeric_exports_roundtrip_through_vm3_and_jit_declare() {
    use object::{Object, ObjectSection};
    let temp = Directory::new();
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/native/scalars/NativeScalars.basproj"),
        out_dir: temp.0.clone(),
        template_path: None,
    })
    .unwrap();
    let bytes = std::fs::read(&output.dll_path).unwrap();
    let file = object::File::parse(bytes.as_slice()).unwrap();
    assert_eq!(file.exports().unwrap().len(), 9);
    assert!(file.section_by_name(".cvbtext").is_some());
    let image = file.section_by_name(".cvbimg").unwrap();
    let parsed =
        oxvba_oxir::OxImage::from_bytes(&image.data().unwrap()[..image.size() as usize]).unwrap();
    assert_eq!(parsed.programs.len(), 1);
    let source = include_str!("support/native_scalar_client.bas")
        .replace("__DLL__", output.dll_path.to_str().unwrap());
    let expected = include_str!("support/native_scalar_expected.txt").trim();
    for config in [oxvba_host::HostConfig::vm3(), oxvba_host::HostConfig::jit()] {
        let mut engine = oxvba_host::Engine::new(config);
        engine.set_host_policy(oxvba_hal::model::HostPolicy::native());
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
fn native_build_refuses_private_byref_state_and_unsupported_reachable_bodies() {
    let temp = Directory::new();
    let cases = [
        (
            "Private Function Selected(ByVal x As Long) As Long\nSelected=x\nEnd Function",
            "no Public procedure",
        ),
        (
            "Public Function Selected(ByRef x As Long) As Long\nSelected=x\nEnd Function",
            "ByVal",
        ),
        (
            "Public counter As Long\nPublic Function Selected(ByVal x As Long) As Long\ncounter=x\nSelected=counter\nEnd Function",
            "module state",
        ),
        (
            "Public Function Selected(ByVal x As String) As String\nSelected=x\nEnd Function",
            "ByVal",
        ),
        (
            "Public Function Selected(ByVal x As Long) As Long\nSelected=Descendant(x)\nEnd Function\nPrivate Function Descendant(ByVal x As Long) As Long\nOn Error Resume Next\nDescendant=x\nEnd Function",
            "Descendant",
        ),
    ];
    for (index, (source, message)) in cases.iter().enumerate() {
        let project = temp.0.join(format!("case{index}.basproj"));
        std::fs::write(
            temp.0.join("Code.bas"),
            format!("Attribute VB_Name = \"Code\"\n{source}\n"),
        )
        .unwrap();
        std::fs::write(&project,"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>Refused</ProjectName></PropertyGroup><ItemGroup><Module Include=\"Code.bas\"/><NativeExport Include=\"Export\"><Module>Code</Module><Procedure>Selected</Procedure></NativeExport></ItemGroup></Project>").unwrap();
        let out = temp.0.join(format!("out{index}"));
        let error = build_native_library(&NativeLibraryBuildOptions {
            project_path: project,
            out_dir: out.clone(),
            template_path: None,
        })
        .unwrap_err()
        .to_string();
        assert!(error.contains(message), "expected {message} in {error}");
        assert!(!out.exists(), "refused build published artifacts");
    }
}
