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
    assert_eq!(file.exports().unwrap().len(), 32);
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

#[test]
fn native_build_refuses_more_than_32_parameters_and_invalid_export_ordinals() {
    let temp = Directory::new();
    let parameters = (0..33)
        .map(|i| format!("ByVal p{i} As Long"))
        .collect::<Vec<_>>()
        .join(", ");
    let cases = [
        (
            format!("Public Function Selected({parameters}) As Long\nSelected=p0\nEnd Function"),
            "",
            "argument/slot bounds",
        ),
        (
            "Public Function Selected() As Long\nSelected=1&\nEnd Function".into(),
            "<Ordinal>0</Ordinal>",
            "ordinal",
        ),
    ];
    for (index, (source, metadata, expected)) in cases.iter().enumerate() {
        std::fs::write(temp.0.join("Code.bas"), source).unwrap();
        let project = temp.0.join("Refused.basproj");
        std::fs::write(&project, format!("<Project><PropertyGroup><OutputType>NativeLibrary</OutputType></PropertyGroup><ItemGroup><Module Include=\"Code.bas\"/><NativeExport Include=\"Selected\"><Module>Code</Module><Procedure>Selected</Procedure>{metadata}</NativeExport></ItemGroup></Project>")).unwrap();
        let out = temp.0.join(format!("out{index}"));
        let error = build_native_library(&NativeLibraryBuildOptions {
            project_path: project,
            out_dir: out.clone(),
            template_path: None,
        })
        .unwrap_err()
        .to_string();
        assert!(error.contains(expected), "{error}");
        assert!(!out.exists());
    }
}

#[test]
fn escaped_xml_export_names_resolve_by_the_decoded_name() {
    use object::Object;
    let temp = Directory::new();
    let project = temp.0.join("Names.basproj");
    std::fs::write(
        temp.0.join("Code&Numbers.bas"),
        "Attribute VB_Name = \"Code\"\nPublic Function Selected() As Long\nSelected=42&\nEnd Function",
    )
    .unwrap();
    std::fs::write(&project,"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>Names</ProjectName></PropertyGroup><ItemGroup><Module Include=\"Code&amp;Numbers.bas\"/><NativeExport Include=\"A&amp;B.C\"><Module>Code</Module><Procedure>Selected</Procedure></NativeExport></ItemGroup></Project>").unwrap();
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: temp.0.join("out"),
        template_path: None,
    })
    .unwrap();
    let bytes = std::fs::read(&output.dll_path).unwrap();
    let file = object::File::parse(bytes.as_slice()).unwrap();
    assert_eq!(file.exports().unwrap()[0].name(), b"A&B.C");
    let source = format!(
        "Private Declare PtrSafe Function ReadName Lib \"{}\" Alias \"A&B.C\" () As Long\nSub Main()\nDim result As Long\nresult=ReadName()\nEnd Sub",
        output.dll_path.display()
    );
    for config in [oxvba_host::HostConfig::vm3(), oxvba_host::HostConfig::jit()] {
        let mut engine = oxvba_host::Engine::new(config);
        engine.set_host_policy(oxvba_hal::model::HostPolicy::native());
        assert!(
            engine
                .execute_source_with_variant_snapshot_clean(&source)
                .unwrap()
                .iter()
                .any(|value| value.as_i32() == Some(42))
        );
    }
}

#[test]
fn discarded_numeric_conversions_still_execute_and_propagate_errors() {
    let temp = Directory::new();
    let project = temp.0.join("Discarded.basproj");
    let mut source = String::from("Attribute VB_Name = \"Code\"\n");
    let mut exports = String::new();
    for parameter in ["Double", "Long", "Boolean"] {
        source.push_str(&format!(
            "Public Function From{parameter}(ByVal value As {parameter}) As Long\nCall CLng(value)\nCall CDbl(value)\nCall CBool(value)\nFrom{parameter}=42&\nEnd Function\n"
        ));
        exports.push_str(&format!(
            "<NativeExport Include=\"From{parameter}\"><Module>Code</Module><Procedure>From{parameter}</Procedure></NativeExport>"
        ));
    }
    std::fs::write(temp.0.join("Code.bas"), source).unwrap();
    std::fs::write(&project, format!("<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>Discarded</ProjectName></PropertyGroup><ItemGroup><Module Include=\"Code.bas\"/>{exports}</ItemGroup></Project>")).unwrap();
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: temp.0.join("out"),
        template_path: None,
    })
    .unwrap();
    let source = format!(
        "Private Declare PtrSafe Function FromDouble Lib \"{dll}\" (ByVal value As Double) As Long\nPrivate Declare PtrSafe Function FromLong Lib \"{dll}\" (ByVal value As Long) As Long\nPrivate Declare PtrSafe Function FromBoolean Lib \"{dll}\" (ByVal value As Boolean) As Long\nSub Main()\nDim observed As String\nobserved=CStr(FromDouble(2.5#)) & \"|\" & CStr(FromDouble(1E+30)) & \"|\" & CStr(FromDouble(-2.5#)) & \"|\" & CStr(FromLong(-2147483647&)) & \"|\" & CStr(FromBoolean(True))\nEnd Sub",
        dll = output.dll_path.display()
    );
    for config in [oxvba_host::HostConfig::vm3(), oxvba_host::HostConfig::jit()] {
        let mut engine = oxvba_host::Engine::new(config);
        engine.set_host_policy(oxvba_hal::model::HostPolicy::native());
        let values = engine
            .execute_source_with_variant_snapshot_clean(&source)
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|value| value.as_bstr())
                .any(|value| value.to_string() == "42|0|42|42|42"),
            "discarded conversions: {values:?}"
        );
    }
}
