#![cfg(all(windows, target_arch = "x86_64"))]

use calcvba_build::{NativeLibraryBuildOptions, build_native_library};
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
    assert_eq!(file.exports().unwrap().len(), 35);
    assert!(file.section_by_name(".cvbtext").is_some());
    let image = file.section_by_name(".cvbimg").unwrap();
    let parsed =
        calcvba_oxir::OxImage::from_bytes(&image.data().unwrap()[..image.size() as usize]).unwrap();
    assert_eq!(parsed.programs.len(), 1);
    let source = include_str!("support/native_scalar_client.bas")
        .replace("__DLL__", output.dll_path.to_str().unwrap());
    let expected = include_str!("support/native_scalar_expected.txt").trim();
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
fn native_build_refuses_private_optional_paramarray_and_unsupported_reachable_bodies() {
    let temp = Directory::new();
    let cases = [
        (
            "Private Function Selected(ByVal x As Long) As Long\nSelected=x\nEnd Function",
            "no Public procedure",
        ),
        (
            "Public Function Selected(Optional ByVal x As Long = 0) As Long\nSelected=x\nEnd Function",
            "optional signature",
        ),
        (
            "Public Function Selected(ParamArray values() As Variant) As Long\nSelected=1&\nEnd Function",
            "ParamArray signature",
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
        if index > 0 {
            assert!(
                error.contains("Code:"),
                "missing source declaration: {error}"
            );
        }
        assert!(!out.exists(), "refused build published artifacts");
    }
}

#[test]
fn every_numeric_boundary_width_matches_both_declare_backends_and_direct_bodies() {
    let temp = Directory::new();
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/native/widths/NativeWide.basproj"),
        out_dir: temp.0.clone(),
        template_path: None,
    })
    .unwrap();
    let client = include_str!("support/native_wide_client.bas")
        .replace("__DLL__", output.dll_path.to_str().unwrap());
    let expected = include_str!("support/native_wide_expected.txt").trim();
    let direct = format!(
        "{}\nPublic Sub Main()\nDim observed As String\nobserved=WideSuite()\nEnd Sub\n",
        include_str!("../../../examples/native/widths/Wide.bas")
    );
    let direct_expected = "255:255:-1:-1:9007199254740994:9007199254740994:81985529216486903:81985529216486903:1.75:1.75:12346:12346:45000.75:45000.75:45106.5";
    for config in [
        calcvba_engine::HostConfig::vm3(),
        calcvba_engine::HostConfig::jit(),
    ] {
        let mut engine = calcvba_engine::Engine::new(config);
        engine.set_host_policy(calcvba_hal::model::HostPolicy::native());
        for (source, expected) in [(&client, expected), (&direct, direct_expected)] {
            let values = engine
                .execute_source_with_variant_snapshot_clean(source)
                .unwrap();
            assert!(
                values
                    .iter()
                    .filter_map(|v| v.as_bstr())
                    .any(|v| v.to_string() == expected),
                "numeric boundary/direct results: {values:?}"
            );
        }
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
    for config in [
        calcvba_engine::HostConfig::vm3(),
        calcvba_engine::HostConfig::jit(),
    ] {
        let mut engine = calcvba_engine::Engine::new(config);
        engine.set_host_policy(calcvba_hal::model::HostPolicy::native());
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
                .filter_map(|value| value.as_bstr())
                .any(|value| value.to_string() == "42|0|42|42|42"),
            "discarded conversions: {values:?}"
        );
    }
}

#[test]
fn native_ownership_state_errors_and_arrays_roundtrip_through_both_declare_backends() {
    let temp = Directory::new();
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/native/interop/NativeInterop.basproj"),
        out_dir: temp.0.clone(),
        template_path: None,
    })
    .unwrap();
    let source = include_str!("support/native_interop_client.bas")
        .replace("__DLL__", output.dll_path.to_str().unwrap());
    let expected = include_str!("support/native_interop_expected.txt").trim();
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
            "native interop values: {values:?}"
        );
    }
}

#[test]
fn the_same_owning_bodies_agree_in_native_vm3_and_jit() {
    let source = include_str!("../../../examples/native/interop/Interop.bas");
    let source = format!(
        "{source}\nPublic Sub Main()\nDim observed As String\nobserved = ValueSuite()\nEnd Sub\n"
    );
    let expected = "19:6:13:24:12:3.5:1.75:False:<omega>:omega!:6:42:42:8.75:ab!:49:-2:3.5:3.75:3.75:6:7.5:left|right:right!:True:9:9:7";
    for config in [
        calcvba_engine::HostConfig::vm3(),
        calcvba_engine::HostConfig::jit(),
    ] {
        let engine = calcvba_engine::Engine::new(config);
        let values = engine
            .execute_source_with_variant_snapshot_clean(&source)
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|v| v.as_bstr())
                .any(|v| v.to_string() == expected),
            "direct bodies: {values:?}"
        );
    }
    // The C and Declare clients execute ValueSuite through its named AOT export
    // and check this same literal, including every mutable result it observes.
}

#[test]
fn ordinary_vba_xll_conversions_use_native_pointers_and_declares() {
    let temp = Directory::new();
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/native/xll/PlainXll.basproj"),
        out_dir: temp.0.clone(),
        template_path: None,
    })
    .unwrap();
    let source = include_str!("support/native_xll_client.bas")
        .replace("__DLL__", output.dll_path.to_str().unwrap());
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
                .any(|v| v.to_string() == "5.75|3.75|9|16385|1|0"),
            "XLOPER12 conversions: {values:?}"
        );
    }
}

#[test]
fn native_variant_fixed_string_records_and_string_redim_roundtrip() {
    let temp = Directory::new();
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/native/records/NativeRecords.basproj"),
        out_dir: temp.0.clone(),
        template_path: None,
    })
    .unwrap();
    let source = include_str!("support/native_record_client.bas")
        .replace("__DLL__", output.dll_path.to_str().unwrap());
    let expected = include_str!("support/native_record_expected.txt").trim();
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
            "{values:?}"
        );
    }
}
