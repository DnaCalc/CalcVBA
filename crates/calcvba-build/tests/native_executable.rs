#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_build::{NativeExecutableBuildOptions, build_native_executable};
use calcvba_engine::{Engine, HostConfig};
use object::Object;
use std::{
    path::PathBuf,
    process::Command,
    sync::atomic::{AtomicUsize, Ordering},
};

struct Directory(PathBuf);
impl Directory {
    fn new() -> Self {
        static NEXT: AtomicUsize = AtomicUsize::new(0);
        let path = std::env::temp_dir().join(format!(
            "calcvba-exe-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        std::fs::create_dir(&path).unwrap();
        Self(path)
    }
    fn project(&self, source: &str) -> PathBuf {
        std::fs::write(self.0.join("Code.bas"), source).unwrap();
        let path = self.0.join("Program.basproj");
        std::fs::write(&path,"<Project><PropertyGroup><OutputType>Exe</OutputType><ProjectName>Program</ProjectName></PropertyGroup><ItemGroup><Module Include=\"Code.bas\"/></ItemGroup></Project>").unwrap();
        path
    }
}
impl Drop for Directory {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

#[test]
fn same_sub_main_runs_on_vm3_jit_and_a_standalone_native_executable_with_early_com() {
    let directory = Directory::new();
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../examples/native/exe");
    for file in ["Hello.bas", "Hello.basproj"] {
        std::fs::copy(fixture.join(file), directory.0.join(file)).unwrap();
    }
    let project = directory.0.join("Hello.basproj");
    let closure = calcvba_project::load_project_closure(&project).unwrap();
    let expected = "native-exe:41:42:1";
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(calcvba_hal::HostPolicy::native());
        let values = engine
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
    let output = build_native_executable(&NativeExecutableBuildOptions {
        project_path: project,
        out_dir: directory.0.join("out"),
        template_path: None,
    })
    .unwrap();
    let bytes = std::fs::read(&output.exe_path).unwrap();
    let pe = object::File::parse(bytes.as_slice()).unwrap();
    assert_eq!(pe.kind(), object::ObjectKind::Executable);
    for import in pe.imports().unwrap() {
        let library = String::from_utf8_lossy(import.library()).to_ascii_lowercase();
        assert!(
            !library.contains("calcvba") && !library.contains("cranelift"),
            "{library}"
        );
    }
    // A user run needs only the produced image and Windows. In particular,
    // neither the CLI nor a Rust toolchain is available through this PATH.
    let system = PathBuf::from(std::env::var_os("SystemRoot").unwrap()).join("System32");
    let result = Command::new(&output.exe_path)
        .current_dir(&directory.0)
        .env("PATH", system)
        .output()
        .unwrap();
    assert!(
        result.status.success(),
        "{:?}: {}",
        result.status,
        String::from_utf8_lossy(&result.stderr)
    );
    assert_eq!(String::from_utf8_lossy(&result.stderr).trim(), expected);
    let mut already_filled = bytes;
    assert!(
        calcvba_pe::patch_executable_entries(&mut already_filled, &[], "CalcVbaExeMain").is_err()
    );
}

#[test]
fn executable_links_the_source_project_diamond_and_preserves_class_identity() {
    let directory = Directory::new();
    let fixtures = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/support/native_projects");
    for file in std::fs::read_dir(fixtures).unwrap() {
        let file = file.unwrap();
        std::fs::copy(file.path(), directory.0.join(file.file_name())).unwrap();
    }
    let project = directory.0.join("Root.basproj");
    let xml = std::fs::read_to_string(&project).unwrap();
    let start = xml.find("<NativeExport ").unwrap();
    let end = xml.find("</NativeExport>").unwrap() + "</NativeExport>".len();
    let xml = format!("{}{}", &xml[..start], &xml[end..]).replace("NativeLibrary", "Exe");
    std::fs::write(&project, xml).unwrap();
    let source = directory.0.join("RootCode.bas");
    let text = std::fs::read_to_string(&source).unwrap().replace(
        "observed = Probe()",
        "observed = Probe()\n    Debug.Print observed",
    );
    std::fs::write(source, text).unwrap();
    let output = build_native_executable(&NativeExecutableBuildOptions {
        project_path: project,
        out_dir: directory.0.join("out"),
        template_path: None,
    })
    .unwrap();
    let result = Command::new(output.exe_path).output().unwrap();
    assert!(result.status.success(), "{:?}", result);
    assert_eq!(
        String::from_utf8_lossy(&result.stderr).trim(),
        "1:2:0:9|3:2:11:Token:False:True:1:6:Leaf#7:4:11:Token:False:True:2:6:Leaf"
    );
}

#[test]
fn main_errors_report_details_and_wrong_template_kind_is_refused() {
    let directory = Directory::new();
    let project = directory
        .project("Public Sub Main()\nErr.Raise 513, \"Exe.Source\", \"native failure\"\nEnd Sub\n");
    let output = build_native_executable(&NativeExecutableBuildOptions {
        project_path: project.clone(),
        out_dir: directory.0.join("out"),
        template_path: None,
    })
    .unwrap();
    let result = Command::new(output.exe_path).output().unwrap();
    assert!(!result.status.success());
    assert!(
        String::from_utf8_lossy(&result.stderr)
            .contains("VBA error 513 (Exe.Source): native failure"),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let error = build_native_executable(&NativeExecutableBuildOptions {
        project_path: project,
        out_dir: directory.0.join("wrong"),
        template_path: Some(
            PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                .join("../../assets/native/calcvba-native-x64.dll"),
        ),
    })
    .unwrap_err()
    .to_string();
    assert!(error.contains("EXE native template"), "{error}");
    assert!(!directory.0.join("wrong").exists());
}

#[test]
fn missing_main_and_unsupported_bodies_publish_no_executable() {
    let directory = Directory::new();
    let project = directory.project("Public Function Answer() As Long\nAnswer=42\nEnd Function\n");
    let options = NativeExecutableBuildOptions {
        project_path: project.clone(),
        out_dir: directory.0.join("out"),
        template_path: None,
    };
    assert!(build_native_executable(&options).is_err());
    assert!(!options.out_dir.exists());
    std::fs::write(
        directory.0.join("Worker.cls"),
        "Public Property Get Value() As Long\nValue=42\nEnd Property\n",
    )
    .unwrap();
    std::fs::write(directory.0.join("Code.bas"),"Sub Main()\nDim item As Worker, member As String\nSet item=New Worker\nmember=\"Value\"\nDebug.Print CallByName(item,member,vbGet)\nEnd Sub\n").unwrap();
    let xml = std::fs::read_to_string(&project).unwrap().replace(
        "</ItemGroup>",
        "<ClassModule Include=\"Worker.cls\"/></ItemGroup>",
    );
    std::fs::write(&project, xml).unwrap();
    let error = build_native_executable(&options).unwrap_err().to_string();
    assert!(error.contains("dynamic project CallByName"), "{error}");
    assert!(!options.out_dir.exists());
}
