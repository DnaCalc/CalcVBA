//! Product CLI path: compile the top-level Exe extension once, remove source,
//! then execute the same artifact through each real backend in another process.
use std::{
    path::PathBuf,
    process::{Command, Output},
};

struct Workspace(PathBuf);
impl Workspace {
    fn new() -> Self {
        let stamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root =
            std::env::temp_dir().join(format!("calcvba-image-{}-{stamp}", std::process::id()));
        std::fs::create_dir(&root).unwrap();
        Self(root)
    }
}
impl Drop for Workspace {
    fn drop(&mut self) {
        std::fs::remove_dir_all(&self.0).unwrap();
    }
}
fn cli(args: &[&str]) -> Output {
    Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
        .args(args)
        .output()
        .unwrap()
}

#[test]
fn top_level_exe_package_runs_without_source_and_rejects_corruption() {
    let workspace = Workspace::new();
    let project = workspace.0.join("Script.basproj");
    let source = workspace.0.join("Script.bas");
    let out = workspace.0.join("out");
    std::fs::write(&project, r#"<Project><PropertyGroup><ProjectName>Script</ProjectName><OutputType>Exe</OutputType></PropertyGroup><ItemGroup><Module Include="Script.bas" /></ItemGroup></Project>"#).unwrap();
    std::fs::write(&source, "Public result As Long\nresult = 42\n").unwrap();
    let build = cli(&[
        "build",
        project.to_str().unwrap(),
        "--target",
        "OxImage",
        "--out-dir",
        out.to_str().unwrap(),
    ]);
    assert!(
        build.status.success(),
        "build: {}",
        String::from_utf8_lossy(&build.stderr)
    );
    let image = out.join("Script.oxi");
    assert!(image.exists());
    std::fs::remove_file(source).unwrap();
    std::fs::remove_file(project).unwrap();
    for backend in ["vm3", "jit"] {
        let run = cli(&[
            "run-image",
            image.to_str().unwrap(),
            "--backend",
            backend,
            "--dump-values",
        ]);
        assert!(
            run.status.success(),
            "{backend}: {}",
            String::from_utf8_lossy(&run.stderr)
        );
        assert!(
            String::from_utf8_lossy(&run.stdout).contains("42"),
            "{backend}: {run:?}"
        );
    }
    std::fs::write(
        &image,
        br#"{"format":"bad","version":2,"entry":0,"programs":[]}"#,
    )
    .unwrap();
    for backend in ["vm3", "jit"] {
        let run = cli(&[
            "run-image",
            image.to_str().unwrap(),
            "--backend",
            backend,
            "--diagnostic-format",
            "json",
        ]);
        assert!(!run.status.success());
        assert!(String::from_utf8_lossy(&run.stderr).contains("RUN-E-IMAGE"));
    }
}
