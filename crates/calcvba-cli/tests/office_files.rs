use std::{
    path::{Path, PathBuf},
    process::Command,
};
fn fixture() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../calcvba-office/tests/fixtures/strings/workbook.xlsm")
}
#[test]
fn inspect_and_folder_write_back_work_in_a_copied_cli() {
    let root = std::env::temp_dir().join(format!("calcvba-office-cli-{}", std::process::id()));
    std::fs::create_dir(&root).unwrap();
    let root = root.canonicalize().unwrap();
    assert!(root.starts_with(std::env::temp_dir().canonicalize().unwrap()));
    let cli = root.join(if cfg!(windows) {
        "calcvba.exe"
    } else {
        "calcvba"
    });
    std::fs::copy(env!("CARGO_BIN_EXE_calcvba-cli"), &cli).unwrap();
    let input = fixture();
    let inspect = Command::new(&cli)
        .arg("inspect-vba")
        .arg(&input)
        .env("PATH", "")
        .output()
        .unwrap();
    assert!(
        inspect.status.success(),
        "{}",
        String::from_utf8_lossy(&inspect.stderr)
    );
    let json: serde_json::Value = serde_json::from_slice(&inspect.stdout).unwrap();
    assert_eq!(json["codepage"], 1252);
    assert_eq!(json["modules"].as_array().unwrap().len(), 3);
    let folder = root.join("source");
    let export = Command::new(&cli)
        .arg("export-vba")
        .arg(&input)
        .arg("--out-dir")
        .arg(&folder)
        .env("PATH", "")
        .output()
        .unwrap();
    assert!(
        export.status.success(),
        "{}",
        String::from_utf8_lossy(&export.stderr)
    );
    let api = folder.join("Api.bas");
    let mut source = std::fs::read_to_string(&api).unwrap();
    source = source.replace(
        "RunProbe = \"ok\"",
        "RunProbe = \"changed by standalone CLI\"",
    );
    std::fs::write(api, source).unwrap();
    let output = root.join("changed.xlsm");
    let written = Command::new(&cli)
        .arg("write-vba")
        .arg(&input)
        .arg("--project")
        .arg(folder.join("VBAProject.basproj"))
        .arg("--out")
        .arg(&output)
        .env("PATH", "")
        .output()
        .unwrap();
    assert!(
        written.status.success(),
        "{}",
        String::from_utf8_lossy(&written.stderr)
    );
    let file = calcvba_office::OfficeFile::open(&output).unwrap();
    assert!(
        file.project
            .modules
            .iter()
            .any(|m| m.source.contains("changed by standalone CLI"))
    );
    let original = std::fs::read(&output).unwrap();
    let refused = Command::new(&cli)
        .arg("write-vba")
        .arg(&input)
        .arg("--project")
        .arg(folder.join("VBAProject.basproj"))
        .arg("--out")
        .arg(&output)
        .output()
        .unwrap();
    assert!(!refused.status.success());
    assert_eq!(std::fs::read(output).unwrap(), original);
    // Windows may retain the just-exited copied executable briefly. Retry only
    // cleanup of this test's owned directory; all editing assertions stay above.
    let mut cleanup = std::fs::remove_dir_all(&root);
    for _ in 0..20 {
        if !cleanup
            .as_ref()
            .is_err_and(|error| error.kind() == std::io::ErrorKind::PermissionDenied)
        {
            break;
        }
        std::thread::sleep(std::time::Duration::from_millis(100));
        cleanup = std::fs::remove_dir_all(&root);
    }
    cleanup.unwrap();
}
