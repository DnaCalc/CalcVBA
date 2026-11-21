use std::{fs, process::Command};

#[test]
fn check_reports_host_shims_and_windows_dependencies_at_source_lines() {
    let path = std::env::temp_dir().join(format!("calcvba-portability-{}.bas", std::process::id()));
    fs::write(
        &path,
        "Private Declare PtrSafe Sub Sleep Lib \"kernel32\" (ByVal ms As Long)\n\
         Sub Example()\n\
         Debug.Print Left$(\"abc\", 1), Now, Dir(\"C:\\temp\\*\")\n\
         Set obj = CreateObject(\"Scripting.Dictionary\")\n\
         Set other = CreateObject(\"Excel.Application\")\n\
         End Sub\n",
    )
    .unwrap();
    let output = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
        .args(["check", "--portability", "--host", "linux"])
        .arg(&path)
        .output()
        .unwrap();
    fs::remove_file(path).unwrap();
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report: serde_json::Value = serde_json::from_slice(&output.stdout).unwrap();
    assert_eq!(report["verdict"], "needs Windows");
    let uses = report["findings"].as_array().unwrap();
    for (feature, line, verdict) in [
        ("Declare", 1, "reject"),
        ("Left", 3, "native"),
        ("Now", 3, "host"),
        ("Dir", 3, "shim"),
        ("Windows path", 3, "shim"),
        ("CreateObject", 4, "replace"),
        ("CreateObject", 5, "reject"),
    ] {
        assert!(
            uses.iter().any(|use_| {
                use_["feature"] == feature && use_["line"] == line && use_["verdict"] == verdict
            }),
            "{feature} {line} {verdict}: {report}"
        );
    }
}
