#![cfg(target_os = "linux")]

use std::{fs, process::Command};

fn assert_probe(case: &str, source: &str, expected: &str) {
    let root = std::env::temp_dir().join(format!("calcvba-{case}-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    let dir = root.join("Data");
    fs::create_dir(&dir).unwrap();
    fs::write(dir.join("Note.txt"), "portable").unwrap();
    let path = root.join("Probe.bas");
    fs::write(&path, source).unwrap();

    let outputs = ["vm3", "jit"].map(|backend| {
        Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
            .args([
                "run",
                "--backend",
                backend,
                "--profile",
                "linux-stdio",
                "--portability-mode",
                "windows-faithful",
            ])
            .arg(&path)
            .output()
            .unwrap()
    });
    fs::remove_dir_all(root).unwrap();
    for (backend, output) in ["vm3", "jit"].into_iter().zip(outputs) {
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(output.status.success(), "{backend}: {stderr}");
        assert_eq!(stderr.trim(), expected, "{backend}");
    }
}

#[test]
fn faithful_mode_uses_the_portable_regexp() {
    assert_probe(
        "regexp",
        include_str!("fixtures/portable-regexp.bas"),
        "True|False",
    );
}

#[test]
fn faithful_mode_parses_xml_with_independent_document_instances() {
    assert_probe(
        "xml",
        include_str!("fixtures/portable-xml.bas"),
        "True|False|True|root|hello & world|True|False|True|root|independent",
    );
}

#[test]
fn faithful_mode_file_system_object_uses_mapped_paths() {
    assert_probe(
        "fso",
        include_str!("fixtures/portable-fso.bas"),
        "True|True|False|False|False",
    );
}

#[test]
fn browser_report_refuses_dom_document_without_a_com_adapter() {
    let root = std::env::temp_dir().join(format!("calcvba-browser-xml-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    let path = root.join("Probe.bas");
    fs::write(&path, include_str!("fixtures/portable-xml.bas")).unwrap();
    let output = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
        .args(["check", "--portability", "--host", "browser"])
        .arg(&path)
        .output()
        .unwrap();
    fs::remove_dir_all(root).unwrap();
    assert!(output.status.success(), "{output:?}");
    let report: serde_json::Value = serde_json::from_slice(&output.stdout).unwrap();
    let findings = report["findings"].as_array().unwrap();
    for feature in ["CreateObject", "GetObject"] {
        assert!(
            findings
                .iter()
                .any(|finding| { finding["feature"] == feature && finding["verdict"] == "reject" }),
            "{report}"
        );
    }
}
