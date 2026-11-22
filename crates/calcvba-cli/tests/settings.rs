#![cfg(target_os = "linux")]
use std::{fs, path::Path, process::Command};

fn run(root: &Path, body: &str, backend: &str) -> String {
    let source = root.join("main.bas");
    fs::write(&source, format!("Sub Main()\n{body}\nEnd Sub\n")).unwrap();
    let output = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
        .args(["run", "--backend", backend, "--profile", "linux-stdio"])
        .arg(&source)
        .env("XDG_CONFIG_HOME", root)
        .output()
        .unwrap();
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(output.status.success(), "{stderr}");
    stderr.trim().to_string()
}

#[test]
fn native_settings_persist_and_delete_across_processes() {
    let root = std::env::temp_dir().join(format!("calcvba-settings-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    let run = |body, backend| run(&root, body, backend);

    run(r#"SaveSetting "Probe", "Section", "Key", "stored""#, "vm3");
    let first = run(
        r#"Debug.Print GetSetting("probe", "SECTION", "key", "missing")"#,
        "jit",
    );
    let second = run(
        r#"Debug.Print GetSetting("probe", "SECTION", "key", "missing")"#,
        "vm3",
    );
    if second == "stored" {
        run(r#"DeleteSetting "PROBE", "section", "key""#, "jit");
    }
    let deleted = run(
        r#"Debug.Print GetSetting("probe", "section", "key", "missing")"#,
        "vm3",
    );
    fs::remove_dir_all(root).unwrap();
    assert_eq!(
        [first.as_str(), second.as_str(), deleted.as_str()],
        ["stored", "stored", "missing"]
    );
}

#[test]
fn native_app_delete_persists_the_partial_deletion_when_a_section_has_a_child() {
    // SS1 captured error 5 after deleting the first two empty sections.
    let actual = [("vm3", "jit"), ("jit", "vm3")].map(|(deleting, reading)| {
        let root = std::env::temp_dir().join(format!(
            "calcvba-settings-nested-{}-{deleting}",
            std::process::id()
        ));
        fs::create_dir(&root).unwrap();
        run(
            &root,
            r#"
    Dim section As Variant
    For Each section In Array(" Space Section ", "MiXeD", "Outer\Inner", "Payload", "Slashes", "Text")
        SaveSetting "Probe", section, "seed", "seed"
        DeleteSetting "Probe", section, "seed"
    Next
"#,
            reading,
        );
        let error = run(
            &root,
            r#"
    On Error Resume Next
    DeleteSetting "PROBE"
    Debug.Print CStr(Err.Number)
"#,
            deleting,
        );
        let remaining = run(
            &root,
            r#"
    Dim section As Variant
    On Error Resume Next
    For Each section In Array(" Space Section ", "MiXeD", "outer\inner", "OUTER", "Payload", "Slashes", "Text")
        Err.Clear
        DeleteSetting "Probe", section
        Debug.Print CStr(Err.Number)
    Next
    Err.Clear
    DeleteSetting "Probe"
    Debug.Print CStr(Err.Number)
    Err.Clear
    DeleteSetting "Probe"
    Debug.Print CStr(Err.Number)
"#,
            reading,
        );
        fs::remove_dir_all(root).unwrap();
        (error, remaining)
    });
    let expected = ("5".to_owned(), "5\n5\n0\n0\n0\n0\n0\n0\n5".to_owned());
    assert_eq!(actual, [expected.clone(), expected]);
}

#[test]
fn native_settings_reads_legacy_rows_through_the_next_write() {
    let root = std::env::temp_dir().join(format!("calcvba-settings-legacy-{}", std::process::id()));
    fs::create_dir_all(root.join("calcvba")).unwrap();
    fs::write(
        root.join("calcvba/settings"),
        r#"[[["probe","outer\\inner","first"],{"key_name":"First","value":"legacy","seq":4}]]"#,
    )
    .unwrap();
    run(
        &root,
        r#"
    SaveSetting "Probe", "Outer\Inner", "FIRST", "updated"
    SaveSetting "Probe", "Outer\Inner", "Second", "new"
"#,
        "jit",
    );
    let values = run(
        &root,
        r#"
    Dim entries As Variant
    entries = GetAllSettings("PROBE", "outer\inner")
    Debug.Print entries(0, 0) & "=" & entries(0, 1)
    Debug.Print entries(1, 0) & "=" & entries(1, 1)
"#,
        "vm3",
    );
    fs::remove_dir_all(root).unwrap();
    assert_eq!(values, "First=updated\nSecond=new");
}
