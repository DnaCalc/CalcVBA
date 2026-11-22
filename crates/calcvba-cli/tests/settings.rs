#![cfg(target_os = "linux")]
use std::{fs, process::Command};

#[test]
fn native_settings_persist_and_delete_across_processes() {
    let root = std::env::temp_dir().join(format!("calcvba-settings-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    let source = root.join("main.bas");
    let run = |body: &str, backend: &str| {
        fs::write(&source, format!("Sub Main()\n{body}\nEnd Sub\n")).unwrap();
        let output = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
            .args(["run", "--backend", backend, "--profile", "linux-stdio"])
            .arg(&source)
            .env("XDG_CONFIG_HOME", &root)
            .output()
            .unwrap();
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(output.status.success(), "{stderr}");
        stderr.trim().to_string()
    };

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
