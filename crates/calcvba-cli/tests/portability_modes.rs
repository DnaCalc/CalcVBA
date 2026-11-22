use std::{fs, process::Command};

#[test]
fn strict_and_faithful_refuse_at_load_on_both_backends() {
    let path = std::env::temp_dir().join(format!(
        "calcvba-portability-modes-{}.bas",
        std::process::id()
    ));
    // The Windows declaration is unreachable. The report must still identify it
    // before executing Main, independent of which backend would run it.
    fs::write(
        &path,
        "Private Declare PtrSafe Sub Sleep Lib \"kernel32\" (ByVal ms As Long)\n\
         Private Declare PtrSafe Sub Absent Lib \"kernel32\" ()\n\
         Sub Main()\n\
         Debug.Print \"entered Main\"\n\
         End Sub\n",
    )
    .unwrap();
    let results = ["vm3", "jit"].map(|backend| {
        ["strict", "windows-faithful"].map(|mode| {
            Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
                .args([
                    "run",
                    "--backend",
                    backend,
                    "--profile",
                    "linux-stdio",
                    "--portability-mode",
                    mode,
                    "--diagnostic-format",
                    "json",
                ])
                .arg(&path)
                .output()
                .unwrap()
        })
    });
    fs::remove_file(path).unwrap();
    for (index, output) in results.into_iter().flatten().enumerate() {
        assert!(!output.status.success());
        assert!(!String::from_utf8_lossy(&output.stdout).contains("entered Main"));
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(!stderr.contains("entered Main"));
        for expected in [
            "PORT-E-UNSUPPORTED",
            "Declare",
            "kernel32",
            if index % 2 == 0 {
                "\"line\": 1"
            } else {
                "\"line\": 2"
            },
            "linux",
        ] {
            assert!(
                stderr.contains(expected),
                "{expected:?} missing from {stderr}"
            );
        }
    }
}

#[test]
#[cfg(target_os = "linux")]
fn faithful_mode_executes_supported_kernel32_declares() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/portable-declare.bas"
    );
    for backend in ["vm3", "jit"] {
        let output = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
            .args([
                "run",
                "--backend",
                backend,
                "--profile",
                "linux-stdio",
                "--portability-mode",
                "windows-faithful",
                path,
            ])
            .output()
            .unwrap();
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(output.status.success(), "{backend}: {stderr}");
        assert_eq!(
            stderr.trim(),
            "True|True|True|True|3|0|41|41|42",
            "{backend}"
        );
    }
}

#[test]
#[cfg(target_os = "linux")]
fn faithful_mode_refuses_known_entry_with_wrong_signature() {
    let path = std::env::temp_dir().join(format!(
        "calcvba-portability-abi-{}.bas",
        std::process::id()
    ));
    fs::write(
        &path,
        "Private Declare PtrSafe Function Sleep Lib \"kernel32\" (ByVal ms As Double) As Long\nSub Main()\nEnd Sub\n",
    )
    .unwrap();
    let output = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
        .args([
            "run",
            "--profile",
            "linux-stdio",
            "--portability-mode",
            "windows-faithful",
        ])
        .arg(&path)
        .output()
        .unwrap();
    fs::remove_file(path).unwrap();
    assert!(!output.status.success(), "{output:?}");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(
        stderr.contains("PORT-E-UNSUPPORTED") && stderr.contains("kernel32!Sleep"),
        "{stderr}"
    );
}

#[test]
fn strict_reports_host_values_and_native_preserves_source_execution() {
    let path = std::env::temp_dir().join(format!(
        "calcvba-portability-host-{}.bas",
        std::process::id()
    ));
    fs::write(
        &path,
        "Sub Main()\nDim value As Date\nvalue = Now\nEnd Sub\n",
    )
    .unwrap();
    let strict = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
        .args([
            "run",
            "--profile",
            "linux-stdio",
            "--portability-mode",
            "strict",
            "--dump-values",
        ])
        .arg(&path)
        .output()
        .unwrap();
    fs::write(&path, "Private Declare PtrSafe Sub Sleep Lib \"kernel32\" (ByVal ms As Long)\nSub Main()\nEnd Sub\n").unwrap();
    let native = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
        .args([
            "run",
            "--profile",
            "linux-stdio",
            "--portability-mode",
            "native",
            "--dump-values",
        ])
        .arg(&path)
        .output()
        .unwrap();
    fs::remove_file(path).unwrap();

    assert!(strict.status.success(), "{strict:?}");
    let warning = String::from_utf8_lossy(&strict.stderr);
    assert!(
        warning.contains("PORT-W-HOST-DEPENDENT") && warning.contains("Now"),
        "{warning}"
    );
    assert!(native.status.success(), "{native:?}");
}

#[test]
fn faithful_mode_allows_empty_path_getobject_for_the_portable_dictionary() {
    let path = std::env::temp_dir().join(format!(
        "calcvba-portability-dict-{}.bas",
        std::process::id()
    ));
    fs::write(
        &path,
        r#"
Sub Main()
Dim value As Object
Set value = GetObject("", "Scripting.Dictionary")
Debug.Print TypeName(value)
End Sub
"#,
    )
    .unwrap();
    let output = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
        .args([
            "run",
            "--profile",
            "linux-stdio",
            "--portability-mode",
            "windows-faithful",
        ])
        .arg(&path)
        .output()
        .unwrap();
    fs::remove_file(path).unwrap();
    let text = String::from_utf8_lossy(&output.stderr);
    assert!(output.status.success(), "{text}");
    assert_eq!(text.trim(), "Dictionary");
}
