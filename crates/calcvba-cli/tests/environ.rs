#![cfg(target_os = "linux")]
use std::{fs, process::Command};

#[test]
fn environ_uses_host_order_and_windows_names_in_faithful_mode() {
    let root = std::env::temp_dir().join(format!("calcvba-environ-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    let temp = root.join("Tmp");
    let config = root.join("Config");
    fs::create_dir(&temp).unwrap();
    fs::create_dir(&config).unwrap();
    let source = root.join("main.bas");
    fs::write(
        &source,
        r#"
Sub Main()
Debug.Print Environ$(1)
Debug.Print Environ$("zFIRST")
Debug.Print Environ$("temp")
Debug.Print Environ$("tmp")
Debug.Print Environ$("userprofile")
Debug.Print Environ$("appdata")
Debug.Print Environ$("computername")
Debug.Print Environ$("username")
End Sub
"#,
    )
    .unwrap();
    let results = ["vm3", "jit"].map(|backend| {
        // `env` provides a known insertion order for the child environment.
        Command::new("env")
            .arg("-i")
            .args([
                "Zfirst=one",
                "Asecond=two",
                "USER=probeuser",
                "HOSTNAME=probehost",
            ])
            .arg(format!("TMPDIR={}", temp.display()))
            .arg(format!("HOME={}", root.display()))
            .arg(format!("XDG_CONFIG_HOME={}", config.display()))
            .arg(env!("CARGO_BIN_EXE_calcvba-cli"))
            .args([
                "run",
                "--backend",
                backend,
                "--profile",
                "linux-stdio",
                "--portability-mode",
                "windows-faithful",
                "--pass-env",
                "Zfirst",
                "--pass-env",
                "Asecond",
            ])
            .arg(&source)
            .current_dir(&root)
            .output()
            .unwrap()
    });
    fs::remove_dir_all(root).unwrap();
    for output in results {
        let text = String::from_utf8_lossy(&output.stderr);
        assert!(output.status.success(), "{text}");
        assert_eq!(
            text.lines().map(str::trim).collect::<Vec<_>>(),
            [
                "Zfirst=one",
                "one",
                r"C:\Tmp",
                r"C:\Tmp",
                "C:\\",
                r"C:\Config",
                "probehost",
                "probeuser"
            ]
        );
    }
}

#[test]
fn environ_handles_integer_overflow_non_ascii_folding_and_outside_root_aliases() {
    let root = std::env::temp_dir().join(format!("calcvba-environ-edge-{}", std::process::id()));
    let _ = fs::remove_dir_all(&root);
    fs::create_dir_all(&root).unwrap();
    let source = root.join("main.bas");
    fs::write(
        &source,
        r#"
Function ErrCode(ByVal idx As Variant) As Long
    On Error Resume Next
    Dim s As String
    s = Environ$(idx)
    ErrCode = Err.Number
    Err.Clear
End Function

Sub Main()
    Dim tempDir As String, probePath As String, back As String
    tempDir = Environ$("TEMP")
    probePath = tempDir & "\probe.txt"
    Open probePath For Output As #1
    Print #1, "ok"
    Close #1
    Open probePath For Input As #1
    Line Input #1, back
    Close #1

    Debug.Print CStr(ErrCode(40000)) & "|" & CStr(ErrCode(-40000)) & "|" & CStr(ErrCode(32000)) & "|" & CStr(ErrCode(0)) & "|" & CStr(ErrCode(256))
    Debug.Print Environ$("café_var") & "|" & Environ$("STRAßE") & "|" & Environ$("STRASSE")
    Debug.Print tempDir & "|" & Environ$("TMP") & "|" & Environ$("USERPROFILE") & "|" & Environ$("APPDATA") & "|" & back
End Sub
"#,
    )
    .unwrap();

    let results = ["vm3", "jit"].map(|backend| {
        Command::new("env")
            .arg("-i")
            .args([
                "CAFÉ_VAR=latte",
                "Straße=road",
                "USER=probeuser",
                "HOSTNAME=probehost",
                // Deliberately outside `root` (the mapped C: drive root).
                "TMPDIR=/tmp",
                "HOME=/outside-home-dir",
                "XDG_CONFIG_HOME=/outside-config-dir",
            ])
            .arg(env!("CARGO_BIN_EXE_calcvba-cli"))
            .args([
                "run",
                "--backend",
                backend,
                "--profile",
                "linux-stdio",
                "--portability-mode",
                "windows-faithful",
                "--pass-env",
                "CAFÉ_VAR",
                "--pass-env",
                "Straße",
            ])
            .arg(&source)
            .current_dir(&root)
            .output()
            .unwrap()
    });
    fs::remove_dir_all(&root).unwrap();

    for (backend, output) in ["vm3", "jit"].into_iter().zip(results) {
        let text = String::from_utf8_lossy(&output.stderr);
        assert!(output.status.success(), "{backend}: {text}");
        let lines: Vec<_> = text.lines().map(str::trim).collect();
        assert_eq!(
            lines,
            ["6|6|5|5|5", "latte|road|", r"C:\|C:\|C:\|C:\|ok"],
            "{backend}"
        );
    }
}

#[test]
fn faithful_environment_synthesizes_windows_names_and_requires_pass_through() {
    let root = std::env::temp_dir().join(format!("calcvba-environ-persona-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    let source = root.join("main.bas");
    fs::write(
        &source,
        r#"Public Sub Main()
    Debug.Print "[" & Environ$("px1_private_marker") & "]|" & Environ$("OS") & "|" & Environ$("PROCESSOR_ARCHITECTURE")
End Sub
"#,
    )
    .unwrap();

    for backend in ["vm3", "jit"] {
        for explicit in [false, true] {
            let mut command = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"));
            command
                .env_clear()
                .env("PX1_PRIVATE_MARKER", "fixture-marker");
            command.args([
                "run",
                "--backend",
                backend,
                "--profile",
                "linux-stdio",
                "--portability-mode",
                "windows-faithful",
            ]);
            if explicit {
                command.args(["--pass-env", "PX1_PRIVATE_MARKER"]);
            }
            let output = command.arg(&source).output().unwrap();
            let stderr = String::from_utf8_lossy(&output.stderr);
            assert!(output.status.success(), "{backend}: {stderr}");
            let value = if explicit { "fixture-marker" } else { "" };
            assert_eq!(
                stderr.trim(),
                format!("[{value}]|Windows_NT|AMD64"),
                "{backend}"
            );
        }
    }
    fs::remove_dir_all(root).unwrap();
}
