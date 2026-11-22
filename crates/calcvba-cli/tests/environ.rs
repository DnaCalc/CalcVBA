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
