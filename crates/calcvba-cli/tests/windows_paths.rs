#![cfg(target_os = "linux")]
use std::{fs, process::Command};

#[test]
fn windows_paths_resolve_per_project_and_preserve_disk_case() {
    let root = std::env::temp_dir().join(format!("calcvba-paths-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    let data = root.join("Data");
    fs::create_dir(&data).unwrap();
    fs::write(data.join("Notes.txt"), b"notes").unwrap();
    let source = root.join("entry.bas");
    fs::write(
        &source,
        r#"
Sub Main()
Dim size As Long, folder As String, listing As String, missing As Long
size = FileLen("c:\DATA\notes.TXT")
ChDir "C:\dAta"
folder = CurDir$
listing = Dir$("*.TXT")
On Error Resume Next
ChDrive "Q:"
missing = Err.Number
Debug.Print size, folder, listing, missing
End Sub
"#,
    )
    .unwrap();
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
            .arg(&source)
            .current_dir(&root)
            .output()
            .unwrap()
    });
    fs::remove_dir_all(root).unwrap();
    for output in outputs {
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(output.status.success(), "{stderr}");
        assert_eq!(
            stderr.split_whitespace().collect::<Vec<_>>(),
            ["5", "C:\\Data", "Notes.txt", "76"]
        );
    }
}
