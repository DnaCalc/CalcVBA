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
    let outputs = ["vm3", "jit"].map(|backend| run_mapped_path_probe(&source, backend));
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

#[test]
fn mapped_read_only_file_can_change_attributes_and_clear_read_only() {
    let root = std::env::temp_dir().join(format!("calcvba-attr-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    let entry = root.join("entry.bas");
    fs::write(
        &entry,
        r#"
Sub Main()
    Open "C:\Notes.txt" For Output As #1
    Close #1
    SetAttr "C:\notes.txt", vbReadOnly
    SetAttr "C:\NOTES.txt", vbReadOnly + vbHidden
    Debug.Print GetAttr("C:\Notes.txt")
    SetAttr "C:\Notes.txt", vbNormal
    Debug.Print GetAttr("C:\Notes.txt")
    Kill "C:\Notes.txt"
End Sub
"#,
    )
    .unwrap();
    let runs = ["vm3", "jit"].map(|backend| run_mapped_path_probe(&entry, backend));
    fs::remove_dir_all(root).unwrap();
    for output in runs {
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(output.status.success(), "{stderr}");
        assert_eq!(stderr.split_whitespace().collect::<Vec<_>>(), ["3", "0"]);
    }
}

#[test]
fn mapped_filelen_uses_windows_directory_size_and_missing_parent_errors() {
    let root = std::env::temp_dir().join(format!("calcvba-dirlen-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    fs::create_dir(root.join("Data")).unwrap();
    let entry = root.join("entry.bas");
    fs::write(
        &entry,
        r#"
Sub Main()
    Debug.Print FileLen("C:\DATA")
    On Error Resume Next
    Debug.Print FileLen("C:\missing\Notes.txt")
    Debug.Print Err.Number
    Err.Clear
    Debug.Print GetAttr("C:\missing\Notes.txt")
    Debug.Print Err.Number
End Sub
"#,
    )
    .unwrap();
    let runs = ["vm3", "jit"].map(|backend| run_mapped_path_probe(&entry, backend));
    fs::remove_dir_all(root).unwrap();
    for output in runs {
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(output.status.success(), "{stderr}");
        assert_eq!(
            stderr.split_whitespace().collect::<Vec<_>>(),
            ["0", "76", "76"]
        );
    }
}

#[test]
fn mapped_name_preserves_the_requested_case() {
    let root = std::env::temp_dir().join(format!("calcvba-case-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    let entry = root.join("entry.bas");
    fs::write(
        &entry,
        r#"
Sub Main()
    Open "C:\Notes.txt" For Output As #1
    Print #1, "notes"
    Close #1
    SetAttr "C:\Notes.txt", vbHidden + vbSystem + vbArchive
    Name "c:\NOTES.txt" As "C:\notes.TXT"
    Debug.Print Dir$("C:\*.txt", vbHidden + vbSystem), GetAttr("C:\NOTES.txt")
    SetAttr "C:\notes.TXT", vbNormal
    Open "C:\other.bin" For Output As #1
    Close #1
    On Error Resume Next
    Name "C:\notes.TXT" As "C:\OTHER.BIN"
    Debug.Print Err.Number
    On Error GoTo 0
    Kill "C:\other.bin"
    Kill "C:\notes.TXT"
End Sub
"#,
    )
    .unwrap();
    let runs = ["vm3", "jit"].map(|backend| run_mapped_path_probe(&entry, backend));
    fs::remove_dir_all(root).unwrap();
    for output in runs {
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(output.status.success(), "{stderr}");
        assert_eq!(
            stderr.split_whitespace().collect::<Vec<_>>(),
            ["notes.TXT", "38", "58"]
        );
    }
}

fn run_mapped_path_probe(path: &std::path::Path, backend: &str) -> std::process::Output {
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
        .arg(path)
        .output()
        .unwrap()
}
