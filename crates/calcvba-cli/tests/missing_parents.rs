#![cfg(any(windows, target_os = "linux"))]
use std::{fs, process::Command};

#[test]
fn open_and_name_never_create_missing_parent_directories() {
    let root = std::env::temp_dir().join(format!("calcvba-parent-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    fs::write(root.join("blocking-file"), b"block").unwrap();
    let source = root.join("main.bas");
    fs::write(
        &source,
        r#"
Sub Main()
Dim outputError As Long, nameError As Long, inputError As Long, nestedError As Long
Dim nameBlockedTarget As Long, nameMissingSource As Long, nameMissingFile As Long, nameBlockedSource As Long
On Error Resume Next
Open "new-output/out.txt" For Output As #1
outputError = Err.Number
Close #1
Err.Clear
Name "original.txt" As "new-name/moved.txt"
nameError = Err.Number
Err.Clear
Open "new-input/in.txt" For Input As #2
inputError = Err.Number
Err.Clear
Open "blocking-file/child/out.txt" For Output As #3
nestedError = Err.Number
Err.Clear
Name "original.txt" As "blocking-file/moved.txt"
nameBlockedTarget = Err.Number
Err.Clear
Name "new-name/original.txt" As "moved.txt"
nameMissingSource = Err.Number
Err.Clear
Name "missing.txt" As "moved.txt"
nameMissingFile = Err.Number
Err.Clear
Name "blocking-file/original.txt" As "moved.txt"
nameBlockedSource = Err.Number
Debug.Print outputError, nameError, inputError, nestedError
Debug.Print nameBlockedTarget, nameMissingSource, nameMissingFile, nameBlockedSource
End Sub
"#,
    )
    .unwrap();
    let profile = if cfg!(windows) {
        "windows-stdio"
    } else {
        "linux-stdio"
    };
    let results = ["vm3", "jit"].map(|backend| {
        let original = root.join("original.txt");
        fs::write(&original, b"keep").unwrap();
        let output = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
            .args(["run", "--backend", backend, "--profile", profile])
            .arg(&source)
            .current_dir(&root)
            .output()
            .unwrap();
        let parents = ["new-output", "new-name", "new-input"].map(|directory| {
            let path = root.join(directory);
            let created = path.is_dir();
            if created {
                fs::remove_dir_all(path).unwrap();
            }
            created
        });
        (output, fs::read(original).ok(), parents)
    });
    fs::remove_dir_all(root).unwrap();

    for (output, original, parents) in results {
        let text = String::from_utf8_lossy(&output.stderr);
        assert!(output.status.success(), "{text}");
        assert_eq!(
            text.split_whitespace().collect::<Vec<_>>(),
            ["76", "53", "76", "76", "53", "53", "53", "53"]
        );
        assert_eq!(parents, [false; 3]);
        assert_eq!(original.as_deref(), Some(b"keep".as_slice()));
    }
}
