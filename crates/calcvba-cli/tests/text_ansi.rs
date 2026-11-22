#![cfg(target_os = "linux")]
use std::{fs, process::Command};

#[test]
fn sequential_vba_files_use_ansi_bytes_for_print_write_and_input() {
    let root = std::env::temp_dir().join(format!("calcvba-ansi-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    let source = root.join("main.bas");
    fs::write(
        &source,
        r#"
Sub Main()
Dim printed As String, written As String
Open "out.txt" For Output As #1
Print #1, "é€"
Write #1, "£€""é"
Close #1
Open "out.txt" For Input As #1
Line Input #1, printed
Input #1, written
Close #1
Debug.Print printed, written
End Sub
"#,
    )
    .unwrap();
    let results = ["vm3", "jit"].map(|backend| {
        let output = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
            .args(["run", "--backend", backend, "--profile", "linux-stdio"])
            .arg(&source)
            .current_dir(&root)
            .output()
            .unwrap();
        let bytes = fs::read(root.join("out.txt")).unwrap();
        (output, bytes)
    });
    fs::remove_dir_all(root).unwrap();
    for (output, bytes) in results {
        assert!(
            output.status.success(),
            "{}",
            String::from_utf8_lossy(&output.stderr)
        );
        assert_eq!(bytes, b"\xe9\x80\r\n\"\xa3\x80\"\"\xe9\"\r\n");
        assert_eq!(
            String::from_utf8_lossy(&output.stderr)
                .split_whitespace()
                .collect::<Vec<_>>(),
            ["é€", "£€\"é"]
        );
    }
}
