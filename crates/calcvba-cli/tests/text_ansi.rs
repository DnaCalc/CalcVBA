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
Dim printed As String, first As String, second As String
Open "out.txt" For Output As #1
Print #1, "é€"
Write #1, "£€""é"
Close #1
Open "out.txt" For Input As #1
Line Input #1, printed
Input #1, first, second
Close #1
Debug.Print printed, first, second
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
            ["é€", "£€", "é"]
        );
    }
}

#[test]
fn basproj_codepage_drives_text_files_and_ansi_library_functions_on_both_backends() {
    let root = std::env::temp_dir().join(format!("calcvba-cp1251-{}", std::process::id()));
    let _ = fs::remove_dir_all(&root);
    fs::create_dir_all(&root).unwrap();
    fs::write(
        root.join("Main.bas"),
        r#"
Public Sub Main()
    Dim printed As String, quoted As String, two As String, binRead As String
    Open "out.txt" For Output As #1
    Print #1, "Привет"
    Write #1, "Мир"
    Close #1

    Open "out.txt" For Input As #1
    Line Input #1, printed
    Input #1, quoted
    Close #1

    Open "out.txt" For Input As #1
    two = Input(2, #1)
    Close #1

    Open "bin.dat" For Binary As #2
    Dim binWrite As String
    binWrite = "Ж"
    Put #2, 1, binWrite
    binRead = Space$(1)
    Get #2, 1, binRead
    Close #2

    Dim rawBytes As String
    rawBytes = StrConv("Привет", vbFromUnicode)
    Debug.Print Chr$(208) & "|" & CStr(Asc("Я")) & "|" & String$(2, 208) & "|" & CStr(LenB(rawBytes)) & "|" & StrConv(rawBytes, vbUnicode) & "|" & printed & "|" & quoted & "|" & two & "|" & binRead
End Sub
"#,
    )
    .unwrap();
    let project = root.join("Cyrillic.basproj");
    fs::write(
        &project,
        r#"<Project Sdk="CalcVBA.Sdk/0.1.0">
<PropertyGroup>
  <ProjectName>Cyrillic</ProjectName>
  <OutputType>Exe</OutputType>
  <Codepage>1251</Codepage>
</PropertyGroup>
<ItemGroup><Module Include="Main.bas" /></ItemGroup>
</Project>"#,
    )
    .unwrap();

    for backend in ["vm3", "jit"] {
        let output = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
            .args([
                "run-project",
                "--backend",
                backend,
                "--profile",
                "linux-stdio",
            ])
            .arg(&project)
            .current_dir(&root)
            .output()
            .unwrap();
        let stderr = String::from_utf8_lossy(&output.stderr);
        assert!(output.status.success(), "{backend}: {stderr}");
        let out_bytes = fs::read(root.join("out.txt")).unwrap();
        // Windows-1251: "Привет" = CF F0 E8 E2 E5 F2, "Мир" = CC E8 F0
        assert_eq!(
            out_bytes, b"\xcf\xf0\xe8\xe2\xe5\xf2\r\n\"\xcc\xe8\xf0\"\r\n",
            "{backend}: out.txt bytes"
        );
        let bin_bytes = fs::read(root.join("bin.dat")).unwrap();
        // Windows-1251: "Ж" = C6
        assert_eq!(bin_bytes, b"\xc6", "{backend}: bin.dat bytes");
        assert_eq!(
            stderr.trim(),
            "Р|223|РР|6|Привет|Привет|Мир|Пр|Ж",
            "{backend}"
        );
    }

    // Unsupported code page is rejected at project load.
    let bad_project = root.join("BadCp.basproj");
    fs::write(
        &bad_project,
        r#"<Project Sdk="CalcVBA.Sdk/0.1.0">
<PropertyGroup>
  <ProjectName>BadCp</ProjectName>
  <OutputType>Exe</OutputType>
  <Codepage>65000</Codepage>
</PropertyGroup>
<ItemGroup><Module Include="Main.bas" /></ItemGroup>
</Project>"#,
    )
    .unwrap();
    let bad = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
        .args(["run-project", "--profile", "linux-stdio"])
        .arg(&bad_project)
        .current_dir(&root)
        .output()
        .unwrap();
    fs::remove_dir_all(&root).unwrap();
    assert!(!bad.status.success(), "expected invalid codepage rejection");
}
