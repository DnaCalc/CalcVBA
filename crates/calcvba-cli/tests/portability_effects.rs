use std::{fs, path::Path, process::Command};

fn inspect(path: &Path) -> serde_json::Value {
    let output = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
        .args(["check", "--portability", "--host", "linux"])
        .arg(path)
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    serde_json::from_slice(&output.stdout).unwrap()
}

#[test]
fn bound_effects_follow_calls_and_intersect_branches() {
    let root = std::env::temp_dir().join(format!(
        "calcvba-portability-effects-{}",
        std::process::id()
    ));
    fs::create_dir_all(&root).unwrap();
    let source = root.join("Probe.bas");
    fs::write(
        &source,
        r#"Public Sub Main()
End Sub
Public Sub AlwaysMail()
    Mail
End Sub
Public Sub MaybeMail(ByVal enabled As Boolean)
    If enabled Then Mail
End Sub
Public Sub EitherBranch(ByVal enabled As Boolean)
    If enabled Then
        Mail
    Else
        AlwaysMail
    End If
End Sub
Private Sub Mail()
    Dim app As Object
    Set app = CreateObject("Outlook.Application")
End Sub
Public Sub Computed(ByVal className As String)
    Dim app As Object
    Set app = CreateObject(className)
End Sub
"#,
    )
    .unwrap();
    let project = root.join("Probe.basproj");
    fs::write(
        &project,
        r#"<Project Sdk="CalcVBA.Sdk/0.1.0">
<PropertyGroup><ProjectName>Probe</ProjectName><OutputType>Exe</OutputType></PropertyGroup>
<ItemGroup><Module Include="Probe.bas" /></ItemGroup>
</Project>"#,
    )
    .unwrap();
    let output_dir = root.join("out");
    let output = Command::new(env!("CARGO_BIN_EXE_calcvba-cli"))
        .arg("build")
        .arg(&project)
        .args(["--target", "OxImage", "--out-dir"])
        .arg(&output_dir)
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    let reports = [inspect(&source), inspect(&output_dir.join("Probe.oxi"))];
    fs::remove_dir_all(&root).unwrap();

    let source_rows = reports[0]["procedures"].as_array().unwrap();
    assert_eq!(source_rows.len(), 6, "{source_rows:?}");
    let source_module = source.display().to_string();
    assert!(source_rows.iter().all(|row| row["module"] == source_module));

    for report in reports {
        assert_eq!(report["effects_complete"], true, "{report}");
        let procedures = report["procedures"].as_array().unwrap();
        let find = |name| {
            procedures
                .iter()
                .find(|row| row["procedure"] == name)
                .unwrap()
        };
        assert_eq!(find("Main")["status"], "runs");
        for name in ["AlwaysMail", "EitherBranch", "Mail"] {
            let row = find(name);
            assert_eq!(row["status"], "cannot_run", "{row}");
            assert_eq!(row["must"][0]["detail"], "Outlook.Application");
            assert_eq!(row["must"][0]["verdict"], "reject");
        }
        let optional = find("MaybeMail");
        assert_eq!(optional["status"], "may_stop", "{optional}");
        assert!(optional["must"].as_array().unwrap().is_empty());
        assert_eq!(optional["may"][0]["detail"], "Outlook.Application");
        let computed = find("Computed");
        assert_eq!(computed["status"], "may_stop", "{computed}");
        assert_eq!(computed["unknown"], true);
        assert!(computed["must"].as_array().unwrap().is_empty());
    }
}

#[test]
fn bound_effects_follow_local_strings_and_gosub() {
    let root = std::env::temp_dir().join(format!(
        "calcvba-portability-strings-{}",
        std::process::id()
    ));
    fs::create_dir_all(&root).unwrap();
    let source = root.join("Probe.bas");
    fs::write(
        &source,
        r#"Option Explicit
Public Sub SameClass(ByVal enabled As Boolean)
    Dim className As String
    Dim app As Object
    If enabled Then
        className = "Outlook." & "Application"
    Else
        className = "Outlook.Application"
    End If
    Set app = CreateObject(className)
End Sub
Public Sub DifferentClasses(ByVal enabled As Boolean)
    Dim className As String
    Dim app As Object
    className = "Scripting.Dictionary"
    If enabled Then className = "Outlook.Application"
    Set app = CreateObject(className)
End Sub
Public Sub Overwritten()
    Dim className As String
    Dim app As Object
    className = "Outlook.Application"
    Mutate className
    Set app = CreateObject(className)
End Sub
Private Sub Mutate(ByRef className As String)
    className = "Scripting.Dictionary"
End Sub
Public Sub DynamicServer(ByVal serverName As String)
    Dim target As Object
    Set target = CreateObject("Scripting.FileSystemObject", serverName)
End Sub
Public Sub DynamicPath(ByVal fileName As String)
    Dim target As Object
    Set target = GetObject(fileName, "Scripting.FileSystemObject")
End Sub
Public Function TempSize() As Long
    Dim path As String
    path = Environ$("TEMP") & "\data.bin"
    TempSize = FileLen(path)
End Function
Public Sub UnreachableWithGoSub()
    GoSub LocalWork
    Exit Sub
Skipped:
    Dim app As Object
    Set app = CreateObject("Outlook.Application")
    Exit Sub
LocalWork:
    Return
End Sub
"#,
    )
    .unwrap();
    let report = inspect(&source);
    fs::remove_dir_all(root).unwrap();
    assert_eq!(report["effects_complete"], true, "{report}");
    let procedures = report["procedures"].as_array().unwrap();
    let find = |name| {
        procedures
            .iter()
            .find(|row| row["procedure"] == name)
            .unwrap()
    };
    let same = find("SameClass");
    assert_eq!(same["status"], "cannot_run", "{same}");
    assert_eq!(same["must"][0]["detail"], "Outlook.Application");
    for name in [
        "DifferentClasses",
        "Overwritten",
        "DynamicServer",
        "DynamicPath",
    ] {
        let row = find(name);
        assert_eq!(row["status"], "may_stop", "{row}");
        assert_eq!(row["unknown"], true, "{row}");
    }
    let path = find("TempSize");
    assert!(
        path["may"].as_array().unwrap().iter().any(|effect| {
            effect["feature"] == "FileLen" && effect["detail"] == r#"Env(TEMP)\data.bin"#
        }),
        "{path}"
    );
    let gosub = find("UnreachableWithGoSub");
    assert_eq!(gosub["status"], "runs", "{gosub}");
}
