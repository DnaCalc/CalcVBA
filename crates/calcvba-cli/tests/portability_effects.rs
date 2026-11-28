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
