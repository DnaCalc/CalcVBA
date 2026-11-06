use std::{
    fs,
    path::PathBuf,
    process::Command,
    sync::atomic::{AtomicUsize, Ordering},
};
static NEXT: AtomicUsize = AtomicUsize::new(0);
struct Directory(PathBuf);
impl Directory {
    fn new() -> Self {
        let path = std::env::temp_dir().join(format!(
            "calcvba-udf-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        fs::create_dir(&path).unwrap();
        assert_eq!(
            path.canonicalize().unwrap().parent(),
            Some(std::env::temp_dir().canonicalize().unwrap().as_path())
        );
        Self(path)
    }
    fn project(&self, tag: &str, source: &str) -> PathBuf {
        fs::write(self.0.join("Functions.bas"), source).unwrap();
        let path = self.0.join("Input.basproj");
        fs::write(&path,format!("<Project><PropertyGroup><ProjectName>Input</ProjectName><OutputType>Library</OutputType></PropertyGroup><ItemGroup><{tag} Include=\"Functions.bas\"/></ItemGroup></Project>")).unwrap();
        path
    }
}
impl Drop for Directory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}
#[test]
fn bound_defaults_and_paramarray_are_emitted_as_ordinary_vba() {
    let directory = Directory::new();
    let project=directory.project("Module","Attribute VB_Name = \"Functions\"\nPrivate Const DefaultNumber As Long=2+5\nPublic Function Value(Optional ByVal number As Long=DefaultNumber,Optional item As Variant) As String\nValue=CStr(number) & CStr(IsMissing(item))\nEnd Function\nPublic Function Sum(ByVal first As Double,ParamArray items() As Variant) As Double\nDim index As Long\nSum=first\nFor index=LBound(items) To UBound(items)\nSum=Sum+CDbl(items(index))\nNext index\nEnd Function\n");
    let out = directory.0.join("generated");
    let source_path = directory.0.join("Functions.bas");
    let mut source = fs::read_to_string(&source_path).unwrap();
    source.push_str("Public Function Defaults(Optional ByVal text As String=vbCrLf, Optional ByVal amount As Currency=123.4567@) As Long\nDefaults=Len(text)\nEnd Function\n");
    fs::write(source_path, source).unwrap();
    let result = Command::new(env!("CARGO_BIN_EXE_calcvba-udf"))
        .args([
            "--project",
            project.to_str().unwrap(),
            "--out",
            out.to_str().unwrap(),
            "--emit-only",
            "--paramarray-slots",
            "3",
        ])
        .output()
        .unwrap();
    assert!(
        result.status.success(),
        "{}",
        String::from_utf8_lossy(&result.stderr)
    );
    let shim = fs::read_to_string(out.join("UdfRuntime.bas")).unwrap();
    assert!(shim.contains("a0=7&"), "folded default missing");
    assert!(shim.contains("a0=ChrW$(13&) & ChrW$(10&)"));
    assert!(shim.contains("a1=UdfCurrency(1234567^)"));
    assert!(shim.contains("result=Functions.Sum(a0,a1,a2,a3)"));
    assert!(shim.contains("Case 0\nresult=Functions.Sum(a0)"));
    let table = fs::read_to_string(out.join("registration.tsv")).unwrap();
    assert!(table.contains("Value\tQQU\tnumber,item\tFunctions.Value"));
    let project = fs::read_to_string(out.join("Udf.basproj")).unwrap();
    assert!(project.contains("<NativeExport Include=\"Udf1\">"));
}
#[test]
fn unsupported_host_modules_and_events_leave_no_output() {
    for (tag, source) in [
        (
            "DocumentModule",
            "Public Function Value() As Long\nEnd Function\n",
        ),
        (
            "FormModule",
            "Public Function Value() As Long\nEnd Function\n",
        ),
        ("ClassModule", "Public Event Changed()\n"),
    ] {
        let directory = Directory::new();
        let project = directory.project(tag, source);
        let out = directory.0.join("generated");
        let result = Command::new(env!("CARGO_BIN_EXE_calcvba-udf"))
            .args([
                "--project",
                project.to_str().unwrap(),
                "--out",
                out.to_str().unwrap(),
                "--emit-only",
            ])
            .output()
            .unwrap();
        assert!(!result.status.success());
        let message = String::from_utf8_lossy(&result.stderr);
        assert!(message.contains("cannot be wrapped"), "{message}");
        assert!(!out.exists());
    }
}
