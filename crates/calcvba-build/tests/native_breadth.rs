#![cfg(all(windows, target_arch = "x86_64"))]

use calcvba_build::{NativeLibraryBuildOptions, build_native_library};
use calcvba_engine::{Engine, HostConfig};
use std::{
    path::PathBuf,
    sync::atomic::{AtomicUsize, Ordering},
};

struct Directory(PathBuf);
impl Directory {
    fn new() -> Self {
        static NEXT: AtomicUsize = AtomicUsize::new(0);
        let path = std::env::temp_dir().join(format!(
            "calcvba-native-breadth-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        std::fs::create_dir_all(&path).unwrap();
        Self(path)
    }
}
impl Drop for Directory {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}
fn evaluate(config: HostConfig, source: &str) -> String {
    let mut engine = Engine::new(config);
    engine.set_host_policy(calcvba_hal::HostPolicy::native());
    let values = engine
        .execute_source_with_variant_snapshot_clean(source)
        .unwrap();
    values
        .first()
        .and_then(|value| value.as_bstr())
        .expect("observed String local")
        .to_string()
}

#[test]
fn conformance_bodies_agree_in_vm3_jit_and_aot() {
    let directory = Directory::new();
    let source = include_str!("../../../examples/native/breadth/Breadth.bas");
    // Captured by visible VBE Compile and execution in one owned Excel instance.
    let cases = [
        ("ActivationSuite", "9:True"),
        ("ForeignObjectSuite", "9:False|9:True"),
        (
            "ErrorBoundsSuite",
            "5:True:Invalid procedure call or argument|5:Invalid procedure call or argument|5:True:Invalid procedure call or argument|-1:Explicit:negative",
        ),
        (
            "ArrayAllocationSuite",
            "resize=10;1,2,3,after=0;replace=10;1,2,3,",
        ),
        ("EraseSuite", "1:10,2:0,3:0,|reset=0error=9|3:5"),
        ("StringSuite", "AΛ:xy:xy:3:cba"),
        (
            "VariantSuite",
            "2:9:4:0:8:True:x:null-branch:True:True:Error 2042:3:True",
        ),
        ("LibrarySuite", "a/b/c:2024-02-29:FF:xxx:2:-1:-2:-1"),
        ("ArraySuite", "103:2=2:8=x:11=True:early"),
        (
            "ErrorSuite",
            "11|True|0|513|Native.Source|detail|help.chm|17|514|detail|0",
        ),
        ("ResumeSuite", "11:10|6:20|13:30|0|5"),
        ("NestedErrorSuite", "519|Nested.Source|nested fault|0"),
        ("MixedSuite", "True|True|True|True|aXYde|0|err=13"),
    ];
    for (name, expected) in cases {
        let (setup, argument) = if name == "ForeignObjectSuite" {
            (
                "Dim foreign As Object\nSet foreign=CreateObject(\"Scripting.Dictionary\")\n",
                "foreign",
            )
        } else {
            ("", "")
        };
        let direct = format!(
            "{source}\nPublic Sub Main()\nDim observed As String\n{setup}observed={name}({argument})\nEnd Sub\n"
        );
        for backend in [HostConfig::vm3(), HostConfig::jit()] {
            assert_eq!(evaluate(backend, &direct), expected, "direct {name}");
        }
    }
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/native/breadth/NativeBreadth.basproj"),
        out_dir: directory.0.clone(),
        template_path: None,
    })
    .unwrap();
    for (name, expected) in cases {
        let (parameters, setup, argument) = if name == "ForeignObjectSuite" {
            (
                "ByVal foreign As LongPtr",
                "Dim foreign As Object\nSet foreign=CreateObject(\"Scripting.Dictionary\")",
                "ObjPtr(foreign)",
            )
        } else {
            ("", "", "")
        };
        let caller = format!(
            r#"
Private Declare PtrSafe Function NativeBody Lib "{}" Alias "{name}" ({parameters}) As LongPtr
Private Declare PtrSafe Function Shutdown Lib "{}" Alias "CalcVbaShutdown" () As Long
Private Declare PtrSafe Function LastError Lib "{}" Alias "CalcVbaGetLastError" (ByRef source As LongPtr, ByRef description As LongPtr) As Long
Private Declare PtrSafe Function BStrLength Lib "oleaut32" Alias "SysStringLen" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Sub FreeBStr Lib "oleaut32" Alias "SysFreeString" (ByVal value As LongPtr)
Private Declare PtrSafe Sub CopyBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal count As LongPtr)
Private Function TakeString(ByVal pointer As LongPtr) As String
    Dim length As Long
    length = BStrLength(pointer)
    TakeString = String$(length, ChrW$(0))
    If length Then CopyBytes StrPtr(TakeString), pointer, length * 2&
    FreeBStr pointer
End Function
Public Sub Main()
    Dim observed As String
    Dim code As Long, source As LongPtr, description As LongPtr
    {setup}
    observed = TakeString(NativeBody({argument}))
    code = LastError(source, description)
    If code <> 0 Then observed = "ERROR " & CStr(code) & ":" & TakeString(source) & ":" & TakeString(description)
    If Shutdown() <> 0 Then Err.Raise 5, , "native shutdown failed"
End Sub
"#,
            output.dll_path.display(),
            output.dll_path.display(),
            output.dll_path.display()
        );
        for backend in [HostConfig::vm3(), HostConfig::jit()] {
            assert_eq!(evaluate(backend, &caller), expected, "AOT {name}");
        }
    }
}

#[test]
fn unhandled_export_faults_do_not_supply_a_later_exports_error_fields() {
    let directory = Directory::new();
    let source = r#"
Public Function FirstError() As Long
    Err.Raise 6, "Sticky.Source", "first failure"
End Function
Public Function SecondError() As Double
    Dim number As Double
    number=1E+100
    SecondError=CLng(number)
End Function
"#;
    let expected = "6:Sticky.Source:first failure|6:True:Overflow";
    let direct = format!(
        r#"{source}
Public Sub Main()
    Dim observed As String, discarded As Double
    On Error Resume Next
    discarded=FirstError()
    observed=CStr(Err.Number) & ":" & Err.Source & ":" & Err.Description
    discarded=SecondError()
    observed=observed & "|" & CStr(Err.Number) & ":" & CStr(Err.Source<>"Sticky.Source") & ":" & Err.Description
End Sub
"#
    );
    for backend in [HostConfig::vm3(), HostConfig::jit()] {
        assert_eq!(evaluate(backend, &direct), expected);
    }
    std::fs::write(directory.0.join("Code.bas"), source).unwrap();
    let project = directory.0.join("FaultScope.basproj");
    std::fs::write(&project, r#"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>FaultScope</ProjectName></PropertyGroup><ItemGroup><Module Include="Code.bas"/><NativeExport Include="FirstError"><Module>Code</Module><Procedure>FirstError</Procedure></NativeExport><NativeExport Include="SecondError"><Module>Code</Module><Procedure>SecondError</Procedure></NativeExport></ItemGroup></Project>"#).unwrap();
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: directory.0.join("out"),
        template_path: None,
    })
    .unwrap();
    let library = output.dll_path.display();
    let caller = format!(
        r#"
Private Declare PtrSafe Function FirstError Lib "{library}" () As Long
Private Declare PtrSafe Function SecondError Lib "{library}" () As Double
Private Declare PtrSafe Function LastError Lib "{library}" Alias "CalcVbaGetLastError" (ByRef source As LongPtr, ByRef description As LongPtr) As Long
Private Declare PtrSafe Function Shutdown Lib "{library}" Alias "CalcVbaShutdown" () As Long
Private Declare PtrSafe Function BStrLength Lib "oleaut32" Alias "SysStringLen" (ByVal value As LongPtr) As Long
Private Declare PtrSafe Sub FreeBStr Lib "oleaut32" Alias "SysFreeString" (ByVal value As LongPtr)
Private Declare PtrSafe Sub CopyBytes Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal count As LongPtr)
Private Function TakeString(ByVal pointer As LongPtr) As String
    Dim length As Long
    length=BStrLength(pointer)
    TakeString=String$(length, ChrW$(0))
    If length Then CopyBytes StrPtr(TakeString), pointer, length*2&
    FreeBStr pointer
End Function
Public Sub Main()
    Dim observed As String, discarded As Double
    Dim number As Long, source As LongPtr, description As LongPtr, text As String
    discarded=FirstError()
    number=LastError(source, description)
    observed=CStr(number) & ":" & TakeString(source) & ":" & TakeString(description)
    discarded=SecondError()
    number=LastError(source, description)
    text=TakeString(source)
    observed=observed & "|" & CStr(number) & ":" & CStr(text<>"Sticky.Source") & ":" & TakeString(description)
    If Shutdown()<>0 Then Err.Raise 5
End Sub
"#
    );
    for backend in [HostConfig::vm3(), HostConfig::jit()] {
        assert_eq!(evaluate(backend, &caller), expected);
    }
}
