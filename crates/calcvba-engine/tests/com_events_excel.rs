//! WithEvents on a COM source behaves as in VBA (design doc, step 8c), checked against
//! out-of-process Excel: the handler runs inside the call that fires the event, a ByVal object
//! argument is the same object the call returns, a ByRef argument is the server's storage
//! (`Cancel = True` keeps the workbook open), an event fired inside a handler runs nested, and
//! `Set x = Nothing` stops delivery. An error a handler leaves unhandled ends the run even under
//! the caller's `On Error Resume Next`. Expected lines are Excel x64's in-process VBA results for
//! the same scenario (`temp/variant-layout-probe/events8`, captured 2026-11-15).
//!
//! Live COM (starts Excel) — every test is `#[ignore]`. Run explicitly:
//! ```text
//! cargo test -p calcvba-engine --test com_events_excel -- --ignored --test-threads=1
//! ```
#![cfg(target_os = "windows")]
#[path = "com_matrix_common.rs"]
mod common;
use common::*;

use calcvba_engine::{Engine, HostConfig};
use calcvba_hal::HostPolicy;
use calcvba_symbol::manifest as sym;

const MAIN: &str = "Public result As String\n\
Sub Main()\n\
    Dim s As New Sink\n\
    result = s.Run()\n\
End Sub\n";

/// Calls into Excel stay late bound (`As Object`); the WithEvents variable is typed, as VBA
/// requires.
const SINK: &str = "Private app As Object\n\
Private WithEvents appEv As Excel.Application\n\
Private mLog As String\n\
Private mCancel As Boolean\n\
Private mDepth As Long\n\
Private mLastWb As Object\n\
Private mRaise As Boolean\n\
\n\
Public Property Let RaiseInHandler(ByVal value As Boolean)\n\
    mRaise = value\n\
End Property\n\
\n\
Private Function IsOpen(ByVal wb As Object) As String\n\
    Dim n As String\n\
    On Error GoTo gone\n\
    n = wb.Name\n\
    IsOpen = \"open\"\n\
    Exit Function\n\
gone:\n\
    IsOpen = \"closed\"\n\
End Function\n\
\n\
Public Function Run() As String\n\
    Dim out As String, wb As Object\n\
    Set app = CreateObject(\"Excel.Application\")\n\
    app.DisplayAlerts = False\n\
    app.EnableEvents = True\n\
    Set appEv = app\n\
    mLog = \"before\"\n\
    Set wb = app.Workbooks.Add\n\
    mLog = mLog & \"after\"\n\
    out = out & \"SYNC \" & mLog & \" same_wb=\" & (mLastWb Is wb) & vbLf\n\
    mLog = \"\"\n\
    mCancel = True\n\
    wb.Close False\n\
    out = out & \"CANCEL \" & mLog & \" \" & IsOpen(wb) & vbLf\n\
    mCancel = False\n\
    mLog = \"\"\n\
    wb.Worksheets(1).Range(\"A1\").Value = 1\n\
    out = out & \"REENTRANT \" & mLog & vbLf\n\
    mLog = \"\"\n\
    wb.Close False\n\
    out = out & \"CLOSE \" & mLog & vbLf\n\
    mLog = \"\"\n\
    Set appEv = Nothing\n\
    Set wb = app.Workbooks.Add\n\
    out = out & \"UNHOOKED log=[\" & mLog & \"]\" & vbLf\n\
    wb.Close False\n\
    app.Quit\n\
    Run = out\n\
End Function\n\
\n\
Public Function RunError() As String\n\
    Dim wb As Object\n\
    Set app = CreateObject(\"Excel.Application\")\n\
    app.DisplayAlerts = False\n\
    app.EnableEvents = True\n\
    Set appEv = app\n\
    mRaise = True\n\
    On Error Resume Next\n\
    Set wb = app.Workbooks.Add\n\
    RunError = \"not ended: err=\" & Err.Number\n\
End Function\n\
\n\
Private Sub appEv_NewWorkbook(ByVal Wb As Excel.Workbook)\n\
    mLog = mLog & \"[new \" & (Wb Is app.ActiveWorkbook) & \"]\"\n\
    Set mLastWb = Wb\n\
    If mRaise Then app.Quit: Err.Raise 1234, \"Sink\", \"handler failed\"\n\
End Sub\n\
\n\
Private Sub appEv_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, Cancel As Boolean)\n\
    mLog = mLog & \"[close in=\" & Cancel & \"]\"\n\
    Cancel = mCancel\n\
End Sub\n\
\n\
Private Sub appEv_SheetChange(ByVal Sh As Object, ByVal Target As Range)\n\
    Dim t As Object, a As String\n\
    Set t = Target\n\
    a = t.Address(False, False)\n\
    mDepth = mDepth + 1\n\
    mLog = mLog & \"[change \" & a & \" d=\" & mDepth\n\
    If a = \"A1\" Then Sh.Range(\"B1\").Value = 2\n\
    mLog = mLog & \" end \" & a & \"]\"\n\
    mDepth = mDepth - 1\n\
End Sub\n";

const EXPECTED: &str = "SYNC before[new True]after same_wb=True\n\
CANCEL [close in=False] open\n\
REENTRANT [change A1 d=1[change B1 d=2 end B1] end A1]\n\
CLOSE [close in=False]\n\
UNHOOKED log=[]\n";

fn manifest(main: &str) -> sym::SymbolProjectManifest {
    sym::SymbolProjectManifest {
        project_name: "Main".to_string(),
        project_kind: sym::ProjectKind::Source,
        modules: [
            ("Main", sym::ModuleKind::Procedural, main.to_string()),
            ("Sink", sym::ModuleKind::Class, SINK.to_string()),
        ]
        .into_iter()
        .map(|(name, kind, source)| sym::ModuleUnit {
            module_name: name.to_string(),
            module_kind: kind,
            attributes: sym::ModuleAttributes::named(name),
            source,
        })
        .collect(),
        references: excel_ref(),
        reference_projects: Vec::new(),
        conditional_constants: std::collections::BTreeMap::new(),
        conditional_compilation_target: Default::default(),
    }
}

fn engine(config: HostConfig) -> Engine {
    let mut engine = Engine::new(config);
    engine.set_host_policy(HostPolicy::native());
    engine
}

#[test]
#[ignore = "live COM; starts Excel"]
fn withevents_on_excel_behave_as_in_vba_on_vm3_and_jit() {
    for (name, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let values = engine(config)
            .execute_manifest_with_variant_snapshot(&manifest(MAIN))
            .unwrap_or_else(|e| panic!("{name}: {e:?}"));
        let out = values
            .iter()
            .find_map(|v| {
                v.as_bstr()
                    .map(|s| s.to_string())
                    .filter(|s| s.contains("SYNC"))
            })
            .unwrap_or_else(|| panic!("{name}: no output in {values:?}"));
        assert_eq!(out, EXPECTED, "{name}");
    }
}

#[test]
#[ignore = "live COM; starts Excel"]
fn an_unhandled_handler_error_ends_the_run_on_vm3_and_jit() {
    let main = "Public result As String\n\
Sub Main()\n\
    Dim s As New Sink\n\
    On Error Resume Next\n\
    result = s.RunError()\n\
End Sub\n";
    for (name, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let error = engine(config)
            .execute_manifest_with_variant_snapshot(&manifest(main))
            .expect_err(name);
        // The manifest path reports an uncaught error by its text ("uncaught error 1234 ...",
        // "VBA error 1234").
        assert!(format!("{error:?}").contains("1234"), "{name}: {error:?}");
    }
}
