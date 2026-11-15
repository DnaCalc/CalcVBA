//! Differential COM matrix — Category V: EVENTS (`WithEvents`, in & out of proc).
//!
//! `WithEvents` is only valid in a class module, so every scenario is a manifest
//! (a `Main` procedural module + a `Sink` class module) run on vm3 via [`run_manifest`].
//! Each event scenario fires the server's `Fire*` trigger; the handler runs inside that call, as
//! in VBA (design doc `runtime-variant-automation-layout-2026-11-13.md`, step 8c), and the
//! captured value(s) are read back into a `Main` global. `OxVba.TestEventServer` is the .NET
//! server in `tools/OxVba.TestEventServer` (`register.ps1` registers it for the current user).
//!
//! These are NOT three-leg differential value tests — an event delivery is a single
//! observable effect — so each asserts the captured vm3 effect directly (with the
//! occasional project-`RaiseEvent` twin as a same-language oracle for the same
//! arity). The connection-point path is exercised the same way the existing
//! `com_office_integration` event smoke test does.
//!
//! Live COM — every test is `#[ignore]`. Run explicitly:
//! ```text
//! cargo test -p calcvba-engine --test com_matrix_events -- --ignored --test-threads=1
//! ```
#![cfg(target_os = "windows")]
#[path = "com_matrix_common.rs"]
mod common;
use common::*;

use calcvba_symbol::manifest::ModuleKind;

// ── V1: OnSimpleEvent (0-arg), 2 fires ───────────────────────────────────────

#[test]
#[ignore = "live COM; run explicitly"]
fn v1_on_simple_event_multi_fire() {
    // 0-arg sink path. FireSimpleEvent x2 -> hits == 2.
    let sink = "Private mHits As Long\n\
         Private WithEvents src As OxVba.TestEventServer\n\
         Public Sub Wire()\n\
         Set src = New OxVba.TestEventServer\n\
         src.FireSimpleEvent\n\
         src.FireSimpleEvent\n\
         DoEvents\n\
         End Sub\n\
         Public Function Hits() As Long\n\
         Hits = mHits\n\
         End Function\n\
         Private Sub src_OnSimpleEvent()\n\
         mHits = mHits + 1\n\
         End Sub\n";
    run_event_case("V1", sink, 2);
}

// ── V2: OnValueChanged (1-arg), 2 fires ──────────────────────────────────────

#[test]
#[ignore = "live COM; run explicitly"]
fn v2_on_value_changed_ordered_lossless() {
    // Ordered lossless delivery. Fire 7 then 42 -> mGot == 42 AND mSum == 49.
    // verdict = IIf(mGot = 42 And mSum = 49, 1, 0).
    let sink = "Private mGot As Long\n\
         Private mSum As Long\n\
         Private WithEvents src As OxVba.TestEventServer\n\
         Public Sub Wire()\n\
         Set src = New OxVba.TestEventServer\n\
         src.FireValueChanged 7\n\
         src.FireValueChanged 42\n\
         DoEvents\n\
         End Sub\n\
         Public Function Result() As Long\n\
         Result = IIf(mGot = 42 And mSum = 49, 1, 0)\n\
         End Function\n\
         Private Sub src_OnValueChanged(ByVal value As Long)\n\
         mGot = value\n\
         mSum = mSum + value\n\
         End Sub\n";
    run_event_case_result("V2", sink, 1);
}

// ── V3 (P0): OnPairChanged (2-arg) arg-order pin ─────────────────────────────

#[test]
#[ignore = "live COM; run explicitly"]
fn v3_on_pair_changed_arg_order_pin() {
    // Positional DISPPARAMS arg-order pin. FirePairChanged 10 raises (10, 11).
    // The in-proc fixture sends plain positional IDispatch args, so rgvarg is
    // last-to-first and the sink must recover declared order. The a/a+1 invariant
    // catches reversal regardless of value:
    // verdict = IIf(mA = 10 And mB = 11 And (mB - mA) = 1, 1, 0).
    let sink = "Private mA As Long\n\
         Private mB As Long\n\
         Private WithEvents src As OxVba.TestEventServer\n\
         Public Sub Wire()\n\
         Set src = New OxVba.TestEventServer\n\
         src.FirePairChanged 10\n\
         DoEvents\n\
         End Sub\n\
         Public Function Result() As Long\n\
         Result = IIf(mA = 10 And mB = 11 And (mB - mA) = 1, 1, 0)\n\
         End Function\n\
         Private Sub src_OnPairChanged(ByVal a As Long, ByVal b As Long)\n\
         mA = a\n\
         mB = b\n\
         End Sub\n";
    run_event_case_result("V3", sink, 1);
}

// ── V4: unsubscribe Set src=Nothing (keep 2nd ref) ───────────────────────────

#[test]
#[ignore = "live COM; run explicitly"]
fn v4_unsubscribe_set_nothing() {
    // Real Unadvise. keepAlive holds a 2nd ref so the server lives; fire 1
    // (delivered), Set src = Nothing (unsubscribe), fire 99 via keepAlive (NOT
    // delivered). verdict = IIf(mGot = 1, 1, 0).
    let sink = "Private mGot As Long\n\
         Private keepAlive As OxVba.TestEventServer\n\
         Private WithEvents src As OxVba.TestEventServer\n\
         Public Sub Wire()\n\
         Set src = New OxVba.TestEventServer\n\
         Set keepAlive = src\n\
         src.FireValueChanged 1\n\
         DoEvents\n\
         Set src = Nothing\n\
         keepAlive.FireValueChanged 99\n\
         DoEvents\n\
         End Sub\n\
         Public Function Result() As Long\n\
         Result = IIf(mGot = 1, 1, 0)\n\
         End Function\n\
         Private Sub src_OnValueChanged(ByVal value As Long)\n\
         mGot = value\n\
         End Sub\n";
    run_event_case_result("V4", sink, 1);
}

// ── V5: re-subscribe / replace source ────────────────────────────────────────

#[test]
#[ignore = "live COM; run explicitly"]
fn v5_replace_source_pre_teardown() {
    // Pre-teardown of the old advise on WithEventsSet. Wire a -> fire 1 (delivered),
    // swap src to b, a.fire 50 (ignored — a is no longer the subscribed source),
    // b.fire 2 (delivered). log == "1;2;". verdict = IIf(mLog = "1;2;", 1, 0).
    let sink = "Private mLog As String\n\
         Private a As OxVba.TestEventServer\n\
         Private b As OxVba.TestEventServer\n\
         Private WithEvents src As OxVba.TestEventServer\n\
         Public Sub Wire()\n\
         Set a = New OxVba.TestEventServer\n\
         Set b = New OxVba.TestEventServer\n\
         Set src = a\n\
         a.FireValueChanged 1\n\
         DoEvents\n\
         Set src = b\n\
         a.FireValueChanged 50\n\
         DoEvents\n\
         b.FireValueChanged 2\n\
         DoEvents\n\
         End Sub\n\
         Public Function Result() As Long\n\
         Result = IIf(mLog = \"1;2;\", 1, 0)\n\
         End Function\n\
         Private Sub src_OnValueChanged(ByVal value As Long)\n\
         mLog = mLog & CStr(value) & \";\"\n\
         End Sub\n";
    run_event_case_result("V5", sink, 1);
}

// ── V6: multi-sink fan-out (2 sinks, 1 source) ───────────────────────────────

#[test]
#[ignore = "live COM; run explicitly"]
fn v6_multi_sink_fan_out() {
    // Subscription -> owner routing. Two Sink instances both wired to the SAME server
    // must each receive the fire. Main builds two sinks sharing one server and reads
    // both captured values. verdict = IIf(r1 = 77 And r2 = 77, 1, 0).
    let main_src = "Public result As Long\n\
         Sub Main()\n\
         Dim srv As New OxVba.TestEventServer\n\
         Dim s1 As New Sink\n\
         Dim s2 As New Sink\n\
         s1.WireTo srv\n\
         s2.WireTo srv\n\
         srv.FireValueChanged 77\n\
         DoEvents\n\
         result = IIf(s1.Got() = 77 And s2.Got() = 77, 1, 0)\n\
         End Sub\n";
    let sink = "Private mGot As Long\n\
         Private WithEvents src As OxVba.TestEventServer\n\
         Public Sub WireTo(ByVal srv As OxVba.TestEventServer)\n\
         Set src = srv\n\
         End Sub\n\
         Public Function Got() As Long\n\
         Got = mGot\n\
         End Function\n\
         Private Sub src_OnValueChanged(ByVal value As Long)\n\
         mGot = value\n\
         End Sub\n";
    run_manifest_case("V6", main_src, sink, 1);
}

// ── V7 (P0): OOP Excel Application.NewWorkbook ───────────────────────────────

#[test]
#[ignore = "live COM; run explicitly"]
fn v7_excel_new_workbook_oop_event() {
    // Out-of-process Excel `Application.NewWorkbook` through an early-bound `Workbooks.Add` with
    // its optional argument omitted. The handler also dereferences the `Wb` object arg (reads
    // `Wb.Name`), so the verdict requires a live workbook, not Nothing or a dead proxy.
    // verdict = IIf(mFired >= 1 And mWbOk >= 1, 1, 0).
    let main_src = "Public result As Long\n\
         Sub Main()\n\
         Dim s As New Sink\n\
         result = s.Run()\n\
         End Sub\n";
    let sink = "Private mFired As Long\n\
         Private mWbOk As Long\n\
         Private app As Excel.Application\n\
         Private WithEvents appEv As Excel.Application\n\
         Public Function Run() As Long\n\
         Set app = CreateObject(\"Excel.Application\")\n\
         app.Visible = False\n\
         app.DisplayAlerts = False\n\
         Set appEv = app\n\
         app.Workbooks.Add\n\
         Dim i As Long\n\
         For i = 1 To 50\n\
         DoEvents\n\
         Next i\n\
         app.Quit\n\
         Run = IIf(mFired >= 1 And mWbOk >= 1, 1, 0)\n\
         End Function\n\
         Private Sub appEv_NewWorkbook(ByVal Wb As Excel.Workbook)\n\
         mFired = mFired + 1\n\
         If Len(Wb.Name) > 0 Then mWbOk = mWbOk + 1\n\
         End Sub\n";
    run_manifest_case_excel("V7", main_src, sink, 1);
}

// ── V8: OOP Excel Workbook.SheetChange (2-arg, Range arg) ────────────────────

#[test]
#[ignore = "live COM; run explicitly"]
fn v8_excel_sheet_change_oop_object_arg() {
    // 2-arg OOP event + object-arg re-entry into a Range from inside the handler. Write to a
    // C-column cell; the handler reads Target.Column (== 3). Excel passes `Sh` and `Target` as
    // named arguments (rgdispidNamedArgs maps them to positions 0 and 1).
    // verdict = IIf(mCol = 3, 1, 0).
    let main_src = "Public result As Long\n\
         Sub Main()\n\
         Dim s As New Sink\n\
         result = s.Run()\n\
         End Sub\n";
    let sink = "Private mCol As Long\n\
         Private app As Excel.Application\n\
         Private wb As Excel.Workbook\n\
         Private WithEvents wbEv As Excel.Workbook\n\
         Public Function Run() As Long\n\
         Set app = CreateObject(\"Excel.Application\")\n\
         app.Visible = False\n\
         app.DisplayAlerts = False\n\
         Set wb = app.Workbooks.Add\n\
         Set wbEv = wb\n\
         wb.Worksheets(1).Range(\"C1\").Value = 5\n\
         Dim i As Long\n\
         For i = 1 To 50\n\
         DoEvents\n\
         Next i\n\
         app.Quit\n\
         Run = IIf(mCol = 3, 1, 0)\n\
         End Function\n\
         Private Sub wbEv_SheetChange(ByVal Sh As Object, ByVal Target As Excel.Range)\n\
         mCol = Target.Column\n\
         End Sub\n";
    run_manifest_case_excel("V8", main_src, sink, 1);
}

// ── V9: re-entrancy — handler re-fires ───────────────────────────────────────

#[test]
#[ignore = "live COM; run explicitly"]
fn v9_handler_re_entrancy() {
    // An event fired inside a handler runs at once, nested inside it, as in VBA: Excel x64 with
    // this server logs "(3(2(11)2)3)" (probe `temp/variant-layout-probe/events8tes`, 2026-11-15).
    // FireValueChanged 3; the handler appends "(v", re-fires v-1 (while v > 1) and calls
    // DoEvents, then appends "v)".
    let sink = "Private mLog As String\n\
         Private WithEvents src As OxVba.TestEventServer\n\
         Public Sub Wire()\n\
         Set src = New OxVba.TestEventServer\n\
         src.FireValueChanged 3\n\
         DoEvents\n\
         End Sub\n\
         Public Function Result() As Long\n\
         Result = IIf(mLog = \"(3(2(11)2)3)\", 1, 0)\n\
         End Function\n\
         Private Sub src_OnValueChanged(ByVal value As Long)\n\
         mLog = mLog & \"(\" & CStr(value)\n\
         If value > 1 Then\n\
         src.FireValueChanged value - 1\n\
         DoEvents\n\
         End If\n\
         mLog = mLog & CStr(value) & \")\"\n\
         End Sub\n";
    run_event_case_result("V9", sink, 1);
}

// ── V10: an unhandled handler error ends the run ──────────────────────────────

#[test]
#[ignore = "live COM; run explicitly"]
fn v10_unhandled_handler_error_ends_the_run() {
    // VBA reports an error a handler leaves unhandled as a top-level run-time error and ends the
    // macro, even under the firing code's `On Error Resume Next` (Excel x64 with this server,
    // probe `temp/variant-layout-probe/events8tes`, 2026-11-15).
    let sink = "Private WithEvents src As OxVba.TestEventServer\n\
         Public Sub Wire()\n\
         Set src = New OxVba.TestEventServer\n\
         On Error Resume Next\n\
         src.FireValueChanged 1\n\
         End Sub\n\
         Public Function Result() As Long\n\
         Result = 1\n\
         End Function\n\
         Private Sub src_OnValueChanged(ByVal value As Long)\n\
         Err.Raise 5\n\
         End Sub\n";
    let main_src = "Public result As Long\n\
         Sub Main()\n\
         Dim s As New Sink\n\
         s.Wire\n\
         result = s.Result()\n\
         End Sub\n";
    let modules = vec![
        ("Main", ModuleKind::Procedural, main_src.to_string()),
        ("Sink", ModuleKind::Class, sink.to_string()),
    ];
    match run_manifest(modules, tes_ref()) {
        Ok(snap) => panic!("V10: the run did not end: {snap:?}"),
        Err(e) if is_typelib_absent(&e) => eprintln!("SKIP V10: {e}"),
        Err(e) => assert!(e.contains("error 5"), "V10: {e}"),
    }
}

// ── V12: subscription lifetime / Class_Terminate Unadvise ────────────────────

#[test]
#[ignore = "live COM; run explicitly"]
fn v12_subscription_lifetime_terminate_unadvise() {
    // Terminate-drain vs Unadvise ordering: a sink released before the fire must not
    // receive delivery and must not AV into the freed sink. Main creates an inner
    // sink, releases it (Set inner = Nothing), THEN fires on a server it still holds.
    // This is purely a NO-AV / reached-completion guard: result is set to 1
    // unconditionally after the fire, and the freed sink's mGot is unobservable (it lives
    // on the released instance, and the snapshot captures only module globals + Main's
    // locals), so non-delivery is NOT independently verified here. verdict = 1.
    let main_src = "Public result As Long\n\
         Sub Main()\n\
         Dim srv As New OxVba.TestEventServer\n\
         Dim inner As Sink\n\
         Set inner = New Sink\n\
         inner.WireTo srv\n\
         Set inner = Nothing\n\
         srv.FireValueChanged 5\n\
         DoEvents\n\
         result = 1\n\
         End Sub\n";
    let sink = "Private mGot As Long\n\
         Private WithEvents src As OxVba.TestEventServer\n\
         Public Sub WireTo(ByVal srv As OxVba.TestEventServer)\n\
         Set src = srv\n\
         End Sub\n\
         Private Sub src_OnValueChanged(ByVal value As Long)\n\
         mGot = value\n\
         End Sub\n";
    run_manifest_case("V12", main_src, sink, 1);
}

// ─────────────────────────────────────────────────────────────────────────────
// Event-case runners. The standard shape is a Main that drives a single Sink and
// reads a result/hit count into `result`; some cases supply their own Main.
// ─────────────────────────────────────────────────────────────────────────────

/// Standard 0/1-arg case: Main calls `s.Wire` then reads `s.Hits()` into `result`.
fn run_event_case(case: &str, sink_src: &str, expected: i64) {
    let main_src = "Public result As Long\n\
         Sub Main()\n\
         Dim s As New Sink\n\
         s.Wire\n\
         result = s.Hits()\n\
         End Sub\n";
    run_manifest_case(case, main_src, sink_src, expected);
}

/// Result-flag case: Main calls `s.Wire` then reads `s.Result()` into `result`.
fn run_event_case_result(case: &str, sink_src: &str, expected: i64) {
    let main_src = "Public result As Long\n\
         Sub Main()\n\
         Dim s As New Sink\n\
         s.Wire\n\
         result = s.Result()\n\
         End Sub\n";
    run_manifest_case(case, main_src, sink_src, expected);
}

/// Run a (Main + Sink) manifest with the TestEventServer typelib reference, env-skip
/// on absent component/typelib, and assert the first Long in the snapshot == expected.
fn run_manifest_case(case: &str, main_src: &str, sink_src: &str, expected: i64) {
    let modules = vec![
        ("Main", ModuleKind::Procedural, main_src.to_string()),
        ("Sink", ModuleKind::Class, sink_src.to_string()),
    ];
    match run_manifest(modules, tes_ref()) {
        Ok(snap) => assert_eq!(
            find_verdict(&snap),
            Some(expected),
            "{case}: event result {:?} != expected {expected} in {snap:?}",
            find_verdict(&snap)
        ),
        Err(e) if is_typelib_absent(&e) => eprintln!("SKIP {case}: {e}"),
        Err(e) => panic!("{case} event manifest failed: {e}"),
    }
}

/// Like [`run_manifest_case`] but carries the Excel typelib reference (OOP events).
fn run_manifest_case_excel(case: &str, main_src: &str, sink_src: &str, expected: i64) {
    let modules = vec![
        ("Main", ModuleKind::Procedural, main_src.to_string()),
        ("Sink", ModuleKind::Class, sink_src.to_string()),
    ];
    match run_manifest(modules, excel_ref()) {
        Ok(snap) => assert_eq!(
            find_verdict(&snap),
            Some(expected),
            "{case}: OOP event result {:?} != expected {expected} in {snap:?}",
            find_verdict(&snap)
        ),
        Err(e) if is_component_absent(&e) => eprintln!("SKIP {case}: {e}"),
        Err(e) if is_typelib_absent(&e) => eprintln!("SKIP {case}: {e}"),
        Err(e) => panic!("{case} OOP event manifest failed: {e}"),
    }
}
