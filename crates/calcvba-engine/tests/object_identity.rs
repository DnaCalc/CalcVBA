//! dc-yqq0.20: a project object is one COM object (design note
//! docs/design/vba-object-is-one-com-object-2026-11-20.md). Each case runs on VM3 and the JIT
//! and expects the answer Excel's VBA gave on Windows (2026-11-21, `temp/objid2` probe):
//!
//! `coll-is=True;coll-call=42;dict-is=True;dict-call=42;coll-term-held=0;coll-term-removed=1;
//! foreign-call=42;term-held=0;term-dropped=1`
//!
//! The early-bound foreign call (dc-yqq0.24, `temp/objid4` probe) calls vtable slots directly:
//! `early-get=21;early-call=42;after-put=5;early-call2=10`. The final release on a foreign thread
//! (dc-yqq0.25) has no Excel answer (a raw Release of a VBA object off its thread is undefined
//! there); CalcVBA runs `Class_Terminate` on the owner thread at the next statement.
//!
//! The live-COM cases need OxVba.TestEventServer (tools/OxVba.TestEventServer/register.ps1).

use calcvba_engine::{Engine, ExecBackend, HostConfig};
use calcvba_hal::model::HostPolicy;
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};

const BACKENDS: [ExecBackend; 2] = [ExecBackend::Vm3, ExecBackend::Jit];

const WIDGET: &str = r#"
Public Value As Long
Public Function Twice() As Long
    Twice = Value * 2
End Function
Private Sub Class_Terminate()
    gTerm = gTerm + 1
End Sub
"#;

fn module(name: &str, kind: ModuleKind, source: &str) -> ModuleUnit {
    ModuleUnit {
        module_name: name.to_string(),
        module_kind: kind,
        attributes: ModuleAttributes::named(name),
        source: source.to_string(),
    }
}

/// Runs `body` as `Sub Main` (with `w` a `Widget` whose Value is 21) and returns `gOut`.
fn run(backend: ExecBackend, body: &str) -> String {
    let main = format!(
        "Public gTerm As Long\nPublic gOut As String\nSub Main()\n    Dim w As New Widget\n    w.Value = 21\n{body}\nEnd Sub\n"
    );
    let manifest = SymbolProjectManifest {
        project_name: "ObjId".to_string(),
        project_kind: ProjectKind::Source,
        modules: vec![
            module("Probe", ModuleKind::Procedural, &main),
            module("Widget", ModuleKind::Class, WIDGET),
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let mut engine = Engine::new(HostConfig { backend });
    engine.set_host_policy(HostPolicy::native());
    let values = engine
        .execute_manifest_with_variant_snapshot(&manifest)
        .unwrap_or_else(|err| panic!("{backend:?}: {err}"));
    values
        .iter()
        .filter_map(|value| value.as_bstr().map(|text| text.as_str().to_string()))
        .find(|text| text.contains('='))
        .unwrap_or_else(|| panic!("{backend:?}: no gOut in {values:?}"))
}

#[test]
fn collection_round_trip_and_terminate_match_excel() {
    let body = r#"
    Dim c As Collection, h As Widget
    Set c = New Collection
    c.Add w, "k"
    gOut = "coll-is=" & CStr(c("k") Is w) & ";coll-call=" & CStr(c("k").Twice())
    Set h = New Widget
    gTerm = 0
    c.Add h, "h"
    Set h = Nothing
    gOut = gOut & ";coll-term-held=" & CStr(gTerm)
    c.Remove "h"
    gOut = gOut & ";coll-term-removed=" & CStr(gTerm)"#;
    for backend in BACKENDS {
        assert_eq!(
            run(backend, body),
            "coll-is=True;coll-call=42;coll-term-held=0;coll-term-removed=1",
            "{backend:?}"
        );
    }
}

#[cfg(target_os = "windows")]
#[test]
fn dictionary_round_trip_matches_excel() {
    // Scripting.Dictionary is a foreign COM object: the instance crosses out and back.
    let body = r#"
    Dim d As Object, back As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "k", w
    Set back = d("k")
    gOut = "dict-is=" & CStr(back Is w) & ";dict-call=" & CStr(back.Twice())"#;
    for backend in BACKENDS {
        assert_eq!(
            run(backend, body),
            "dict-is=True;dict-call=42",
            "{backend:?}"
        );
    }
}

#[cfg(target_os = "windows")]
#[test]
#[ignore = "live COM: needs OxVba.TestEventServer (tools/OxVba.TestEventServer/register.ps1)"]
fn foreign_holder_keeps_the_object_alive_as_excel_does() {
    let body = r#"
    Dim s As Object, h As Widget
    Set s = CreateObject("OxVba.TestEventServer")
    Set h = New Widget
    gTerm = 0
    s.Hold h
    Set h = Nothing
    gOut = "term-held=" & CStr(gTerm)
    s.Drop
    gOut = gOut & ";term-dropped=" & CStr(gTerm)"#;
    for backend in BACKENDS {
        assert_eq!(
            run(backend, body),
            "term-held=0;term-dropped=1",
            "{backend:?}"
        );
    }
}

#[cfg(target_os = "windows")]
#[test]
#[ignore = "live COM: needs OxVba.TestEventServer (tools/OxVba.TestEventServer/register.ps1)"]
fn foreign_late_bound_call_matches_excel() {
    // .NET `InvokeMember` calls the object through its own IDispatch.
    let body = r#"
    Dim s As Object
    Set s = CreateObject("OxVba.TestEventServer")
    On Error Resume Next
    gOut = "foreign-call=" & CStr(s.CallMember(w, "Twice"))
    If Err.Number <> 0 Then gOut = "foreign-call=ERR" & Err.Number"#;
    for backend in BACKENDS {
        assert_eq!(run(backend, body), "foreign-call=42", "{backend:?}");
    }
}

#[cfg(target_os = "windows")]
#[test]
#[ignore = "live COM: needs OxVba.TestEventServer (tools/OxVba.TestEventServer/register.ps1)"]
fn foreign_early_bound_slot_calls_match_excel() {
    // Widget's dual vtable, as in Excel: Value get = slot 7, Value let = 8, Twice = 9.
    let body = r#"
    Dim s As Object
    Set s = CreateObject("OxVba.TestEventServer")
    On Error Resume Next
    gOut = "early-get=" & CStr(s.CallSlotGetLong(w, 7))
    If Err.Number <> 0 Then gOut = gOut & "ERR" & Err.Number: Err.Clear
    gOut = gOut & ";early-call=" & CStr(s.CallSlotGetLong(w, 9))
    If Err.Number <> 0 Then gOut = gOut & "ERR" & Err.Number: Err.Clear
    s.CallSlotPutLong w, 8, 5
    If Err.Number <> 0 Then gOut = gOut & ";early-put=ERR" & Err.Number: Err.Clear
    gOut = gOut & ";after-put=" & CStr(w.Value) & ";early-call2=" & CStr(s.CallSlotGetLong(w, 9))
    If Err.Number <> 0 Then gOut = gOut & "ERR" & Err.Number: Err.Clear"#;
    for backend in BACKENDS {
        assert_eq!(
            run(backend, body),
            "early-get=21;early-call=42;after-put=5;early-call2=10",
            "{backend:?}"
        );
    }
}

#[cfg(target_os = "windows")]
#[test]
#[ignore = "live COM: needs OxVba.TestEventServer (tools/OxVba.TestEventServer/register.ps1)"]
fn foreign_thread_final_release_terminates_on_the_owner_thread() {
    let body = r#"
    Dim s As Object, h As Widget
    Set s = CreateObject("OxVba.TestEventServer")
    Set h = New Widget
    gTerm = 0
    s.Hold h
    Set h = Nothing
    gOut = "term-held=" & CStr(gTerm)
    s.DropOnOtherThread
    gOut = gOut & ";term-dropped=" & CStr(gTerm)"#;
    for backend in BACKENDS {
        assert_eq!(
            run(backend, body),
            "term-held=0;term-dropped=1",
            "{backend:?}"
        );
    }
}

#[cfg(target_os = "windows")]
#[test]
#[ignore = "live COM: needs OxVba.TestEventServer (tools/OxVba.TestEventServer/register.ps1)"]
fn foreign_get_type_info_describe_dispatch_matches_excel() {
    // dc-yqq0.29: OxVba.TestEventServer.DescribeDispatch inspects GetTypeInfo(0) and its dual
    // TKIND_INTERFACE half via COM interop, matching docs/design/vba-class-dispids-excel-2026-11-21.md.
    let w7 = r#"
Public F1 As Long
Public F2 As Variant
Public Property Get Zed() As Long
End Property
Public Property Get Alpha() As Long
End Property
Public Sub M()
End Sub
"#;
    let main = r#"
Public gOut As String
Sub Main()
    Dim x As New W7, s As Object
    Set s = CreateObject("OxVba.TestEventServer")
    gOut = s.DescribeDispatch(x)
End Sub
"#;
    let manifest = SymbolProjectManifest {
        project_name: "ObjIdTypeInfo".to_string(),
        project_kind: ProjectKind::Source,
        modules: vec![
            module("Probe", ModuleKind::Procedural, main),
            module("W7", ModuleKind::Class, w7),
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    for backend in BACKENDS {
        let mut engine = Engine::new(HostConfig { backend });
        engine.set_host_policy(HostPolicy::native());
        let values = engine
            .execute_manifest_with_variant_snapshot(&manifest)
            .unwrap_or_else(|err| panic!("{backend:?}: {err}"));
        let desc = values
            .iter()
            .filter_map(|v| v.as_bstr().map(|text| text.as_str().to_string()))
            .find(|text| text.contains("TKIND_DISPATCH"))
            .unwrap_or_else(|| panic!("{backend:?}: no DescribeDispatch output in {values:?}"));
        assert!(
            desc.contains("type W7 kind=TKIND_DISPATCH flags=0x1C2"),
            "{desc}"
        );
        assert!(desc.contains("funcs=8 vars=0 impl=1 vtbl=120"), "{desc}");
        assert!(desc.contains("TYPE_E_ELEMENTNOTFOUND"), "{desc}");
        assert!(
            desc.contains(
                "func memid=0x40030000 invkind=2 oVft=56 params=0 opt=0 flags=0x0 ret=3 F1"
            ),
            "{desc}"
        );
        assert!(
            desc.contains(
                "func memid=0x40030001 invkind=8 oVft=88 params=1 opt=0 flags=0x0 ret=24 F2"
            ),
            "{desc}"
        );
        assert!(
            desc.contains(
                "func memid=0x68030001 invkind=2 oVft=96 params=0 opt=0 flags=0x0 ret=3 Zed"
            ),
            "{desc}"
        );
        assert!(
            desc.contains(
                "func memid=0x68030000 invkind=2 oVft=104 params=0 opt=0 flags=0x0 ret=3 Alpha"
            ),
            "{desc}"
        );
        assert!(
            desc.contains(
                "func memid=0x60030002 invkind=1 oVft=112 params=0 opt=0 flags=0x0 ret=24 M"
            ),
            "{desc}"
        );
        assert!(
            desc.contains("--dual\ntype W7 kind=TKIND_INTERFACE flags=0x1C2"),
            "{desc}"
        );
        assert!(
            desc.contains(
                "func memid=0x60030002 invkind=1 oVft=112 params=0 opt=0 flags=0x0 ret=25 M"
            ),
            "{desc}"
        );
    }
}
