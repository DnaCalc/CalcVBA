//! dc-yqq0.20: a project object is one COM object (design note
//! docs/design/vba-object-is-one-com-object-2026-11-20.md). Each case runs on VM3 and the JIT
//! and expects the answer Excel's VBA gave on Windows (2026-11-21, `temp/objid2` probe):
//!
//! `coll-is=True;coll-call=42;dict-is=True;dict-call=42;coll-term-held=0;coll-term-removed=1;
//! foreign-call=42;term-held=0;term-dropped=1`
//!
//! Today CalcVBA matches every case except the foreign late-bound call, which raises error 5 until
//! the object answers IDispatch itself (dc-yqq0.23). The early-bound foreign call (dc-yqq0.24) and
//! the final release on a foreign thread (dc-yqq0.25) join here with those beads.

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
