//! `Declare` calls over the runtime cells (design doc, step 9), on VM3 and the JIT: the callee
//! gets the caller's storage where VBA passes it, String arguments are ANSI temporaries copied
//! back into the variable, ByVal arguments take VBA's coercion, objects are interface pointers,
//! and `AddressOf` works on both backends. The String expectations follow Excel x64 (probe
//! `temp/variant-layout-probe/declare9`, 2026-11-15: `CharUpperA` on a ByVal String variable
//! changes the variable, a character outside the code page arriving best-fit).
#![cfg(all(target_os = "windows", target_arch = "x86_64"))]

use calcvba_engine::{Engine, HostConfig};
use calcvba_hal::model::HostPolicy;
use calcvba_runtime::Variant;

fn run(source: &str) -> Vec<(&'static str, Vec<Variant>)> {
    [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())]
        .into_iter()
        .map(|(name, config)| {
            let mut engine = Engine::new(config);
            engine.set_host_policy(HostPolicy::native());
            let values = engine
                .execute_source_with_variant_snapshot_clean(source)
                .unwrap_or_else(|error| panic!("{name}: {error:?}"));
            (name, values)
        })
        .collect()
}

fn text(values: &[Variant]) -> String {
    values
        .iter()
        .find_map(|value| value.as_bstr().map(|text| text.to_string()))
        .unwrap_or_default()
}

#[test]
fn address_of_runs_its_callback_on_both_backends() {
    let source = "Private Declare PtrSafe Function CallPtr4 Lib \"user32\" Alias \"CallWindowProcW\" (ByVal proc As LongPtr, ByVal a0 As LongPtr, ByVal a1 As LongPtr, ByVal a2 As LongPtr, ByVal a3 As LongPtr) As LongPtr\n\
Public result As String\n\
Public seen As String\n\
Sub Main()\n\
    Dim r As LongPtr\n\
    r = CallPtr4(AddressOf Callback, 11, 22, 33, 44)\n\
    result = seen & \"|\" & CStr(r)\n\
End Sub\n\
Function Callback(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long) As LongPtr\n\
    seen = CStr(hWnd) & \",\" & CStr(uMsg) & \",\" & CStr(idEvent) & \",\" & CStr(dwTime)\n\
    Callback = 99\n\
End Function\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "11,22,33,44|99", "{name}");
    }
}

#[test]
fn byval_scalars_take_vba_coercion() {
    // MulDiv(a, 1, 1) returns the Long it received: CLng's banker's rounding, not truncation.
    let source = "Private Declare PtrSafe Function Echo Lib \"kernel32\" Alias \"MulDiv\" (ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long\n\
Public result As String\n\
Sub Main()\n\
    Dim d As Double\n\
    d = 2.5: result = CStr(Echo(d, 1, 1))\n\
    d = 3.5: result = result & \",\" & CStr(Echo(d, 1, 1))\n\
    result = result & \",\" & CStr(Echo(\"7\", 1, 1))\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "2,4,7", "{name}");
    }
}

#[test]
fn a_byref_array_element_is_the_element_itself() {
    // RtlMoveMemory a(0), src(0), 16 writes 16 bytes from a(0) on, as in VBA (no copy the size
    // of one element).
    let source = "Private Declare PtrSafe Sub CopyMemory Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByRef dst As Any, ByRef src As Any, ByVal n As LongPtr)\n\
Public result As String\n\
Sub Main()\n\
    Dim a(0 To 15) As Byte, src(0 To 15) As Byte, i As Long\n\
    For i = 0 To 15: src(i) = i + 100: Next\n\
    CopyMemory a(0), src(0), 16\n\
    result = CStr(a(0)) & \",\" & CStr(a(7)) & \",\" & CStr(a(15))\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "100,107,115", "{name}");
    }
}

#[test]
fn a_byval_string_variable_gets_the_ansi_buffer_back() {
    // CharUpperA uppercases the ANSI buffer in place; VBA converts it back into the variable.
    // Alpha is outside the ANSI code page: best fit makes it `a`, then `A`.
    let source = "Private Declare PtrSafe Function CharUpperA Lib \"user32\" (ByVal s As String) As LongPtr\n\
Public result As String\n\
Sub Main()\n\
    Dim s As String\n\
    s = \"ab\" & ChrW(&H3B1) & \"c\"\n\
    CharUpperA s\n\
    result = s\n\
    CharUpperA \"literal\"\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "ABAC", "{name}");
    }
}

#[test]
fn a_record_without_strings_is_its_own_storage() {
    // VBA passes such a UDT's memory itself (the address `VarPtr` gives); one with a String is
    // an ANSI copy, converted back.
    let source = "Private Declare PtrSafe Function EncodeRef Lib \"kernel32\" Alias \"EncodePointer\" (ByRef p As Any) As LongPtr
Private Declare PtrSafe Function DecodePointer Lib \"kernel32\" (ByVal p As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByRef dst As Any, ByRef src As Any, ByVal n As LongPtr)
Private Type TNum
    a As Long
    b As Double
    c(1 To 2) As Integer
End Type
Private Type TText
    a As Long
    s As String
End Type
Public result As String
Sub Main()
    Dim n As TNum, m As TNum, t As TText
    m.a = 7: m.b = 2.5: m.c(2) = 9
    CopyMemory n, m, LenB(n)
    t.s = \"abc\"
    result = CStr(DecodePointer(EncodeRef(n)) = VarPtr(n)) & \",\" & CStr(n.a) & \",\" & CStr(n.b) & \",\" & CStr(n.c(2))
    result = result & \",\" & CStr(DecodePointer(EncodeRef(t)) = VarPtr(t)) & \",\" & t.s
End Sub
";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "True,7,2.5,9,False,abc", "{name}");
    }
}

#[test]
fn vb_null_string_is_a_null_bstr() {
    // `EncodePointer`/`DecodePointer` hand back the ANSI pointer the callee got: NULL for
    // vbNullString (and a String variable assigned it, or never assigned), not for "".
    let source = "Private Declare PtrSafe Function EncodeText Lib \"kernel32\" Alias \"EncodePointer\" (ByVal s As String) As LongPtr
Private Declare PtrSafe Function DecodePointer Lib \"kernel32\" (ByVal p As LongPtr) As LongPtr
Public result As String
Sub Main()
    Dim s As String, t As String, v As Variant
    s = \"x\": s = vbNullString: v = vbNullString
    result = CStr(StrPtr(vbNullString)) & \",\" & CStr(StrPtr(s)) & \",\" & CStr(StrPtr(t)) & \",\" & CStr(StrPtr(v))
    result = result & \",\" & CStr(StrPtr(\"\") <> 0) & \",\" & CStr(vbNullString = \"\") & \",\" & CStr(Len(vbNullString)) & \",\" & CStr(VarType(v))
    result = result & \",\" & CStr(DecodePointer(EncodeText(vbNullString))) & \",\" & CStr(DecodePointer(EncodeText(s)))
    result = result & \",\" & CStr(DecodePointer(EncodeText(\"\")) <> 0) & \",\" & CStr(StrPtr(vbNullString & \"\") <> 0)
End Sub
";
    for (name, values) in run(source) {
        assert_eq!(
            text(&values),
            "0,0,0,0,True,True,0,8,0,0,True,True",
            "{name}"
        );
    }
}

#[test]
fn a_byval_object_is_its_interface_pointer() {
    // EncodePointer/DecodePointer round-trip the value the callee received.
    let source = "Private Declare PtrSafe Function EncodeObject Lib \"kernel32\" Alias \"EncodePointer\" (ByVal o As Object) As LongPtr\n\
Private Declare PtrSafe Function DecodePointer Lib \"kernel32\" (ByVal p As LongPtr) As LongPtr\n\
Public result As String\n\
Sub Main()\n\
    Dim o As Object\n\
    Set o = CreateObject(\"Scripting.Dictionary\")\n\
    result = CStr(DecodePointer(EncodeObject(o)) = ObjPtr(o)) & \",\" & CStr(DecodePointer(EncodeObject(Nothing)))\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "True,0", "{name}");
    }
}

#[test]
fn a_record_holding_strings_is_passed_in_its_ansi_layout() {
    // VBA passes a UDT with Strings to a Declare in its ANSI layout: String * 4 is four ANSI
    // bytes (lstrlenA sees 4 before the zero Byte), String is an ANSI BSTR (LPSTR), and writes
    // into the ANSI buffer convert back into the Unicode record.
    let source = "Private Type TText\n    f As String * 4\n    z As Byte\n    s As String\n    n As Long\nEnd Type\n\
Private Declare PtrSafe Function StrLenA Lib \"kernel32\" Alias \"lstrlenA\" (ByRef r As TText) As Long\n\
Private Declare PtrSafe Sub FillMemory Lib \"kernel32\" Alias \"RtlFillMemory\" (ByRef r As TText, ByVal n As LongPtr, ByVal b As Byte)\n\
Public result As String\n\
Sub Main()\n\
    Dim t As TText\n\
    t.f = \"abcd\": t.s = \"h\" & ChrW(233) & \"llo\": t.n = 7\n\
    result = CStr(StrLenA(t))\n\
    FillMemory t, 2, 65\n\
    result = result & \"|\" & t.f & \"|\" & t.s & \"|\" & CStr(t.n)\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "4|AAcd|h\u{e9}llo|7", "{name}");
    }
}

#[test]
fn a_udt_array_is_passed_as_an_fadf_record_safearray_and_copied_back() {
    let source = "Private Type TItem\n    id As Long\n    tag As String * 4\n    label As String\n    extra As Variant\nEnd Type\n\
Private Declare PtrSafe Sub ReadPsa Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByRef dst As LongPtr, ByRef arr() As TItem, ByVal n As LongPtr)\n\
Public result As String\n\
Sub Main()\n\
    Dim items(1 To 2) As TItem, grid(1 To 2, 0 To 1) As TItem\n\
    Dim psa As LongPtr, psa2 As LongPtr\n\
    items(1).id = 11: items(1).tag = \"ABCD\": items(1).label = \"first\": items(1).extra = 100&\n\
    items(2).id = 22: items(2).tag = \"WXYZ\": items(2).label = \"second\": items(2).extra = \"val\"\n\
    grid(2, 1).id = 99: grid(2, 1).tag = \"GRID\": grid(2, 1).label = \"cell\": grid(2, 1).extra = \"ok\"\n\
    ReadPsa psa, items, 8\n\
    ReadPsa psa2, grid, 8\n\
    result = CStr(psa <> 0) & \"|\" & CStr(items(1).id) & \",\" & items(1).tag & \",\" & items(1).label & \",\" & CStr(items(1).extra) & \"|\" & CStr(items(2).id) & \",\" & items(2).tag & \",\" & items(2).label & \",\" & CStr(items(2).extra) & \"|\" & CStr(psa2 <> 0) & \",\" & CStr(grid(2, 1).id) & \",\" & grid(2, 1).tag & \",\" & grid(2, 1).label & \",\" & CStr(grid(2, 1).extra)\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(
            text(&values),
            "True|11,ABCD,first,100|22,WXYZ,second,val|True,99,GRID,cell,ok",
            "{name}"
        );
    }
}
