//! `Declare` over the runtime cells off Windows (design doc, step 9), on VM3 and the JIT, against
//! the C library: ByRef arguments are the caller's own storage, a ByVal String is an ANSI
//! temporary copied back into the variable, and ByVal scalars take VBA's coercion. These mirror
//! `declare_cells.rs` (Windows), with libc entry points in place of kernel32's.
#![cfg(all(target_os = "linux", target_pointer_width = "64"))]

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

const MEMCPY: &str = "Private Declare PtrSafe Sub CopyMemory Lib \"libc.so.6\" Alias \"memcpy\" (ByRef dst As Any, ByRef src As Any, ByVal n As LongPtr)\n";

#[test]
fn byref_arguments_are_the_callers_storage() {
    let source = format!(
        "{MEMCPY}Public result As String\n\
Sub Main()\n\
    Dim a As Long, b As Long, d As Double, e As Double\n\
    b = 42: e = 2.5\n\
    CopyMemory a, b, 4\n\
    CopyMemory d, e, 8\n\
    result = CStr(a) & \"|\" & CStr(d)\n\
End Sub\n"
    );
    for (name, values) in run(&source) {
        assert_eq!(text(&values), "42|2.5", "{name}");
    }
}

#[test]
fn a_byref_array_element_is_the_element_itself() {
    let source = format!(
        "{MEMCPY}Public result As String\n\
Sub Main()\n\
    Dim a(0 To 3) As Long, x As Long\n\
    x = 7\n\
    CopyMemory a(2), x, 4\n\
    CopyMemory a(0), a(2), 8\n\
    result = a(0) & \",\" & a(1) & \",\" & a(2) & \",\" & a(3)\n\
End Sub\n"
    );
    for (name, values) in run(&source) {
        assert_eq!(text(&values), "7,0,7,0", "{name}");
    }
}

#[test]
fn varptr_addresses_reach_the_variables() {
    let source = "Private Declare PtrSafe Sub CopyPtr Lib \"libc.so.6\" Alias \"memcpy\" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal n As LongPtr)\n\
Public result As String\n\
Sub Main()\n\
    Dim a As Long, b As Long, v As Variant\n\
    b = 99\n\
    CopyPtr VarPtr(a), VarPtr(b), 4\n\
    v = 3&\n\
    CopyPtr VarPtr(v) + 8, VarPtr(b), 4\n\
    result = CStr(a) & \"|\" & CStr(v) & \"|\" & TypeName(v)\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "99|99|Long", "{name}");
    }
}

#[test]
fn a_byval_string_variable_gets_the_ansi_buffer_back() {
    // memset writes into the ANSI temporary; VBA converts it back into the variable.
    let source = "Private Declare PtrSafe Function FillA Lib \"libc.so.6\" Alias \"memset\" (ByVal s As String, ByVal c As Long, ByVal n As LongPtr) As LongPtr\n\
Private Declare PtrSafe Function StrLenA Lib \"libc.so.6\" Alias \"strlen\" (ByVal s As String) As LongPtr\n\
Public result As String\n\
Sub Main()\n\
    Dim s As String\n\
    s = \"abcde\"\n\
    FillA s, 65, 3\n\
    result = s & \"|\" & CStr(StrLenA(s)) & \"|\" & CStr(StrLenA(\"\"))\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "AAAde|5|0", "{name}");
    }
}

#[test]
fn byval_scalars_take_vba_coercion() {
    // abs(CLng(2.5)) is 2 and abs(CLng(3.5)) is 4: banker's rounding, not truncation.
    let source = "Private Declare PtrSafe Function AbsL Lib \"libc.so.6\" Alias \"abs\" (ByVal a As Long) As Long\n\
Public result As String\n\
Sub Main()\n\
    Dim d As Double\n\
    d = 2.5: result = CStr(AbsL(d))\n\
    d = -3.5: result = result & \",\" & CStr(AbsL(d))\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "2,4", "{name}");
    }
}

#[test]
fn a_record_without_strings_is_its_own_storage() {
    let source = format!(
        "Private Type TPair\n    a As Long\n    b As Long\nEnd Type\n\
{MEMCPY}Public result As String\n\
Sub Main()\n\
    Dim p As TPair, q As TPair\n\
    q.a = 5: q.b = 6\n\
    CopyMemory p, q, 8\n\
    result = CStr(p.a) & \",\" & CStr(p.b)\n\
End Sub\n"
    );
    for (name, values) in run(&source) {
        assert_eq!(text(&values), "5,6", "{name}");
    }
}

#[test]
fn strptr_has_the_bstr_length_prefix() {
    // Plan defect 16: `StrPtr(s) - 4` holds `LenB(s)`, as in Excel.
    let source = "Private Declare PtrSafe Sub CopyPtr Lib \"libc.so.6\" Alias \"memcpy\" (ByRef dst As Any, ByVal src As LongPtr, ByVal n As LongPtr)\n\
Public result As String\n\
Sub Main()\n\
    Dim s As String, n As Long, w As Integer\n\
    s = \"hello\"\n\
    CopyPtr n, StrPtr(s) - 4, 4\n\
    CopyPtr w, StrPtr(s), 2\n\
    result = CStr(n) & \"|\" & CStr(LenB(s)) & \"|\" & ChrW(w)\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "10|10|h", "{name}");
    }
}

#[test]
fn variants_of_the_integer_types_vba_does_not_declare_keep_their_subtype() {
    // A COM callee can hand back VT_I1, VT_UI2, VT_UI4 or VT_UI8; Office VBA keeps them in the
    // Variant (VarType 16, 18, 19, 21) and CStr converts them (Excel: COM datatype oracle,
    // 16:-5, 18:65000, 19:4000000000, 21:9000000000). The cells are written in place here.
    let source = "Private Declare PtrSafe Sub CopyPtr Lib \"libc.so.6\" Alias \"memcpy\" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal n As LongPtr)\n\
Public result As String\n\
Function Typed(ByVal vt As Integer, ByVal payload As LongLong) As String\n\
    Dim v As Variant\n\
    v = 0&\n\
    CopyPtr VarPtr(v) + 8, VarPtr(payload), 8\n\
    CopyPtr VarPtr(v), VarPtr(vt), 2\n\
    Typed = CStr(VarType(v)) & \":\" & CStr(v)\n\
End Function\n\
Sub Main()\n\
    result = Typed(16, 251) & \"|\" & Typed(18, 65000) & \"|\" & Typed(19, 4000000000^) & \"|\" & Typed(21, 9000000000^)\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(
            text(&values),
            "16:-5|18:65000|19:4000000000|21:9000000000",
            "{name}"
        );
    }
}

#[test]
fn an_empty_array_literal_has_vt_zero_before_its_descriptor() {
    // Excel (design doc step 6): `Array()` has vt 0 at psa-4; `Array(1)` has VT_VARIANT (12).
    let source = "Private Declare PtrSafe Sub CopyPtr Lib \"libc.so.6\" Alias \"memcpy\" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal n As LongPtr)\n\
Public result As String\n\
Function Prefix(ByRef v As Variant) As String\n\
    Dim psa As LongPtr, vt As Long, features As Integer\n\
    CopyPtr VarPtr(psa), VarPtr(v) + 8, 8\n\
    CopyPtr VarPtr(vt), psa - 4, 4\n\
    CopyPtr VarPtr(features), psa + 2, 2\n\
    Prefix = CStr(vt) & \":\" & Hex$(features) & \":\" & CStr(UBound(v) - LBound(v) + 1)\n\
End Function\n\
Sub Main()\n\
    Dim a As Variant, b As Variant\n\
    a = Array(): b = Array(1, 2)\n\
    result = Prefix(a) & \"|\" & Prefix(b)\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "0:880:0|12:880:2", "{name}");
    }
}
