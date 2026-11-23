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

#[test]
fn a_record_holding_strings_is_passed_in_its_ansi_layout() {
    // VBA passes a UDT with Strings to a Declare in the ANSI layout: a String * 4 is four ANSI
    // bytes (strlen sees 4 before the zero Byte), a String a char* to an ANSI copy, and what the
    // callee writes comes back into the record.
    let source = "Private Type TText\n    f As String * 4\n    z As Byte\n    s As String\n    n As Long\nEnd Type\n\
Private Declare PtrSafe Function StrLenA Lib \"libc.so.6\" Alias \"strlen\" (ByRef r As TText) As LongPtr\n\
Private Declare PtrSafe Function FillA Lib \"libc.so.6\" Alias \"memset\" (ByRef r As TText, ByVal c As Long, ByVal n As LongPtr) As LongPtr\n\
Public result As String\n\
Sub Main()\n\
    Dim t As TText\n\
    t.f = \"abcd\": t.s = \"h\" & ChrW(233) & \"llo\": t.n = 7\n\
    result = CStr(StrLenA(t))\n\
    FillA t, 65, 2\n\
    result = result & \"|\" & t.f & \"|\" & t.s & \"|\" & CStr(t.n)\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "4|AAcd|h\u{e9}llo|7", "{name}");
    }
}

#[test]
fn a_udt_array_is_passed_as_an_fadf_record_safearray_and_copied_back() {
    // Passing a UDT array to a Declare off Windows projects it to an FADF_RECORD SAFEARRAY
    // backed by our portable registration-free IRecordInfo; both 1D and 2D UDT arrays with
    // Strings and Variants cross the Declare boundary and copy back cleanly on VM3 and JIT.
    let source = "Private Type TItem\n    id As Long\n    tag As String * 4\n    label As String\n    extra As Variant\nEnd Type\n\
Private Declare PtrSafe Sub ReadPsa Lib \"libc.so.6\" Alias \"memcpy\" (ByRef dst As LongPtr, ByRef arr() As TItem, ByVal n As LongPtr)\n\
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

#[test]
fn call_site_byval_overrides_byref_as_any_declare_parameters() {
    // Wasabi's LoadThunk / SecureZeroString / sock_select and Riff's DispCallFunc pass
    // `ByVal pMem`, `ByVal StrPtr(s)`, `ByVal 0&`, and `ByVal "abc"` to `ByRef ... As Any`
    // parameters, while `CopyMemory opcodes(16), VarPtr(x), 8` passes a non-lvalue WITHOUT
    // `ByVal` so the pointer value itself is copied from a temporary cell.
    let source = "Private Declare PtrSafe Sub CopyMemoryFromPtr Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByRef dest As Any, ByVal src As LongPtr, ByVal size As Long)\n\
Private Declare PtrSafe Sub CopyMemory Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByRef dest As Any, ByRef src As Any, ByVal size As Long)\n\
Private Declare PtrSafe Function StrLenAny Lib \"libc.so.6\" Alias \"strlen\" (ByRef s As Any) As LongPtr\n\
Private Declare PtrSafe Function MemCpyPtr Lib \"libc.so.6\" Alias \"memcpy\" (ByRef dst As Any, ByRef src As Any, ByVal n As LongPtr) As LongPtr\n\
Public result As String\n\
Private Function GetBytes() As Byte()\n\
    Dim opcodes(0 To 59) As Byte\n\
    opcodes(0) = &H56: opcodes(1) = &H57: opcodes(59) = &HC3\n\
    GetBytes = opcodes\n\
End Function\n\
Private Sub CopyIntoPtr(ByRef opcodes() As Byte, ByVal pDest As LongPtr)\n\
    CopyMemoryFromPtr ByVal pDest, VarPtr(opcodes(LBound(opcodes))), UBound(opcodes) - LBound(opcodes) + 1\n\
End Sub\n\
Sub Main()\n\
    Dim dst(0 To 59) As Byte\n\
    CopyIntoPtr GetBytes(), VarPtr(dst(0))\n\
    Dim s As String\n\
    s = \"AB\"\n\
    Dim zeros(0 To 3) As Byte\n\
    CopyMemory ByVal StrPtr(s), zeros(0), LenB(s)\n\
    Dim nullRet As LongPtr\n\
    nullRet = MemCpyPtr(ByVal 0&, dst(0), 0)\n\
    Dim ptrRet As LongPtr\n\
    ptrRet = MemCpyPtr(ByVal VarPtr(dst(0)), dst(0), 0)\n\
    Dim textLen As LongPtr\n\
    textLen = StrLenAny(ByVal \"hello\")\n\
    Dim marker As Long\n\
    Dim rawPtrBytes(0 To 7) As Byte\n\
    Dim copiedPtr As LongPtr\n\
    CopyMemory rawPtrBytes(0), VarPtr(marker), 8\n\
    CopyMemory copiedPtr, rawPtrBytes(0), 8\n\
    result = CStr(dst(0)) & \",\" & CStr(dst(1)) & \",\" & CStr(dst(59)) & \"|\" & CStr(AscW(Left$(s, 1))) & \",\" & CStr(AscW(Right$(s, 1))) & \"|\" & CStr(nullRet) & \",\" & CStr(ptrRet = VarPtr(dst(0))) & \"|\" & CStr(textLen) & \"|\" & CStr(copiedPtr = VarPtr(marker))\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "86,87,195|0,0|0,True|5|True", "{name}");
    }
}

#[test]
fn wasabi_safe_array_len_and_wsadata_record_marshal_across_declare() {
    // Wasabi's SafeArrayLen uses `VarPtrArray` (VBE7 Alias "VarPtr") on a `ByRef arr() As Byte`
    // parameter and reads the SAFEARRAY* through `CopyMemoryFromPtr`, and `WSAStartup` passes
    // a `WSADATA` UDT with fixed Byte arrays and a LongPtr field.
    let source = "Private Declare PtrSafe Function VarPtrArray Lib \"VBE7\" Alias \"VarPtr\" (ByRef arr() As Byte) As LongPtr\n\
Private Declare PtrSafe Sub CopyMemoryFromPtr Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByRef dest As Any, ByVal src As LongPtr, ByVal size As Long)\n\
Private Declare PtrSafe Sub FillWsa Lib \"libc.so.6\" Alias \"memset\" (ByRef wsa As WSADATA, ByVal c As Long, ByVal n As LongPtr)\n\
Private Type WSADATA\n\
    wVersion As Integer\n\
    wHighVersion As Integer\n\
    szDescription(0 To 256) As Byte\n\
    szSystemStatus(0 To 128) As Byte\n\
    iMaxSockets As Integer\n\
    iMaxUdpDg As Integer\n\
    lpVendorInfo As LongPtr\n\
End Type\n\
Public result As String\n\
Private Function SafeArrayLen(ByRef arr() As Byte) As Long\n\
    Dim pSA As LongPtr\n\
    Dim ptr As LongPtr\n\
    ptr = VarPtrArray(arr)\n\
    If ptr = 0 Then Exit Function\n\
    CopyMemoryFromPtr pSA, ptr, LenB(pSA)\n\
    If pSA = 0 Then Exit Function\n\
    SafeArrayLen = UBound(arr) - LBound(arr) + 1\n\
End Function\n\
Sub Main()\n\
    Dim b() As Byte\n\
    Dim n0 As Long, n1 As Long, n2 As Long\n\
    n0 = SafeArrayLen(b)\n\
    ReDim b(0 To 11)\n\
    n1 = SafeArrayLen(b)\n\
    Erase b\n\
    n2 = SafeArrayLen(b)\n\
    Dim wsa As WSADATA\n\
    FillWsa wsa, &H22, 6\n\
    result = CStr(n0) & \",\" & CStr(n1) & \",\" & CStr(n2) & \"|\" & Hex$(wsa.wVersion) & \",\" & Hex$(wsa.wHighVersion) & \",\" & CStr(wsa.szDescription(0)) & \",\" & CStr(wsa.szDescription(1)) & \",\" & CStr(wsa.szDescription(2)) & \"|\" & CStr(LenB(wsa))\n\
End Sub\n";
    for (name, values) in run(source) {
        assert_eq!(text(&values), "0,12,0|2222,2222,34,34,0|408", "{name}");
    }
}
