//! `ReDim` and `ReDim Preserve` keep an array's descriptor, as VBA's `SafeArrayRedim` does: the
//! `SAFEARRAY*` in the variable stays the same and `pvData` is new, for typed, String and Variant
//! arrays and for `ReDim Preserve` of a Variant holding an array. A rank change, and a plain
//! `ReDim` of a Variant, get a new descriptor; `Erase` leaves the slot null. Source and expected
//! output are Excel x64's (`temp/variant-layout-probe/redim7c/SlotProbe.bas`, captured
//! 2026-11-14 in one owned Excel; the same lines as `RedimProbe.bas` through `ArrPtr`). The probe
//! reads memory through `CopyMemory`, so it runs with the native Declare lane. It finds an array's
//! `SAFEARRAY*` slot through a ByRef Variant parameter (`VT_BYREF | VT_ARRAY`, payload = the
//! slot) rather than VBE7's `VarPtr`, which loads only inside Office; in Excel that slot is
//! `ArrPtr`'s for Long and Variant arrays, and for a String array another slot holding the same
//! descriptor.
#![cfg(windows)]

use calcvba_engine::{Engine, HostConfig};
use calcvba_hal::HostPolicy;

const MODULE: &str = r#"
Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal n As LongPtr)

Private Function Ptr(ByVal addr As LongPtr) As LongPtr
    Dim p As LongPtr
    CopyMemory VarPtr(p), addr, 8
    Ptr = p
End Function

Private Function SlotOf(v As Variant) As LongPtr
    SlotOf = Ptr(VarPtr(v) + 8)
End Function

Private Function Feat(ByVal psa As LongPtr) As String
    Dim f As Integer
    If psa = 0 Then Feat = "null": Exit Function
    CopyMemory VarPtr(f), psa + 2, 2
    Feat = "&H" & Hex$(f)
End Function

Public Function RunAll() As String
    Dim out As String, d() As Long, s() As String, va() As Variant, v As Variant
    Dim psa As LongPtr, pv As LongPtr, slot As LongPtr

    ReDim d(0 To 2)
    slot = SlotOf(d): psa = Ptr(slot): pv = Ptr(psa + 16)
    ReDim Preserve d(0 To 9)
    out = out & "PRESERVE_GROW psa_same=" & (psa = Ptr(SlotOf(d))) & " pv_same=" & (pv = Ptr(Ptr(SlotOf(d)) + 16)) & " slot_same=" & (slot = SlotOf(d)) & vbLf
    psa = Ptr(SlotOf(d)): pv = Ptr(psa + 16)
    ReDim Preserve d(0 To 1)
    out = out & "PRESERVE_SHRINK psa_same=" & (psa = Ptr(SlotOf(d))) & " pv_same=" & (pv = Ptr(Ptr(SlotOf(d)) + 16)) & vbLf
    psa = Ptr(SlotOf(d)): pv = Ptr(psa + 16)
    ReDim d(0 To 1)
    out = out & "REDIM_SAME psa_same=" & (psa = Ptr(SlotOf(d))) & " pv_same=" & (pv = Ptr(Ptr(SlotOf(d)) + 16)) & vbLf
    psa = Ptr(SlotOf(d)): pv = Ptr(psa + 16)
    ReDim d(0 To 50)
    out = out & "REDIM_GROW psa_same=" & (psa = Ptr(SlotOf(d))) & " pv_same=" & (pv = Ptr(Ptr(SlotOf(d)) + 16)) & vbLf
    psa = Ptr(SlotOf(d))
    ReDim d(1 To 2, 1 To 3)
    out = out & "REDIM_RANK psa_same=" & (psa = Ptr(SlotOf(d))) & " feat=" & Feat(Ptr(SlotOf(d))) & vbLf
    psa = Ptr(SlotOf(d))
    Erase d
    out = out & "ERASE slot=" & Ptr(SlotOf(d)) & vbLf
    ReDim d(0 To 1)
    out = out & "REDIM_AFTER_ERASE nonnull=" & (Ptr(SlotOf(d)) <> 0) & vbLf

    ReDim s(0 To 2): s(1) = "keep"
    psa = Ptr(SlotOf(s))
    ReDim Preserve s(0 To 5)
    out = out & "STR_PRESERVE psa_same=" & (psa = Ptr(SlotOf(s))) & " s1=" & s(1) & vbLf
    psa = Ptr(SlotOf(s))
    ReDim s(0 To 5)
    out = out & "STR_REDIM psa_same=" & (psa = Ptr(SlotOf(s))) & " s1=[" & s(1) & "]" & vbLf

    ReDim va(0 To 2)
    psa = Ptr(SlotOf(va))
    ReDim Preserve va(0 To 5)
    out = out & "VAR_PRESERVE psa_same=" & (psa = Ptr(SlotOf(va))) & vbLf

    v = Array(1, 2, 3)
    psa = Ptr(VarPtr(v) + 8)
    ReDim Preserve v(0 To 5)
    out = out & "VARIANT_PRESERVE psa_same=" & (psa = Ptr(VarPtr(v) + 8)) & " vt=" & VarType(v) & vbLf
    psa = Ptr(VarPtr(v) + 8)
    ReDim v(0 To 5)
    out = out & "VARIANT_REDIM psa_same=" & (psa = Ptr(VarPtr(v) + 8)) & " vt=" & VarType(v) & vbLf
    RunAll = out
End Function


"#;

const EXPECTED: &str = "PRESERVE_GROW psa_same=True pv_same=False slot_same=True
PRESERVE_SHRINK psa_same=True pv_same=False
REDIM_SAME psa_same=True pv_same=False
REDIM_GROW psa_same=True pv_same=False
REDIM_RANK psa_same=False feat=&H80
ERASE slot=0
REDIM_AFTER_ERASE nonnull=True
STR_PRESERVE psa_same=True s1=keep
STR_REDIM psa_same=True s1=[]
VAR_PRESERVE psa_same=True
VARIANT_PRESERVE psa_same=True vt=8204
VARIANT_REDIM psa_same=False vt=8204
";

#[test]
fn redim_keeps_the_descriptor_on_vm3_and_jit() {
    let source = format!(
        "Public result As String{nl}{MODULE}{nl}Sub Main(){nl}    result = RunAll(){nl}End Sub{nl}",
        nl = "\n"
    );
    for (name, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(HostPolicy::native());
        let values = engine
            .execute_source_with_variant_snapshot_clean(&source)
            .unwrap_or_else(|e| panic!("{name}: {e:?}"));
        let out = values
            .iter()
            .find_map(|v| {
                v.as_bstr()
                    .map(|s| s.to_string())
                    .filter(|s| s.contains("PRESERVE_GROW"))
            })
            .unwrap_or_else(|| panic!("{name}: no output in {values:?}"));
        assert_eq!(out, EXPECTED, "{name}");
    }
}
