//! Arrays built by VBA code keep VBA's memory layout on both backends: elements stored first
//! dimension fastest and `rgsabound` listing the last dimension first. The expected numbers
//! are Excel's (`temp/variant-layout-probe/arr6/ArrProbe.bas`, captured 2026-11-14 in one owned
//! Excel x64): `f(0 To 1, 0 To 2) As Long` has `f(1,0)` at 4 and `f(0,1)` at 8 and
//! `rgsabound = {3,0},{2,0}`; `For Each` walks `00 10 01 11 02 12`.
use calcvba_engine::{Engine, HostConfig};
use calcvba_runtime::{Variant, safe_array::SafeArray};

const SOURCE: &str = r#"
Private Type TPair
    a As Long
    b As Integer
End Type

Dim f(0 To 1, 0 To 2) As Long
Dim d() As Long
Dim f3(0 To 1, 0 To 1, 0 To 1) As Byte
Dim t(0 To 1, 0 To 1) As TPair
Dim v As Variant
Dim order As String

Sub Main()
    Dim i As Long, j As Long, k As Long, item As Variant
    For i = 0 To 1
        For j = 0 To 2
            f(i, j) = 10 * i + j
        Next j
    Next i
    ReDim d(1 To 2, 0 To 3)
    For i = 1 To 2
        For j = 0 To 3
            d(i, j) = 10 * i + j
        Next j
    Next i
    ReDim Preserve d(1 To 2, 0 To 5)
    d(2, 5) = 99
    For i = 0 To 1
        For j = 0 To 1
            For k = 0 To 1
                f3(i, j, k) = 100 * i + 10 * j + k
            Next k
        Next j
    Next i
    For i = 0 To 1
        For j = 0 To 1
            t(i, j).a = 10 * i + j
            t(i, j).b = -1
        Next j
    Next i
    ReDim v(0 To 1, 0 To 2)
    For i = 0 To 1
        For j = 0 To 2
            v(i, j) = i & j
        Next j
    Next i
    For Each item In v
        order = order & item & " "
    Next item
End Sub
"#;

struct Descriptor {
    /// `rgsabound` as stored: `(cElements, lLbound)`, the last dimension first.
    rgsabound: Vec<(u32, i32)>,
    pv_data: *const u8,
}

/// Read the descriptor fields at their Automation offsets (x64: `pvData` at 16, `rgsabound`
/// at 24; 32-bit: 12 and 16).
fn descriptor(array: &SafeArray) -> Descriptor {
    let psa = array.raw_safearray_ptr().cast::<u8>();
    let pointer = std::mem::size_of::<usize>();
    let pv_offset = 12usize.next_multiple_of(pointer);
    // SAFETY: `array` is a live descriptor with `cDims` bounds; the offsets are the
    // Automation SAFEARRAY header layout this module's header follows.
    unsafe {
        let dims = psa.cast::<u16>().read() as usize;
        let pv_data = psa.add(pv_offset).cast::<*const u8>().read();
        let bounds = psa.add(pv_offset + pointer);
        let rgsabound = (0..dims)
            .map(|i| {
                let bound = bounds.add(8 * i);
                (
                    bound.cast::<u32>().read_unaligned(),
                    bound.add(4).cast::<i32>().read_unaligned(),
                )
            })
            .collect();
        Descriptor { rgsabound, pv_data }
    }
}

fn i32s(pv: *const u8, stride: usize, count: usize) -> Vec<i32> {
    // SAFETY: callers pass a live payload holding `count` elements of `stride` bytes whose
    // first four bytes are a Long.
    (0..count)
        .map(|i| unsafe { pv.add(i * stride).cast::<i32>().read_unaligned() })
        .collect()
}

fn array(values: &[Variant], index: usize) -> SafeArray {
    values[index]
        .as_safearray()
        .unwrap_or_else(|| panic!("global {index} is an array: {:?}", values[index]))
}

#[test]
fn vba_arrays_store_the_first_dimension_fastest_on_vm3_and_jit() {
    for (name, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let engine = Engine::new(config);
        let values = engine
            .execute_source_with_variant_snapshot_clean(SOURCE)
            .unwrap_or_else(|e| panic!("{name}: {e:?}"));

        let f = array(&values, 0);
        let fd = descriptor(&f);
        assert_eq!(fd.rgsabound, vec![(3, 0), (2, 0)], "{name} f rgsabound");
        assert_eq!(
            i32s(fd.pv_data, 4, 6),
            vec![0, 10, 1, 11, 2, 12],
            "{name} f payload"
        );

        let d = array(&values, 1);
        let dd = descriptor(&d);
        assert_eq!(dd.rgsabound, vec![(6, 0), (2, 1)], "{name} d rgsabound");
        assert_eq!(
            i32s(dd.pv_data, 4, 12),
            vec![10, 20, 11, 21, 12, 22, 13, 23, 0, 0, 0, 99],
            "{name} d payload after ReDim Preserve"
        );

        let f3 = array(&values, 2);
        let f3d = descriptor(&f3);
        assert_eq!(f3d.rgsabound, vec![(2, 0); 3], "{name} f3 rgsabound");
        // SAFETY: the Byte payload holds eight elements.
        let bytes = unsafe { std::slice::from_raw_parts(f3d.pv_data, 8) };
        assert_eq!(
            bytes,
            [0, 100, 10, 110, 1, 101, 11, 111],
            "{name} f3 payload"
        );

        let t = array(&values, 3);
        let td = descriptor(&t);
        assert_eq!(td.rgsabound, vec![(2, 0), (2, 0)], "{name} t rgsabound");
        assert_eq!(
            i32s(td.pv_data, 8, 4),
            vec![0, 10, 1, 11],
            "{name} t payload (LenB(TPair) = 8)"
        );

        let v = array(&values, 4);
        let vd = descriptor(&v);
        assert_eq!(vd.rgsabound, vec![(3, 0), (2, 0)], "{name} v rgsabound");
        let order: Vec<String> = (0..6)
            .map(|i| {
                v.variant_element(i)
                    .unwrap()
                    .as_bstr()
                    .expect("String element")
                    .to_string()
            })
            .collect();
        assert_eq!(
            order,
            ["00", "10", "01", "11", "02", "12"],
            "{name} v storage"
        );

        assert_eq!(
            values[5].as_bstr().map(|s| s.to_string()).as_deref(),
            Some("00 10 01 11 02 12 "),
            "{name} For Each order"
        );
    }
}
