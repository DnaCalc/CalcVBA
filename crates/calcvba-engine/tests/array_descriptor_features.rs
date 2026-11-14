//! The descriptors VBA code builds carry the `fFeatures`, `cbElements` and hidden bytes Excel
//! writes, on both backends. Expected values are Excel x64's
//! (`temp/variant-layout-probe/arr6b/FeatProbe.bas`, captured 2026-11-14 in one owned Excel):
//! a fixed-size array adds `FADF_STATIC | FADF_FIXEDSIZE` (`&H12`); typed arrays have
//! `FADF_HAVEVARTYPE` with the vt at `psa-4`; String `&H180`, Variant `&H880`; object arrays
//! `FADF_HAVEIID | FADF_DISPATCH` (`&H440`) with the IID at `psa-16`; UDT arrays name no element
//! type (`0`, fixed `&H12`); and a copy of a fixed array into a Variant is dynamic (`&H80`).
use calcvba_engine::{Engine, HostConfig};
use calcvba_runtime::{Variant, safe_array::SafeArray};

const SOURCE: &str = r#"
Private Type TPair
    a As Long
    b As Integer
End Type

Private Type TStr
    s As String
    n As Long
End Type

Dim fl(0 To 1) As Long
Dim dl() As Long
Dim fv(0 To 1) As Variant
Dim dv() As Variant
Dim fs(0 To 1) As String
Dim ds() As String
Dim fo(0 To 1) As Object
Dim dob() As Object
Dim fu(0 To 1) As TPair
Dim du() As TPair
Dim dts() As TStr
Dim db() As Byte
Dim v As Variant

Sub Main()
    ReDim dl(0 To 1)
    ReDim dv(0 To 1)
    ReDim ds(0 To 1)
    ReDim dob(0 To 1)
    ReDim du(0 To 1)
    ReDim dts(0 To 1)
    ReDim db(0 To 1)
    v = fl
End Sub
"#;

struct Descriptor {
    features: u16,
    cb_elements: u32,
    hidden: [u8; 16],
}

fn descriptor(array: &SafeArray) -> Descriptor {
    let psa = array.raw_safearray_ptr().cast::<u8>();
    let mut hidden = [0u8; 16];
    // SAFETY: `array` is a live descriptor with oleaut32's 16 hidden bytes before it;
    // `fFeatures` is at 2 and `cbElements` at 4.
    unsafe {
        hidden.copy_from_slice(core::slice::from_raw_parts(psa.sub(16), 16));
        Descriptor {
            features: psa.add(2).cast::<u16>().read_unaligned(),
            cb_elements: psa.add(4).cast::<u32>().read_unaligned(),
            hidden,
        }
    }
}

fn array(values: &[Variant], index: usize) -> SafeArray {
    values[index]
        .as_safearray()
        .unwrap_or_else(|| panic!("global {index} is an array: {:?}", values[index]))
}

const IID_IDISPATCH: [u8; 16] = [0, 4, 2, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46];

#[test]
fn vba_array_descriptors_carry_the_features_excel_writes_on_vm3_and_jit() {
    let pointer = core::mem::size_of::<usize>() as u32;
    let variant = core::mem::size_of::<Variant>() as u32;
    for (name, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let engine = Engine::new(config);
        let values = engine
            .execute_source_with_variant_snapshot_clean(SOURCE)
            .unwrap_or_else(|e| panic!("{name}: {e:?}"));
        let vt = |d: &Descriptor| u32::from_le_bytes(d.hidden[12..].try_into().unwrap());
        // (global, fFeatures, cbElements, vt at psa-4 when FADF_HAVEVARTYPE)
        let typed: [(&str, usize, u16, u32, u32); 9] = [
            ("fl", 0, 0x92, 4, 3),
            ("dl", 1, 0x80, 4, 3),
            ("fv", 2, 0x892, variant, 0x0C),
            ("dv", 3, 0x880, variant, 0x0C),
            ("fs", 4, 0x192, pointer, 8),
            ("ds", 5, 0x180, pointer, 8),
            ("db", 11, 0x80, 1, 0x11),
            ("v = fl", 12, 0x80, 4, 3),
            ("fl again", 0, 0x92, 4, 3),
        ];
        for (label, index, features, cb, element_vt) in typed {
            let d = descriptor(&array(&values, index));
            assert_eq!(
                (d.features, d.cb_elements, vt(&d)),
                (features, cb, element_vt),
                "{name} {label}"
            );
        }
        for (label, index, features) in [("fo", 6, 0x452), ("dob", 7, 0x440)] {
            let d = descriptor(&array(&values, index));
            assert_eq!(
                (d.features, d.cb_elements, d.hidden),
                (features, pointer, IID_IDISPATCH),
                "{name} {label}"
            );
        }
        for (label, index, features, cb) in [
            ("fu", 8, 0x12, 8),
            ("du", 9, 0, 8),
            ("dts", 10, 0, 2 * pointer),
        ] {
            let d = descriptor(&array(&values, index));
            assert_eq!(
                (d.features, d.cb_elements, d.hidden),
                (features, cb, [0; 16]),
                "{name} {label}"
            );
        }
        assert_eq!(values[12].core().raw_vt(), 0x2003, "{name} v = fl tag");
    }
}
