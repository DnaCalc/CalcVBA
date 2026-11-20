//! The descriptors VBA code builds carry the `fFeatures`, `cbElements` and hidden bytes Excel
//! writes, on both backends. Expected values are Excel x64's
//! (`temp/variant-layout-probe/arr6b/FeatProbe.bas`, captured 2026-11-14 in one owned Excel):
//! a fixed-size array adds `FADF_STATIC | FADF_FIXEDSIZE` (`&H12`); typed arrays have
//! `FADF_HAVEVARTYPE` with the vt at `psa-4`; String `&H180`, Variant `&H880`; object arrays
//! `FADF_HAVEIID | FADF_DISPATCH` (`&H440`) with the IID at `psa-16`; UDT arrays name no element
//! type (`0`, fixed `&H12`); and a copy of a fixed array into a Variant is dynamic (`&H80`).
use calcvba_engine::{Engine, HostConfig};
use calcvba_runtime::{
    Variant,
    com::{Guid, IID_VBA_COLLECTION},
    oleaut,
    safe_array::SafeArray,
};
use calcvba_symbol::manifest as sym;

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

fn read_iid_via_oleaut(array: &SafeArray) -> [u8; 16] {
    let mut guid = Guid::default();
    // SAFETY: `array` is a live FADF_HAVEIID descriptor and `guid` is writable.
    let hr = unsafe { oleaut::safe_array_get_iid(array.raw_safearray_ptr().cast(), &raw mut guid) };
    assert_eq!(hr, 0, "SafeArrayGetIID failed with HRESULT {hr:#010X}");
    guid.to_le_bytes()
}

#[test]
fn class_typed_arrays_carry_class_iid_not_idispatch_on_vm3_and_jit() {
    let pointer = core::mem::size_of::<usize>() as u32;
    let project_name = "ClassIidProject";
    let widget_iid = Guid::deterministic(project_name, "IWidget").to_le_bytes();
    let gadget_iid = Guid::deterministic(project_name, "IGadget").to_le_bytes();
    let collection_iid = IID_VBA_COLLECTION.to_le_bytes();
    assert_ne!(widget_iid, IID_IDISPATCH);
    assert_ne!(gadget_iid, IID_IDISPATCH);
    assert_ne!(widget_iid, gadget_iid);
    assert_ne!(collection_iid, IID_IDISPATCH);

    let main_source = r#"
Option Explicit
Dim fo(0 To 1) As Object
Dim dob() As Object
Dim fc(0 To 1) As Collection
Dim dc() As Collection
Dim fw(0 To 1) As Widget
Dim dw() As Widget
Dim dg() As Gadget
Dim vc As Variant
Dim vw As Variant
Dim mismatch_err As Long

Sub Main()
    ReDim dob(0 To 1)
    ReDim dc(0 To 1)
    ReDim dw(0 To 1) As widget
    ReDim dg(0 To 1)
    vc = fc
    ReDim Preserve vc(0 To 2)
    ReDim vw(0 To 1) As Widget
    ReDim Preserve vw(0 To 2)
    On Error Resume Next
    ReDim Preserve vw(0 To 3) As Gadget
    mismatch_err = Err.Number
    On Error GoTo 0
End Sub
"#;
    let manifest = sym::SymbolProjectManifest {
        project_name: project_name.to_string(),
        project_kind: sym::ProjectKind::Source,
        modules: vec![
            sym::ModuleUnit {
                module_name: "Module1".to_string(),
                module_kind: sym::ModuleKind::Procedural,
                attributes: sym::ModuleAttributes::named("Module1"),
                source: main_source.to_string(),
            },
            sym::ModuleUnit {
                module_name: "Widget".to_string(),
                module_kind: sym::ModuleKind::Class,
                attributes: sym::ModuleAttributes::named("Widget"),
                source: "Public Tag As Long\n".to_string(),
            },
            sym::ModuleUnit {
                module_name: "Gadget".to_string(),
                module_kind: sym::ModuleKind::Class,
                attributes: sym::ModuleAttributes::named("Gadget"),
                source: "Public Code As Long\n".to_string(),
            },
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };

    for (name, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let engine = Engine::new(config);
        let values = engine
            .execute_project_closure_with_variant_snapshot(std::slice::from_ref(&manifest))
            .unwrap_or_else(|e| panic!("{name}: {e:?}"));
        let expected: [(&str, usize, u16, [u8; 16]); 9] = [
            ("fo (Object fixed)", 0, 0x452, IID_IDISPATCH),
            ("dob (Object dynamic)", 1, 0x440, IID_IDISPATCH),
            ("fc (Collection fixed)", 2, 0x452, collection_iid),
            ("dc (Collection dynamic)", 3, 0x440, collection_iid),
            ("fw (Widget fixed)", 4, 0x452, widget_iid),
            ("dw (Widget dynamic)", 5, 0x440, widget_iid),
            ("dg (Gadget dynamic)", 6, 0x440, gadget_iid),
            (
                "vc (Variant copy + Preserve Collection)",
                7,
                0x440,
                collection_iid,
            ),
            ("vw (Variant ReDim + Preserve Widget)", 8, 0x440, widget_iid),
        ];
        for (label, index, features, expected_iid) in expected {
            let arr = array(&values, index);
            let d = descriptor(&arr);
            assert_eq!(
                (d.features, d.cb_elements, d.hidden),
                (features, pointer, expected_iid),
                "{name} {label} (psa-16)"
            );
            assert_eq!(
                arr.iid(),
                Some(expected_iid),
                "{name} {label} (SafeArray::iid)"
            );
            assert_eq!(
                read_iid_via_oleaut(&arr),
                expected_iid,
                "{name} {label} (SafeArrayGetIID)"
            );
        }
        assert_eq!(
            values[9].as_i32(),
            Some(13),
            "{name} ReDim Preserve across different classes must raise error 13"
        );
    }
}
