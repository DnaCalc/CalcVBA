//! Differential parity tests for Excel host-stub execution and VBA operator precedence.

use calcvba_differential::{
    Canon,
    excel_stub::{ExcelStubTypeLibs, excel_project_reference},
    parity::{compare_entry, compare_entry_with_excel},
};
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};

fn compile_plain(source: &str) -> calcvba_oxir::OxProgram {
    let manifest = SymbolProjectManifest {
        project_name: "VBAProject".into(),
        project_kind: ProjectKind::Source,
        modules: vec![ModuleUnit {
            module_name: "Probe".into(),
            module_kind: ModuleKind::Procedural,
            attributes: ModuleAttributes::named("Probe"),
            source: source.into(),
        }],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let core = calcvba_bind::bind_program(&manifest, &calcvba_symbol::CatalogTypeLibResolver)
        .expect("bind fixture");
    calcvba_oxir::elaborate::elaborate(&core).expect("elaborate fixture")
}

fn compile_excel(source: &str) -> calcvba_oxir::OxProgram {
    let manifest = SymbolProjectManifest {
        project_name: "VBAProject".into(),
        project_kind: ProjectKind::Source,
        modules: vec![ModuleUnit {
            module_name: "Probe".into(),
            module_kind: ModuleKind::Procedural,
            attributes: ModuleAttributes::named("Probe"),
            source: source.into(),
        }],
        references: vec![excel_project_reference()],
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let core =
        calcvba_bind::bind_program(&manifest, &ExcelStubTypeLibs).expect("bind excel fixture");
    calcvba_oxir::elaborate::elaborate(&core).expect("elaborate excel fixture")
}

#[test]
fn power_operator_is_left_associative_like_excel_vba() {
    let programs = [compile_plain(
        r#"
Public A As Double
Public B As Double
Public C As Double
Public D As Double
Sub Main()
    A = -2 ^ 2
    B = (-2) ^ 2
    C = 2 ^ 3 ^ 2
    D = 2 ^ -2
End Sub
"#,
    )];
    let comparison = compare_entry(&programs);
    assert!(comparison.differences().is_empty(), "{comparison:?}");
    let expected = [
        Canon::Raw {
            tag: 5,
            bytes: (-4.0_f64).to_ne_bytes(),
            reserved: [0, 0, 0],
        },
        Canon::Raw {
            tag: 5,
            bytes: (4.0_f64).to_ne_bytes(),
            reserved: [0, 0, 0],
        },
        Canon::Raw {
            tag: 5,
            bytes: (64.0_f64).to_ne_bytes(),
            reserved: [0, 0, 0],
        },
        Canon::Raw {
            tag: 5,
            bytes: (0.25_f64).to_ne_bytes(),
            reserved: [0, 0, 0],
        },
    ];
    for run in [&comparison.vm3, &comparison.jit] {
        let values = run.outcome.result.as_ref().expect("completed run");
        assert_eq!(values.as_slice(), &expected, "{run:?}");
    }
}

#[test]
fn excel_stub_compares_range_worksheet_and_worksheetfunction_on_vm3_and_jit() {
    let programs = [compile_excel(
        r#"
Public AddrR1C1 As String
Public IntersectNothing As Boolean
Public AreasCount As Long
Public FirstAreaCols As Long
Public ResizedAddr As String
Public FoundAddr As String
Public Rounded As Double
Public ValueVarType As Long
Public Value2VarType As Long

Sub Main()
    Dim ws As Worksheet
    Dim r As Range
    Dim u As Range
    Dim hit As Range

    Set ws = ThisWorkbook.Worksheets(1)
    AddrR1C1 = ws.Range("D5").Address(False, False, xlR1C1, False, ws.Range("B2"))
    IntersectNothing = (Application.Intersect(ws.Range("A1:B2"), ws.Range("D4:E5")) Is Nothing)
    Set u = Union(ws.Range("A1:B2"), ws.Range("D1:F2"))
    AreasCount = u.Areas.Count
    FirstAreaCols = u.Columns.Count
    ResizedAddr = ws.Range("B2:D4").Resize(ColumnSize:=1).Address(False, False)

    ws.Range("A1").Value2 = "x"
    ws.Range("A2").Value2 = "y"
    Set r = ws.Range("A1:A2")
    Set hit = r.Find(What:="x", After:=ws.Range("A1"), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    FoundAddr = hit.Address(False, False)

    Rounded = Application.WorksheetFunction.Round(2.5, 0)

    ws.Range("C1").Value2 = 45000
    ws.Range("C1").NumberFormat = "yyyy-mm-dd"
    ValueVarType = VarType(ws.Range("C1").Value)
    Value2VarType = VarType(ws.Range("C1").Value2)
End Sub
"#,
    )];
    let comparison = compare_entry_with_excel(&programs);
    assert!(comparison.differences().is_empty(), "{comparison:?}");
    for run in [&comparison.vm3, &comparison.jit] {
        assert!(run.outcome.unsupported.is_none(), "{run:?}");
        assert!(run.outcome.result.is_ok(), "{run:?}");
        assert_eq!(run.outcome.err.number, 0, "{run:?}");
        let values = run.outcome.result.as_ref().unwrap();
        assert_eq!(
            values[0],
            Canon::StringBytes(
                "R[3]C[2]"
                    .encode_utf16()
                    .flat_map(u16::to_le_bytes)
                    .collect()
            )
        );
        assert_eq!(
            values[1],
            Canon::Raw {
                tag: 11,
                bytes: [255, 255, 0, 0, 0, 0, 0, 0],
                reserved: [0, 0, 0],
            }
        );
        assert_eq!(
            values[2],
            Canon::Raw {
                tag: 3,
                bytes: 2_i32
                    .to_ne_bytes()
                    .into_iter()
                    .chain([0; 4])
                    .collect::<Vec<_>>()
                    .try_into()
                    .unwrap(),
                reserved: [0, 0, 0],
            }
        );
        assert_eq!(
            values[3],
            Canon::Raw {
                tag: 3,
                bytes: 2_i32
                    .to_ne_bytes()
                    .into_iter()
                    .chain([0; 4])
                    .collect::<Vec<_>>()
                    .try_into()
                    .unwrap(),
                reserved: [0, 0, 0],
            }
        );
        assert_eq!(
            values[4],
            Canon::StringBytes("B2:B4".encode_utf16().flat_map(u16::to_le_bytes).collect())
        );
        assert_eq!(
            values[5],
            Canon::StringBytes("A1".encode_utf16().flat_map(u16::to_le_bytes).collect())
        );
        assert_eq!(
            values[6],
            Canon::Raw {
                tag: 5,
                bytes: 3.0_f64.to_ne_bytes(),
                reserved: [0, 0, 0],
            }
        );
        assert_eq!(
            values[7],
            Canon::Raw {
                tag: 3,
                bytes: 7_i32
                    .to_ne_bytes()
                    .into_iter()
                    .chain([0; 4])
                    .collect::<Vec<_>>()
                    .try_into()
                    .unwrap(),
                reserved: [0, 0, 0],
            }
        );
        assert_eq!(
            values[8],
            Canon::Raw {
                tag: 3,
                bytes: 5_i32
                    .to_ne_bytes()
                    .into_iter()
                    .chain([0; 4])
                    .collect::<Vec<_>>()
                    .try_into()
                    .unwrap(),
                reserved: [0, 0, 0],
            }
        );
    }
}
