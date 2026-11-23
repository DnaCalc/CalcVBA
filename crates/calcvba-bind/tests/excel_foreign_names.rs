//! Which member accesses compile against Excel's real type library, as Excel VBA decides them
//! (probe temp/variant-layout-probe/tw11, Excel 16.0 build 20527, one fresh workbook per line).
//!
//! A bracketed name after a dot (`x.[A1]`) is the declared member of that name, else the
//! interface's own `DISPID_EVALUATE` member when it declares one (`Worksheet`, `_Application`),
//! else a late call when the interface is extensible (`Workbook`, `Range`: it compiles and
//! raises 438), else "Method or data member not found". An unknown plain name compiles only on
//! an extensible interface. Document modules (`ThisWorkbook`, `Sheet1`) are closed types: the
//! project's class, never extensible, even when their host interface is.
#![cfg(windows)]

use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, ProjectReference, SymbolProjectManifest,
};
use calcvba_symbol::{CatalogTypeLibResolver, TypeLibResolver};
use std::cell::RefCell;
use std::collections::{BTreeMap, HashMap};

/// The catalog resolver, remembering each answer: one load of Excel's library for all cases.
#[derive(Default)]
struct Cached(RefCell<HashMap<String, Option<calcvba_com::TypeLibMetadataBlob>>>);
impl TypeLibResolver for Cached {
    fn resolve(
        &self,
        request: &calcvba_com::TypeLibResolveRequest,
    ) -> Option<calcvba_com::TypeLibMetadataBlob> {
        let key = format!("{request:?}");
        if let Some(blob) = self.0.borrow().get(&key) {
            return blob.clone();
        }
        let blob = CatalogTypeLibResolver.resolve(request);
        self.0.borrow_mut().insert(key, blob.clone());
        blob
    }
    fn interface_iid(
        &self,
        request: &calcvba_com::TypeLibResolveRequest,
        blob: &calcvba_com::TypeLibMetadataBlob,
    ) -> Option<calcvba_com::ComInterfaceIid> {
        CatalogTypeLibResolver.interface_iid(request, blob)
    }
}

/// `(case, statement, compiles in Excel VBA)`.
const CASES: &[(&str, &str, bool)] = &[
    ("tw_a1", "v = ThisWorkbook.[A1]", false),
    ("tw_foo", "v = ThisWorkbook.Foo", false),
    ("tw_name", "v = ThisWorkbook.[Name]", true),
    ("app_tw_a1", "v = Application.ThisWorkbook.[A1]", true),
    ("app_tw_foo", "v = Application.ThisWorkbook.Foo", true),
    ("xl_tw_a1", "v = Excel.ThisWorkbook.[A1]", true),
    (
        "wb_a1",
        "Dim wb As Workbook: Set wb = ThisWorkbook: v = wb.[A1]",
        true,
    ),
    (
        "wb_foo",
        "Dim wb As Workbook: Set wb = ThisWorkbook: v = wb.Foo",
        true,
    ),
    (
        "xlwb_a1",
        "Dim wb As Excel.Workbook: Set wb = ThisWorkbook: v = wb.[A1]",
        true,
    ),
    (
        "ws_a1",
        "Dim ws As Worksheet: Set ws = ActiveSheet: v = ws.[A1]",
        true,
    ),
    (
        "ws_foo",
        "Dim ws As Worksheet: Set ws = ActiveSheet: v = ws.Foo",
        false,
    ),
    ("sheet1_a1", "v = Sheet1.[A1]", true),
    ("sheet1_foo", "v = Sheet1.Foo", false),
    ("sheets_foo", "v = Sheets.Foo", false),
    ("app_foo", "v = Application.Foo", true),
    ("range_foo", "v = Range(\"A1\").Foo", true),
    ("range_a1", "v = Range(\"A1\").[A1]", true),
    ("names_foo", "v = Names.Foo", true),
];

fn document(name: &str, host: &str) -> ModuleUnit {
    let mut attributes = ModuleAttributes::named(name);
    attributes.vb_predeclared_id = true;
    attributes.vb_exposed = true;
    attributes.host_document_type = Some(host.into());
    ModuleUnit {
        module_name: name.into(),
        module_kind: ModuleKind::Document,
        attributes,
        source: String::new(),
    }
}

fn manifest(statement: &str) -> SymbolProjectManifest {
    let source = format!(
        "Option Explicit\r\nFunction Run() As String\r\n    Dim v As Variant\r\n    {statement}\r\nEnd Function\r\n"
    );
    SymbolProjectManifest {
        project_name: "VBAProject".into(),
        project_kind: ProjectKind::Source,
        modules: vec![
            document("ThisWorkbook", "Excel.Workbook"),
            document("Sheet1", "Excel.Worksheet"),
            ModuleUnit {
                module_name: "Probe".into(),
                module_kind: ModuleKind::Procedural,
                attributes: ModuleAttributes::named("Probe"),
                source,
            },
        ],
        references: vec![ProjectReference::TypeLibrary {
            name: "Excel".into(),
            guid: Some("{00020813-0000-0000-C000-000000000046}".into()),
            version_major: Some(1),
            version_minor: Some(9),
            lcid: Some(0),
            import_lib: None,
        }],
        reference_projects: Vec::new(),
        conditional_constants: BTreeMap::new(),
        conditional_compilation_target: Default::default(),
        codepage: None,
    }
}

#[test]
fn member_accesses_compile_as_excel_vba_decides() {
    let excel = calcvba_com::TypeLibResolveRequest {
        reference_name: "Excel".into(),
        requested_coclass: None,
        importlib_hint: None,
        libid_hint: Some("{00020813-0000-0000-C000-000000000046}".into()),
        major_version_hint: Some(1),
        minor_version_hint: Some(9),
        lcid_hint: Some(0),
    };
    let resolver = Cached::default();
    if resolver.resolve(&excel).is_none() {
        eprintln!("Excel's type library is not registered here; skipped");
        return;
    }
    let mut wrong = Vec::new();
    for (case, statement, compiles) in CASES {
        let outcome = calcvba_bind::bind_program(&manifest(statement), &resolver);
        if outcome.is_ok() != *compiles {
            wrong.push(format!(
                "{case}: Excel VBA {} it; CalcVBA {}",
                if *compiles { "compiles" } else { "refuses" },
                match &outcome {
                    Ok(_) => "compiles it".to_string(),
                    Err(err) => format!("refuses it ({err})"),
                }
            ));
        }
    }
    assert!(wrong.is_empty(), "{}", wrong.join("\n"));
}
