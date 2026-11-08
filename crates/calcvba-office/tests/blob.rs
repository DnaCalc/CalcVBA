//! The blob route a host takes: it keeps the package and hands over only `vbaProject.bin`.
use calcvba_office::{ModuleKind, OfficeFile, Project};
use std::{io::Read, path::PathBuf};

fn part(case: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join(format!("tests/fixtures/{case}/workbook.xlsm"));
    let mut zip = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
    let mut bytes = Vec::new();
    zip.by_name("xl/vbaProject.bin")
        .unwrap()
        .read_to_end(&mut bytes)
        .unwrap();
    bytes
}

#[test]
fn a_blob_reads_as_the_package_reads() {
    for case in ["dictionary", "collection", "events"] {
        let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join(format!("tests/fixtures/{case}/workbook.xlsm"));
        let office = OfficeFile::open(path).unwrap();
        let blob = Project::read_bin(&part(case)).unwrap();
        assert_eq!(blob.properties, office.project.properties, "{case}");
        assert_eq!(blob.references, office.project.references, "{case}");
        assert_eq!(blob.modules, office.project.modules, "{case}");
        assert_eq!(
            blob.to_manifest().unwrap().modules,
            office.project.to_manifest().unwrap().modules,
            "{case}"
        );
    }
}

#[test]
fn put_module_replaces_code_and_adds_standard_and_class_modules() {
    let mut project = Project::read_bin(&part("dictionary")).unwrap();
    let before = project.modules.len();
    project
        .put_module(
            "Api",
            ModuleKind::Standard,
            "Public Function RunProbe() As String\n    RunProbe = \"changed\"\nEnd Function\n",
        )
        .unwrap();
    project
        .put_module(
            "Extra",
            ModuleKind::Standard,
            "Public Function Twice(x As Double) As Double\r\n    Twice = 2 * x\r\nEnd Function\r\n",
        )
        .unwrap();
    project
        .put_module("Holder", ModuleKind::Class, "Public Value As Variant\r\n")
        .unwrap();
    let written = Project::read_bin(&project.to_bin().unwrap()).unwrap();
    assert_eq!(written.modules.len(), before + 2);
    let module = |name: &str| written.modules.iter().find(|m| m.name == name).unwrap();
    assert_eq!(module("Api").kind, ModuleKind::Standard);
    assert!(module("Api").code().contains("RunProbe = \"changed\""));
    assert_eq!(module("Api").attribute("VB_Name"), Some("\"Api\""));
    assert_eq!(module("Extra").kind, ModuleKind::Standard);
    assert!(module("Extra").code().contains("Twice = 2 * x"));
    assert_eq!(module("Holder").kind, ModuleKind::Class);
    assert_eq!(module("Holder").attribute("VB_Exposed"), Some("False"));
    assert!(module("Holder").code().contains("Public Value As Variant"));
    // The document modules and references are the original's.
    let original = Project::read_bin(&part("dictionary")).unwrap();
    assert_eq!(written.references, original.references);
    for document in original
        .modules
        .iter()
        .filter(|m| m.kind == ModuleKind::Document)
    {
        let written = module(&document.name);
        assert_eq!(
            (&written.kind, &written.source, &written.document_version),
            (&document.kind, &document.source, &document.document_version)
        );
    }
    let manifest = written.to_manifest().unwrap();
    assert!(manifest.modules.iter().any(|m| m.module_name == "Holder"));
}

#[test]
fn put_module_refuses_attributes_kind_changes_and_new_documents() {
    let mut project = Project::read_bin(&part("dictionary")).unwrap();
    assert!(
        project
            .put_module(
                "Api",
                ModuleKind::Standard,
                "Attribute VB_Name = \"Api\"\r\n"
            )
            .is_err()
    );
    assert!(project.put_module("Api", ModuleKind::Class, "").is_err());
    assert!(
        project
            .put_module("Sheet9", ModuleKind::Document, "")
            .is_err()
    );
    assert!(
        project
            .put_module("Bad/Name", ModuleKind::Standard, "")
            .is_err()
    );
    // A refused change leaves the project as it was.
    assert_eq!(
        project.modules,
        Project::read_bin(&part("dictionary")).unwrap().modules
    );
}
