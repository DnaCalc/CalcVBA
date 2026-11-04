use calcvba_office::{ModuleKind, OfficeFile, Project};
use std::{
    io::{Cursor, Read},
    path::PathBuf,
};
fn stream(cfb: &mut cfb::CompoundFile<Cursor<Vec<u8>>>, name: &str) -> Vec<u8> {
    let mut bytes = Vec::new();
    cfb.open_stream(name)
        .unwrap()
        .read_to_end(&mut bytes)
        .unwrap();
    bytes
}
#[test]
fn all_four_native_workbook_formats_keep_documents_forms_and_control_references() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/formats");
    for extension in ["xlsm", "xlsb", "xls", "xlam"] {
        let office = OfficeFile::open(root.join(format!("FormsWorkbook.{extension}"))).unwrap();
        let project = &office.project;
        let form = project
            .modules
            .iter()
            .find(|m| m.kind == ModuleKind::Form)
            .unwrap();
        assert_eq!(form.name, "ExampleForm");
        assert!(form.source.contains("Answer = 42"));
        assert_eq!(
            project
                .modules
                .iter()
                .filter(|m| m.kind == ModuleKind::Document)
                .count(),
            2
        );
        let forms = project
            .references
            .iter()
            .find(|r| r.name == "MSForms")
            .unwrap();
        assert!(matches!(
            forms.kind,
            calcvba_office::ReferenceKind::Control { .. }
        ));
        let output = office.to_bytes().unwrap();
        let reloaded = OfficeFile::read(&output, &Default::default()).unwrap();
        assert_eq!(project.properties, reloaded.project.properties);
        assert_eq!(project.references, reloaded.project.references);
        assert_eq!(
            project.to_manifest().unwrap(),
            reloaded.project.to_manifest().unwrap()
        );
    }
}
#[test]
fn generated_form_pair_preserves_the_vbe_designer_storage_and_source() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/formats");
    let office = OfficeFile::open(root.join("FormsWorkbook.xlsm")).unwrap();
    let dir = std::env::temp_dir().join(format!("calcvba-form-export-{}", std::process::id()));
    let path = office.project.export_folder(&dir).unwrap();
    let original = std::fs::read(root.join("ExampleForm.frx")).unwrap();
    let generated = std::fs::read(dir.join("ExampleForm.frx")).unwrap();
    let native = cfb::CompoundFile::open(Cursor::new(original[24..].to_vec())).unwrap();
    let mut written = cfb::CompoundFile::open(Cursor::new(generated[24..].to_vec())).unwrap();
    assert_eq!(native.root_entry().clsid(), written.root_entry().clsid());
    // VBE Export reserializes the form. The folder must preserve the exact
    // designer streams from the saved workbook, including their padding.
    let mut package = zip::ZipArchive::new(Cursor::new(
        std::fs::read(root.join("FormsWorkbook.xlsm")).unwrap(),
    ))
    .unwrap();
    let mut bin = Vec::new();
    package
        .by_name("xl/vbaProject.bin")
        .unwrap()
        .read_to_end(&mut bin)
        .unwrap();
    let mut saved = cfb::CompoundFile::open(Cursor::new(bin)).unwrap();
    for entry in saved
        .walk_storage("/ExampleForm")
        .unwrap()
        .filter(|e| e.is_stream() && e.name() != "\u{3}VBFrame")
        .collect::<Vec<_>>()
    {
        assert_eq!(
            stream(&mut saved, &entry.path().to_string_lossy()),
            stream(&mut written, &format!("/{}", entry.name()))
        );
    }
    let imported = Project::from_folder(path).unwrap();
    assert_eq!(
        office.project.to_manifest().unwrap(),
        imported.to_manifest().unwrap()
    );
    let source = std::fs::read_to_string(dir.join("ExampleForm.frm")).unwrap();
    assert!(source.contains("Café form"));
    assert!(source.contains("OleObjectBlob"));
    std::fs::remove_dir_all(dir).unwrap();
}
