use calcvba_office::{Error, OfficeFile, Project};
use std::path::PathBuf;
struct Scratch(PathBuf);
impl Scratch {
    fn new() -> Self {
        static NEXT: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
        let root = std::env::temp_dir().join(format!(
            "calcvba-office-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
        ));
        std::fs::create_dir(&root).unwrap();
        Self(root)
    }
}
impl Drop for Scratch {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}
#[test]
fn workbook_folder_workbook_preserves_all_sources_references_and_properties() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    let scratch = Scratch::new();
    for case in std::fs::read_dir(root).unwrap() {
        let path = case.unwrap().path();
        if !path.is_dir() {
            continue;
        }
        let mut office = OfficeFile::open(path.join("workbook.xlsm")).unwrap();
        let original = &office.project;
        let folder = scratch.0.join(path.file_name().unwrap());
        let manifest = original.export_folder(&folder).unwrap();
        let xml = std::fs::read_to_string(&manifest).unwrap();
        assert!(xml.starts_with("<Project Sdk=\"CalcVBA.Sdk\">"));
        assert!(!xml.contains("<VB"));
        let imported =
            Project::from_folder(&manifest).unwrap_or_else(|e| panic!("{}: {e}", path.display()));
        assert_eq!(original.properties, imported.properties);
        assert_eq!(original.references, imported.references);
        assert_eq!(
            original.to_manifest().unwrap(),
            imported.to_manifest().unwrap()
        );
        office.project = imported;
        let bytes = office.to_bytes().unwrap();
        let loaded = OfficeFile::read(&bytes, &Default::default()).unwrap();
        assert_eq!(office.project.properties, loaded.project.properties);
        assert_eq!(office.project.references, loaded.project.references);
        assert_eq!(
            office.project.to_manifest().unwrap(),
            loaded.project.to_manifest().unwrap()
        );
        // The ordinary compiler's .basproj loader sees the same source roles,
        // names, host documents and header attributes as the Office reader.
        let closure = calcvba_project::load_project_closure(&manifest).unwrap();
        let compiled = closure.last().unwrap();
        for expected in &loaded.project.to_manifest().unwrap().modules {
            let actual = compiled
                .modules
                .iter()
                .find(|m| m.module_name == expected.module_name)
                .unwrap();
            assert_eq!(actual.module_kind, expected.module_kind);
            assert_eq!(actual.attributes, expected.attributes);
        }
    }
}
#[test]
fn utf8_disk_text_is_checked_against_the_project_codepage_at_write_back() {
    let path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/strings/workbook.xlsm");
    let mut office = OfficeFile::open(path).unwrap();
    let scratch = Scratch::new();
    let manifest = office
        .project
        .export_folder(scratch.0.join("project"))
        .unwrap();
    let api = manifest.parent().unwrap().join("Api.bas");
    let mut source = std::fs::read_to_string(&api).unwrap();
    source = source.replacen(
        "Attribute VB_Name = \"Api\"",
        "Attribute VB_Name = \"Api\"\r\nAttribute VB_Description = \"\"\"quoted\"\"\"",
        1,
    );
    source.push_str("\r\n' café and € are representable in codepage 1252\r\n");
    std::fs::write(&api, &source).unwrap();
    office.project = Project::from_folder(&manifest).unwrap();
    assert_eq!(
        office
            .project
            .modules
            .iter()
            .find(|m| m.name == "Api")
            .unwrap()
            .description,
        "\"quoted\""
    );
    let bytes = office.to_bytes().unwrap();
    assert!(
        OfficeFile::read(&bytes, &Default::default())
            .unwrap()
            .project
            .modules
            .iter()
            .any(|m| m.source.contains("café"))
    );
    source.push_str("' 漢 cannot be written in codepage 1252\r\n");
    std::fs::write(&api, &source).unwrap();
    office.project = Project::from_folder(&manifest).unwrap();
    let output = scratch.0.join("unrepresentable.xlsm");
    assert!(matches!(
        office.save_as(&output),
        Err(Error::Encoding(1252))
    ));
    assert!(!output.exists());
}

#[test]
fn numeric_project_arguments_and_quoted_properties_survive_the_folder_round_trip() {
    let path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/formats/MetadataWorkbook.xlsm");
    let mut office = OfficeFile::open(path).unwrap();
    assert_eq!(
        office.project.properties.constants,
        "Enabled = -1 : Number = 255"
    );
    assert_eq!(office.project.properties.description, "Café \"project\"");
    assert_eq!(office.project.properties.help_file, "Guide.chm");
    assert_eq!(office.project.properties.help_context, 1873);
    let description = office
        .project
        .project_stream
        .lines()
        .find(|s| s.starts_with("Description="))
        .unwrap()
        .to_owned();
    let scratch = Scratch::new();
    let manifest = office
        .project
        .export_folder(scratch.0.join("project"))
        .unwrap();
    office.project = Project::from_folder(manifest).unwrap();
    let written = OfficeFile::read(&office.to_bytes().unwrap(), &Default::default()).unwrap();
    assert_eq!(written.project.properties, office.project.properties);
    assert_eq!(
        written
            .project
            .project_stream
            .lines()
            .find(|s| s.starts_with("Description=")),
        Some(description.as_str())
    );
    assert_eq!(
        written.project.to_manifest().unwrap().conditional_constants["Number"],
        255
    );
}
