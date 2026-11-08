//! The blob route a host takes: it keeps the package and hands over only `vbaProject.bin`.
use calcvba_office::{Error, ModuleKind, OfficeFile, Project, compression};
use std::{
    io::{Cursor, Read, Write},
    path::PathBuf,
};

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

/// The blob with one stream replaced.
fn with_stream(blob: &[u8], path: &str, bytes: &[u8]) -> Vec<u8> {
    let mut cfb = cfb::CompoundFile::open(Cursor::new(blob.to_vec())).unwrap();
    cfb.create_stream(path).unwrap().write_all(bytes).unwrap();
    cfb.into_inner().into_inner()
}

fn stream(blob: &[u8], path: &str) -> Vec<u8> {
    let mut cfb = cfb::CompoundFile::open(Cursor::new(blob.to_vec())).unwrap();
    let mut bytes = Vec::new();
    cfb.open_stream(path)
        .unwrap()
        .read_to_end(&mut bytes)
        .unwrap();
    bytes
}

// Ported from DNA Calc's calc-doc reader (removed when calc-doc began handing over the part as
// a blob): a truncated or padded directory and an unknown code page are errors, never panics.
#[test]
fn truncated_directories_and_an_unknown_code_page_are_errors_without_panics() {
    let blob = part("dictionary");
    let directory = compression::decompress(&stream(&blob, "/VBA/dir"), 1 << 20).unwrap();
    assert!(
        Project::read_bin(&with_stream(
            &blob,
            "/VBA/dir",
            &compression::compress(&directory).unwrap()
        ))
        .is_ok()
    );
    for end in 0..directory.len() {
        let truncated = with_stream(
            &blob,
            "/VBA/dir",
            &compression::compress(&directory[..end]).unwrap(),
        );
        assert!(Project::read_bin(&truncated).is_err(), "{end}");
    }
    // Bytes after the terminator are refused; zero padding is not (a raw final chunk carries it,
    // MS-OVBA 2.4.1.3.10).
    let mut trailing = directory.clone();
    trailing.push(1);
    assert!(
        Project::read_bin(&with_stream(
            &blob,
            "/VBA/dir",
            &compression::compress(&trailing).unwrap()
        ))
        .is_err()
    );
    let code_page_at = directory
        .windows(8)
        .position(|bytes| bytes[..6] == [3, 0, 2, 0, 0, 0])
        .unwrap()
        + 6;
    let mut unknown = directory;
    unknown[code_page_at..code_page_at + 2].copy_from_slice(&u16::MAX.to_le_bytes());
    assert!(matches!(
        Project::read_bin(&with_stream(
            &blob,
            "/VBA/dir",
            &compression::compress(&unknown).unwrap()
        )),
        Err(Error::Codepage(u16::MAX))
    ));
}

// Ported from the same reader: malformed compressed containers are refused before copying.
#[test]
fn malformed_compressed_streams_are_refused_before_copying() {
    for bytes in [&[][..], &[0], &[1, 0], &[1, 0, 0], &[1, 2, 0xb0, 1, 0, 0]] {
        assert!(compression::decompress(bytes, 4096).is_err(), "{bytes:?}");
    }
    // A literal A and an overlapping copy of three A bytes.
    let compressed = [1, 3, 0xb0, 2, b'A', 0, 0];
    assert_eq!(compression::decompress(&compressed, 4).unwrap(), b"AAAA");
    assert!(matches!(
        compression::decompress(&compressed, 3),
        Err(Error::Limit(_))
    ));
}

// Ported from the same reader: PROJECT's module declarations agree with the directory.
#[test]
fn project_declarations_that_conflict_with_the_directory_are_refused() {
    let blob = part("dictionary");
    let text = String::from_utf8(stream(&blob, "/PROJECT")).unwrap();
    assert!(text.contains("Module=Api"), "{text}");
    for replacement in [
        "Module=Api\r\nModule=Api",
        "Module=Api\r\nClass=Api",
        "Module=Api\r\nClass=Missing",
        "Module=Api\r\nBaseClass=Api",
        "Module=Api\r\nDocument=Api/&H00000000",
    ] {
        let changed = with_stream(
            &blob,
            "/PROJECT",
            text.replace("Module=Api", replacement).as_bytes(),
        );
        assert!(Project::read_bin(&changed).is_err(), "{replacement}");
    }
}
