use calcvba_office::{ModuleKind, OfficeFile, ReadLimits};
use std::{
    io::{Cursor, Read},
    path::PathBuf,
};

fn fixtures() -> Vec<PathBuf> {
    let mut result =
        std::fs::read_dir(PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures"))
            .unwrap()
            .map(|p| p.unwrap().path())
            .filter(|p| p.is_dir())
            .collect::<Vec<_>>();
    result.sort();
    result
}
fn normalize(source: &str) -> String {
    source.lines().collect::<Vec<_>>().join("\n").trim().into()
}

#[test]
fn all_saved_round_14_15_projects_match_their_vbe_snapshots() {
    let cases = fixtures();
    assert_eq!(cases.len(), 33);
    for case in cases {
        let office = OfficeFile::open(case.join("workbook.xlsm"))
            .unwrap_or_else(|e| panic!("{}: {e}", case.display()));
        let project = &office.project;
        let expected: serde_json::Value =
            serde_json::from_slice(&std::fs::read(case.join("project.json")).unwrap()).unwrap();
        assert_eq!(
            project.properties.name,
            expected["name"].as_str().unwrap(),
            "{}",
            case.display()
        );
        assert_eq!(
            project.modules.len(),
            expected["modules"].as_array().unwrap().len()
        );
        for actual in &project.modules {
            let expected = expected["modules"]
                .as_array()
                .unwrap()
                .iter()
                .find(|m| m["name"] == actual.name)
                .unwrap_or_else(|| panic!("{}: {}", case.display(), actual.name));
            assert_eq!(
                normalize(&actual.code()),
                normalize(expected["source"].as_str().unwrap()),
                "{} / {}",
                case.display(),
                actual.name
            );
            assert_eq!(
                actual.kind,
                match expected["kind"].as_u64().unwrap() {
                    1 => ModuleKind::Standard,
                    2 => ModuleKind::Class,
                    100 => ModuleKind::Document,
                    kind => panic!("{kind}"),
                }
            );
            if actual.kind == ModuleKind::Document {
                assert_eq!(
                    actual.host_document_type().unwrap().unwrap(),
                    expected["host_type"].as_str().unwrap()
                );
            }
        }
        let refs = project
            .references
            .iter()
            .map(|r| r.type_library().unwrap())
            .collect::<Vec<_>>();
        let expected_refs = expected["references"].as_array().unwrap();
        assert_eq!(expected_refs[0]["name"], "VBA");
        assert_eq!(expected_refs[1]["name"], "Excel");
        assert_eq!(
            refs.len(),
            expected_refs.len() - 2,
            "{} references: {refs:?}",
            case.display()
        );
        for (actual, expected) in refs.iter().zip(&expected_refs[2..]) {
            let calcvba_symbol::manifest::ProjectReference::TypeLibrary {
                name,
                guid,
                version_major,
                version_minor,
                ..
            } = actual
            else {
                panic!("typelib")
            };
            assert_eq!(name, expected["name"].as_str().unwrap());
            assert_eq!(
                guid.as_ref().unwrap().to_lowercase(),
                expected["guid"]
                    .as_str()
                    .unwrap()
                    .trim_matches(['{', '}'])
                    .to_lowercase()
            );
            assert_eq!(
                u64::from(version_major.unwrap()),
                expected["major"].as_u64().unwrap()
            );
            // Office can resolve a compatible newer minor version (the saved
            // Office reference requests 2.0; this Excel installation uses 2.8).
            assert!(u64::from(version_minor.unwrap()) <= expected["minor"].as_u64().unwrap());
        }
    }
}

#[test]
fn source_rewrite_drops_compiled_caches_and_preserves_workbook_parts() {
    for case in fixtures() {
        let input = std::fs::read(case.join("workbook.xlsm")).unwrap();
        let mut office = OfficeFile::read(&input, &ReadLimits::default()).unwrap();
        let api = office
            .project
            .modules
            .iter_mut()
            .find(|m| m.kind == ModuleKind::Standard)
            .unwrap();
        api.source.push_str("\r\nPublic Function CalcVbaWriteProbe() As Long\r\n CalcVbaWriteProbe = 1873\r\nEnd Function\r\n");
        let output = office
            .to_bytes()
            .unwrap_or_else(|e| panic!("{}: {e}", case.display()));
        let reloaded = OfficeFile::read(&output, &ReadLimits::default()).unwrap();
        assert_eq!(
            office.project.to_manifest().unwrap(),
            reloaded.project.to_manifest().unwrap()
        );
        let mut before = zip::ZipArchive::new(Cursor::new(&input)).unwrap();
        let mut after = zip::ZipArchive::new(Cursor::new(&output)).unwrap();
        assert_eq!(before.len(), after.len());
        for index in 0..before.len() {
            let mut part = before.by_index(index).unwrap();
            let mut replacement = after.by_name(part.name()).unwrap();
            let mut a = Vec::new();
            let mut b = Vec::new();
            part.read_to_end(&mut a).unwrap();
            replacement.read_to_end(&mut b).unwrap();
            if part.name() != "xl/vbaProject.bin" {
                assert_eq!(a, b, "{}: {}", case.display(), part.name());
            } else {
                let mut cfb = cfb::CompoundFile::open(Cursor::new(b)).unwrap();
                assert!(!cfb.walk().any(|e| e.name().starts_with("__SRP_")));
                let mut cache = Vec::new();
                cfb.open_stream("/VBA/_VBA_PROJECT")
                    .unwrap()
                    .read_to_end(&mut cache)
                    .unwrap();
                assert_eq!(cache.len(), 7);
                assert_eq!(&cache[..4], &[0xcc, 0x61, 0xff, 0xff]);
            }
        }
    }
}
