use calcvba_office::{Error, OfficeFile, ReadLimits};
use std::{
    io::{Cursor, Read, Write},
    path::PathBuf,
};
fn workbook() -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/strings/workbook.xlsm"),
    )
    .unwrap()
}
fn project_bin() -> Vec<u8> {
    let mut zip = zip::ZipArchive::new(Cursor::new(workbook())).unwrap();
    let mut bin = Vec::new();
    zip.by_name("xl/vbaProject.bin")
        .unwrap()
        .read_to_end(&mut bin)
        .unwrap();
    bin
}
#[test]
fn limits_and_document_identity_are_checked_before_saving() {
    let bytes = workbook();
    let limits = ReadLimits {
        project_bytes: 32,
        ..Default::default()
    };
    assert!(matches!(
        OfficeFile::read(&bytes, &limits),
        Err(Error::Limit(_))
    ));
    let limits = ReadLimits {
        modules: 1,
        ..Default::default()
    };
    assert!(matches!(
        OfficeFile::read(&bytes, &limits),
        Err(Error::Limit("module count"))
    ));
    let mut file = OfficeFile::read(&bytes, &Default::default()).unwrap();
    file.project.modules.retain(|m| m.name != "Sheet1");
    assert!(matches!(file.to_bytes(), Err(Error::Unsupported(_))));
}
#[test]
fn locked_project_sources_are_refused() {
    // MS-OVBA 2.4.3.2, the published seed/key example with fUserProtected=1.
    let seed = 7u8;
    let key = 0xdfu8;
    let mut encrypted = vec![seed, seed ^ 2, seed ^ key];
    let mut enc1 = encrypted[2];
    let mut enc2 = encrypted[1];
    let mut previous = key;
    for plain in [7, 7, 7, 4, 0, 0, 0, 1, 0, 0, 0] {
        let byte = plain ^ enc2.wrapping_add(previous);
        encrypted.push(byte);
        enc2 = enc1;
        enc1 = byte;
        previous = plain;
    }
    let hex = encrypted
        .iter()
        .map(|b| format!("{b:02X}"))
        .collect::<String>();
    let mut cfb = cfb::CompoundFile::open(Cursor::new(project_bin())).unwrap();
    let mut project = String::new();
    cfb.open_stream("/PROJECT")
        .unwrap()
        .read_to_string(&mut project)
        .unwrap();
    let project = project
        .lines()
        .map(|line| {
            if line.starts_with("CMG=") {
                format!("CMG=\"{hex}\"")
            } else {
                line.into()
            }
        })
        .collect::<Vec<_>>()
        .join("\r\n");
    cfb.create_stream("/PROJECT")
        .unwrap()
        .write_all(project.as_bytes())
        .unwrap();
    assert!(matches!(
        calcvba_office::Project::read_bin(&cfb.into_inner().into_inner()),
        Err(Error::Locked)
    ));
}
#[test]
fn malformed_directory_offsets_and_duplicate_module_names_are_refused() {
    let mut cfb = cfb::CompoundFile::open(Cursor::new(project_bin())).unwrap();
    let mut bytes = Vec::new();
    cfb.open_stream("/VBA/dir")
        .unwrap()
        .read_to_end(&mut bytes)
        .unwrap();
    let mut dir = calcvba_office::compression::decompress(&bytes, 1 << 20).unwrap();
    let signature = [0x31, 0, 4, 0, 0, 0];
    let offset = dir
        .windows(signature.len())
        .position(|s| s == signature)
        .unwrap()
        + signature.len();
    dir[offset..offset + 4].copy_from_slice(&u32::MAX.to_le_bytes());
    let dir = calcvba_office::compression::compress(&dir).unwrap();
    cfb.create_stream("/VBA/dir")
        .unwrap()
        .write_all(&dir)
        .unwrap();
    assert!(calcvba_office::Project::read_bin(&cfb.into_inner().into_inner()).is_err());
    let mut file = OfficeFile::read(&workbook(), &Default::default()).unwrap();
    file.project.modules.push(file.project.modules[0].clone());
    assert!(file.to_bytes().is_err());
}

#[test]
fn unsupported_project_compilation_arguments_are_refused_before_output() {
    let mut file = OfficeFile::read(&workbook(), &Default::default()).unwrap();
    let output = std::env::temp_dir().join(format!(
        "calcvba-invalid-constants-{}.xlsm",
        std::process::id()
    ));
    for constants in [
        "Text = \"a:b\"",
        "Huge = 40000",
        "Same = 1 : same = 2",
        "Bogus-Name = 1",
    ] {
        file.project.properties.constants = constants.into();
        assert!(
            matches!(file.save_as(&output), Err(Error::Invalid(_))),
            "{constants}"
        );
        assert!(!output.exists());
    }
}
