use std::{
    fs::{self, OpenOptions},
    io::{self, Write},
    os::unix::fs::{MetadataExt, OpenOptionsExt},
    path::{Path, PathBuf},
    sync::atomic::{AtomicUsize, Ordering},
    time::UNIX_EPOCH,
};

fn store() -> io::Result<PathBuf> {
    Ok(super::settings_file::settings_path()?.with_file_name("windows-file-attributes"))
}

pub(super) fn read(path: &Path) -> io::Result<Option<i32>> {
    read_from(&store()?, path)
}

pub(super) fn write(path: &Path, bits: u8) -> io::Result<()> {
    write_to(&store()?, path, bits)
}

fn sidecar(root: &Path, path: &Path) -> io::Result<PathBuf> {
    let file = fs::metadata(path)?;
    // Device and inode preserve attributes across Name and hard links. Include a
    // birth time where available to distinguish a file after the inode is reused.
    let birth = file
        .created()
        .ok()
        .and_then(|time| time.duration_since(UNIX_EPOCH).ok())
        .map_or(0, |time| time.as_nanos());
    Ok(root.join(format!("{:x}-{:x}-{birth:x}", file.dev(), file.ino())))
}

fn read_from(root: &Path, path: &Path) -> io::Result<Option<i32>> {
    match fs::read(sidecar(root, path)?) {
        Ok(bytes) => match bytes.as_slice() {
            [bits] if bits & !0x26 == 0 => Ok(Some(i32::from(*bits))),
            _ => Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "invalid Windows file attributes",
            )),
        },
        Err(error) if error.kind() == io::ErrorKind::NotFound => Ok(None),
        Err(error) => Err(error),
    }
}

fn write_to(root: &Path, path: &Path, bits: u8) -> io::Result<()> {
    static SERIAL: AtomicUsize = AtomicUsize::new(0);
    let destination = sidecar(root, path)?;
    fs::create_dir_all(root)?;
    let (temporary, mut file) = loop {
        let serial = SERIAL.fetch_add(1, Ordering::Relaxed);
        let temporary = root.join(format!(".attributes-{}-{serial}.tmp", std::process::id()));
        match OpenOptions::new()
            .write(true)
            .create_new(true)
            .mode(0o600)
            .open(&temporary)
        {
            Ok(file) => break (temporary, file),
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error),
        }
    };
    let result = file.write_all(&[bits]).and_then(|()| file.sync_all());
    drop(file);
    let result = result.and_then(|()| fs::rename(&temporary, destination));
    if result.is_err() {
        let _ = fs::remove_file(temporary);
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn attribute_sidecar_survives_rename_and_replaces_an_existing_value() {
        let root = std::env::temp_dir().join(format!("calcvba-sidecar-{}", std::process::id()));
        let store = root.join("store");
        let first = root.join("file.txt");
        let renamed = root.join("renamed.txt");
        fs::create_dir(&root).unwrap();
        fs::write(&first, b"bytes").unwrap();
        assert_eq!(read_from(&store, &first).unwrap(), None);
        write_to(&store, &first, 6).unwrap();
        fs::rename(&first, &renamed).unwrap();
        assert_eq!(read_from(&store, &renamed).unwrap(), Some(6));
        write_to(&store, &renamed, 32).unwrap();
        assert_eq!(read_from(&store, &renamed).unwrap(), Some(32));
        assert_eq!(fs::read(&renamed).unwrap(), b"bytes");
        fs::remove_dir_all(root).unwrap();
    }
}
