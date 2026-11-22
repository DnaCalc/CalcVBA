//! File storage for native VBA settings.
use super::process::{SettingEntry, SettingsState};
#[cfg(target_os = "linux")]
use std::os::unix::fs::OpenOptionsExt;
use std::{
    fs::{self, File, OpenOptions},
    io,
    path::{Path, PathBuf},
    sync::atomic::{AtomicUsize, Ordering},
};

pub(super) fn settings_path() -> io::Result<PathBuf> {
    let config = std::env::var_os("XDG_CONFIG_HOME")
        .map(PathBuf::from)
        .filter(|path| path.is_absolute())
        .or_else(|| {
            std::env::var_os("HOME")
                .map(PathBuf::from)
                .filter(|path| path.is_absolute())
                .map(|home| home.join(".config"))
        });
    config
        .map(|dir| dir.join("calcvba/settings"))
        .ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::NotFound,
                "no XDG_CONFIG_HOME or HOME for VBA settings",
            )
        })
}

#[cfg(target_os = "linux")]
pub(super) fn lock(path: &Path, write: bool) -> io::Result<Option<File>> {
    let lock_path = path.with_extension("lock");
    let lock = if write {
        if let Some(dir) = path.parent() {
            fs::create_dir_all(dir)?;
        }
        OpenOptions::new()
            .read(true)
            .write(true)
            .create(true)
            .truncate(false)
            .mode(0o600)
            .open(lock_path)?
    } else {
        match File::open(lock_path) {
            Ok(file) => file,
            Err(error) if error.kind() == io::ErrorKind::NotFound => return Ok(None),
            Err(error) => return Err(error),
        }
    };
    if write {
        lock.lock()?;
    } else {
        lock.lock_shared()?;
    }
    Ok(Some(lock))
}

type Row = ((String, String, String), SettingEntry);

pub(super) fn read(path: &Path) -> io::Result<SettingsState> {
    let file = match File::open(path) {
        Ok(file) => file,
        Err(error) if error.kind() == io::ErrorKind::NotFound => {
            return Ok(SettingsState::default());
        }
        Err(error) => return Err(error),
    };
    let rows: Vec<Row> = serde_json::from_reader(std::io::BufReader::new(file))?;
    let next_seq = rows.iter().map(|(_, entry)| entry.seq).max().unwrap_or(0);
    Ok(SettingsState {
        next_seq,
        values: rows.into_iter().collect(),
    })
}

pub(super) fn write(path: &Path, state: &SettingsState) -> io::Result<()> {
    static ID: AtomicUsize = AtomicUsize::new(0);
    #[cfg(target_os = "linux")]
    let writer = std::process::id();
    #[cfg(target_os = "wasi")]
    let writer = {
        if let Some(dir) = path.parent() {
            fs::create_dir_all(dir)?;
        }
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_nanos()
    };
    let (temporary, file) = loop {
        let serial = ID.fetch_add(1, Ordering::Relaxed);
        let name = format!("settings.{writer}.{serial}.tmp");
        let temporary = path.with_file_name(name);
        let mut options = OpenOptions::new();
        options.write(true).create_new(true);
        #[cfg(target_os = "linux")]
        options.mode(0o600);
        match options.open(&temporary) {
            Ok(file) => break (temporary, file),
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => {}
            Err(error) => return Err(error),
        }
    };
    let result = (|| {
        let rows: Vec<_> = state.values.iter().collect();
        serde_json::to_writer(&file, &rows)?;
        file.sync_all()?;
        fs::rename(&temporary, path)
    })();
    if result.is_err() {
        let _ = fs::remove_file(&temporary);
    }
    result
}
