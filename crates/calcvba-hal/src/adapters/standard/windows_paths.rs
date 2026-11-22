//! Project-local Windows drive and directory state for host-backed file operations.
use std::{
    collections::BTreeMap,
    fs, io,
    path::{Path, PathBuf},
};

use crate::{
    error::{HalError, HalResult},
    model::CapabilityId,
};

use super::StandardHostServices;

#[derive(Debug)]
pub(super) struct WindowsPaths {
    drive: char,
    directories: BTreeMap<char, PathBuf>,
}

impl Default for WindowsPaths {
    fn default() -> Self {
        Self {
            drive: 'C',
            directories: BTreeMap::new(),
        }
    }
}

impl WindowsPaths {
    fn current_directory(&self, drive: char) -> &Path {
        self.directories
            .get(&drive)
            .map(PathBuf::as_path)
            .unwrap_or(Path::new(""))
    }

    fn resolve(
        &self,
        roots: &BTreeMap<char, PathBuf>,
        shares: &BTreeMap<(String, String), PathBuf>,
        text: &str,
    ) -> io::Result<(Option<char>, PathBuf, PathBuf)> {
        let normalized = text.replace('\\', "/");
        let (drive, root, path) = if let Some(unc) = normalized.strip_prefix("//") {
            let mut parts = unc.splitn(3, '/');
            let server = parts.next().unwrap_or_default();
            let share = parts.next().unwrap_or_default();
            let root = shares
                .get(&(server.to_lowercase(), share.to_lowercase()))
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::NotFound,
                        format!("UNC share {server}/{share} is unmapped"),
                    )
                })?;
            (
                None,
                fs::canonicalize(root)?,
                parts.next().unwrap_or_default(),
            )
        } else {
            let (drive, path) = match normalized.as_bytes() {
                [letter, b':', ..] if letter.is_ascii_alphabetic() => {
                    ((*letter as char).to_ascii_uppercase(), &normalized[2..])
                }
                _ => (self.drive, normalized.as_str()),
            };
            (Some(drive), resolve_root(roots, drive)?, path)
        };
        let mut segments: Vec<String> = if let Some(drive) = drive
            && !path.starts_with('/')
        {
            self.current_directory(drive)
                .iter()
                .map(|part| part.to_string_lossy().into_owned())
                .collect()
        } else {
            Vec::new()
        };
        for part in path.split('/') {
            match part {
                "" | "." => {}
                ".." => {
                    segments.pop();
                }
                part if part.contains([':', '\0']) => {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "invalid Windows path",
                    ));
                }
                part => segments.push(part.into()),
            }
        }

        let mut absolute = root.clone();
        let mut relative = PathBuf::new();
        let last_segment = segments.len().saturating_sub(1);
        for (index, name) in segments.into_iter().enumerate() {
            if name.contains(['*', '?']) {
                if index != last_segment {
                    return Err(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "wildcards are only supported in the mapped file name",
                    ));
                }
                absolute.push(&name);
                relative.push(name);
                continue;
            }
            let name = matching_entry(&absolute, &name)?;
            absolute.push(&name);
            relative.push(name);
            match fs::canonicalize(&absolute) {
                Ok(path) if !path.starts_with(&root) => {
                    return Err(io::Error::new(
                        io::ErrorKind::PermissionDenied,
                        "path leaves the mapped drive",
                    ));
                }
                Err(error)
                    if error.kind() != io::ErrorKind::NotFound
                        || fs::symlink_metadata(&absolute).is_ok() =>
                {
                    return Err(error);
                }
                _ => {}
            }
        }
        Ok((drive, absolute, relative))
    }
}

fn resolve_root(roots: &BTreeMap<char, PathBuf>, drive: char) -> io::Result<PathBuf> {
    let root = roots.get(&drive).ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::NotFound,
            format!("drive {drive}: is unmapped"),
        )
    })?;
    fs::canonicalize(root)
}

fn matching_entry(directory: &Path, name: &str) -> io::Result<std::ffi::OsString> {
    let names = match fs::read_dir(directory) {
        Ok(names) => names,
        Err(error)
            if matches!(
                error.kind(),
                io::ErrorKind::NotFound | io::ErrorKind::NotADirectory
            ) =>
        {
            return Ok(name.into());
        }
        Err(error) => return Err(error),
    };
    let folded = name.to_lowercase();
    let mut existing: Option<std::ffi::OsString> = None;
    for entry in names {
        let candidate = entry?.file_name();
        if candidate.to_string_lossy().to_lowercase() == folded {
            if let Some(first) = existing {
                return Err(io::Error::new(
                    io::ErrorKind::AlreadyExists,
                    format!("both {:?} and {:?} match {name:?}", first, candidate),
                ));
            }
            existing = Some(candidate);
        }
    }
    Ok(existing.unwrap_or_else(|| name.into()))
}

impl StandardHostServices {
    pub(super) fn uses_windows_paths(&self) -> bool {
        !cfg!(windows)
            && self.native_fs_enabled()
            && !(self.policy.windows_drive_roots.is_empty()
                && self.policy.windows_unc_roots.is_empty())
    }

    pub(super) fn map_path(
        &self,
        text: &str,
        capability: CapabilityId,
        operation: &'static str,
    ) -> HalResult<PathBuf> {
        if !self.uses_windows_paths() || text.is_empty() {
            return Ok(PathBuf::from(text));
        }
        self.windows_paths
            .borrow()
            .resolve(
                &self.policy.windows_drive_roots,
                &self.policy.windows_unc_roots,
                text,
            )
            .map(|(_, path, _)| path)
            .map_err(|error| self.windows_path_error(capability, operation, error))
    }

    pub(super) fn mapped_cur_dir(&self, letter: Option<char>) -> HalResult<String> {
        let state = self.windows_paths.borrow();
        let drive = letter.map_or(state.drive, |c| c.to_ascii_uppercase());
        resolve_root(&self.policy.windows_drive_roots, drive).map_err(|error| {
            self.windows_path_error(CapabilityId::FileSystemIo, "curdir", error)
        })?;
        let suffix = state
            .current_directory(drive)
            .to_string_lossy()
            .replace('/', "\\");
        Ok(format!("{drive}:\\{suffix}"))
    }

    pub(super) fn mapped_ch_dir(&self, text: &str) -> HalResult<()> {
        let mut state = self.windows_paths.borrow_mut();
        let (drive, path, relative) = state
            .resolve(
                &self.policy.windows_drive_roots,
                &self.policy.windows_unc_roots,
                text,
            )
            .map_err(|error| self.windows_path_error(CapabilityId::FileSystemIo, "chdir", error))?;
        if !path.is_dir() {
            return Err(self.windows_path_error(
                CapabilityId::FileSystemIo,
                "chdir",
                io::Error::new(
                    io::ErrorKind::NotFound,
                    format!("{} is not a directory", path.display()),
                ),
            ));
        }
        let drive = drive.ok_or_else(|| {
            self.windows_path_error(
                CapabilityId::FileSystemIo,
                "chdir",
                io::Error::new(
                    io::ErrorKind::NotFound,
                    "UNC shares do not have a drive working directory",
                ),
            )
        })?;
        state.directories.insert(drive, relative);
        Ok(())
    }

    pub(super) fn mapped_ch_drive(&self, letter: char) -> HalResult<()> {
        let drive = letter.to_ascii_uppercase();
        resolve_root(&self.policy.windows_drive_roots, drive).map_err(|error| {
            self.windows_path_error(CapabilityId::FileSystemIo, "chdrive", error)
        })?;
        self.windows_paths.borrow_mut().drive = drive;
        Ok(())
    }

    fn windows_path_error(
        &self,
        capability: CapabilityId,
        operation: &'static str,
        error: io::Error,
    ) -> HalError {
        let code = match error.kind() {
            io::ErrorKind::NotFound | io::ErrorKind::NotADirectory => 76,
            io::ErrorKind::InvalidInput => 52,
            io::ErrorKind::PermissionDenied => 70,
            _ => 75,
        };
        HalError::adapter_fault(self.profile, capability, operation, error.to_string())
            .with_host_error_code(code)
    }
}
