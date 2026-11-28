//! Project-local Windows drive and directory state for host-backed file operations.
use std::{
    collections::BTreeMap,
    fs, io,
    path::{Path, PathBuf},
};

use crate::{
    error::{HalError, HalResult},
    model::{CapabilityId, HostPolicy, WindowsMountAccess},
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
        policy: &HostPolicy,
        text: &str,
        write: bool,
    ) -> io::Result<(Option<char>, PathBuf, PathBuf)> {
        let normalized = text.replace('\\', "/");
        let (drive, share, path) = if let Some(unc) = normalized.strip_prefix("//") {
            let mut parts = unc.splitn(3, '/');
            let server = parts.next().unwrap_or_default();
            let share = parts.next().unwrap_or_default();
            let root = policy
                .windows_unc_roots
                .get(&(server.to_lowercase(), share.to_lowercase()))
                .ok_or_else(|| {
                    io::Error::new(
                        io::ErrorKind::NotFound,
                        format!("UNC share {server}/{share} is unmapped"),
                    )
                })?;
            (None, Some(root), parts.next().unwrap_or_default())
        } else {
            let (drive, path) = match normalized.as_bytes() {
                [letter, b':', ..] if letter.is_ascii_alphabetic() => {
                    ((*letter as char).to_ascii_uppercase(), &normalized[2..])
                }
                _ => (self.drive, normalized.as_str()),
            };
            (Some(drive), None, path)
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

        let mount =
            drive.and_then(|drive| {
                policy
                    .windows_drive_mounts
                    .iter()
                    .filter(|mount| mount.drive == drive && mount.prefix.len() <= segments.len())
                    .filter(|mount| {
                        mount.prefix.iter().zip(&segments).all(|(prefix, segment)| {
                            prefix.to_lowercase() == segment.to_lowercase()
                        })
                    })
                    .max_by_key(|mount| mount.prefix.len())
            });
        let (root, prefix) = if let Some(mount) = mount {
            if write && mount.access == WindowsMountAccess::Read {
                return Err(io::Error::new(
                    io::ErrorKind::ReadOnlyFilesystem,
                    "Windows mount is read-only",
                ));
            }
            (mapped_root(&mount.root)?, mount.prefix.as_slice())
        } else if let Some(share) = share {
            (mapped_root(share)?, &[][..])
        } else {
            (
                resolve_root(
                    &policy.windows_drive_roots,
                    drive.expect("drive path has a drive"),
                )?,
                &[][..],
            )
        };
        let mut absolute = root.clone();
        let mut relative: PathBuf = prefix.iter().collect();
        let last_segment = segments.len().saturating_sub(1);
        for (index, name) in segments.into_iter().enumerate().skip(prefix.len()) {
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
            // WASI resolves files through its preopened capabilities. Rust's
            // canonicalize is unavailable there.
            if cfg!(target_os = "wasi") {
                continue;
            }
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

fn mapped_root(path: &Path) -> io::Result<PathBuf> {
    if cfg!(target_os = "wasi") {
        let metadata = fs::metadata(path)?;
        if metadata.is_dir() {
            return Ok(path.to_owned());
        }
        return Err(io::Error::from(io::ErrorKind::NotADirectory));
    }
    fs::canonicalize(path)
}

fn resolve_root(roots: &BTreeMap<char, PathBuf>, drive: char) -> io::Result<PathBuf> {
    let root = roots.get(&drive).ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::NotFound,
            format!("drive {drive}: is unmapped"),
        )
    })?;
    mapped_root(root)
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
                && self.policy.windows_drive_mounts.is_empty()
                && self.policy.windows_unc_roots.is_empty())
    }

    pub(super) fn map_path(
        &self,
        text: &str,
        capability: CapabilityId,
        operation: &'static str,
        write: bool,
    ) -> HalResult<PathBuf> {
        if !self.uses_windows_paths() || text.is_empty() {
            return Ok(PathBuf::from(text));
        }
        self.windows_paths
            .borrow()
            .resolve(&self.policy, text, write)
            .map(|(_, path, _)| path)
            .map_err(|error| self.windows_path_error(capability, operation, error))
    }

    pub(super) fn mapped_cur_dir(&self, letter: Option<char>) -> HalResult<String> {
        let state = self.windows_paths.borrow();
        let drive = letter.map_or(state.drive, |c| c.to_ascii_uppercase());
        self.require_mapped_drive(drive, "curdir")?;
        let suffix = state
            .current_directory(drive)
            .to_string_lossy()
            .replace('/', "\\");
        Ok(format!("{drive}:\\{suffix}"))
    }

    pub(super) fn mapped_ch_dir(&self, text: &str) -> HalResult<()> {
        let mut state = self.windows_paths.borrow_mut();
        let (drive, path, relative) = state
            .resolve(&self.policy, text, false)
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
        self.require_mapped_drive(drive, "chdrive")?;
        self.windows_paths.borrow_mut().drive = drive;
        Ok(())
    }

    fn require_mapped_drive(&self, drive: char, operation: &'static str) -> HalResult<()> {
        if self
            .policy
            .windows_drive_mounts
            .iter()
            .any(|mount| mount.drive == drive && mapped_root(&mount.root).is_ok())
        {
            return Ok(());
        }
        resolve_root(&self.policy.windows_drive_roots, drive)
            .map(|_| ())
            .map_err(|error| self.windows_path_error(CapabilityId::FileSystemIo, operation, error))
    }

    #[cfg(not(windows))]
    pub(super) fn windows_temp_dir(&self) -> Option<String> {
        let host_path = if cfg!(target_os = "wasi") {
            std::env::var_os("TMPDIR").map_or_else(|| PathBuf::from("/tmp"), PathBuf::from)
        } else {
            std::env::temp_dir()
        };
        self.windows_path_for_host(&host_path)
            .or_else(|| self.fallback_windows_dir(&[&["Temp"]]))
    }

    #[cfg(not(windows))]
    pub(super) fn windows_path_for_host(&self, path: &Path) -> Option<String> {
        let absolute = fs::canonicalize(path).unwrap_or_else(|_| path.to_path_buf());
        self.policy
            .windows_drive_roots
            .iter()
            .map(|(&drive, root)| (drive, &[][..], root))
            .chain(
                self.policy
                    .windows_drive_mounts
                    .iter()
                    .map(|mount| (mount.drive, mount.prefix.as_slice(), &mount.root)),
            )
            .filter_map(|(drive, prefix, path)| {
                let root = mapped_root(path).ok()?;
                let suffix = absolute.strip_prefix(&root).ok()?;
                Some((root.components().count(), drive, prefix, suffix))
            })
            .max_by_key(|(depth, _, prefix, _)| (*depth, prefix.len()))
            .map(|(_, drive, prefix, suffix)| {
                let mut relative: PathBuf = prefix.iter().collect();
                if !suffix.as_os_str().is_empty() {
                    relative.push(suffix);
                }
                format!(
                    "{drive}:\\{}",
                    relative.to_string_lossy().replace('/', "\\")
                )
            })
    }

    #[cfg(not(windows))]
    pub(super) fn fallback_windows_dir(&self, candidates: &[&[&str]]) -> Option<String> {
        let (&drive, root_path) = self
            .policy
            .windows_drive_roots
            .get_key_value(&'C')
            .or_else(|| self.policy.windows_drive_roots.iter().next())?;
        let root = mapped_root(root_path).ok()?;
        for segments in candidates {
            let mut candidate = root.clone();
            let mut windows_parts = Vec::with_capacity(segments.len());
            let mut valid = true;
            for segment in *segments {
                if segment.is_empty()
                    || segment.contains(['/', '\\', ':', '\0'])
                    || matches!(*segment, "." | "..")
                {
                    valid = false;
                    break;
                }
                let Ok(matched) = matching_entry(&candidate, segment) else {
                    valid = false;
                    break;
                };
                candidate.push(&matched);
                windows_parts.push(matched.to_string_lossy().into_owned());
            }
            if valid && candidate.is_dir() {
                return Some(format!("{drive}:\\{}", windows_parts.join("\\")));
            }
        }
        Some(format!("{drive}:\\"))
    }

    pub(super) fn windows_path_error(
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
