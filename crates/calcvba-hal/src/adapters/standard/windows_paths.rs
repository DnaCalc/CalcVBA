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
    directories: BTreeMap<char, (char, PathBuf)>,
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
            .map(|(_, path)| path.as_path())
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
            let key = (server.to_lowercase(), share.to_lowercase());
            let root = policy.windows_unc_roots.get(&key).ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::NotFound,
                    format!("UNC share {server}/{share} is unmapped"),
                )
            })?;
            if write && policy.windows_unc_read_only.contains(&key) {
                return Err(io::Error::new(
                    io::ErrorKind::ReadOnlyFilesystem,
                    "Windows mount is read-only",
                ));
            }
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
                part => {
                    let name = if part.contains(['*', '?']) {
                        part
                    } else {
                        part.trim_end_matches([' ', '.'])
                    };
                    if name.is_empty() || name.contains(['<', '>', ':', '"', '|', '\0']) {
                        return Err(io::Error::new(
                            io::ErrorKind::InvalidInput,
                            "invalid Windows path",
                        ));
                    }
                    segments.push(name.into());
                }
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
        let relative = prefix.iter().chain(&segments[prefix.len()..]).collect();
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
                continue;
            }
            let name = matching_entry(&absolute, &name)?;
            absolute.push(&name);
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
            && (self.policy.windows_faithful
                || !self.policy.windows_drive_roots.is_empty()
                || !self.policy.windows_drive_mounts.is_empty()
                || !self.policy.windows_unc_roots.is_empty())
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
        if !matches!(operation, "dir" | "kill") && text.contains(['*', '?']) {
            return Err(self.invalid_windows_path(capability, operation));
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
        let (spelling, directory) = state
            .directories
            .get(&drive)
            .map_or((drive, Path::new("")), |(letter, path)| {
                (*letter, path.as_path())
            });
        let suffix = directory.to_string_lossy().replace('/', "\\");
        Ok(format!("{spelling}:\\{suffix}"))
    }

    pub(super) fn mapped_ch_dir(&self, text: &str) -> HalResult<()> {
        if text.is_empty() || text.contains(['*', '?']) {
            return Err(self.invalid_windows_path(CapabilityId::FileSystemIo, "chdir"));
        }
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
        let spelling = match text.as_bytes() {
            [letter, b':', _, ..] if letter.is_ascii_alphabetic() => *letter as char,
            _ => state
                .directories
                .get(&drive)
                .map_or(drive, |(letter, _)| *letter),
        };
        state.directories.insert(drive, (spelling, relative));
        Ok(())
    }

    pub(super) fn mapped_ch_drive(&self, letter: char) -> HalResult<()> {
        if !letter.is_ascii_alphabetic() {
            return Err(HalError::adapter_fault(
                self.profile,
                CapabilityId::FileSystemIo,
                "chdrive",
                "invalid drive letter",
            )
            .with_host_error_code(5));
        }
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
        let state = self.windows_paths.borrow();
        let drives = self
            .policy
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
                let mut relative: PathBuf = prefix.iter().collect();
                if !suffix.as_os_str().is_empty() {
                    relative.push(suffix);
                }
                let windows = format!(
                    "{drive}:\\{}",
                    relative.to_string_lossy().replace('/', "\\")
                );
                Some((root.components().count(), prefix.len(), windows))
            });
        let shares = self
            .policy
            .windows_unc_roots
            .iter()
            .filter_map(|((server, share), path)| {
                let root = mapped_root(path).ok()?;
                let suffix = absolute.strip_prefix(&root).ok()?;
                let windows = format!(
                    r"\\{server}\{share}\{}",
                    suffix.to_string_lossy().replace('/', "\\")
                );
                Some((root.components().count(), 0, windows))
            });
        shares
            .chain(drives)
            .filter(|(_, _, windows)| {
                let Ok((_, resolved, _)) = state.resolve(&self.policy, windows, false) else {
                    return false;
                };
                fs::canonicalize(&resolved).unwrap_or(resolved) == absolute
            })
            .max_by_key(|(depth, prefix, _)| (*depth, *prefix))
            .map(|(_, _, windows)| windows)
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
        let code = match (operation, error.kind()) {
            ("chdrive", io::ErrorKind::NotFound | io::ErrorKind::NotADirectory) => 68,
            (_, io::ErrorKind::NotFound | io::ErrorKind::NotADirectory) => 76,
            ("chdir" | "mkdir" | "rmdir", io::ErrorKind::InvalidInput) => 76,
            (_, io::ErrorKind::InvalidInput) => 52,
            (_, io::ErrorKind::PermissionDenied) => 70,
            _ => 75,
        };
        HalError::adapter_fault(self.profile, capability, operation, error.to_string())
            .with_host_error_code(code)
    }

    fn invalid_windows_path(&self, capability: CapabilityId, operation: &'static str) -> HalError {
        self.windows_path_error(
            capability,
            operation,
            io::Error::new(io::ErrorKind::InvalidInput, "invalid Windows path"),
        )
    }
}

#[cfg(all(test, target_os = "linux"))]
mod tests {
    use super::*;
    use crate::model::HalProfileId;

    #[test]
    fn reversed_environment_paths_skip_shadowed_directories() {
        let root = std::env::temp_dir().join(format!("calcvba-env-shadow-{}", std::process::id()));
        let private_drive = root.join("private");
        let shadowed = private_drive.join("Temp");
        let mounted = root.join("mounted");
        fs::create_dir_all(&shadowed).unwrap();
        fs::create_dir_all(&mounted).unwrap();
        let policy = HostPolicy::native()
            .with_windows_drive_root('C', &private_drive)
            .with_windows_drive_mount(r"C:\Temp", &mounted, WindowsMountAccess::ReadWrite)
            .unwrap();
        let host = StandardHostServices::new(HalProfileId::Linux, policy);
        let from_shadowed = host.windows_path_for_host(&shadowed);
        let from_mounted = host.windows_path_for_host(&mounted);
        fs::remove_dir_all(root).unwrap();
        assert_eq!(from_shadowed, None);
        assert_eq!(from_mounted.as_deref(), Some(r"C:\Temp"));
    }
}
