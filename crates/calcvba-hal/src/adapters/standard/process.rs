use super::filesystem::dir_listing;
use crate::{
    error::{HalError, HalResult},
    model::CapabilityId,
    traits::ProcessEnvHal,
};
use calcvba_runtime::{
    VarType, Variant,
    bstr::BStr,
    safe_array::{SafeArray, SafeArrayBound, VT_BSTR_VALUE},
    variant_to_vba_string,
};
use std::collections::BTreeMap;

use super::StandardHostServices;

#[derive(Debug, Clone, Default)]
pub(super) struct DirSearchState {
    pub(super) remaining: Vec<String>,
    /// A pattern has been given, so a bare `Dir()` continues it; before that (or after the
    /// listing ended) a bare `Dir()` is error 5.
    pub(super) active: bool,
}

#[derive(Debug, Clone)]
#[cfg_attr(
    any(target_os = "linux", target_family = "wasm"),
    derive(serde::Serialize, serde::Deserialize)
)]
pub(super) struct SettingEntry {
    key_name: String,
    value: String,
    /// Creation order: the registry lists a section's values in the order they were first
    /// written, and a rewrite keeps the original position (P_GetAllSettings.5).
    pub(super) seq: u64,
}

type SectionValues = BTreeMap<String, SettingEntry>;

#[derive(Debug, Clone, Default)]
#[cfg_attr(
    any(target_os = "linux", target_family = "wasm"),
    derive(serde::Serialize, serde::Deserialize)
)]
pub(super) struct SettingsState {
    apps: BTreeMap<String, BTreeMap<String, SectionValues>>,
    pub(super) next_seq: u64,
}

impl SettingsState {
    pub(super) fn create_section(&mut self, app: String, section: &str) -> &mut SectionValues {
        let sections = self.apps.entry(app).or_default();
        // Persist parent sections even when they have no values of their own.
        for (end, _) in section.match_indices('\\') {
            sections.entry(section[..end].to_owned()).or_default();
        }
        sections.entry(section.to_owned()).or_default()
    }

    fn section(&self, app: &str, section: &str) -> Option<&SectionValues> {
        self.apps
            .get(&normalize_setting_name(app))?
            .get(&normalize_setting_name(section))
    }
}

/// The executable portion of a command line, without launching it.
fn shell_program(command: &str) -> &str {
    let command = command.trim_start();
    if let Some(quoted) = command.strip_prefix('"') {
        quoted.split('"').next().unwrap_or_default()
    } else if std::path::Path::new(command).is_file() {
        command
    } else {
        command.split_whitespace().next().unwrap_or_default()
    }
}

fn shell_program_exists(command: &str) -> bool {
    let program = shell_program(command);
    if program.is_empty() {
        return false;
    }
    let exists = |path: &std::path::Path| {
        path.is_file()
            || (cfg!(windows) && path.extension().is_none() && path.with_extension("exe").is_file())
    };
    let path = std::path::Path::new(program);
    if exists(path) {
        return true;
    }
    if program.contains(['/', '\\', ':']) {
        return false;
    }
    std::env::var_os("PATH").is_some_and(|paths| {
        std::env::split_paths(&paths).any(|directory| exists(&directory.join(program)))
    })
}

fn shell_launch_error(profile: crate::model::HalProfileId, error: std::io::Error) -> HalError {
    let number = match error.kind() {
        std::io::ErrorKind::NotFound => 53,
        std::io::ErrorKind::PermissionDenied => 70,
        _ => 5,
    };
    HalError::adapter_fault(
        profile,
        CapabilityId::ProcessEnv,
        "shell",
        error.to_string(),
    )
    .with_host_error_code(number)
}

#[cfg(windows)]
fn native_windows_shell(command: &str, style: i32) -> std::io::Result<u32> {
    use windows_sys::Win32::{
        Foundation::CloseHandle,
        System::Threading::{
            CreateProcessW, PROCESS_INFORMATION, STARTF_USESHOWWINDOW, STARTUPINFOW,
        },
    };
    let program = shell_program(command).to_ascii_lowercase();
    // CreateProcess requires the command interpreter for batch files. Preserve their
    // async launch while passing the same startup window style as an executable.
    let line = if program.ends_with(".cmd") || program.ends_with(".bat") {
        if !shell_program_exists(command) {
            return Err(std::io::Error::from(std::io::ErrorKind::NotFound));
        }
        format!("cmd.exe /c \"{command}\"")
    } else {
        command.to_string()
    };
    let mut command_line: Vec<u16> = line.encode_utf16().chain(Some(0)).collect();
    // SAFETY: all fields of these Win32 structs accept zero initialization.
    let mut startup: STARTUPINFOW = unsafe { std::mem::zeroed() };
    startup.cb = std::mem::size_of::<STARTUPINFOW>() as u32;
    startup.dwFlags = STARTF_USESHOWWINDOW;
    startup.wShowWindow = style as u16;
    // SAFETY: all fields of PROCESS_INFORMATION accept zero initialization.
    let mut process: PROCESS_INFORMATION = unsafe { std::mem::zeroed() };
    // SAFETY: buffers and structs stay live for the synchronous API call. The command
    // buffer is writable and terminated; null security/env/directory arguments select
    // the caller defaults. No handle inheritance or callback is requested.
    let created = unsafe {
        CreateProcessW(
            std::ptr::null(),
            command_line.as_mut_ptr(),
            std::ptr::null(),
            std::ptr::null(),
            0,
            0,
            std::ptr::null(),
            std::ptr::null(),
            &startup,
            &mut process,
        )
    };
    if created == 0 {
        return Err(std::io::Error::last_os_error());
    }
    let task = process.dwProcessId;
    // SAFETY: successful CreateProcess transferred these two owned handles. Closing
    // them releases our references without waiting for or terminating the child.
    unsafe {
        CloseHandle(process.hThread);
        CloseHandle(process.hProcess);
    }
    Ok(task)
}

/// The virtual host has a small, stable environment, like its clock and process model.
/// Native hosts read their own environment instead.
fn deterministic_environment(
    profile: crate::model::HalProfileId,
) -> &'static [(&'static str, &'static str)] {
    if profile == crate::model::HalProfileId::Windows {
        &[
            ("COMSPEC", "C:\\Windows\\System32\\cmd.exe"),
            ("PATH", "C:\\Windows\\System32"),
            ("TEMP", "C:\\Temp"),
            ("WINDIR", "C:\\Windows"),
        ]
    } else {
        &[("PATH", "/usr/bin:/bin"), ("TEMP", "/tmp")]
    }
}

fn environment_value(value: &std::ffi::OsStr) -> Variant {
    #[cfg(windows)]
    {
        use std::os::windows::ffi::OsStrExt;
        Variant::from_utf16_units(&value.encode_wide().collect::<Vec<_>>())
    }
    #[cfg(not(windows))]
    {
        Variant::from_string(value.to_string_lossy().into_owned())
    }
}

impl ProcessEnvHal for StandardHostServices {
    fn command_variant(&self) -> HalResult<Variant> {
        let capability = CapabilityId::ProcessEnv;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "command"));
        }
        if self.native_process_enabled() && !self.policy.deterministic_mode {
            let args = std::env::args().skip(1).collect::<Vec<_>>().join(" ");
            return Ok(Variant::from_string(args));
        }
        Ok(Variant::from_string(BStr::empty()))
    }

    fn get_setting_variant(
        &self,
        appname: Variant,
        section: Variant,
        key: Variant,
        default: Variant,
    ) -> HalResult<Variant> {
        let capability = CapabilityId::ProcessEnv;
        let op = "get_setting";
        if !self.supports_settings() {
            return Err(self.unsupported(capability, op));
        }
        let appname = self.setting_text(&appname, capability, op, "appname")?;
        let section = self.setting_text(&section, capability, op, "section")?;
        let key = self.setting_text(&key, capability, op, "key")?;
        let default = self.setting_text(&default, capability, op, "default")?;
        #[cfg(windows)]
        if self.native_process_enabled() {
            let value = super::settings_registry::get(&appname, &section, &key)
                .map_err(|error| self.registry_fault(op, error))?;
            return Ok(Variant::from_string(value.unwrap_or(default)));
        }
        let value = self.with_settings(op, false, |state| {
            Ok(state
                .section(&appname, &section)
                .and_then(|values| values.get(&normalize_setting_name(&key)))
                .map_or(default, |entry| entry.value.clone()))
        })?;
        Ok(Variant::from_string(value))
    }

    fn get_all_settings_variant(&self, appname: Variant, section: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::ProcessEnv;
        let op = "get_all_settings";
        if !self.supports_settings() {
            return Err(self.unsupported(capability, op));
        }
        let appname = self.setting_text(&appname, capability, op, "appname")?;
        let section = self.setting_text(&section, capability, op, "section")?;
        if section.is_empty() {
            return Err(settings_missing_fault(
                self.profile,
                capability,
                op,
                "section name is empty",
            ));
        }
        #[cfg(windows)]
        if self.native_process_enabled() {
            let entries = super::settings_registry::all(&appname, &section)
                .map_err(|error| self.registry_fault(op, error))?;
            return self.settings_array(entries);
        }
        let mut entries: Vec<SettingEntry> = self.with_settings(op, false, |state| {
            Ok(state
                .section(&appname, &section)
                .map(|values| values.values().cloned().collect())
                .unwrap_or_default())
        })?;
        entries.sort_by_key(|entry| entry.seq);
        self.settings_array(entries.into_iter().map(|e| (e.key_name, e.value)).collect())
    }

    fn save_setting_variant(
        &self,
        appname: Variant,
        section: Variant,
        key: Variant,
        setting: Variant,
    ) -> HalResult<Variant> {
        let capability = CapabilityId::ProcessEnv;
        let op = "save_setting";
        if !self.supports_settings() {
            return Err(self.unsupported(capability, op));
        }
        let appname = self.setting_text(&appname, capability, op, "appname")?;
        let section = self.setting_text(&section, capability, op, "section")?;
        let key_name = self.setting_text(&key, capability, op, "key")?;
        let value = self.setting_text(&setting, capability, op, "setting")?;
        #[cfg(windows)]
        if self.native_process_enabled() {
            super::settings_registry::save(&appname, &section, &key_name, &value)
                .map_err(|error| self.registry_fault(op, error))?;
            return Ok(Variant::empty());
        }
        self.with_settings(op, true, |state| {
            let mut next_seq = state.next_seq;
            let values = state.create_section(
                normalize_setting_name(&appname),
                &normalize_setting_name(&section),
            );
            let name = normalize_setting_name(&key_name);
            if let Some(entry) = values.get_mut(&name) {
                entry.value = value;
            } else {
                next_seq += 1;
                values.insert(
                    name,
                    SettingEntry {
                        key_name,
                        value,
                        seq: next_seq,
                    },
                );
                state.next_seq = next_seq;
            }
            Ok(Variant::empty())
        })
    }

    fn delete_setting_variant(
        &self,
        appname: Variant,
        section: Variant,
        key: Variant,
    ) -> HalResult<Variant> {
        let capability = CapabilityId::ProcessEnv;
        let op = "delete_setting";
        if !self.supports_settings() {
            return Err(self.unsupported(capability, op));
        }
        let appname = self.setting_text(&appname, capability, op, "appname")?;
        #[cfg(windows)]
        if self.native_process_enabled() {
            let optional_text = |value: &Variant, field| {
                (value.vtype() != VarType::Empty)
                    .then(|| self.setting_text(value, capability, op, field))
                    .transpose()
            };
            let section = optional_text(&section, "section")?;
            let key = optional_text(&key, "key")?;
            super::settings_registry::delete(&appname, section.as_deref(), key.as_deref())
                .map_err(|error| self.registry_fault(op, error))?;
            return Ok(Variant::empty());
        }
        let missing = |message| settings_missing_fault(self.profile, capability, op, message);
        let app = normalize_setting_name(&appname);
        if section.vtype() == VarType::Empty {
            if key.vtype() != VarType::Empty {
                return Err(missing("a key needs a section"));
            }
            let blocked = self.with_settings(op, true, |state| {
                let sections = state
                    .apps
                    .get_mut(&app)
                    .ok_or_else(|| missing("setting application does not exist"))?;
                let first_parent = sections
                    .keys()
                    .filter_map(|path| path.split_once('\\').map(|(parent, _)| parent))
                    .min()
                    .map(str::to_owned);
                if let Some(parent) = first_parent {
                    sections.retain(|name, _| name >= &parent);
                    return Ok(true);
                }
                state.apps.remove(&app);
                Ok(false)
            })?;
            // The preceding section deletions persist when Excel reports error 5.
            return if blocked {
                Err(missing("setting section has subkeys"))
            } else {
                Ok(Variant::empty())
            };
        }

        let section = self.setting_text(&section, capability, op, "section")?;
        let section = normalize_setting_name(&section);
        self.with_settings(op, true, |state| {
            if key.vtype() == VarType::Empty {
                let section_missing = || missing("setting section does not exist");
                let sections = state.apps.get_mut(&app).ok_or_else(section_missing)?;
                let prefix = format!("{section}\\");
                if sections.keys().any(|name| name.starts_with(&prefix)) {
                    let values = sections.get_mut(&section).ok_or_else(section_missing)?;
                    if values.is_empty() {
                        return Err(section_missing());
                    }
                    values.clear();
                } else if sections.remove(&section).is_none() {
                    return Err(section_missing());
                }
                return Ok(Variant::empty());
            }
            let key = self.setting_text(&key, capability, op, "key")?;
            if state
                .apps
                .get_mut(&app)
                .and_then(|sections| sections.get_mut(&section))
                .and_then(|values| values.remove(&normalize_setting_name(&key)))
                .is_none()
            {
                return Err(missing("setting key does not exist"));
            }
            Ok(Variant::empty())
        })
    }

    fn shell_variant(&self, command: Variant, _window_style: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::ProcessEnv;
        if self.wasi_native() {
            return Err(self.unsupported(capability, "shell"));
        }
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "shell"));
        }
        if !self.policy.allow_process_spawn {
            return Err(self.denied(capability, "shell"));
        }
        if self.native_process_enabled()
            && let Some(text) = command.as_bstr()
            && !text.as_str().trim().is_empty()
        {
            let command_text = text.as_str();
            #[cfg(windows)]
            {
                let style =
                    self.variant_to_i32(&_window_style, capability, "shell", "window_style")?;
                let task = native_windows_shell(&command_text, style)
                    .map_err(|error| shell_launch_error(self.profile, error))?;
                return Ok(Variant::from_f64(f64::from(task)));
            }
            #[cfg(not(windows))]
            {
                // An interpreter would return a task ID and fail later for a missing
                // executable. VBA reports the failure before returning from Shell.
                if !shell_program_exists(&command_text) {
                    return Err(shell_launch_error(
                        self.profile,
                        std::io::ErrorKind::NotFound.into(),
                    ));
                }
                let child = self
                    .spawn_probe_shell_process_text(&command_text)
                    .map_err(|error| shell_launch_error(self.profile, error))?;
                return Ok(Variant::from_f64(f64::from(child.id())));
            }
        }
        if self.native_process_enabled() {
            let command = self
                .variant_to_i32(&command, capability, "shell", "command")
                .unwrap_or(0);
            if command != 0 {
                let child = self.spawn_probe_shell_process(command).map_err(|err| {
                    HalError::adapter_fault(
                        self.profile,
                        capability,
                        "shell",
                        format!("failed to spawn probe shell process: {err}"),
                    )
                })?;
                let child_id = i32::try_from(child.id()).unwrap_or(i32::MAX).max(1);
                return Ok(Variant::from_f64(f64::from(child_id)));
            }
        }
        let command = match command.as_bstr() {
            Some(text) => i32::from(!text.as_str().trim().is_empty()),
            None => self
                .variant_to_i32(&command, capability, "shell", "command")
                .unwrap_or(0),
        };
        Ok(Variant::from_i32(if command == 0 { 0 } else { 1 }))
    }

    fn environ_variant(&self, key: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::ProcessEnv;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "environ"));
        }
        let virtual_table = deterministic_environment(self.profile);
        if let Some(text) = key.as_bstr() {
            let text = text.as_str();
            let name = text.trim_matches(' ');
            if name.is_empty() {
                return Err(HalError::adapter_fault(
                    self.profile,
                    capability,
                    "environ",
                    "Invalid procedure call or argument",
                )
                .with_host_error_code(5));
            }
            if self.native_process_enabled() || self.wasi_native() {
                #[cfg(any(target_os = "linux", target_os = "wasi"))]
                if self.uses_windows_paths() {
                    let value = self
                        .windows_environment()
                        .into_iter()
                        .find(|(key, _)| {
                            key.to_str()
                                .is_some_and(|key| key.eq_ignore_ascii_case(name))
                        })
                        .map(|(_, value)| value)
                        .unwrap_or_default();
                    return Ok(environment_value(&value));
                }
                let value = std::env::var_os(name).unwrap_or_default();
                return Ok(environment_value(&value));
            }
            let value = virtual_table
                .iter()
                .find(|(key, _)| key.eq_ignore_ascii_case(name))
                .map(|(_, value)| *value)
                .unwrap_or_default();
            return Ok(Variant::from_string(value));
        }
        // MS-VBAL: non-string keys are Let-coercible to Long, with ties-to-even rounding.
        let index = calcvba_eval::arith::coerce_for_byref_target(&key, VarType::Long as u16)
            .map_err(|error| {
                HalError::adapter_fault(self.profile, capability, "environ", error.message)
                    .with_host_error_code(error.code)
            })?
            .as_i32()
            .expect("Long coercion returns a Long");
        // Legacy VB's environment index range; additional boundary cases are queued for Excel.
        if !(1..=255).contains(&index) {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "environ",
                "Invalid procedure call or argument",
            )
            .with_host_error_code(5));
        }
        if self.native_process_enabled() || self.wasi_native() {
            #[cfg(any(target_os = "linux", target_os = "wasi"))]
            let variables: Vec<_> = if self.uses_windows_paths() {
                self.windows_environment()
            } else {
                std::env::vars_os().collect()
            };
            #[cfg(not(any(target_os = "linux", target_os = "wasi")))]
            let variables: Vec<_> = std::env::vars_os().collect();
            let Some((name, value)) = variables.get((index - 1) as usize) else {
                return Ok(Variant::from_string(""));
            };
            let mut entry = name.clone();
            entry.push("=");
            entry.push(value);
            return Ok(environment_value(&entry));
        }
        let entry = virtual_table
            .get((index - 1) as usize)
            .map(|(name, value)| format!("{name}={value}"))
            .unwrap_or_default();
        Ok(Variant::from_string(entry))
    }

    fn dir_variant(&self, path: Variant, attrs: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::ProcessEnv;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "dir"));
        }
        if self.native_process_enabled() || self.wasi_native() {
            let is_continuation = match path.vtype() {
                VarType::Empty | VarType::Null => true,
                VarType::String => false,
                _ => {
                    self.variant_to_i32(&path, capability, "dir", "path")
                        .unwrap_or(1)
                        == 0
                }
            };
            if is_continuation {
                let mut state = self.borrow_dir(capability, "dir")?;
                if !state.active {
                    return Err(HalError::adapter_fault(
                        self.profile,
                        capability,
                        "dir",
                        "Dir() has no previous pattern",
                    )
                    .with_host_error_code(5));
                }
                let next = if state.remaining.is_empty() {
                    state.active = false;
                    String::new()
                } else {
                    state.remaining.remove(0)
                };
                return Ok(Variant::from_string(next));
            }

            // Explicit Dir("") starts a fresh search in the current directory;
            // only the omitted pathname continues the previous search.
            let target = if path.as_bstr().is_some_and(|text| text.is_empty()) {
                if self.uses_windows_paths() {
                    self.map_path("*", capability, "dir")?
                } else {
                    std::env::current_dir()
                        .map(|dir| dir.join("*"))
                        .map_err(|err| {
                            HalError::adapter_fault(
                                self.profile,
                                capability,
                                "dir",
                                err.to_string(),
                            )
                        })?
                }
            } else {
                self.variant_to_path(&path, capability, "dir", "path")?
            };
            let attrs = self
                .variant_to_i32(&attrs, capability, "dir", "attributes")
                .unwrap_or(0);
            let raw = path.as_bstr().map(|text| text.as_str().to_string());
            let mut names = dir_listing(
                raw.as_deref().unwrap_or(""),
                &target,
                attrs,
                self.uses_windows_paths(),
            )
            .map_err(|err| {
                HalError::adapter_fault(
                    self.profile,
                    capability,
                    "dir",
                    format!("failed to enumerate dir path {}: {err}", target.display()),
                )
            })?;
            let first = names.first().cloned().unwrap_or_default();
            let mut state = self.borrow_dir(capability, "dir")?;
            state.active = !names.is_empty();
            state.remaining = if names.len() > 1 {
                names.drain(1..).collect()
            } else {
                Vec::new()
            };
            return Ok(Variant::from_string(first));
        }
        let out = match path.vtype() {
            VarType::Empty | VarType::Null => 0,
            VarType::String => {
                i32::from(path.as_bstr().map(|text| !text.is_empty()).unwrap_or(false))
            }
            _ => i32::from(
                self.variant_to_i32(&path, capability, "dir", "path")
                    .unwrap_or(0)
                    != 0,
            ),
        };
        Ok(Variant::from_i32(out))
    }
}

impl StandardHostServices {
    fn settings_array(&self, entries: Vec<(String, String)>) -> HalResult<Variant> {
        let capability = CapabilityId::ProcessEnv;
        let op = "get_all_settings";
        if entries.is_empty() {
            return Ok(Variant::empty());
        }

        // `a(i, 0)` is a key and `a(i, 1)` its value; storage is first dimension fastest,
        // so every key comes before every value.
        let mut values = Vec::with_capacity(entries.len() * 2);
        let mut settings = Vec::with_capacity(entries.len());
        for (name, value) in entries {
            values.push(Variant::from_string(name));
            settings.push(Variant::from_string(value));
        }
        values.append(&mut settings);
        let bounds = vec![
            SafeArrayBound {
                lower: 0,
                count: u32::try_from(values.len() / 2).map_err(|_| {
                    HalError::adapter_fault(
                        self.profile,
                        capability,
                        op,
                        "settings array row count exceeds SAFEARRAY bounds",
                    )
                })?,
            },
            SafeArrayBound { lower: 0, count: 2 },
        ];
        let array = SafeArray::from_typed_variants_nd(bounds, VT_BSTR_VALUE, values)
            .map_err(|message| HalError::adapter_fault(self.profile, capability, op, message))?;
        Ok(Variant::from_safearray(array))
    }

    #[cfg(windows)]
    fn registry_fault(&self, op: &'static str, error: std::io::Error) -> HalError {
        HalError::adapter_fault(
            self.profile,
            CapabilityId::ProcessEnv,
            op,
            error.to_string(),
        )
        .with_host_error_code(5)
    }

    fn supports_settings(&self) -> bool {
        self.browser_settings || self.supports(CapabilityId::ProcessEnv)
    }

    fn with_settings<T>(
        &self,
        operation: &'static str,
        write: bool,
        operation_body: impl FnOnce(&mut SettingsState) -> HalResult<T>,
    ) -> HalResult<T> {
        let _ = write;
        let capability = CapabilityId::ProcessEnv;
        #[cfg(any(target_os = "linux", target_os = "wasi"))]
        if self.native_process_enabled() || self.wasi_native() {
            use super::settings_file;
            let fault = |error: std::io::Error| {
                HalError::adapter_fault(self.profile, capability, operation, error.to_string())
                    .with_host_error_code(5)
            };
            let location = settings_file::settings_path().map_err(fault)?;
            #[cfg(target_os = "linux")]
            let _lock = settings_file::lock(&location, write).map_err(fault)?;
            let mut state = settings_file::read(&location).map_err(fault)?;
            let answer = operation_body(&mut state)?;
            if write {
                settings_file::write(&location, &state).map_err(fault)?;
            }
            return Ok(answer);
        }
        let mut state = self.borrow_settings(capability, operation)?;
        operation_body(&mut state)
    }

    #[cfg(any(target_os = "linux", target_os = "wasi"))]
    fn windows_environment(&self) -> Vec<(std::ffi::OsString, std::ffi::OsString)> {
        use std::{env, path::PathBuf};
        let mut vars: Vec<_> = env::vars_os().collect();
        let temp = if cfg!(target_os = "wasi") {
            env::var_os("TMPDIR").map_or_else(|| PathBuf::from("/tmp"), PathBuf::from)
        } else {
            env::temp_dir()
        };
        let temp = self.windows_path_for_host(&temp);
        let home = env::var_os("HOME").map(PathBuf::from);
        let config = env::var_os("XDG_CONFIG_HOME")
            .map(PathBuf::from)
            .filter(|path| path.is_absolute())
            .or_else(|| home.as_ref().map(|path| path.join(".config")));
        let hostname = env::var_os("HOSTNAME").or_else(|| {
            if cfg!(target_os = "wasi") {
                return None;
            }
            std::fs::read_to_string("/proc/sys/kernel/hostname")
                .ok()
                .map(|name| name.trim_end().into())
        });
        let aliases = [
            ("TEMP", temp.clone().map(Into::into)),
            ("TMP", temp.map(Into::into)),
            (
                "USERPROFILE",
                home.as_deref()
                    .and_then(|path| self.windows_path_for_host(path))
                    .map(Into::into),
            ),
            (
                "APPDATA",
                config
                    .as_deref()
                    .and_then(|path| self.windows_path_for_host(path))
                    .map(Into::into),
            ),
            ("COMPUTERNAME", hostname),
            (
                "USERNAME",
                env::var_os("USER").or_else(|| env::var_os("LOGNAME")),
            ),
        ];
        for (key, value) in aliases {
            if let Some(value) = value
                && !vars.iter().any(|(old, _)| {
                    old.to_str()
                        .is_some_and(|old| old.eq_ignore_ascii_case(key))
                })
            {
                vars.push((key.into(), value));
            }
        }
        vars
    }

    fn setting_text(
        &self,
        value: &Variant,
        capability: CapabilityId,
        op: &'static str,
        field: &'static str,
    ) -> HalResult<String> {
        variant_to_vba_string(value)
            .map(|text| text.as_str().to_string())
            .map_err(|message| {
                HalError::adapter_fault(self.profile, capability, op, format!("{field}: {message}"))
            })
    }
}

fn normalize_setting_name(value: &str) -> String {
    value.to_ascii_lowercase()
}

fn settings_missing_fault(
    profile: crate::model::HalProfileId,
    capability: CapabilityId,
    op: &'static str,
    message: &'static str,
) -> HalError {
    HalError::adapter_fault(profile, capability, op, message).with_host_error_code(5)
}
