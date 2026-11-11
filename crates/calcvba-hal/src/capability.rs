//! Capability policy for the host layer.
//!
//! [`PolicyHost`] wraps any [`HostServices`] and gates the operations that reach the
//! operating system (files, processes, environment, settings, `Declare`, COM activation,
//! `SendKeys`/`AppActivate`) behind eight [`Capability`] values. Both execution backends
//! reach those services only through the host they are given, so one wrapper covers VM3 and
//! the JIT. In Safe mode a denied operation raises the nearest real VBA run-time error
//! (70, 53 or 429) with its own `Err.Description`, and the [`HalError`] keeps the stable code
//! `HAL-E-POLICY-DENIED` so a harness can tell a policy denial from the real error. Every
//! request, allowed or denied, goes into a bounded audit log that the harness drains per case.
//!
//! This is a policy gate, not a sandbox. Design record:
//! `docs/design/capability-policy-review-2026-11-11.md`.
//!
//! Gated and logged: `Open` (read or write by mode), `Kill`/`MkDir`/`RmDir`/`Name`/`FileCopy`/
//! `SetAttr`/`ChDir`/`ChDrive`, `Dir`/`FileLen`/`GetAttr`/`FileDateTime`/`CurDir`, `Shell`,
//! `Environ`/`Command`, the four settings functions, every `Declare` call, `SendKeys`/
//! `AppActivate`, and COM activation (`CreateObject`, `GetObject`, `New` of an external class)
//! through [`HostServices::authorize_com_activation`]. Handle operations (`Print #`, `Get`,
//! `EOF`), `MsgBox`/`InputBox`, `Now`/`Timer`, `Debug.Print` and `DoEvents` pass through
//! unlogged: after a denied open no handle exists. Typed COM member calls after activation
//! are not visible to the policy.

use std::cell::RefCell;
use std::collections::BTreeMap;
use std::rc::Rc;
use std::sync::Arc;

use calcvba_runtime::{BindingHandle, DynLinkSymbol, RuntimeCallArgument, Variant};

use crate::error::{HalError, HalResult, HostErrorInfo};
use crate::model::{CapabilityId, HalDescriptor, HalProfileId, HostPolicy};
use crate::traits::{
    ComHal, ConsoleHal, DiagnosticsHal, DynLinkDescriptorView, DynamicLinkHal, EventPumpHal,
    FileSystemHal, HostServices, ProcessEnvHal, ProjectCatalogHal, ProjectMutationHal,
    ProjectReferenceHal, TimeLocaleHal, UiInteractionHal,
};

/// The stable code carried by every policy denial.
pub const POLICY_DENIED_CODE: &str = "HAL-E-POLICY-DENIED";

/// What a gated operation needs. One capability per family of operations.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum Capability {
    FsRead,
    FsWrite,
    Process,
    Env,
    Settings,
    Ffi,
    ComActivate,
    UiKeys,
}

impl Capability {
    pub const ALL: [Capability; 8] = [
        Capability::FsRead,
        Capability::FsWrite,
        Capability::Process,
        Capability::Env,
        Capability::Settings,
        Capability::Ffi,
        Capability::ComActivate,
        Capability::UiKeys,
    ];

    /// The stable dotted name used in reports and CLI output.
    pub const fn name(self) -> &'static str {
        match self {
            Capability::FsRead => "fs.read",
            Capability::FsWrite => "fs.write",
            Capability::Process => "process",
            Capability::Env => "env",
            Capability::Settings => "settings",
            Capability::Ffi => "ffi",
            Capability::ComActivate => "com.activate",
            Capability::UiKeys => "ui.keys",
        }
    }

    pub fn from_name(name: &str) -> Option<Self> {
        Self::ALL.into_iter().find(|c| c.name() == name)
    }

    /// The VBA run-time error and `Err.Description` raised when Safe mode denies it.
    /// 70 is Excel's number for a blocked `SendKeys` or a locked file, 53 for a `Declare`
    /// whose library is absent, 429 for a class that cannot be created.
    pub const fn safe_error(self) -> (i32, &'static str) {
        match self {
            Capability::Ffi => (53, "File not found"),
            Capability::ComActivate => (429, "ActiveX component can't create object"),
            _ => (70, "Permission denied"),
        }
    }

    const fn hal_capability(self) -> CapabilityId {
        match self {
            Capability::FsRead | Capability::FsWrite => CapabilityId::FileSystemIo,
            Capability::Process | Capability::Env | Capability::Settings => {
                CapabilityId::ProcessEnv
            }
            Capability::Ffi => CapabilityId::DynamicLinking,
            Capability::ComActivate => CapabilityId::ComActivationDispatch,
            Capability::UiKeys => CapabilityId::UiInteraction,
        }
    }

    const fn bit(self) -> u8 {
        1 << (self as u8)
    }
}

impl std::fmt::Display for Capability {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(self.name())
    }
}

/// A set of capabilities.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct CapabilitySet(u8);

impl CapabilitySet {
    pub const EMPTY: CapabilitySet = CapabilitySet(0);
    pub const ALL: CapabilitySet = CapabilitySet(0xFF);

    pub fn of(caps: &[Capability]) -> Self {
        Self(caps.iter().fold(0, |acc, c| acc | c.bit()))
    }
    pub fn contains(self, cap: Capability) -> bool {
        self.0 & cap.bit() != 0
    }
    pub fn insert(&mut self, cap: Capability) {
        self.0 |= cap.bit();
    }
    pub fn is_empty(self) -> bool {
        self.0 == 0
    }
    pub fn iter(self) -> impl Iterator<Item = Capability> {
        Capability::ALL
            .into_iter()
            .filter(move |c| self.contains(*c))
    }
    pub fn union(self, other: CapabilitySet) -> CapabilitySet {
        CapabilitySet(self.0 | other.0)
    }
    pub fn difference(self, other: CapabilitySet) -> CapabilitySet {
        CapabilitySet(self.0 & !other.0)
    }
    pub fn intersection(self, other: CapabilitySet) -> CapabilitySet {
        CapabilitySet(self.0 & other.0)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Decision {
    Allowed,
    Denied,
}

/// One request that reached the policy.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuditRecord {
    pub seq: u64,
    pub capability: Capability,
    /// The HAL operation name (`open`, `kill`, `shell`, `invoke_descriptor_arguments`, ...).
    pub op: &'static str,
    /// Path, command, variable name, `lib!symbol` or ProgID, truncated to [`DETAIL_MAX`] characters.
    pub detail: String,
    pub decision: Decision,
    pub err_number: Option<i32>,
}

/// Exact per-`(capability, detail)` totals; these keep counting after the record list is full.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuditCount {
    pub capability: Capability,
    pub detail: String,
    pub allowed: u64,
    pub denied: u64,
}

/// What [`AuditHandle::drain`] returns.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct AuditSnapshot {
    /// The first [`MAX_RECORDS`] full records since the last drain.
    pub records: Vec<AuditRecord>,
    pub counts: Vec<AuditCount>,
    pub total: u64,
    pub denied: u64,
    /// Records past the cap (counted in `counts`, not kept in `records`).
    pub dropped_records: u64,
    /// Requests whose `(capability, detail)` key was past the key cap (counted only in
    /// `total` and `denied`).
    pub dropped_keys: u64,
}

impl AuditSnapshot {
    /// The capabilities requested since the last drain, allowed or denied.
    pub fn requested(&self) -> CapabilitySet {
        let mut set = CapabilitySet::EMPTY;
        for count in &self.counts {
            set.insert(count.capability);
        }
        set
    }
    pub fn denied_set(&self) -> CapabilitySet {
        let mut set = CapabilitySet::EMPTY;
        for count in &self.counts {
            if count.denied > 0 {
                set.insert(count.capability);
            }
        }
        set
    }
    pub fn denied_any(&self) -> bool {
        self.denied > 0
    }
}

pub const DETAIL_MAX: usize = 260;
pub const MAX_RECORDS: usize = 256;
pub const MAX_KEYS: usize = 1024;

#[derive(Default)]
struct AuditState {
    seq: u64,
    records: Vec<AuditRecord>,
    counts: BTreeMap<(Capability, String), (u64, u64)>,
    total: u64,
    denied: u64,
    dropped_records: u64,
    dropped_keys: u64,
    denials_ever: u64,
}

/// Shared handle on the audit log. The runtime is single-threaded on the owning VBA thread, so
/// the log is an `Rc<RefCell<..>>`; clone the handle before the host is moved into an engine.
#[derive(Clone, Default)]
pub struct AuditHandle(Rc<RefCell<AuditState>>);

impl AuditHandle {
    fn record(
        &self,
        capability: Capability,
        op: &'static str,
        detail: String,
        decision: Decision,
        err: Option<i32>,
    ) {
        let mut state = self.0.borrow_mut();
        state.seq += 1;
        state.total += 1;
        let denied = decision == Decision::Denied;
        if denied {
            state.denied += 1;
            state.denials_ever += 1;
        }
        let key = (capability, detail.clone());
        if let Some(entry) = state.counts.get_mut(&key) {
            if denied {
                entry.1 += 1;
            } else {
                entry.0 += 1;
            }
        } else if state.counts.len() < MAX_KEYS {
            state
                .counts
                .insert(key, if denied { (0, 1) } else { (1, 0) });
        } else {
            state.dropped_keys += 1;
        }
        if state.records.len() < MAX_RECORDS {
            let seq = state.seq;
            state.records.push(AuditRecord {
                seq,
                capability,
                op,
                detail,
                decision,
                err_number: err,
            });
        } else {
            state.dropped_records += 1;
        }
    }

    /// Take everything logged since the last drain and start a fresh window.
    pub fn drain(&self) -> AuditSnapshot {
        let mut state = self.0.borrow_mut();
        let counts = std::mem::take(&mut state.counts)
            .into_iter()
            .map(|((capability, detail), (allowed, denied))| AuditCount {
                capability,
                detail,
                allowed,
                denied,
            })
            .collect();
        let snapshot = AuditSnapshot {
            records: std::mem::take(&mut state.records),
            counts,
            total: state.total,
            denied: state.denied,
            dropped_records: state.dropped_records,
            dropped_keys: state.dropped_keys,
        };
        state.seq = 0;
        state.total = 0;
        state.denied = 0;
        state.dropped_records = 0;
        state.dropped_keys = 0;
        snapshot
    }

    /// Denials since this handle was created, across drains. A harness uses it to mark a session
    /// tainted: a denial can leave module state that changes later cases.
    pub fn denials_ever(&self) -> u64 {
        self.0.borrow().denials_ever
    }
}

fn truncate(detail: String) -> String {
    if detail.chars().count() > DETAIL_MAX {
        detail.chars().take(DETAIL_MAX).collect()
    } else {
        detail
    }
}

fn text(value: &Variant) -> String {
    if let Some(s) = value.as_bstr() {
        return s.as_str().to_string();
    }
    if let Some(n) = value.as_i32() {
        return n.to_string();
    }
    String::new()
}

fn settings_detail(parts: &[&Variant]) -> String {
    parts.iter().map(|v| text(v)).collect::<Vec<_>>().join("/")
}

fn declare_detail(descriptor: &DynLinkDescriptorView<'_>) -> String {
    format!("{}!{}", descriptor.library, descriptor.alias)
}

/// `Open` mode codes: 0 is `For Input`; every other mode can write.
fn open_capability(mode: &Variant) -> Capability {
    let code = mode
        .as_i32()
        .or_else(|| mode.as_i16().map(i32::from))
        .or_else(|| mode.as_u8().map(i32::from));
    match code {
        Some(code) if code & 0xFFFF == 0 => Capability::FsRead,
        _ => Capability::FsWrite,
    }
}

/// The policy choice an application passes: Safe denies and audits every capability, Full
/// passes everything through (still audited). "Not at all" is the embedder's choice not to run.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PolicyMode {
    #[default]
    Safe,
    Full,
}

impl PolicyMode {
    pub const fn name(self) -> &'static str {
        match self {
            PolicyMode::Safe => "safe",
            PolicyMode::Full => "full",
        }
    }
    pub fn from_name(name: &str) -> Option<Self> {
        match name {
            "safe" => Some(PolicyMode::Safe),
            "full" => Some(PolicyMode::Full),
            _ => None,
        }
    }
}

/// Wraps a [`HostServices`] with the capability policy and the audit log.
pub struct PolicyHost {
    inner: Arc<dyn HostServices>,
    deny: CapabilitySet,
    audit: AuditHandle,
    /// ProgID prefixes of classes the embedder itself serves (its object model). Activating one
    /// reaches no operating-system service, so it is neither gated nor logged.
    host_progid_prefixes: Vec<String>,
}

impl PolicyHost {
    pub fn with_mode(inner: Arc<dyn HostServices>, mode: PolicyMode) -> Self {
        match mode {
            PolicyMode::Safe => Self::safe(inner),
            PolicyMode::Full => Self::full(inner),
        }
    }

    /// Safe mode: all eight capabilities denied (and logged).
    pub fn safe(inner: Arc<dyn HostServices>) -> Self {
        Self::with_denied(inner, CapabilitySet::ALL)
    }

    /// Full mode: everything passes through, and is logged.
    pub fn full(inner: Arc<dyn HostServices>) -> Self {
        Self::with_denied(inner, CapabilitySet::EMPTY)
    }

    pub fn with_denied(inner: Arc<dyn HostServices>, deny: CapabilitySet) -> Self {
        Self {
            inner,
            deny,
            audit: AuditHandle::default(),
            host_progid_prefixes: Vec::new(),
        }
    }

    /// Classes whose ProgID starts with `prefix` (compared without regard to case) are the
    /// embedder's own objects: activation passes through unlogged.
    pub fn with_host_progid_prefix(mut self, prefix: &str) -> Self {
        self.host_progid_prefixes.push(prefix.to_ascii_lowercase());
        self
    }

    /// The host as the shared handle an engine or embedder holds.
    #[expect(
        clippy::arc_with_non_send_sync,
        reason = "Execution stays on the owning VBA thread, like the HostBuilder hosts."
    )]
    pub fn into_host(self) -> Arc<dyn HostServices> {
        Arc::new(self)
    }

    /// Log into a handle the caller already holds, so several hosts (one per loaded module) can
    /// share one audit log.
    pub fn with_audit(mut self, audit: AuditHandle) -> Self {
        self.audit = audit;
        self
    }

    pub fn audit(&self) -> AuditHandle {
        self.audit.clone()
    }

    pub fn denied_capabilities(&self) -> CapabilitySet {
        self.deny
    }

    fn denial(&self, capability: Capability, op: &'static str) -> HalError {
        let (number, description) = capability.safe_error();
        let mut error =
            HalError::policy_denied(self.inner.profile(), capability.hal_capability(), op)
                .with_host_error_code(number)
                .with_host_error_info(HostErrorInfo {
                    description: Some(description.to_string()),
                    ..HostErrorInfo::default()
                });
        // The library path reads `message` for `Err.Description`; keep both paths identical.
        error.message = description.to_string();
        error
    }

    fn check(&self, capability: Capability, op: &'static str, detail: String) -> HalResult<()> {
        let detail = truncate(detail);
        if self.deny.contains(capability) {
            let (number, _) = capability.safe_error();
            self.audit
                .record(capability, op, detail, Decision::Denied, Some(number));
            Err(self.denial(capability, op))
        } else {
            self.audit
                .record(capability, op, detail, Decision::Allowed, None);
            Ok(())
        }
    }

    fn check_open(&self, op: &'static str, path: &Variant, mode: &Variant) -> HalResult<()> {
        self.check(open_capability(mode), op, text(path))
    }
}

impl HostServices for PolicyHost {
    fn profile(&self) -> HalProfileId {
        self.inner.profile()
    }
    fn descriptor(&self) -> HalDescriptor {
        self.inner.descriptor()
    }
    fn policy(&self) -> &HostPolicy {
        self.inner.policy()
    }
    fn console(&self) -> &dyn ConsoleHal {
        self.inner.console()
    }
    fn ui(&self) -> &dyn UiInteractionHal {
        self
    }
    fn events(&self) -> &dyn EventPumpHal {
        self.inner.events()
    }
    fn fs(&self) -> &dyn FileSystemHal {
        self
    }
    fn process(&self) -> &dyn ProcessEnvHal {
        self
    }
    fn com(&self) -> &dyn ComHal {
        // COM is gated at activation through `authorize_com_activation`; every other
        // COM method is the inner host's own, so no forwarder can be forgotten.
        self.inner.com()
    }
    fn time_locale(&self) -> &dyn TimeLocaleHal {
        self.inner.time_locale()
    }
    fn dynlink(&self) -> &dyn DynamicLinkHal {
        self
    }
    fn diag(&self) -> &dyn DiagnosticsHal {
        self.inner.diag()
    }
    fn project_catalog(&self) -> Option<&dyn ProjectCatalogHal> {
        self.inner.project_catalog()
    }
    fn project_references(&self) -> Option<&dyn ProjectReferenceHal> {
        self.inner.project_references()
    }
    fn project_mutation(&self) -> Option<&dyn ProjectMutationHal> {
        self.inner.project_mutation()
    }
    fn authorize_com_activation(&self, op: &'static str, target: &str) -> HalResult<()> {
        let lowered = target.to_ascii_lowercase();
        if self
            .host_progid_prefixes
            .iter()
            .any(|prefix| lowered.starts_with(prefix.as_str()))
        {
            return self.inner.authorize_com_activation(op, target);
        }
        self.check(Capability::ComActivate, op, target.to_string())?;
        self.inner.authorize_com_activation(op, target)
    }
}

impl FileSystemHal for PolicyHost {
    fn open_variant(&self, path: Variant, mode: Variant) -> HalResult<Variant> {
        self.check_open("open", &path, &mode)?;
        self.inner.fs().open_variant(path, mode)
    }

    fn open_with_record_len(
        &self,
        path: Variant,
        mode: Variant,
        record_len: Variant,
    ) -> HalResult<Variant> {
        self.check_open("open_with_record_len", &path, &mode)?;
        self.inner.fs().open_with_record_len(path, mode, record_len)
    }

    fn close_variant(&self, handle: Variant) -> HalResult<Variant> {
        self.inner.fs().close_variant(handle)
    }

    fn kill_variant(&self, path: Variant) -> HalResult<Variant> {
        self.check(Capability::FsWrite, "kill", text(&path))?;
        self.inner.fs().kill_variant(path)
    }

    fn mkdir_variant(&self, path: Variant) -> HalResult<Variant> {
        self.check(Capability::FsWrite, "mkdir", text(&path))?;
        self.inner.fs().mkdir_variant(path)
    }

    fn rmdir_variant(&self, path: Variant) -> HalResult<Variant> {
        self.check(Capability::FsWrite, "rmdir", text(&path))?;
        self.inner.fs().rmdir_variant(path)
    }

    fn cur_dir_variant(&self, drive: Variant) -> HalResult<Variant> {
        self.check(Capability::FsRead, "cur_dir", text(&drive))?;
        self.inner.fs().cur_dir_variant(drive)
    }

    fn ch_dir_variant(&self, path: Variant) -> HalResult<Variant> {
        self.check(Capability::FsWrite, "ch_dir", text(&path))?;
        self.inner.fs().ch_dir_variant(path)
    }

    fn file_len_variant(&self, path: Variant) -> HalResult<Variant> {
        self.check(Capability::FsRead, "file_len", text(&path))?;
        self.inner.fs().file_len_variant(path)
    }

    fn file_copy_variant(&self, source: Variant, dest: Variant) -> HalResult<Variant> {
        self.check(
            Capability::FsWrite,
            "file_copy",
            format!("{} -> {}", text(&source), text(&dest)),
        )?;
        self.inner.fs().file_copy_variant(source, dest)
    }

    fn get_attr_variant(&self, path: Variant) -> HalResult<Variant> {
        self.check(Capability::FsRead, "get_attr", text(&path))?;
        self.inner.fs().get_attr_variant(path)
    }

    fn set_attr_variant(&self, path: Variant, attributes: Variant) -> HalResult<Variant> {
        self.check(Capability::FsWrite, "set_attr", text(&path))?;
        self.inner.fs().set_attr_variant(path, attributes)
    }

    fn ch_drive_variant(&self, drive: Variant) -> HalResult<Variant> {
        self.check(Capability::FsWrite, "ch_drive", text(&drive))?;
        self.inner.fs().ch_drive_variant(drive)
    }

    fn file_date_time_variant(&self, path: Variant) -> HalResult<Variant> {
        self.check(Capability::FsRead, "file_date_time", text(&path))?;
        self.inner.fs().file_date_time_variant(path)
    }

    fn seek_variant(&self, handle: Variant, position: Variant) -> HalResult<Variant> {
        self.inner.fs().seek_variant(handle, position)
    }

    fn eof_variant(&self, handle: Variant) -> HalResult<Variant> {
        self.inner.fs().eof_variant(handle)
    }

    fn lof_variant(&self, handle: Variant) -> HalResult<Variant> {
        self.inner.fs().lof_variant(handle)
    }

    fn free_file_variant(&self, range_selector: Variant) -> HalResult<Variant> {
        self.inner.fs().free_file_variant(range_selector)
    }

    fn read_bytes_variant(&self, handle: Variant, count: Variant) -> HalResult<Variant> {
        self.inner.fs().read_bytes_variant(handle, count)
    }

    fn input_chars_variant(&self, handle: Variant, count: Variant) -> HalResult<Variant> {
        self.inner.fs().input_chars_variant(handle, count)
    }

    fn write_bytes_variant(&self, handle: Variant, data: Variant) -> HalResult<Variant> {
        self.inner.fs().write_bytes_variant(handle, data)
    }

    fn print_line_variant(&self, handle: Variant, data: Variant) -> HalResult<Variant> {
        self.inner.fs().print_line_variant(handle, data)
    }

    fn print_column_variant(&self, handle: Variant) -> HalResult<Variant> {
        self.inner.fs().print_column_variant(handle)
    }

    fn print_width_variant(&self, handle: Variant) -> HalResult<Variant> {
        self.inner.fs().print_width_variant(handle)
    }

    fn input_fields_variant(&self, handle: Variant, count: Variant) -> HalResult<Variant> {
        self.inner.fs().input_fields_variant(handle, count)
    }

    fn line_input_variant(&self, handle: Variant) -> HalResult<Variant> {
        self.inner.fs().line_input_variant(handle)
    }

    fn loc_variant(&self, handle: Variant) -> HalResult<Variant> {
        self.inner.fs().loc_variant(handle)
    }

    fn put_record_variant(
        &self,
        handle: Variant,
        position: Variant,
        value: Variant,
        fixed: Variant,
    ) -> HalResult<Variant> {
        self.inner
            .fs()
            .put_record_variant(handle, position, value, fixed)
    }

    fn get_record_variant(
        &self,
        handle: Variant,
        position: Variant,
        vtype: Variant,
        str_len: Variant,
    ) -> HalResult<Variant> {
        self.inner
            .fs()
            .get_record_variant(handle, position, vtype, str_len)
    }

    fn width_variant(&self, handle: Variant, width: Variant) -> HalResult<Variant> {
        self.inner.fs().width_variant(handle, width)
    }

    fn name_variant(&self, old_path: Variant, new_path: Variant) -> HalResult<Variant> {
        self.check(
            Capability::FsWrite,
            "name",
            format!("{} -> {}", text(&old_path), text(&new_path)),
        )?;
        self.inner.fs().name_variant(old_path, new_path)
    }

    fn lock_variant(&self, handle: Variant, start: Variant, end: Variant) -> HalResult<Variant> {
        self.inner.fs().lock_variant(handle, start, end)
    }

    fn unlock_variant(&self, handle: Variant, start: Variant, end: Variant) -> HalResult<Variant> {
        self.inner.fs().unlock_variant(handle, start, end)
    }
}

impl ProcessEnvHal for PolicyHost {
    fn command_variant(&self) -> HalResult<Variant> {
        self.check(Capability::Env, "command", "Command".to_string())?;
        self.inner.process().command_variant()
    }

    fn get_setting_variant(
        &self,
        appname: Variant,
        section: Variant,
        key: Variant,
        default: Variant,
    ) -> HalResult<Variant> {
        self.check(
            Capability::Settings,
            "get_setting",
            settings_detail(&[&appname, &section, &key]),
        )?;
        self.inner
            .process()
            .get_setting_variant(appname, section, key, default)
    }

    fn get_all_settings_variant(&self, appname: Variant, section: Variant) -> HalResult<Variant> {
        self.check(
            Capability::Settings,
            "get_all_settings",
            settings_detail(&[&appname, &section]),
        )?;
        self.inner
            .process()
            .get_all_settings_variant(appname, section)
    }

    fn save_setting_variant(
        &self,
        appname: Variant,
        section: Variant,
        key: Variant,
        setting: Variant,
    ) -> HalResult<Variant> {
        self.check(
            Capability::Settings,
            "save_setting",
            settings_detail(&[&appname, &section, &key]),
        )?;
        self.inner
            .process()
            .save_setting_variant(appname, section, key, setting)
    }

    fn delete_setting_variant(
        &self,
        appname: Variant,
        section: Variant,
        key: Variant,
    ) -> HalResult<Variant> {
        self.check(
            Capability::Settings,
            "delete_setting",
            settings_detail(&[&appname, &section, &key]),
        )?;
        self.inner
            .process()
            .delete_setting_variant(appname, section, key)
    }

    fn shell_variant(&self, command: Variant, window_style: Variant) -> HalResult<Variant> {
        self.check(Capability::Process, "shell", text(&command))?;
        self.inner.process().shell_variant(command, window_style)
    }

    fn environ_variant(&self, key: Variant) -> HalResult<Variant> {
        self.check(Capability::Env, "environ", text(&key))?;
        self.inner.process().environ_variant(key)
    }

    fn dir_variant(&self, path: Variant, attrs: Variant) -> HalResult<Variant> {
        self.check(Capability::FsRead, "dir", text(&path))?;
        self.inner.process().dir_variant(path, attrs)
    }
}

impl UiInteractionHal for PolicyHost {
    fn msg_box_variant(&self, prompt: Variant, style: Variant) -> HalResult<Variant> {
        self.inner.ui().msg_box_variant(prompt, style)
    }

    fn input_box_variant(&self, prompt: Variant, default_value: Variant) -> HalResult<Variant> {
        self.inner.ui().input_box_variant(prompt, default_value)
    }

    fn send_keys_variant(&self, keys: Variant, wait: Variant) -> HalResult<Variant> {
        self.check(Capability::UiKeys, "send_keys", String::new())?;
        self.inner.ui().send_keys_variant(keys, wait)
    }

    fn app_activate_variant(&self, title: Variant, wait: Variant) -> HalResult<Variant> {
        self.check(Capability::UiKeys, "app_activate", String::new())?;
        self.inner.ui().app_activate_variant(title, wait)
    }
}

impl DynamicLinkHal for PolicyHost {
    fn bind_descriptor(&self, descriptor: &DynLinkDescriptorView<'_>) -> HalResult<BindingHandle> {
        self.inner.dynlink().bind_descriptor(descriptor)
    }

    fn invoke_bound_variants(
        &self,
        binding: BindingHandle,
        args: &[Variant],
    ) -> HalResult<(Variant, Vec<Variant>)> {
        self.check(
            Capability::Ffi,
            "invoke_bound_variants",
            "bound-call".to_string(),
        )?;
        self.inner.dynlink().invoke_bound_variants(binding, args)
    }

    fn invoke_descriptor_variants(
        &self,
        descriptor: &DynLinkDescriptorView<'_>,
        args: &[Variant],
    ) -> HalResult<(Variant, Vec<Variant>)> {
        self.check(
            Capability::Ffi,
            "invoke_descriptor_variants",
            declare_detail(descriptor),
        )?;
        self.inner
            .dynlink()
            .invoke_descriptor_variants(descriptor, args)
    }

    fn invoke_descriptor_arguments(
        &self,
        descriptor: &DynLinkDescriptorView<'_>,
        args: &[RuntimeCallArgument],
    ) -> HalResult<(Variant, Vec<Variant>)> {
        self.check(
            Capability::Ffi,
            "invoke_descriptor_arguments",
            declare_detail(descriptor),
        )?;
        self.inner
            .dynlink()
            .invoke_descriptor_arguments(descriptor, args)
    }

    fn invoke_symbol_variant(&self, symbol: DynLinkSymbol, arg: &Variant) -> HalResult<Variant> {
        self.check(Capability::Ffi, "invoke_symbol", "symbol-call".to_string())?;
        self.inner.dynlink().invoke_symbol_variant(symbol, arg)
    }

    fn last_dll_error(&self) -> i32 {
        self.inner.dynlink().last_dll_error()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::cell::RefCell;

    /// An inner host that records which HAL method reached it.
    struct Recorder {
        calls: RefCell<Vec<&'static str>>,
        null: Arc<dyn HostServices>,
    }
    impl Recorder {
        fn log(&self, name: &'static str) {
            self.calls.borrow_mut().push(name);
        }
    }
    impl HostServices for Recorder {
        fn profile(&self) -> HalProfileId {
            self.null.profile()
        }
        fn descriptor(&self) -> HalDescriptor {
            self.null.descriptor()
        }
        fn policy(&self) -> &HostPolicy {
            self.null.policy()
        }
        fn console(&self) -> &dyn ConsoleHal {
            self.null.console()
        }
        fn ui(&self) -> &dyn UiInteractionHal {
            self
        }
        fn events(&self) -> &dyn EventPumpHal {
            self.null.events()
        }
        fn fs(&self) -> &dyn FileSystemHal {
            self
        }
        fn process(&self) -> &dyn ProcessEnvHal {
            self
        }
        fn com(&self) -> &dyn ComHal {
            self.null.com()
        }
        fn time_locale(&self) -> &dyn TimeLocaleHal {
            self.null.time_locale()
        }
        fn dynlink(&self) -> &dyn DynamicLinkHal {
            self
        }
        fn diag(&self) -> &dyn DiagnosticsHal {
            self.null.diag()
        }
    }

    impl FileSystemHal for Recorder {
        fn open_variant(&self, _path: Variant, _mode: Variant) -> HalResult<Variant> {
            self.log("open_variant");
            Ok(Variant::from_i32(7))
        }
        fn open_with_record_len(
            &self,
            _path: Variant,
            _mode: Variant,
            _record_len: Variant,
        ) -> HalResult<Variant> {
            self.log("open_with_record_len");
            Ok(Variant::from_i32(7))
        }
        fn close_variant(&self, _handle: Variant) -> HalResult<Variant> {
            self.log("close_variant");
            Ok(Variant::from_i32(7))
        }
        fn kill_variant(&self, _path: Variant) -> HalResult<Variant> {
            self.log("kill_variant");
            Ok(Variant::from_i32(7))
        }
        fn mkdir_variant(&self, _path: Variant) -> HalResult<Variant> {
            self.log("mkdir_variant");
            Ok(Variant::from_i32(7))
        }
        fn rmdir_variant(&self, _path: Variant) -> HalResult<Variant> {
            self.log("rmdir_variant");
            Ok(Variant::from_i32(7))
        }
        fn cur_dir_variant(&self, _drive: Variant) -> HalResult<Variant> {
            self.log("cur_dir_variant");
            Ok(Variant::from_i32(7))
        }
        fn ch_dir_variant(&self, _path: Variant) -> HalResult<Variant> {
            self.log("ch_dir_variant");
            Ok(Variant::from_i32(7))
        }
        fn file_len_variant(&self, _path: Variant) -> HalResult<Variant> {
            self.log("file_len_variant");
            Ok(Variant::from_i32(7))
        }
        fn file_copy_variant(&self, _source: Variant, _dest: Variant) -> HalResult<Variant> {
            self.log("file_copy_variant");
            Ok(Variant::from_i32(7))
        }
        fn get_attr_variant(&self, _path: Variant) -> HalResult<Variant> {
            self.log("get_attr_variant");
            Ok(Variant::from_i32(7))
        }
        fn set_attr_variant(&self, _path: Variant, _attributes: Variant) -> HalResult<Variant> {
            self.log("set_attr_variant");
            Ok(Variant::from_i32(7))
        }
        fn ch_drive_variant(&self, _drive: Variant) -> HalResult<Variant> {
            self.log("ch_drive_variant");
            Ok(Variant::from_i32(7))
        }
        fn file_date_time_variant(&self, _path: Variant) -> HalResult<Variant> {
            self.log("file_date_time_variant");
            Ok(Variant::from_i32(7))
        }
        fn seek_variant(&self, _handle: Variant, _position: Variant) -> HalResult<Variant> {
            self.log("seek_variant");
            Ok(Variant::from_i32(7))
        }
        fn eof_variant(&self, _handle: Variant) -> HalResult<Variant> {
            self.log("eof_variant");
            Ok(Variant::from_i32(7))
        }
        fn lof_variant(&self, _handle: Variant) -> HalResult<Variant> {
            self.log("lof_variant");
            Ok(Variant::from_i32(7))
        }
        fn free_file_variant(&self, _range_selector: Variant) -> HalResult<Variant> {
            self.log("free_file_variant");
            Ok(Variant::from_i32(7))
        }
        fn read_bytes_variant(&self, _handle: Variant, _count: Variant) -> HalResult<Variant> {
            self.log("read_bytes_variant");
            Ok(Variant::from_i32(7))
        }
        fn input_chars_variant(&self, _handle: Variant, _count: Variant) -> HalResult<Variant> {
            self.log("input_chars_variant");
            Ok(Variant::from_i32(7))
        }
        fn write_bytes_variant(&self, _handle: Variant, _data: Variant) -> HalResult<Variant> {
            self.log("write_bytes_variant");
            Ok(Variant::from_i32(7))
        }
        fn print_line_variant(&self, _handle: Variant, _data: Variant) -> HalResult<Variant> {
            self.log("print_line_variant");
            Ok(Variant::from_i32(7))
        }
        fn print_column_variant(&self, _handle: Variant) -> HalResult<Variant> {
            self.log("print_column_variant");
            Ok(Variant::from_i32(7))
        }
        fn print_width_variant(&self, _handle: Variant) -> HalResult<Variant> {
            self.log("print_width_variant");
            Ok(Variant::from_i32(7))
        }
        fn input_fields_variant(&self, _handle: Variant, _count: Variant) -> HalResult<Variant> {
            self.log("input_fields_variant");
            Ok(Variant::from_i32(7))
        }
        fn line_input_variant(&self, _handle: Variant) -> HalResult<Variant> {
            self.log("line_input_variant");
            Ok(Variant::from_i32(7))
        }
        fn loc_variant(&self, _handle: Variant) -> HalResult<Variant> {
            self.log("loc_variant");
            Ok(Variant::from_i32(7))
        }
        fn put_record_variant(
            &self,
            _handle: Variant,
            _position: Variant,
            _value: Variant,
            _fixed: Variant,
        ) -> HalResult<Variant> {
            self.log("put_record_variant");
            Ok(Variant::from_i32(7))
        }
        fn get_record_variant(
            &self,
            _handle: Variant,
            _position: Variant,
            _vtype: Variant,
            _str_len: Variant,
        ) -> HalResult<Variant> {
            self.log("get_record_variant");
            Ok(Variant::from_i32(7))
        }
        fn width_variant(&self, _handle: Variant, _width: Variant) -> HalResult<Variant> {
            self.log("width_variant");
            Ok(Variant::from_i32(7))
        }
        fn name_variant(&self, _old_path: Variant, _new_path: Variant) -> HalResult<Variant> {
            self.log("name_variant");
            Ok(Variant::from_i32(7))
        }
        fn lock_variant(
            &self,
            _handle: Variant,
            _start: Variant,
            _end: Variant,
        ) -> HalResult<Variant> {
            self.log("lock_variant");
            Ok(Variant::from_i32(7))
        }
        fn unlock_variant(
            &self,
            _handle: Variant,
            _start: Variant,
            _end: Variant,
        ) -> HalResult<Variant> {
            self.log("unlock_variant");
            Ok(Variant::from_i32(7))
        }
    }
    impl ProcessEnvHal for Recorder {
        fn command_variant(&self) -> HalResult<Variant> {
            self.log("command_variant");
            Ok(Variant::from_i32(7))
        }
        fn get_setting_variant(
            &self,
            _appname: Variant,
            _section: Variant,
            _key: Variant,
            _default: Variant,
        ) -> HalResult<Variant> {
            self.log("get_setting_variant");
            Ok(Variant::from_i32(7))
        }
        fn get_all_settings_variant(
            &self,
            _appname: Variant,
            _section: Variant,
        ) -> HalResult<Variant> {
            self.log("get_all_settings_variant");
            Ok(Variant::from_i32(7))
        }
        fn save_setting_variant(
            &self,
            _appname: Variant,
            _section: Variant,
            _key: Variant,
            _setting: Variant,
        ) -> HalResult<Variant> {
            self.log("save_setting_variant");
            Ok(Variant::from_i32(7))
        }
        fn delete_setting_variant(
            &self,
            _appname: Variant,
            _section: Variant,
            _key: Variant,
        ) -> HalResult<Variant> {
            self.log("delete_setting_variant");
            Ok(Variant::from_i32(7))
        }
        fn shell_variant(&self, _command: Variant, _window_style: Variant) -> HalResult<Variant> {
            self.log("shell_variant");
            Ok(Variant::from_i32(7))
        }
        fn environ_variant(&self, _key: Variant) -> HalResult<Variant> {
            self.log("environ_variant");
            Ok(Variant::from_i32(7))
        }
        fn dir_variant(&self, _path: Variant, _attrs: Variant) -> HalResult<Variant> {
            self.log("dir_variant");
            Ok(Variant::from_i32(7))
        }
    }
    impl UiInteractionHal for Recorder {
        fn msg_box_variant(&self, _prompt: Variant, _style: Variant) -> HalResult<Variant> {
            self.log("msg_box_variant");
            Ok(Variant::from_i32(7))
        }
        fn input_box_variant(
            &self,
            _prompt: Variant,
            _default_value: Variant,
        ) -> HalResult<Variant> {
            self.log("input_box_variant");
            Ok(Variant::from_i32(7))
        }
        fn send_keys_variant(&self, _keys: Variant, _wait: Variant) -> HalResult<Variant> {
            self.log("send_keys_variant");
            Ok(Variant::from_i32(7))
        }
        fn app_activate_variant(&self, _title: Variant, _wait: Variant) -> HalResult<Variant> {
            self.log("app_activate_variant");
            Ok(Variant::from_i32(7))
        }
    }
    impl DynamicLinkHal for Recorder {
        fn bind_descriptor(
            &self,
            _descriptor: &DynLinkDescriptorView<'_>,
        ) -> HalResult<BindingHandle> {
            self.log("bind_descriptor");
            Ok(BindingHandle::from(1))
        }
        fn invoke_bound_variants(
            &self,
            _binding: BindingHandle,
            _args: &[Variant],
        ) -> HalResult<(Variant, Vec<Variant>)> {
            self.log("invoke_bound_variants");
            Ok((Variant::from_i32(7), Vec::new()))
        }
        fn invoke_descriptor_variants(
            &self,
            _descriptor: &DynLinkDescriptorView<'_>,
            _args: &[Variant],
        ) -> HalResult<(Variant, Vec<Variant>)> {
            self.log("invoke_descriptor_variants");
            Ok((Variant::from_i32(7), Vec::new()))
        }
        fn invoke_descriptor_arguments(
            &self,
            _descriptor: &DynLinkDescriptorView<'_>,
            _args: &[RuntimeCallArgument],
        ) -> HalResult<(Variant, Vec<Variant>)> {
            self.log("invoke_descriptor_arguments");
            Ok((Variant::from_i32(7), Vec::new()))
        }
        fn invoke_symbol_variant(
            &self,
            _symbol: DynLinkSymbol,
            _arg: &Variant,
        ) -> HalResult<Variant> {
            self.log("invoke_symbol_variant");
            Ok(Variant::from_i32(7))
        }
        fn last_dll_error(&self) -> i32 {
            self.log("last_dll_error");
            42
        }
    }

    fn v() -> Variant {
        Variant::from_string("x")
    }

    fn view() -> DynLinkDescriptorView<'static> {
        DynLinkDescriptorView {
            descriptor_id: 1,
            declared_name: "F",
            library: "lib",
            alias: "F",
            ordinal_alias: false,
            symbol: 1.into(),
            marshal_lane: "m1-native-ffi",
            calling_convention: "platform-default",
            selection_policy: "case-insensitive-canonical",
            param_count: 0,
            param_types: &[],
            param_by_ref: &[],
            return_type: None,
            return_record_layout: None,
        }
    }

    #[expect(clippy::arc_with_non_send_sync, reason = "test host; single-threaded")]
    fn recorder() -> Arc<Recorder> {
        Arc::new(Recorder {
            calls: RefCell::new(Vec::new()),
            null: crate::adapters::null::NullHostServices::boxed(HostPolicy::default()),
        })
    }

    type Call = Box<dyn Fn(&dyn HostServices) -> bool>;

    /// Every method of the four wrapped facets as `(name, is_gated, call)`; the wrapper
    /// must forward each to the inner host, and deny the gated ones in Safe mode.
    fn all_calls() -> Vec<(&'static str, bool, Call)> {
        vec![
            (
                "open_variant",
                true,
                Box::new(|host| host.fs().open_variant(v(), v()).is_ok()),
            ),
            (
                "open_with_record_len",
                true,
                Box::new(|host| host.fs().open_with_record_len(v(), v(), v()).is_ok()),
            ),
            (
                "close_variant",
                false,
                Box::new(|host| host.fs().close_variant(v()).is_ok()),
            ),
            (
                "kill_variant",
                true,
                Box::new(|host| host.fs().kill_variant(v()).is_ok()),
            ),
            (
                "mkdir_variant",
                true,
                Box::new(|host| host.fs().mkdir_variant(v()).is_ok()),
            ),
            (
                "rmdir_variant",
                true,
                Box::new(|host| host.fs().rmdir_variant(v()).is_ok()),
            ),
            (
                "cur_dir_variant",
                true,
                Box::new(|host| host.fs().cur_dir_variant(v()).is_ok()),
            ),
            (
                "ch_dir_variant",
                true,
                Box::new(|host| host.fs().ch_dir_variant(v()).is_ok()),
            ),
            (
                "file_len_variant",
                true,
                Box::new(|host| host.fs().file_len_variant(v()).is_ok()),
            ),
            (
                "file_copy_variant",
                true,
                Box::new(|host| host.fs().file_copy_variant(v(), v()).is_ok()),
            ),
            (
                "get_attr_variant",
                true,
                Box::new(|host| host.fs().get_attr_variant(v()).is_ok()),
            ),
            (
                "set_attr_variant",
                true,
                Box::new(|host| host.fs().set_attr_variant(v(), v()).is_ok()),
            ),
            (
                "ch_drive_variant",
                true,
                Box::new(|host| host.fs().ch_drive_variant(v()).is_ok()),
            ),
            (
                "file_date_time_variant",
                true,
                Box::new(|host| host.fs().file_date_time_variant(v()).is_ok()),
            ),
            (
                "seek_variant",
                false,
                Box::new(|host| host.fs().seek_variant(v(), v()).is_ok()),
            ),
            (
                "eof_variant",
                false,
                Box::new(|host| host.fs().eof_variant(v()).is_ok()),
            ),
            (
                "lof_variant",
                false,
                Box::new(|host| host.fs().lof_variant(v()).is_ok()),
            ),
            (
                "free_file_variant",
                false,
                Box::new(|host| host.fs().free_file_variant(v()).is_ok()),
            ),
            (
                "read_bytes_variant",
                false,
                Box::new(|host| host.fs().read_bytes_variant(v(), v()).is_ok()),
            ),
            (
                "input_chars_variant",
                false,
                Box::new(|host| host.fs().input_chars_variant(v(), v()).is_ok()),
            ),
            (
                "write_bytes_variant",
                false,
                Box::new(|host| host.fs().write_bytes_variant(v(), v()).is_ok()),
            ),
            (
                "print_line_variant",
                false,
                Box::new(|host| host.fs().print_line_variant(v(), v()).is_ok()),
            ),
            (
                "print_column_variant",
                false,
                Box::new(|host| host.fs().print_column_variant(v()).is_ok()),
            ),
            (
                "print_width_variant",
                false,
                Box::new(|host| host.fs().print_width_variant(v()).is_ok()),
            ),
            (
                "input_fields_variant",
                false,
                Box::new(|host| host.fs().input_fields_variant(v(), v()).is_ok()),
            ),
            (
                "line_input_variant",
                false,
                Box::new(|host| host.fs().line_input_variant(v()).is_ok()),
            ),
            (
                "loc_variant",
                false,
                Box::new(|host| host.fs().loc_variant(v()).is_ok()),
            ),
            (
                "put_record_variant",
                false,
                Box::new(|host| host.fs().put_record_variant(v(), v(), v(), v()).is_ok()),
            ),
            (
                "get_record_variant",
                false,
                Box::new(|host| host.fs().get_record_variant(v(), v(), v(), v()).is_ok()),
            ),
            (
                "width_variant",
                false,
                Box::new(|host| host.fs().width_variant(v(), v()).is_ok()),
            ),
            (
                "name_variant",
                true,
                Box::new(|host| host.fs().name_variant(v(), v()).is_ok()),
            ),
            (
                "lock_variant",
                false,
                Box::new(|host| host.fs().lock_variant(v(), v(), v()).is_ok()),
            ),
            (
                "unlock_variant",
                false,
                Box::new(|host| host.fs().unlock_variant(v(), v(), v()).is_ok()),
            ),
            (
                "command_variant",
                true,
                Box::new(|host| host.process().command_variant().is_ok()),
            ),
            (
                "get_setting_variant",
                true,
                Box::new(|host| {
                    host.process()
                        .get_setting_variant(v(), v(), v(), v())
                        .is_ok()
                }),
            ),
            (
                "get_all_settings_variant",
                true,
                Box::new(|host| host.process().get_all_settings_variant(v(), v()).is_ok()),
            ),
            (
                "save_setting_variant",
                true,
                Box::new(|host| {
                    host.process()
                        .save_setting_variant(v(), v(), v(), v())
                        .is_ok()
                }),
            ),
            (
                "delete_setting_variant",
                true,
                Box::new(|host| host.process().delete_setting_variant(v(), v(), v()).is_ok()),
            ),
            (
                "shell_variant",
                true,
                Box::new(|host| host.process().shell_variant(v(), v()).is_ok()),
            ),
            (
                "environ_variant",
                true,
                Box::new(|host| host.process().environ_variant(v()).is_ok()),
            ),
            (
                "dir_variant",
                true,
                Box::new(|host| host.process().dir_variant(v(), v()).is_ok()),
            ),
            (
                "msg_box_variant",
                false,
                Box::new(|host| host.ui().msg_box_variant(v(), v()).is_ok()),
            ),
            (
                "input_box_variant",
                false,
                Box::new(|host| host.ui().input_box_variant(v(), v()).is_ok()),
            ),
            (
                "send_keys_variant",
                true,
                Box::new(|host| host.ui().send_keys_variant(v(), v()).is_ok()),
            ),
            (
                "app_activate_variant",
                true,
                Box::new(|host| host.ui().app_activate_variant(v(), v()).is_ok()),
            ),
            (
                "bind_descriptor",
                false,
                Box::new(|host| host.dynlink().bind_descriptor(&view()).is_ok()),
            ),
            (
                "invoke_bound_variants",
                true,
                Box::new(|host| {
                    host.dynlink()
                        .invoke_bound_variants(BindingHandle::from(1), &[])
                        .is_ok()
                }),
            ),
            (
                "invoke_descriptor_variants",
                true,
                Box::new(|host| {
                    host.dynlink()
                        .invoke_descriptor_variants(&view(), &[])
                        .is_ok()
                }),
            ),
            (
                "invoke_descriptor_arguments",
                true,
                Box::new(|host| {
                    host.dynlink()
                        .invoke_descriptor_arguments(&view(), &[])
                        .is_ok()
                }),
            ),
            (
                "invoke_symbol_variant",
                true,
                Box::new(|host| {
                    host.dynlink()
                        .invoke_symbol_variant(DynLinkSymbol::from(1), &v())
                        .is_ok()
                }),
            ),
            (
                "last_dll_error",
                false,
                Box::new(|host| {
                    let _ = host.dynlink().last_dll_error();
                    true
                }),
            ),
        ]
    }

    #[test]
    fn full_mode_forwards_every_method_to_the_inner_host() {
        for (name, _, call) in all_calls() {
            let inner = recorder();
            let host = PolicyHost::full(inner.clone());
            assert!(call(&host), "{name} failed through the wrapper");
            assert_eq!(
                inner.calls.borrow().as_slice(),
                [name],
                "{name} did not reach the inner host"
            );
        }
    }

    #[test]
    fn safe_mode_denies_gated_methods_and_forwards_the_rest() {
        for (name, gated, call) in all_calls() {
            let inner = recorder();
            let host = PolicyHost::safe(inner.clone());
            let ok = call(&host);
            if gated {
                assert!(!ok, "{name} should be denied in Safe mode");
                assert!(
                    inner.calls.borrow().is_empty(),
                    "{name} reached the inner host"
                );
                assert_eq!(host.audit().drain().denied, 1, "{name}");
            } else {
                assert!(ok, "{name} should pass through in Safe mode");
                assert_eq!(inner.calls.borrow().as_slice(), [name]);
                assert_eq!(host.audit().drain().total, 0, "{name} must not be logged");
            }
        }
    }

    #[test]
    fn com_activation_is_gated_and_audited() {
        let inner = recorder();
        let safe = PolicyHost::safe(inner.clone());
        let err = safe
            .authorize_com_activation("create_object", "Scripting.Dictionary")
            .expect_err("denied");
        assert_eq!(err.host_error_code, Some(429));
        assert_eq!(err.message, "ActiveX component can't create object");
        let full = PolicyHost::full(inner);
        full.authorize_com_activation("get_object", "Excel.Application")
            .expect("allowed");
        let snapshot = full.audit().drain();
        assert_eq!(snapshot.records.len(), 1);
        assert_eq!(snapshot.records[0].detail, "Excel.Application");
        assert_eq!(snapshot.records[0].decision, Decision::Allowed);
    }

    #[test]
    fn embedder_classes_pass_unlogged() {
        let host = PolicyHost::safe(recorder()).with_host_progid_prefix("DnaCalc.");
        host.authorize_com_activation("create_object", "DnaCalc.Application")
            .expect("host class");
        assert_eq!(host.audit().drain().total, 0);
        assert!(
            host.authorize_com_activation("create_object", "Scripting.Dictionary")
                .is_err()
        );
    }

    #[test]
    fn audit_log_is_bounded_with_exact_counts() {
        let host = PolicyHost::safe(recorder());
        let p = Variant::from_string("C:/x");
        for _ in 0..(MAX_RECORDS * 3) {
            let _ = host.fs().kill_variant(p.clone());
        }
        for i in 0..(MAX_KEYS + 5) {
            let _ = host
                .fs()
                .kill_variant(Variant::from_string(format!("p{i}").as_str()));
        }
        let long = "a".repeat(DETAIL_MAX * 2);
        let _ = host.fs().kill_variant(Variant::from_string(long.as_str()));
        let handle = host.audit();
        let snapshot = handle.drain();
        assert_eq!(snapshot.records.len(), MAX_RECORDS);
        assert!(snapshot.dropped_records > 0);
        assert_eq!(snapshot.total, (MAX_RECORDS * 3 + MAX_KEYS + 5 + 1) as u64);
        assert_eq!(snapshot.counts.len(), MAX_KEYS);
        assert!(snapshot.dropped_keys > 0);
        let kept = snapshot
            .counts
            .iter()
            .find(|c| c.detail == "C:/x")
            .expect("exact count");
        assert_eq!(kept.denied, (MAX_RECORDS * 3) as u64);
        assert!(
            snapshot
                .records
                .iter()
                .all(|r| r.detail.chars().count() <= DETAIL_MAX)
        );
        assert_eq!(handle.drain().total, 0, "drain resets the window");
        assert!(handle.denials_ever() > 0);
    }
}
