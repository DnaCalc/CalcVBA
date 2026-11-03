//! Contract model for HAL capabilities, profile descriptors, and host policy.

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum HalProfileId {
    Windows,
    Linux,
    MacOs,
    Wasm,
    Null,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum HalRuntimeClass {
    HostNative,
    WindowsStdio,
    WindowsGui,
    WindowsHeadless,
    LinuxStdio,
    LinuxHeadless,
    MacOsGui,
    MacOsHeadless,
    WasmWasiLocal,
    WasmBrowserSandbox,
    NullFloor,
}

impl HalRuntimeClass {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::HostNative => "host-native",
            Self::WindowsStdio => "windows-stdio",
            Self::WindowsGui => "windows-gui",
            Self::WindowsHeadless => "windows-headless",
            Self::LinuxStdio => "linux-stdio",
            Self::LinuxHeadless => "linux-headless",
            Self::MacOsGui => "macos-gui",
            Self::MacOsHeadless => "macos-headless",
            Self::WasmWasiLocal => "wasi-local",
            Self::WasmBrowserSandbox => "browser-sandbox",
            Self::NullFloor => "null-floor",
        }
    }

    pub const fn default_for(profile: HalProfileId, wasm_runtime_class: WasmRuntimeClass) -> Self {
        match profile {
            HalProfileId::Windows => Self::HostNative,
            HalProfileId::Linux => Self::HostNative,
            HalProfileId::MacOs => Self::HostNative,
            HalProfileId::Wasm => match wasm_runtime_class {
                WasmRuntimeClass::Wasi => Self::WasmWasiLocal,
                WasmRuntimeClass::BrowserSandbox => Self::WasmBrowserSandbox,
            },
            HalProfileId::Null => Self::NullFloor,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CapabilityId {
    ConsoleIo,
    UiInteraction,
    EventPump,
    FileSystemIo,
    ProcessEnv,
    ComActivationDispatch,
    TimeLocale,
    DynamicLinking,
    DiagnosticsTelemetry,
    ProjectCatalog,
    ProjectReferenceProvider,
    ProjectMutation,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CapabilityDescriptor {
    pub id: CapabilityId,
    pub supported: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HalDescriptor {
    pub profile: HalProfileId,
    pub runtime_class: &'static str,
    pub capabilities: Vec<CapabilityDescriptor>,
}

impl HalDescriptor {
    pub fn capability(&self, id: CapabilityId) -> Option<&CapabilityDescriptor> {
        self.capabilities.iter().find(|entry| entry.id == id)
    }

    pub fn supports(&self, id: CapabilityId) -> bool {
        self.capability(id).is_some_and(|entry| entry.supported)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UiVirtualizationMode {
    Disabled,
    ScriptedResponses,
    FailOnPrompt,
    HostCallback,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnsupportedFeatureMode {
    CompileTime,
    Runtime,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
/// Transport preference for dynamic COM calls. A compiler-bound dual member
/// always uses its declared vtable; this preference never erases static binding.
/// Capability denial remains controlled by allow_com_activation/profile support.
pub enum ComInvocationStrategy {
    DispatchOnly,
    PreferVtable,
}

impl ComInvocationStrategy {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::DispatchOnly => "dispatch-only",
            Self::PreferVtable => "prefer-vtable",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WasmRuntimeClass {
    Wasi,
    BrowserSandbox,
}

impl WasmRuntimeClass {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Wasi => "wasi",
            Self::BrowserSandbox => "browser-sandbox",
        }
    }
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HostPolicy {
    pub runtime_class: Option<HalRuntimeClass>,
    pub allow_interaction: bool,
    pub allow_process_spawn: bool,
    pub allow_filesystem_mutation: bool,
    pub allow_dynamic_link: bool,
    pub allow_com_activation: bool,
    pub deterministic_mode: bool,
    pub ui_virtualization: UiVirtualizationMode,
    pub unsupported_feature_mode: UnsupportedFeatureMode,
    pub com_invocation_strategy: ComInvocationStrategy,
    pub wasm_runtime_class: WasmRuntimeClass,
}

impl HostPolicy {
    /// Enable the native host services for ordinary Office-style execution.
    pub fn native() -> Self {
        Self {
            runtime_class: None,
            allow_interaction: true,
            allow_process_spawn: true,
            allow_filesystem_mutation: true,
            allow_dynamic_link: true,
            allow_com_activation: true,
            deterministic_mode: false,
            ui_virtualization: UiVirtualizationMode::Disabled,
            unsupported_feature_mode: UnsupportedFeatureMode::Runtime,
            // Normal-operation default: in-process dual-interface members dispatch
            // through the COM vtable (slot = oVft / live-pointer-size, bound-checked
            // against the FDUAL partner interface's cbSizeVft); out-of-process
            // marshaling proxies (IID_IProxyManager) and ineligible shapes fall back
            // to IDispatch::Invoke for genuinely dynamic calls.
            com_invocation_strategy: ComInvocationStrategy::PreferVtable,
            wasm_runtime_class: WasmRuntimeClass::Wasi,
        }
    }

    pub fn with_wasm_runtime_class(mut self, runtime_class: WasmRuntimeClass) -> Self {
        self.wasm_runtime_class = runtime_class;
        self
    }

    pub fn with_runtime_class(mut self, runtime_class: HalRuntimeClass) -> Self {
        self.runtime_class = Some(runtime_class);
        self
    }

    pub fn with_com_invocation_strategy(mut self, strategy: ComInvocationStrategy) -> Self {
        self.com_invocation_strategy = strategy;
        self
    }
}

impl Default for HostPolicy {
    fn default() -> Self {
        Self {
            runtime_class: None,
            allow_interaction: false,
            allow_process_spawn: true,
            allow_filesystem_mutation: false,
            allow_dynamic_link: false,
            allow_com_activation: true,
            deterministic_mode: true,
            ui_virtualization: UiVirtualizationMode::ScriptedResponses,
            unsupported_feature_mode: UnsupportedFeatureMode::Runtime,
            com_invocation_strategy: ComInvocationStrategy::DispatchOnly,
            wasm_runtime_class: WasmRuntimeClass::Wasi,
        }
    }
}

pub fn host_backed_profile_matches_host(profile: HalProfileId) -> bool {
    match profile {
        HalProfileId::Windows => cfg!(target_os = "windows"),
        HalProfileId::Linux => cfg!(target_os = "linux"),
        _ => false,
    }
}

pub fn native_host_profile() -> HalProfileId {
    if cfg!(target_os = "windows") {
        HalProfileId::Windows
    } else if cfg!(target_os = "linux") {
        HalProfileId::Linux
    } else if cfg!(target_os = "macos") {
        HalProfileId::MacOs
    } else {
        HalProfileId::Null
    }
}

pub fn host_backed_mode_active(profile: HalProfileId, policy: &HostPolicy) -> bool {
    !policy.deterministic_mode && host_backed_profile_matches_host(profile)
}

#[cfg(test)]
mod tests {
    use super::{
        ComInvocationStrategy, HalProfileId, HostPolicy, UiVirtualizationMode,
        UnsupportedFeatureMode, WasmRuntimeClass, host_backed_profile_matches_host,
    };
    #[test]
    fn native_host_allows_platform_operations() {
        let policy = HostPolicy::native();
        assert!(policy.allow_interaction);
        assert!(policy.allow_process_spawn);
        assert!(policy.allow_filesystem_mutation);
        assert!(policy.allow_dynamic_link);
        assert!(policy.allow_com_activation);
        assert!(!policy.deterministic_mode);
        assert_eq!(policy.ui_virtualization, UiVirtualizationMode::Disabled);
        assert_eq!(
            policy.unsupported_feature_mode,
            UnsupportedFeatureMode::Runtime
        );
        // Interactive (normal-operation) default prefers the in-process COM vtable.
        assert_eq!(
            policy.com_invocation_strategy,
            ComInvocationStrategy::PreferVtable
        );
        assert_eq!(policy.wasm_runtime_class, WasmRuntimeClass::Wasi);
    }

    #[test]
    fn com_invocation_strategy_override_is_available() {
        let policy =
            HostPolicy::native().with_com_invocation_strategy(ComInvocationStrategy::PreferVtable);
        assert_eq!(
            policy.com_invocation_strategy,
            ComInvocationStrategy::PreferVtable
        );
        assert_eq!(policy.com_invocation_strategy.as_str(), "prefer-vtable");
    }

    #[test]
    fn wasm_runtime_class_override_is_available() {
        let policy =
            HostPolicy::default().with_wasm_runtime_class(WasmRuntimeClass::BrowserSandbox);
        assert_eq!(policy.wasm_runtime_class, WasmRuntimeClass::BrowserSandbox);
    }

    #[test]
    fn host_backed_profile_match_function_is_stable() {
        let windows = host_backed_profile_matches_host(HalProfileId::Windows);
        let linux = host_backed_profile_matches_host(HalProfileId::Linux);
        let macos = host_backed_profile_matches_host(HalProfileId::MacOs);
        let wasm = host_backed_profile_matches_host(HalProfileId::Wasm);
        let null = host_backed_profile_matches_host(HalProfileId::Null);
        assert!(!macos);
        assert!(!wasm);
        assert!(!null);
        if cfg!(target_os = "windows") {
            assert!(windows);
            assert!(!linux);
        } else if cfg!(target_os = "linux") {
            assert!(!windows);
            assert!(linux);
        } else {
            assert!(!windows);
            assert!(!linux);
        }
    }
}
