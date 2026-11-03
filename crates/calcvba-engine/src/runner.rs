use calcvba_hal::model::{HalProfileId, HalRuntimeClass, HostPolicy, WasmRuntimeClass};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuntimeProfileId {
    WindowsGui,
    WindowsStdio,
    WindowsHeadless,
    LinuxStdio,
    WasmWasiLocal,
    WasmBrowserSandbox,
    NullFloor,
    MacOsHeadless,
}

impl RuntimeProfileId {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::WindowsGui => "windows-gui",
            Self::WindowsStdio => "windows-stdio",
            Self::WindowsHeadless => "windows-headless",
            Self::LinuxStdio => "linux-stdio",
            Self::WasmWasiLocal => "wasm-wasi-local",
            Self::WasmBrowserSandbox => "wasm-browser-sandbox",
            Self::NullFloor => "null-floor",
            Self::MacOsHeadless => "macos-headless",
        }
    }

    pub fn parse(raw: &str) -> Result<Self, String> {
        match raw.trim().to_ascii_lowercase().as_str() {
            "windows-gui" => Ok(Self::WindowsGui),
            "windows-stdio" => Ok(Self::WindowsStdio),
            "windows-headless" => Ok(Self::WindowsHeadless),
            "linux-stdio" => Ok(Self::LinuxStdio),
            "wasm-wasi-local" => Ok(Self::WasmWasiLocal),
            "wasm-browser-sandbox" => Ok(Self::WasmBrowserSandbox),
            "null-floor" => Ok(Self::NullFloor),
            "macos-headless" => Ok(Self::MacOsHeadless),
            other => Err(format!("invalid runtime profile: {other}")),
        }
    }

    pub const fn hal_profile(self) -> HalProfileId {
        match self {
            Self::WindowsGui | Self::WindowsStdio | Self::WindowsHeadless => HalProfileId::Windows,
            Self::LinuxStdio => HalProfileId::Linux,
            Self::WasmWasiLocal | Self::WasmBrowserSandbox => HalProfileId::Wasm,
            Self::NullFloor => HalProfileId::Null,
            Self::MacOsHeadless => HalProfileId::MacOs,
        }
    }

    pub const fn runtime_class(self) -> HalRuntimeClass {
        match self {
            Self::WindowsGui => HalRuntimeClass::WindowsGui,
            Self::WindowsStdio => HalRuntimeClass::WindowsStdio,
            Self::WindowsHeadless => HalRuntimeClass::WindowsHeadless,
            Self::LinuxStdio => HalRuntimeClass::LinuxStdio,
            Self::WasmWasiLocal => HalRuntimeClass::WasmWasiLocal,
            Self::WasmBrowserSandbox => HalRuntimeClass::WasmBrowserSandbox,
            Self::NullFloor => HalRuntimeClass::NullFloor,
            Self::MacOsHeadless => HalRuntimeClass::MacOsHeadless,
        }
    }

    pub const fn default_for_hal_profile(profile: HalProfileId) -> Self {
        match profile {
            HalProfileId::Windows => Self::WindowsStdio,
            HalProfileId::Linux => Self::LinuxStdio,
            HalProfileId::MacOs => Self::MacOsHeadless,
            HalProfileId::Wasm => Self::WasmWasiLocal,
            HalProfileId::Null => Self::NullFloor,
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RunnerBootstrapOptions {
    pub profile: Option<String>,
}
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RunnerBootstrapFallbacks {
    pub profile: Option<String>,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ResolvedRunnerBootstrap {
    pub runtime_profile: RuntimeProfileId,
    pub policy: HostPolicy,
}
impl ResolvedRunnerBootstrap {
    pub fn fingerprint(&self) -> String {
        format!(
            "profile={};hal_profile={:?};runtime_class={}",
            self.runtime_profile.as_str(),
            self.runtime_profile.hal_profile(),
            self.runtime_profile.runtime_class().as_str()
        )
    }
}
pub fn resolve_runner_bootstrap(
    options: &RunnerBootstrapOptions,
    env_get: impl Fn(&str) -> Option<String>,
) -> Result<ResolvedRunnerBootstrap, String> {
    resolve_runner_bootstrap_with_fallbacks(options, &RunnerBootstrapFallbacks::default(), env_get)
}
pub fn resolve_runner_bootstrap_with_fallbacks(
    options: &RunnerBootstrapOptions,
    fallbacks: &RunnerBootstrapFallbacks,
    env_get: impl Fn(&str) -> Option<String>,
) -> Result<ResolvedRunnerBootstrap, String> {
    let selected = options
        .profile
        .clone()
        .or_else(|| env_get("OXVBA_PROFILE"))
        .or_else(|| fallbacks.profile.clone());
    let runtime_profile = selected
        .as_deref()
        .map(RuntimeProfileId::parse)
        .transpose()?
        .unwrap_or_else(|| {
            RuntimeProfileId::default_for_hal_profile(calcvba_hal::model::native_host_profile())
        });
    let mut policy = HostPolicy::native();
    policy.runtime_class = Some(runtime_profile.runtime_class());
    match runtime_profile {
        RuntimeProfileId::WasmWasiLocal => policy.wasm_runtime_class = WasmRuntimeClass::Wasi,
        RuntimeProfileId::WasmBrowserSandbox => {
            policy.wasm_runtime_class = WasmRuntimeClass::BrowserSandbox;
            policy.allow_interaction = false;
        }
        _ => {}
    }
    Ok(ResolvedRunnerBootstrap {
        runtime_profile,
        policy,
    })
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn cli_profile_overrides_environment_and_project() {
        let result = resolve_runner_bootstrap_with_fallbacks(
            &RunnerBootstrapOptions {
                profile: Some("linux-stdio".into()),
            },
            &RunnerBootstrapFallbacks {
                profile: Some("null-floor".into()),
            },
            |_| Some("windows-headless".into()),
        )
        .unwrap();
        assert_eq!(result.runtime_profile, RuntimeProfileId::LinuxStdio);
    }
    #[test]
    fn environment_overrides_project_and_project_supplies_fallback() {
        let fallback = RunnerBootstrapFallbacks {
            profile: Some("linux-stdio".into()),
        };
        assert_eq!(
            resolve_runner_bootstrap_with_fallbacks(
                &RunnerBootstrapOptions::default(),
                &fallback,
                |_| Some("windows-headless".into())
            )
            .unwrap()
            .runtime_profile,
            RuntimeProfileId::WindowsHeadless
        );
        assert_eq!(
            resolve_runner_bootstrap_with_fallbacks(
                &RunnerBootstrapOptions::default(),
                &fallback,
                |_| None
            )
            .unwrap()
            .runtime_profile,
            RuntimeProfileId::LinuxStdio
        );
    }
    #[test]
    fn default_profile_matches_the_current_platform() {
        let result =
            resolve_runner_bootstrap(&RunnerBootstrapOptions::default(), |_| None).unwrap();
        assert_eq!(
            result.runtime_profile,
            RuntimeProfileId::default_for_hal_profile(calcvba_hal::model::native_host_profile())
        );
    }
    #[test]
    fn browser_profile_selects_the_browser_adapter_without_native_prompts() {
        let result = resolve_runner_bootstrap(
            &RunnerBootstrapOptions {
                profile: Some("wasm-browser-sandbox".into()),
            },
            |_| None,
        )
        .unwrap();
        assert_eq!(
            result.policy.wasm_runtime_class,
            WasmRuntimeClass::BrowserSandbox
        );
        assert!(!result.policy.allow_interaction);
    }
}
