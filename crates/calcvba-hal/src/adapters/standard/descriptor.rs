use crate::model::{
    CapabilityDescriptor, CapabilityId, HalDescriptor, HalProfileId, HalRuntimeClass, HostPolicy,
    WasmRuntimeClass,
};

pub(crate) fn descriptor_for_profile(
    profile: HalProfileId,
    runtime_class: HalRuntimeClass,
    policy: &HostPolicy,
) -> HalDescriptor {
    HalDescriptor {
        profile,
        runtime_class: runtime_class.as_str(),
        capabilities: capability_matrix(profile, runtime_class, policy.wasm_runtime_class),
    }
}

fn capability_matrix(
    profile: HalProfileId,
    runtime_class: HalRuntimeClass,
    wasm_runtime_class: WasmRuntimeClass,
) -> Vec<CapabilityDescriptor> {
    use CapabilityId as C;
    let supports_console = matches!(
        runtime_class,
        HalRuntimeClass::WindowsStdio | HalRuntimeClass::LinuxStdio
    );
    let mut out = Vec::new();
    let mut push = |id: C, supported: bool| {
        out.push(CapabilityDescriptor { id, supported });
    };

    match profile {
        HalProfileId::Windows => {
            push(C::ConsoleIo, supports_console);
            push(C::UiInteraction, true);
            push(C::EventPump, true);
            push(C::FileSystemIo, true);
            push(C::ProcessEnv, true);
            push(C::ComActivationDispatch, true);
            push(C::TimeLocale, true);
            push(C::DynamicLinking, true);
            push(C::DiagnosticsTelemetry, true);
            push(C::ProjectCatalog, false);
            push(C::ProjectReferenceProvider, false);
            push(C::ProjectMutation, false);
        }
        HalProfileId::Linux => {
            push(C::ConsoleIo, supports_console);
            push(C::UiInteraction, true);
            push(C::EventPump, true);
            push(C::FileSystemIo, true);
            push(C::ProcessEnv, true);
            push(C::ComActivationDispatch, false);
            push(C::TimeLocale, true);
            push(C::DynamicLinking, true);
            push(C::DiagnosticsTelemetry, true);
            push(C::ProjectCatalog, false);
            push(C::ProjectReferenceProvider, false);
            push(C::ProjectMutation, false);
        }
        HalProfileId::MacOs => {
            push(C::ConsoleIo, false);
            push(C::UiInteraction, true);
            push(C::EventPump, true);
            push(C::FileSystemIo, true);
            push(C::ProcessEnv, true);
            push(C::ComActivationDispatch, false);
            push(C::TimeLocale, true);
            push(C::DynamicLinking, true);
            push(C::DiagnosticsTelemetry, true);
            push(C::ProjectCatalog, false);
            push(C::ProjectReferenceProvider, false);
            push(C::ProjectMutation, false);
        }
        HalProfileId::Wasm => match runtime_class {
            HalRuntimeClass::WasmWasiLocal => {
                push(C::ConsoleIo, false);
                push(C::UiInteraction, true);
                push(C::EventPump, true);
                push(C::FileSystemIo, false);
                push(C::ProcessEnv, false);
                push(C::ComActivationDispatch, false);
                push(C::TimeLocale, true);
                push(C::DynamicLinking, false);
                push(C::DiagnosticsTelemetry, true);
                push(C::ProjectCatalog, false);
                push(C::ProjectReferenceProvider, false);
                push(C::ProjectMutation, false);
            }
            HalRuntimeClass::WasmBrowserSandbox => {
                push(C::ConsoleIo, false);
                push(C::UiInteraction, false);
                push(C::EventPump, true);
                push(C::FileSystemIo, false);
                push(C::ProcessEnv, false);
                push(C::ComActivationDispatch, false);
                push(C::TimeLocale, true);
                push(C::DynamicLinking, false);
                push(C::DiagnosticsTelemetry, true);
                push(C::ProjectCatalog, false);
                push(C::ProjectReferenceProvider, false);
                push(C::ProjectMutation, false);
            }
            _ => match wasm_runtime_class {
                WasmRuntimeClass::Wasi => {
                    push(C::ConsoleIo, false);
                    push(C::UiInteraction, true);
                    push(C::EventPump, true);
                    push(C::FileSystemIo, false);
                    push(C::ProcessEnv, false);
                    push(C::ComActivationDispatch, false);
                    push(C::TimeLocale, true);
                    push(C::DynamicLinking, false);
                    push(C::DiagnosticsTelemetry, true);
                    push(C::ProjectCatalog, false);
                    push(C::ProjectReferenceProvider, false);
                    push(C::ProjectMutation, false);
                }
                WasmRuntimeClass::BrowserSandbox => {
                    push(C::ConsoleIo, false);
                    push(C::UiInteraction, false);
                    push(C::EventPump, true);
                    push(C::FileSystemIo, false);
                    push(C::ProcessEnv, false);
                    push(C::ComActivationDispatch, false);
                    push(C::TimeLocale, true);
                    push(C::DynamicLinking, false);
                    push(C::DiagnosticsTelemetry, true);
                    push(C::ProjectCatalog, false);
                    push(C::ProjectReferenceProvider, false);
                    push(C::ProjectMutation, false);
                }
            },
        },
        HalProfileId::Null => {
            push(C::ConsoleIo, false);
            push(C::UiInteraction, false);
            push(C::EventPump, false);
            push(C::FileSystemIo, false);
            push(C::ProcessEnv, false);
            push(C::ComActivationDispatch, false);
            push(C::TimeLocale, true);
            push(C::DynamicLinking, false);
            push(C::DiagnosticsTelemetry, true);
            push(C::ProjectCatalog, false);
            push(C::ProjectReferenceProvider, false);
            push(C::ProjectMutation, false);
        }
    }
    out
}
