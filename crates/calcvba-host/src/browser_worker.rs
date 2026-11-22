use crate::{Visibility, portability, sym};
use calcvba_hal::{
    adapters::builder::HostBuilder,
    callbacks::HostCallbacks,
    model::{HalProfileId, HalRuntimeClass, HostPolicy, UiVirtualizationMode},
};
use calcvba_runtime::{print_display_text, regional};
use std::sync::Arc;

#[link(wasm_import_module = "calcvba_browser")]
unsafe extern "C" {
    fn report(kind: u32, text: *const u8, len: usize);
    fn message(prompt: *const u8, len: usize, style: i32) -> i32;
    fn input(
        prompt: *const u8,
        len: usize,
        default: *const u8,
        default_len: usize,
        output: *mut u8,
        capacity: usize,
    ) -> usize;
    fn local_time_millis() -> f64;
}

fn emit(kind: u32, text: &str) {
    // SAFETY: the import copies the string synchronously from the live Wasm memory.
    unsafe { report(kind, text.as_ptr(), text.len()) }
}

struct BrowserCallbacks;

impl HostCallbacks for BrowserCallbacks {
    fn on_msg_box(&self, prompt: &str, style: i32) -> i32 {
        // SAFETY: prompt is live for the synchronous import, which copies before waiting.
        unsafe { message(prompt.as_ptr(), prompt.len(), style) }
    }

    fn on_input_box(&self, prompt: &str, default: &str) -> String {
        let mut buffer = [0u8; 1024];
        // SAFETY: both input strings and the output buffer are live throughout the
        // import. The JS adapter copies the inputs and writes at most capacity bytes.
        let n = unsafe {
            input(
                prompt.as_ptr(),
                prompt.len(),
                default.as_ptr(),
                default.len(),
                buffer.as_mut_ptr(),
                buffer.len(),
            )
        };
        String::from_utf8_lossy(&buffer[..n.min(buffer.len())]).into_owned()
    }

    fn on_status_bar(&self, _text: &str) {}

    fn on_debug_print(&self, text: &str) {
        emit(0, text);
    }

    fn on_local_time_serial(&self) -> Option<f64> {
        // SAFETY: the import takes no pointers and returns the browser's local clock in milliseconds.
        Some(25569.0 + unsafe { local_time_millis() } / 86_400_000.0)
    }
}

#[unsafe(no_mangle)]
pub extern "C" fn calcvba_source_alloc(len: usize) -> *mut u8 {
    Box::into_raw(vec![0; len].into_boxed_slice()).cast()
}

/// Consumes a source buffer allocated by `calcvba_source_alloc` in this Wasm instance.
///
/// # Safety
/// The pointer and length must identify a live allocation from that export. This
/// call takes it once, after the worker has written the source as UTF-8.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn calcvba_run_source(ptr: *mut u8, len: usize, dialogs: bool, en_za: bool) {
    // SAFETY: the JS worker hands back the matching, initialized boxed slice exactly once.
    let bytes = unsafe { Box::from_raw(std::ptr::slice_from_raw_parts_mut(ptr, len)) };
    let result = std::str::from_utf8(&bytes)
        .map_err(|e| e.to_string())
        .and_then(|source| {
            let settings = if en_za {
                regional::RegionalSettings::en_za()
            } else {
                regional::RegionalSettings::en_us()
            };
            regional::with_settings(settings, || execute(source, dialogs))
        });
    match result {
        Ok(text) => emit(1, &text),
        Err(error) => emit(2, &error),
    }
}

fn execute(source: &str, dialogs: bool) -> Result<String, String> {
    let manifest = sym::SymbolProjectManifest {
        project_name: "VBAProject".into(),
        project_kind: sym::ProjectKind::Source,
        modules: vec![sym::ModuleUnit {
            module_name: "Entry".into(),
            module_kind: sym::ModuleKind::Procedural,
            attributes: sym::ModuleAttributes::named("Entry"),
            source: source.into(),
        }],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    portability::Mode::WindowsFaithful
        .check_with_browser_dialogs(
            std::slice::from_ref(&manifest),
            portability::Host::Browser,
            dialogs,
        )
        .map_err(|error| error.render_human())?;
    let mut policy = HostPolicy::native();
    policy.allow_interaction = dialogs;
    policy.ui_virtualization = UiVirtualizationMode::HostCallback;
    let host = HostBuilder::new()
        .profile(HalProfileId::Wasm)
        .runtime_class(HalRuntimeClass::WasmBrowserSandbox)
        .policy(policy)
        .callbacks(Arc::new(BrowserCallbacks))
        .build();
    let (core, procedures) = calcvba_bind::bind_program_with_procedures(
        &manifest,
        &calcvba_symbol::CatalogTypeLibResolver,
    )
    .map_err(|error| error.to_string())?;
    let program = calcvba_oxir::elaborate::elaborate(&core).map_err(|error| error.to_string())?;
    let entry = procedures
        .iter()
        .find(|proc| {
            proc.name.eq_ignore_ascii_case("Main")
                && proc.visibility == Visibility::Public
                && program.funcs[proc.index].param_count == 0
        })
        .ok_or_else(|| "A parameterless public Main is required".to_string())?;
    let mut vm = calcvba_vm3::Vm3::activate(&program, &*host).map_err(|error| error.to_string())?;
    let result = vm
        .invoke_function_values(calcvba_oxir::FuncId(entry.index), Vec::new())
        .map_err(|error| error.to_string())?;
    Ok(print_display_text(&result))
}
