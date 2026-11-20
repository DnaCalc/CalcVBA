//! CalcVBA hosted in-process, for Excel and any other Automation host
//! (`docs/design/excel-inprocess-loaders-2026-11-15.md`).
//!
//! A Windows x64 DLL with a C ABI over Automation types, so real VBA calls it through
//! `Declare PtrSafe` and an XLL through the same functions. Nothing here names Excel: the
//! loader supplies the host's objects as application objects (`calcvba_set_app_object`), and
//! the project's own type-library references say what those objects are.
//!
//! A project is its modules and references; `calcvba_compile` picks the backend (VM3 or the
//! JIT) and `calcvba_run` runs one public `Function` through the whole engine: a driver
//! module's `Sub Main` stores the function's value, which the run hands back. Every project
//! belongs to the thread that created it (Excel's main thread); a call from another thread
//! fails with `RPC_E_WRONG_THREAD` and touches nothing. Strings arrive as `BSTR` (`StrPtr`
//! from VBA), objects as `IDispatch*`, results as `VARIANT*`.
#![cfg(all(windows, target_arch = "x86_64"))]

use calcvba_runtime::ObjectRefExt;
pub mod xll;

use calcvba_engine::{Engine, HostConfig, SnapshotOutcome};
use calcvba_hal::model::HostPolicy;
use calcvba_runtime::Variant;
use calcvba_symbol::manifest as sym;
use std::cell::RefCell;
use std::collections::HashMap;
use std::sync::atomic::{AtomicIsize, Ordering};
use windows_sys::Win32::Foundation::SysStringLen;
use windows_sys::Win32::System::Variant::{VARIANT, VariantClear};

const S_OK: i32 = 0;
const E_FAIL: i32 = 0x8000_4005_u32 as i32;
const E_INVALIDARG: i32 = 0x8007_0057_u32 as i32;
const E_POINTER: i32 = 0x8000_4003_u32 as i32;
const E_HANDLE: i32 = 0x8007_0006_u32 as i32;
const RPC_E_WRONG_THREAD: i32 = 0x8001_010E_u32 as i32;

/// The driver module `calcvba_run` adds: its `Main` is the engine's entry.
const DRIVER: &str = "CalcVBAEntry";

/// One `IDispatch` reference the project holds.
struct Dispatch(*mut core::ffi::c_void);

impl Dispatch {
    /// Takes a new reference to `pointer`.
    ///
    /// # Safety
    /// `pointer` is a live `IDispatch`.
    unsafe fn add_ref(pointer: *mut core::ffi::c_void) -> Self {
        // SAFETY: a live interface, per the caller.
        unsafe { vtable_call(pointer, 1) };
        Self(pointer)
    }

    /// Another reference, for a consumer that takes one over.
    fn retained(&self) -> *mut core::ffi::c_void {
        // SAFETY: this value's reference keeps the interface live.
        unsafe { vtable_call(self.0, 1) };
        self.0
    }
}

impl Drop for Dispatch {
    fn drop(&mut self) {
        // SAFETY: the reference this value owns.
        unsafe { vtable_call(self.0, 2) };
    }
}

/// Calls IUnknown slot `slot` (1 AddRef, 2 Release) of `pointer`.
///
/// # Safety
/// `pointer` is a live COM interface.
unsafe fn vtable_call(pointer: *mut core::ffi::c_void, slot: usize) -> u32 {
    type Method = unsafe extern "system" fn(*mut core::ffi::c_void) -> u32;
    // SAFETY: a COM interface starts with its vtable; slots 1 and 2 are AddRef and Release.
    unsafe {
        let vtable = *(pointer as *const *const Method);
        (*vtable.add(slot))(pointer)
    }
}

/// One project.
struct Entry {
    modules: Vec<sym::ModuleUnit>,
    references: Vec<sym::ProjectReference>,
    app_objects: Vec<(String, Dispatch)>,
    /// Document modules' host objects, by module name.
    documents: Vec<(String, Dispatch)>,
    engine: Option<(Engine, bool)>,
    error: (i32, String),
}

static NEXT_HANDLE: AtomicIsize = AtomicIsize::new(1);

thread_local! {
    static ENTRIES: RefCell<HashMap<isize, Entry>> = RefCell::new(HashMap::new());
}

/// Runs `f` on project `handle`, recording a failure as its last error.
fn with_entry(handle: isize, f: impl FnOnce(&mut Entry) -> Result<(), (i32, String)>) -> i32 {
    ENTRIES.with(|entries| {
        let mut entries = entries.borrow_mut();
        let Some(entry) = entries.get_mut(&handle) else {
            // A project of another thread is not in this thread's table.
            return if handle > 0 && handle < NEXT_HANDLE.load(Ordering::Relaxed) {
                RPC_E_WRONG_THREAD
            } else {
                E_HANDLE
            };
        };
        match f(entry) {
            Ok(()) => {
                entry.error = (0, String::new());
                S_OK
            }
            Err(error) => {
                entry.error = error;
                E_FAIL
            }
        }
    })
}

/// The text of a `BSTR` argument (null is "").
///
/// # Safety
/// `text` is null or a live `BSTR`.
unsafe fn bstr_text(text: *const u16) -> String {
    if text.is_null() {
        return String::new();
    }
    // SAFETY: a live BSTR carries its length.
    let units = unsafe { std::slice::from_raw_parts(text, SysStringLen(text) as usize) };
    String::from_utf16_lossy(units)
}

/// A module's code as the VBE exports it, from `Attribute VB_Name` on: the exported header
/// (`VERSION`, `BEGIN … END`) goes; member attributes such as `VB_UserMemId` stay. Plain code
/// (no header) is unchanged.
fn module_text(source: &str) -> String {
    let first = source.lines().next().unwrap_or("").trim_start();
    if !first.to_ascii_lowercase().starts_with("version ") {
        return source.to_string();
    }
    let lines: Vec<&str> = source.lines().collect();
    match lines.iter().position(|line| {
        line.trim_start()
            .to_ascii_lowercase()
            .starts_with("attribute vb_name")
    }) {
        Some(at) => lines[at..].join("\r\n"),
        None => source.to_string(),
    }
}

fn failure(message: impl Into<String>) -> (i32, String) {
    (0, message.into())
}

/// Moves `value` into the caller's `VARIANT`, releasing what it held.
///
/// # Safety
/// `target` is a live `VARIANT`.
unsafe fn store_variant(target: *mut VARIANT, value: Variant) {
    // SAFETY: clearing releases the old contents; the cell's bytes (and what they own) then
    // move into it.
    unsafe {
        VariantClear(target);
        let value = core::mem::ManuallyDrop::new(value);
        let bytes = value.to_wire_bytes();
        core::ptr::copy_nonoverlapping(bytes.as_ptr(), target.cast::<u8>(), bytes.len());
    }
}

/// Creates an empty project; `handle` receives its handle.
///
/// # Safety
/// `handle` is a writable `isize`.
#[unsafe(no_mangle)]
pub unsafe extern "system" fn calcvba_new(handle: *mut isize) -> i32 {
    if handle.is_null() {
        return E_POINTER;
    }
    let id = NEXT_HANDLE.fetch_add(1, Ordering::Relaxed);
    let entry = Entry {
        modules: Vec::new(),
        references: Vec::new(),
        app_objects: Vec::new(),
        documents: Vec::new(),
        engine: None,
        error: (0, String::new()),
    };
    ENTRIES.with(|entries| entries.borrow_mut().insert(id, entry));
    // SAFETY: checked writable above.
    unsafe { handle.write(id) };
    S_OK
}

/// Adds a module: `kind` 1 is a standard module, 2 a class module.
///
/// # Safety
/// `name` and `source` are null or live `BSTR`s.
#[unsafe(no_mangle)]
pub unsafe extern "system" fn calcvba_add_module(
    handle: isize,
    name: *const u16,
    kind: i32,
    source: *const u16,
) -> i32 {
    // SAFETY: the caller's BSTRs.
    let (name, source) = unsafe { (bstr_text(name), module_text(&bstr_text(source))) };
    let module_kind = match kind {
        1 => sym::ModuleKind::Procedural,
        2 => sym::ModuleKind::Class,
        _ => return E_INVALIDARG,
    };
    with_entry(handle, |entry| {
        if name.eq_ignore_ascii_case(DRIVER) {
            return Err(failure(format!("`{DRIVER}` is the driver module's name")));
        }
        entry.modules.push(sym::ModuleUnit {
            attributes: sym::ModuleAttributes::named(&name),
            module_name: name,
            module_kind,
            source,
        });
        Ok(())
    })
}

/// Adds a document module `name` (`ThisWorkbook`, `Sheet1`) whose host object is `document`,
/// of type `host_type` as the project's references name it (`Excel.Workbook`). The module's
/// default instance carries the object; members of the host type reach it.
///
/// # Safety
/// `name`, `host_type` and `source` are null or live `BSTR`s; `document` is a live `IDispatch`.
#[unsafe(no_mangle)]
pub unsafe extern "system" fn calcvba_add_document(
    handle: isize,
    name: *const u16,
    host_type: *const u16,
    source: *const u16,
    document: *mut core::ffi::c_void,
) -> i32 {
    if document.is_null() {
        return E_POINTER;
    }
    // SAFETY: the caller's BSTRs.
    let (name, host_type, source) = unsafe {
        (
            bstr_text(name),
            bstr_text(host_type),
            module_text(&bstr_text(source)),
        )
    };
    with_entry(handle, |entry| {
        if name.eq_ignore_ascii_case(DRIVER) {
            return Err(failure(format!("`{DRIVER}` is the driver module's name")));
        }
        if host_type.is_empty() {
            return Err(failure(format!("document `{name}` needs its host type")));
        }
        // SAFETY: a live IDispatch; the project takes its own reference.
        let object = unsafe { Dispatch::add_ref(document) };
        if let Some((engine, _)) = &entry.engine {
            register_document(engine, &name, &object)?;
        }
        let mut attributes = sym::ModuleAttributes::named(&name);
        attributes.vb_predeclared_id = true;
        attributes.vb_exposed = true;
        attributes.host_document_type = Some(host_type);
        entry.modules.push(sym::ModuleUnit {
            attributes,
            module_name: name.clone(),
            module_kind: sym::ModuleKind::Document,
            source,
        });
        entry.documents.push((name, object));
        Ok(())
    })
}

/// Adds a type-library reference (`Excel`, `{00020813-…}`, 1, 9).
///
/// # Safety
/// `name` and `libid` are null or live `BSTR`s.
#[unsafe(no_mangle)]
pub unsafe extern "system" fn calcvba_add_reference(
    handle: isize,
    name: *const u16,
    libid: *const u16,
    major: i32,
    minor: i32,
) -> i32 {
    // SAFETY: the caller's BSTRs.
    let (name, libid) = unsafe { (bstr_text(name), bstr_text(libid)) };
    with_entry(handle, |entry| {
        entry.references.push(sym::ProjectReference::TypeLibrary {
            name,
            guid: (!libid.is_empty()).then_some(libid),
            version_major: u16::try_from(major).ok(),
            version_minor: u16::try_from(minor).ok(),
            lcid: Some(0),
            import_lib: None,
        });
        Ok(())
    })
}

/// Registers `object` as the project's instance of application object `class` (a CLSID or
/// ProgID): Excel's loaders pass `Application` for `Excel.Global`. The project keeps its own
/// reference until it closes.
///
/// # Safety
/// `class` is null or a live `BSTR`; `object` is a live `IDispatch`.
#[unsafe(no_mangle)]
pub unsafe extern "system" fn calcvba_set_app_object(
    handle: isize,
    class: *const u16,
    object: *mut core::ffi::c_void,
) -> i32 {
    if object.is_null() {
        return E_POINTER;
    }
    // SAFETY: the caller's BSTR.
    let class = unsafe { bstr_text(class) };
    with_entry(handle, |entry| {
        // SAFETY: a live IDispatch; the project takes its own reference.
        let object = unsafe { Dispatch::add_ref(object) };
        if let Some((engine, _)) = &entry.engine {
            register(engine, &class, &object)?;
        }
        entry.app_objects.push((class, object));
        Ok(())
    })
}

fn register_document(
    engine: &Engine,
    module: &str,
    object: &Dispatch,
) -> Result<(), (i32, String)> {
    // SAFETY: a retained reference, which the registry takes over.
    unsafe {
        engine
            .host_services()
            .com()
            .register_document_object(module, object.retained())
    }
    .map_err(|error| failure(error.to_string()))
}

fn register(engine: &Engine, class: &str, object: &Dispatch) -> Result<(), (i32, String)> {
    // SAFETY: a retained reference, which the registry takes over.
    unsafe {
        engine
            .host_services()
            .com()
            .register_app_object(class, object.retained())
    }
    .map_err(|error| failure(error.to_string()))
}

/// Chooses the backend (0 VM3, 1 the JIT) for the project's runs.
#[unsafe(no_mangle)]
pub extern "system" fn calcvba_compile(handle: isize, backend: i32) -> i32 {
    let (config, jit) = match backend {
        0 => (HostConfig::vm3(), false),
        1 => (HostConfig::jit(), true),
        _ => return E_INVALIDARG,
    };
    with_entry(handle, |entry| {
        let mut engine = Engine::new(config);
        engine.set_host_policy(HostPolicy::native());
        for (class, object) in &entry.app_objects {
            register(&engine, class, object)?;
        }
        for (module, object) in &entry.documents {
            register_document(&engine, module, object)?;
        }
        entry.engine = Some((engine, jit));
        Ok(())
    })
}

/// Runs the public `Function` `module.entry` (no arguments) with the project's modules, and
/// stores its value in `result` (cleared first).
///
/// # Safety
/// `module` and `entry` are null or live `BSTR`s; `result` is a live `VARIANT`.
#[unsafe(no_mangle)]
pub unsafe extern "system" fn calcvba_run(
    handle: isize,
    module: *const u16,
    entry_name: *const u16,
    result: *mut VARIANT,
) -> i32 {
    if result.is_null() {
        return E_POINTER;
    }
    // SAFETY: the caller's BSTRs.
    let (module, entry_name) = unsafe { (bstr_text(module), bstr_text(entry_name)) };
    let mut value = None;
    let status = with_entry(handle, |entry| {
        value = Some(run(entry, &module, &entry_name)?);
        Ok(())
    });
    if status == S_OK
        && let Some(value) = value
    {
        // SAFETY: `result` is a live VARIANT.
        unsafe { store_variant(result, value) };
    }
    status
}

/// Runs `module.entry_name()` with the project's modules, on its compiled backend.
fn run(entry: &Entry, module: &str, entry_name: &str) -> Result<Variant, (i32, String)> {
    let (engine, jit) = entry
        .engine
        .as_ref()
        .ok_or_else(|| failure("the project is not compiled"))?;
    let driver = format!(
        "Public CalcVBAResult As Variant\r\nPublic Sub Main()\r\n    CalcVBAResult = {module}.{entry_name}()\r\nEnd Sub\r\n"
    );
    let mut modules = vec![sym::ModuleUnit {
        module_name: DRIVER.into(),
        module_kind: sym::ModuleKind::Procedural,
        attributes: sym::ModuleAttributes::named(DRIVER),
        source: driver,
    }];
    modules.extend(entry.modules.iter().cloned());
    let manifest = sym::SymbolProjectManifest {
        project_name: "CalcVBAProject".into(),
        project_kind: sym::ProjectKind::Source,
        modules,
        references: entry.references.clone(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let outcome = if *jit {
        engine.execute_manifest_snapshot_with_err_jit(&manifest)
    } else {
        engine.execute_manifest_snapshot_with_err_vm3(&manifest)
    };
    let value = match outcome {
        // The driver's global is the first module-level global.
        SnapshotOutcome::Completed { mut values, .. } if !values.is_empty() => {
            values.swap_remove(0)
        }
        SnapshotOutcome::Completed { .. } => return Err(failure("the run returned no snapshot")),
        SnapshotOutcome::Raised { err } => return Err((err.number, err.description)),
        SnapshotOutcome::Unsupported(what) => {
            return Err(failure(format!("unsupported: {what}")));
        }
        SnapshotOutcome::Failed(message) => return Err(failure(message)),
    };
    if value
        .as_object_ref()
        .is_some_and(|object| object.is_compat_object())
    {
        return Err(failure(
            "a project object cannot be returned to the host yet",
        ));
    }
    Ok(value)
}

/// The last failure of project `handle`: its VBA error number (0 when it was not a VBA
/// run-time error), and its description as a String in `description` (cleared first).
///
/// # Safety
/// `number` is null or a writable `i32`; `description` is null or a live `VARIANT`.
#[unsafe(no_mangle)]
pub unsafe extern "system" fn calcvba_last_error(
    handle: isize,
    number: *mut i32,
    description: *mut VARIANT,
) -> i32 {
    ENTRIES.with(|entries| {
        let entries = entries.borrow();
        let Some(entry) = entries.get(&handle) else {
            return E_HANDLE;
        };
        // SAFETY: the caller's writable slots; the description cell moves into the VARIANT.
        unsafe {
            if !number.is_null() {
                number.write(entry.error.0);
            }
            if !description.is_null() {
                store_variant(description, Variant::from_string(entry.error.1.as_str()));
            }
        }
        S_OK
    })
}

/// Closes project `handle`, releasing its engine and objects (application objects included).
#[unsafe(no_mangle)]
pub extern "system" fn calcvba_close(handle: isize) -> i32 {
    ENTRIES.with(|entries| match entries.borrow_mut().remove(&handle) {
        Some(entry) => {
            drop(entry);
            S_OK
        }
        None => E_HANDLE,
    })
}
