//! Prebuilt CalcVBA WrappedComServer COM host.
//!
//! Reusable in-process COM DLL host over `ProjectRuntimeSession`: class factory activation,
//! per-user registration, and generated type-library registration. Project class instances carry
//! their own `ClassDualVtbl` and table-driven `IDispatch` directly on `CompatObjectBase`
//! (dc-yqq0.23 / dc-yqq0.26), so `IClassFactory::CreateInstance` returns the project object itself
//! with no wrapper allocation or wrapper map.

#![cfg(target_os = "windows")]
#![allow(non_snake_case)]
#![allow(unsafe_op_in_unsafe_fn)]

use std::cell::RefCell;
use std::env;
use std::ffi::c_void;
use std::fs;
use std::path::PathBuf;
use std::ptr;
use std::rc::Rc;
use std::sync::OnceLock;
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering, fence};

use calcvba_build::{ComClassDescriptor, ComServerDescriptor};
use calcvba_engine::{Engine, ExecBackend, HostConfig, ProjectRuntimeSession, RuntimeProfileId};
use calcvba_runtime::events::{MemberFault, MemberInvoker};
use calcvba_runtime::{
    ObjectRef, RuntimeMemberDescriptor, RuntimeMemberInvokeKind, Utf16String, Variant,
};
use windows_sys::Win32::Foundation::{ERROR_SUCCESS, HMODULE};
use windows_sys::Win32::System::Com::{SYS_WIN64, SetErrorInfo};
use windows_sys::Win32::System::LibraryLoader::GetModuleFileNameW;
use windows_sys::Win32::System::Ole::{
    CreateErrorInfo, LoadTypeLibEx, REGKIND_NONE, RegisterTypeLibForUser, UnRegisterTypeLibForUser,
};
use windows_sys::Win32::System::Registry::{
    HKEY, HKEY_CURRENT_USER, KEY_WRITE, REG_DWORD, REG_SZ, RegCloseKey, RegCreateKeyExW,
    RegDeleteKeyW, RegDeleteTreeW, RegSetValueExW,
};
use windows_sys::Win32::System::SystemServices::DLL_PROCESS_ATTACH;
use windows_sys::core::GUID;

mod error_info;
use error_info::with_project_session;

struct HostArtifacts {
    /// The runtime artifact (`OxImage`) loaded into the session — the `.oxi` sidecar next to
    /// the COM host DLL.
    oxi_path: PathBuf,
    descriptor_path: PathBuf,
    tlb_path: PathBuf,
}

static ARTIFACTS: OnceLock<Result<HostArtifacts, String>> = OnceLock::new();

fn artifacts() -> Result<&'static HostArtifacts, String> {
    ARTIFACTS
        .get_or_init(resolve_host_artifacts)
        .as_ref()
        .map_err(|err| err.clone())
}

fn resolve_host_artifacts() -> Result<HostArtifacts, String> {
    // SAFETY: module_path reads the current DLL module path into owned Rust storage.
    let module_path_text = unsafe {
        module_path().map_err(|hr| format!("GetModuleFileNameW failed: 0x{:08X}", hr as u32))?
    };
    let module_path = PathBuf::from(&module_path_text);
    let parent = module_path
        .parent()
        .ok_or_else(|| format!("COM host module path has no parent: {module_path_text}"))?
        .to_path_buf();
    let stem = module_path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .ok_or_else(|| format!("COM host module path has no file stem: {module_path_text}"))?
        .to_string();
    Ok(HostArtifacts {
        oxi_path: parent.join(format!("{stem}.oxi")),
        descriptor_path: parent.join(format!("{stem}.comserver.json")),
        tlb_path: parent.join(format!("{stem}.tlb")),
    })
}

fn artifact_descriptor_json() -> Result<String, String> {
    let path = artifacts()?.descriptor_path.clone();
    fs::read_to_string(&path)
        .map_err(|err| format!("failed to read COM descriptor `{}`: {err}", path.display()))
}

fn artifact_image_bytes() -> Result<Vec<u8>, String> {
    let path = artifacts()?.oxi_path.clone();
    fs::read(&path).map_err(|err| format!("failed to read OxVBA image `{}`: {err}", path.display()))
}

fn artifact_tlb_path_string() -> Result<String, String> {
    Ok(artifacts()?.tlb_path.display().to_string())
}

const S_OK: i32 = 0;
const S_FALSE: i32 = 1;
const E_NOINTERFACE: i32 = 0x8000_4002u32 as i32;
const E_POINTER: i32 = 0x8000_4003u32 as i32;
const E_FAIL: i32 = 0x8000_4005u32 as i32;
const CLASS_E_NOAGGREGATION: i32 = 0x8004_0110u32 as i32;
const CLASS_E_CLASSNOTAVAILABLE: i32 = 0x8004_0111u32 as i32;
const SELFREG_E_CLASS: i32 = 0x8004_0201u32 as i32;

const IID_IUNKNOWN: GUID = GUID {
    data1: 0x0000_0000,
    data2: 0x0000,
    data3: 0x0000,
    data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};
const IID_ICLASSFACTORY: GUID = GUID {
    data1: 0x0000_0001,
    data2: 0x0000,
    data3: 0x0000,
    data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};
const IID_IDISPATCH: GUID = GUID {
    data1: 0x0002_0400,
    data2: 0x0000,
    data3: 0x0000,
    data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};
const IID_IERRORINFO: GUID = GUID {
    data1: 0x1CF2_B120,
    data2: 0x547D,
    data3: 0x101B,
    data4: [0x8E, 0x65, 0x08, 0x00, 0x2B, 0x2B, 0xD1, 0x19],
};

static GLOBAL_REF_COUNT: AtomicU32 = AtomicU32::new(0);
static SESSION_ACTIVE: AtomicBool = AtomicBool::new(false);
static DESCRIPTOR: OnceLock<Result<ComServerDescriptor, String>> = OnceLock::new();
static mut MODULE_HANDLE: HMODULE = ptr::null_mut();

enum SessionSlot {
    Uninitialized,
    Running,
    Ready(Box<ProjectRuntimeSession>),
}

/// Between-runs [`MemberInvoker`] installed on every project instance created by this COM host's
/// session so incoming `IDispatch::Invoke` calls directly on `CompatObjectBase` enter the owning
/// `ProjectRuntimeSession`. Re-entrant calls during an active VBA run hit the active VM3/JIT
/// `ExecutorScope` before consulting this invoker.
struct ComhostSessionInvoker;

impl MemberInvoker for ComhostSessionInvoker {
    fn invoke_member(
        &mut self,
        object: ObjectRef,
        member: &'static RuntimeMemberDescriptor,
        invoke_kind: RuntimeMemberInvokeKind,
        args: &mut [Variant],
        aliases: &[usize],
    ) -> Result<Variant, MemberFault> {
        let mut outcome: Option<Result<Variant, MemberFault>> = None;
        match with_session(|session| {
            outcome = Some(MemberInvoker::invoke_member(
                session,
                object,
                member,
                invoke_kind,
                args,
                aliases,
            ));
            Ok(())
        }) {
            Ok(()) => outcome.unwrap_or_else(|| Err(MemberFault::from_number(5))),
            Err(message) => Err(MemberFault {
                number: 5,
                description: message.into(),
                source: descriptor()
                    .map(|d| d.project_name.as_str())
                    .unwrap_or("CalcVBA")
                    .into(),
                help_file: Utf16String::new(),
                help_context: 0,
            }),
        }
    }
}

thread_local! {
    static SESSION: RefCell<SessionSlot> = const { RefCell::new(SessionSlot::Uninitialized) };
    static SESSION_INVOKER: Rc<RefCell<dyn MemberInvoker>> =
        Rc::new(RefCell::new(ComhostSessionInvoker));
}

#[repr(C)]
struct IClassFactoryVtbl {
    query_interface: unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut c_void) -> u32,
    release: unsafe extern "system" fn(*mut c_void) -> u32,
    create_instance:
        unsafe extern "system" fn(*mut c_void, *mut c_void, *const GUID, *mut *mut c_void) -> i32,
    lock_server: unsafe extern "system" fn(*mut c_void, i32) -> i32,
}

#[repr(C)]
struct IUnknownVtbl {
    query_interface: unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut c_void) -> u32,
    release: unsafe extern "system" fn(*mut c_void) -> u32,
}

#[repr(C)]
struct CreateErrorInfoVtbl {
    query_interface: unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut c_void) -> u32,
    release: unsafe extern "system" fn(*mut c_void) -> u32,
    set_guid: unsafe extern "system" fn(*mut c_void, *const GUID) -> i32,
    set_source: unsafe extern "system" fn(*mut c_void, *const u16) -> i32,
    set_description: unsafe extern "system" fn(*mut c_void, *const u16) -> i32,
    set_help_file: unsafe extern "system" fn(*mut c_void, *const u16) -> i32,
    set_help_context: unsafe extern "system" fn(*mut c_void, u32) -> i32,
}

#[repr(C)]
struct ClassFactory {
    vtbl: *const IClassFactoryVtbl,
    ref_count: AtomicU32,
    class_index: usize,
}

static CLASS_FACTORY_VTBL: IClassFactoryVtbl = IClassFactoryVtbl {
    query_interface: factory_query_interface,
    add_ref: factory_add_ref,
    release: factory_release,
    create_instance: factory_create_instance,
    lock_server: factory_lock_server,
};

#[unsafe(no_mangle)]
unsafe extern "system" fn DllMain(module: HMODULE, reason: u32, _reserved: *mut c_void) -> i32 {
    if reason == DLL_PROCESS_ATTACH {
        MODULE_HANDLE = module;
    }
    1
}

#[unsafe(no_mangle)]
unsafe extern "system" fn DllGetClassObject(
    rclsid: *const GUID,
    riid: *const GUID,
    out: *mut *mut c_void,
) -> i32 {
    if rclsid.is_null() || riid.is_null() || out.is_null() {
        return E_POINTER;
    }
    *out = ptr::null_mut();

    let Ok(descriptor) = descriptor() else {
        return E_FAIL;
    };
    let Some((class_index, _class)) = descriptor
        .classes
        .iter()
        .enumerate()
        .find(|(_, class)| class.creatable && guid_matches_text(&*rclsid, &class.clsid))
    else {
        return CLASS_E_CLASSNOTAVAILABLE;
    };

    let factory = allocate_factory(class_index);
    let hr = factory_query_interface(factory, riid, out);
    factory_release(factory);
    hr
}

#[unsafe(no_mangle)]
unsafe extern "system" fn DllCanUnloadNow() -> i32 {
    if !SESSION_ACTIVE.load(Ordering::Acquire) && GLOBAL_REF_COUNT.load(Ordering::Acquire) == 0 {
        S_OK
    } else {
        S_FALSE
    }
}

#[unsafe(no_mangle)]
unsafe extern "system" fn DllRegisterServer() -> i32 {
    match register_server() {
        Ok(()) => S_OK,
        Err(hr) => hr,
    }
}

#[unsafe(no_mangle)]
unsafe extern "system" fn DllUnregisterServer() -> i32 {
    match unregister_server() {
        Ok(()) => S_OK,
        Err(hr) => hr,
    }
}

unsafe extern "system" fn factory_query_interface(
    this: *mut c_void,
    riid: *const GUID,
    out: *mut *mut c_void,
) -> i32 {
    if this.is_null() || riid.is_null() || out.is_null() {
        return E_POINTER;
    }
    *out = ptr::null_mut();
    if guid_eq(&*riid, &IID_IUNKNOWN) || guid_eq(&*riid, &IID_ICLASSFACTORY) {
        factory_add_ref(this);
        *out = this;
        S_OK
    } else {
        E_NOINTERFACE
    }
}

unsafe extern "system" fn factory_add_ref(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    let factory = this.cast::<ClassFactory>();
    (*factory).ref_count.fetch_add(1, Ordering::Relaxed) + 1
}

unsafe extern "system" fn factory_release(this: *mut c_void) -> u32 {
    if this.is_null() {
        return 0;
    }
    let factory = this.cast::<ClassFactory>();
    let previous = (*factory).ref_count.fetch_sub(1, Ordering::Release);
    if previous == 0 {
        return 0;
    }
    let remaining = previous - 1;
    if remaining == 0 {
        fence(Ordering::Acquire);
        GLOBAL_REF_COUNT.fetch_sub(1, Ordering::AcqRel);
        drop(Box::from_raw(factory));
    }
    remaining
}

unsafe extern "system" fn factory_create_instance(
    this: *mut c_void,
    outer: *mut c_void,
    riid: *const GUID,
    out: *mut *mut c_void,
) -> i32 {
    if this.is_null() || riid.is_null() || out.is_null() {
        return E_POINTER;
    }
    *out = ptr::null_mut();
    if !outer.is_null() {
        return CLASS_E_NOAGGREGATION;
    }

    let factory = &*this.cast::<ClassFactory>();
    let Ok(descriptor) = descriptor() else {
        return E_FAIL;
    };
    let Some(class) = descriptor.classes.get(factory.class_index) else {
        return E_FAIL;
    };

    let runtime_object =
        match with_project_session(|session| session.create_class_instance(&class.class_name)) {
            Ok(object) => object,
            Err(hr) => return hr,
        };

    // Return the project object itself (one COM object, no wrapper): its `ClassDualVtbl` answers
    // `IUnknown`, `IDispatch`, and the class's default dual interface IID directly on `&primary`.
    let raw = runtime_object.raw_iunknown();
    let hr = query_interface(raw.cast(), &*riid, out);
    drop(runtime_object);
    hr
}

unsafe extern "system" fn factory_lock_server(_this: *mut c_void, lock: i32) -> i32 {
    if lock != 0 {
        GLOBAL_REF_COUNT.fetch_add(1, Ordering::AcqRel);
    } else {
        GLOBAL_REF_COUNT.fetch_sub(1, Ordering::AcqRel);
    }
    S_OK
}

unsafe fn allocate_factory(class_index: usize) -> *mut c_void {
    GLOBAL_REF_COUNT.fetch_add(1, Ordering::AcqRel);
    Box::into_raw(Box::new(ClassFactory {
        vtbl: &CLASS_FACTORY_VTBL,
        ref_count: AtomicU32::new(1),
        class_index,
    }))
    .cast()
}

fn descriptor() -> Result<&'static ComServerDescriptor, i32> {
    DESCRIPTOR
        .get_or_init(|| {
            artifact_descriptor_json()
                .and_then(|text| serde_json::from_str(&text).map_err(|err| err.to_string()))
        })
        .as_ref()
        .map_err(|_| E_FAIL)
}

fn comhost_host_config() -> Result<HostConfig, String> {
    // The project's RuntimeFlavor (recorded in the descriptor) picks the executor;
    // OXVBA_BACKEND overrides it.
    let backend = match env::var("OXVBA_BACKEND") {
        Ok(value) => ExecBackend::parse(&value)
            .ok_or_else(|| format!("invalid OXVBA_BACKEND value {value:?}; expected vm3 or jit"))?,
        Err(env::VarError::NotPresent) if descriptor().is_ok_and(|d| d.jit) => ExecBackend::Jit,
        Err(env::VarError::NotPresent) => ExecBackend::Vm3,
        Err(err) => return Err(format!("could not read OXVBA_BACKEND: {err}")),
    };
    Ok(HostConfig { backend })
}

fn with_session<R>(
    f: impl FnOnce(&mut ProjectRuntimeSession) -> Result<R, String>,
) -> Result<R, String> {
    SESSION.with(|slot| {
        let previous = std::mem::replace(&mut *slot.borrow_mut(), SessionSlot::Running);
        let mut session = match previous {
            SessionSlot::Ready(session) => session,
            SessionSlot::Running => {
                return Err("COM re-entry requires a suspended VBA session".into());
            }
            SessionSlot::Uninitialized => {
                let initialized = (|| {
                    let image_bytes = artifact_image_bytes()?;
                    let mut engine = Engine::new(comhost_host_config()?)
                        .with_runtime_profile(RuntimeProfileId::WindowsHeadless);
                    engine.enable_host_native_runtime();
                    let mut session = engine
                        .prepare_image_session_bytes(&image_bytes)
                        .map_err(|err| err.to_string())?;
                    SESSION_INVOKER.with(|invoker| {
                        session.set_member_invoker(Rc::downgrade(invoker));
                    });
                    SESSION_ACTIVE.store(true, Ordering::Release);
                    Ok(Box::new(session))
                })();
                match initialized {
                    Ok(session) => session,
                    Err(error) => {
                        *slot.borrow_mut() = SessionSlot::Uninitialized;
                        return Err(error);
                    }
                }
            }
        };
        let result = f(&mut session);
        *slot.borrow_mut() = SessionSlot::Ready(session);
        result
    })
}

fn class_implements_interface_named(class: &ComClassDescriptor, name: &str) -> bool {
    class
        .implemented_interfaces
        .iter()
        .any(|interface| interface.name.eq_ignore_ascii_case(name))
}

unsafe fn register_server() -> Result<(), i32> {
    let module_path = module_path()?;
    let descriptor = descriptor()?;
    register_typelib()?;
    cleanup_imported_interface_registration(descriptor);
    for class in descriptor.classes.iter().filter(|class| class.creatable) {
        let clsid_key = format!("Software\\Classes\\CLSID\\{{{}}}", class.clsid);
        set_key_default(
            &clsid_key,
            class.description.as_deref().unwrap_or(&class.class_name),
        )?;
        set_key_default(&format!("{clsid_key}\\InprocServer32"), &module_path)?;
        set_key_value(
            &format!("{clsid_key}\\InprocServer32"),
            "ThreadingModel",
            "Apartment",
        )?;
        set_key_default(&format!("{clsid_key}\\ProgID"), &class.prog_id)?;
        set_key_default(
            &format!("{clsid_key}\\TypeLib"),
            &format!("{{{}}}", descriptor.libid),
        )?;
        set_key_default(
            &format!("{clsid_key}\\Version"),
            &format!("{}.{}", descriptor.version_major, descriptor.version_minor),
        )?;
        set_key_default(&format!("{clsid_key}\\Programmable"), "")?;
        set_key_default(
            &format!("Software\\Classes\\{}", class.prog_id),
            &class.class_name,
        )?;
        set_key_default(
            &format!("Software\\Classes\\{}\\CLSID", class.prog_id),
            &format!("{{{}}}", class.clsid),
        )?;
        if class_implements_interface_named(class, "IDTExtensibility2") {
            let addin_key = format!(
                "Software\\Microsoft\\Office\\Excel\\Addins\\{}",
                class.prog_id
            );
            let friendly_name = class.description.as_deref().unwrap_or(&class.class_name);
            set_key_default(&addin_key, friendly_name)?;
            set_key_value(&addin_key, "FriendlyName", friendly_name)?;
            set_key_value(&addin_key, "Description", friendly_name)?;
            set_key_dword(&addin_key, "LoadBehavior", 3)?;
        }
    }
    Ok(())
}

unsafe fn unregister_server() -> Result<(), i32> {
    let descriptor = descriptor()?;
    for class in descriptor.classes.iter().filter(|class| class.creatable) {
        let clsid_key = format!("Software\\Classes\\CLSID\\{{{}}}", class.clsid);
        delete_tree(&clsid_key);
        delete_tree(&format!("Software\\Classes\\{}", class.prog_id));
        if class_implements_interface_named(class, "IDTExtensibility2") {
            delete_tree(&format!(
                "Software\\Microsoft\\Office\\Excel\\Addins\\{}",
                class.prog_id
            ));
        }
    }
    unregister_typelib(descriptor);
    Ok(())
}

unsafe fn cleanup_imported_interface_registration(descriptor: &ComServerDescriptor) {
    for interface in descriptor
        .classes
        .iter()
        .flat_map(|class| class.implemented_interfaces.iter())
    {
        // Implemented imported interfaces (for example Excel.IRtdServer) are
        // owned by their source type library. RegisterTypeLibForUser registers
        // every interface present in the generated typelib, so remove the
        // user-scope Interface key for these foreign IIDs after registration;
        // Excel and other hosts then resolve the canonical interface metadata.
        delete_tree(&format!(
            "Software\\Classes\\Interface\\{{{}}}",
            interface.iid
        ));
    }
}

unsafe fn register_typelib() -> Result<(), i32> {
    let path_w = match artifact_tlb_path_string() {
        Ok(path) => wide_null(&path),
        Err(_) => return Err(SELFREG_E_CLASS),
    };
    let mut typelib: *mut c_void = ptr::null_mut();
    let hr = LoadTypeLibEx(path_w.as_ptr(), REGKIND_NONE, &mut typelib);
    if hr < 0 || typelib.is_null() {
        return Err(SELFREG_E_CLASS);
    }
    let hr = RegisterTypeLibForUser(typelib, path_w.as_ptr(), ptr::null());
    release_unknown(typelib);
    if hr < 0 { Err(SELFREG_E_CLASS) } else { Ok(()) }
}

unsafe fn unregister_typelib(descriptor: &ComServerDescriptor) {
    let Some(libid) = parse_guid_text(&descriptor.libid) else {
        return;
    };
    UnRegisterTypeLibForUser(
        &libid,
        descriptor.version_major,
        descriptor.version_minor,
        0,
        SYS_WIN64,
    );
}

unsafe fn set_key_default(path: &str, value: &str) -> Result<(), i32> {
    set_key_value_raw(path, None, value)
}

unsafe fn set_key_value(path: &str, name: &str, value: &str) -> Result<(), i32> {
    set_key_value_raw(path, Some(name), value)
}

unsafe fn set_key_dword(path: &str, name: &str, value: u32) -> Result<(), i32> {
    let path_w = wide_null(path);
    let mut key: HKEY = ptr::null_mut();
    let status = RegCreateKeyExW(
        HKEY_CURRENT_USER,
        path_w.as_ptr(),
        0,
        ptr::null(),
        0,
        KEY_WRITE,
        ptr::null(),
        &mut key,
        ptr::null_mut(),
    );
    if status != ERROR_SUCCESS {
        return Err(SELFREG_E_CLASS);
    }

    let name_w = wide_null(name);
    let bytes = value.to_le_bytes();
    let status = RegSetValueExW(
        key,
        name_w.as_ptr(),
        0,
        REG_DWORD,
        bytes.as_ptr(),
        bytes.len() as u32,
    );
    RegCloseKey(key);
    if status == ERROR_SUCCESS {
        Ok(())
    } else {
        Err(SELFREG_E_CLASS)
    }
}

unsafe fn set_key_value_raw(path: &str, name: Option<&str>, value: &str) -> Result<(), i32> {
    let path_w = wide_null(path);
    let mut key: HKEY = ptr::null_mut();
    let status = RegCreateKeyExW(
        HKEY_CURRENT_USER,
        path_w.as_ptr(),
        0,
        ptr::null(),
        0,
        KEY_WRITE,
        ptr::null(),
        &mut key,
        ptr::null_mut(),
    );
    if status != ERROR_SUCCESS {
        return Err(SELFREG_E_CLASS);
    }

    let name_w = name.map(wide_null);
    let value_w = wide_null(value);
    let status = RegSetValueExW(
        key,
        name_w.as_ref().map_or(ptr::null(), |text| text.as_ptr()),
        0,
        REG_SZ,
        value_w.as_ptr().cast::<u8>(),
        (value_w.len() * std::mem::size_of::<u16>()) as u32,
    );
    RegCloseKey(key);
    if status == ERROR_SUCCESS {
        Ok(())
    } else {
        Err(SELFREG_E_CLASS)
    }
}

unsafe fn delete_tree(path: &str) {
    let path_w = wide_null(path);
    RegDeleteTreeW(HKEY_CURRENT_USER, path_w.as_ptr());
    RegDeleteKeyW(HKEY_CURRENT_USER, path_w.as_ptr());
}

unsafe fn module_path() -> Result<String, i32> {
    let mut buffer = vec![0u16; 32768];
    let len = GetModuleFileNameW(MODULE_HANDLE, buffer.as_mut_ptr(), buffer.len() as u32);
    if len == 0 {
        return Err(SELFREG_E_CLASS);
    }
    Ok(String::from_utf16_lossy(&buffer[..len as usize]))
}

fn guid_eq(left: &GUID, right: &GUID) -> bool {
    left.data1 == right.data1
        && left.data2 == right.data2
        && left.data3 == right.data3
        && left.data4 == right.data4
}

fn guid_matches_text(guid: &GUID, text: &str) -> bool {
    guid_to_string(guid).eq_ignore_ascii_case(text.trim_matches(|ch| ch == '{' || ch == '}'))
}

fn parse_guid_text(text: &str) -> Option<GUID> {
    let text = text.trim_matches(|ch| ch == '{' || ch == '}');
    let parts: Vec<&str> = text.split('-').collect();
    if parts.len() != 5 || parts[3].len() != 4 || parts[4].len() != 12 {
        return None;
    }
    let data1 = u32::from_str_radix(parts[0], 16).ok()?;
    let data2 = u16::from_str_radix(parts[1], 16).ok()?;
    let data3 = u16::from_str_radix(parts[2], 16).ok()?;
    let tail = format!("{}{}", parts[3], parts[4]);
    if tail.len() != 16 {
        return None;
    }
    let mut data4 = [0u8; 8];
    for index in 0..8 {
        data4[index] = u8::from_str_radix(&tail[index * 2..index * 2 + 2], 16).ok()?;
    }
    Some(GUID {
        data1,
        data2,
        data3,
        data4,
    })
}

fn guid_to_string(guid: &GUID) -> String {
    format!(
        "{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}",
        guid.data1,
        guid.data2,
        guid.data3,
        guid.data4[0],
        guid.data4[1],
        guid.data4[2],
        guid.data4[3],
        guid.data4[4],
        guid.data4[5],
        guid.data4[6],
        guid.data4[7],
    )
}

unsafe fn release_unknown(ptr: *mut c_void) {
    if ptr.is_null() {
        return;
    }
    let vtbl = *(ptr.cast::<*const IUnknownVtbl>());
    if !vtbl.is_null() {
        ((*vtbl).release)(ptr);
    }
}

unsafe fn query_interface(ptr: *mut c_void, iid: &GUID, out: *mut *mut c_void) -> i32 {
    if ptr.is_null() || out.is_null() {
        return E_POINTER;
    }
    let vtbl = *(ptr.cast::<*const IUnknownVtbl>());
    if vtbl.is_null() {
        return E_POINTER;
    }
    ((*vtbl).query_interface)(ptr, iid, out)
}

fn wide_null(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}
