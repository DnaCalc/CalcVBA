//! STA class factories and Automation over compiled native vtable entries.
use calcvba_rt_abi::native_export::{E_WRONG_THREAD, NativeComServer};
use std::{
    ffi::c_void,
    ptr,
    sync::atomic::{AtomicU32, Ordering},
};
use windows_sys::{
    Win32::{
        System::{
            Com::*,
            LibraryLoader::{GetModuleHandleW, GetProcAddress},
            Ole::{DispGetIDsOfNames, DispInvoke, LoadTypeLibEx, REGKIND_NONE},
            Threading::GetCurrentThreadId,
            Variant::VARIANT,
        },
        UI::WindowsAndMessaging::*,
    },
    core::GUID,
};

const E_NOINTERFACE: i32 = 0x80004002u32 as i32;
const E_POINTER: i32 = 0x80004003u32 as i32;
const CLASS_E_NOAGGREGATION: i32 = 0x80040110u32 as i32;
const DISPOSE: u32 = WM_APP + 27;
const UNKNOWN: GUID = GUID::from_u128(0x00000000_0000_0000_c000_000000000046);
const DISPATCH: GUID = GUID::from_u128(0x00020400_0000_0000_c000_000000000046);
const FACTORY: GUID = GUID::from_u128(0x00000001_0000_0000_c000_000000000046);
type Construct = unsafe extern "system" fn(*mut *mut c_void) -> i32;
type ReleaseInstance = unsafe extern "system" fn(*mut *mut c_void) -> i32;

fn guid(text: &str) -> Result<GUID, String> {
    let digits = text
        .chars()
        .filter(|c| !matches!(c, '{' | '}' | '-'))
        .collect::<String>();
    u128::from_str_radix(&digits, 16)
        .map(GUID::from_u128)
        .map_err(|_| format!("invalid GUID {text}"))
}
fn equal(a: &GUID, b: &GUID) -> bool {
    a.data1 == b.data1 && a.data2 == b.data2 && a.data3 == b.data3 && a.data4 == b.data4
}
unsafe fn add_ref_raw(object: *mut c_void) {
    // SAFETY: caller retains a COM interface; IUnknown is its first three slots.
    unsafe {
        let table = *(object.cast::<*const usize>());
        let call: unsafe extern "system" fn(*mut c_void) -> u32 =
            std::mem::transmute(*table.add(1));
        call(object);
    }
}
unsafe fn release_raw(object: *mut c_void) {
    if object.is_null() {
        return;
    }
    // SAFETY: caller transfers one retained COM interface reference.
    unsafe {
        let table = *(object.cast::<*const usize>());
        let call: unsafe extern "system" fn(*mut c_void) -> u32 =
            std::mem::transmute(*table.add(2));
        call(object);
    }
}
fn entry(name: &str) -> Result<usize, String> {
    let name = std::ffi::CString::new(name).map_err(|e| e.to_string())?;
    // SAFETY: the current executable and its patched export table stay loaded
    // throughout the server; no foreign library or procedure token is resolved.
    unsafe {
        GetProcAddress(GetModuleHandleW(ptr::null()), name.as_ptr().cast())
            .map(|value| value as usize)
            .ok_or_else(|| format!("missing compiled export {name:?}"))
    }
}
struct Class {
    iid: GUID,
    clsid: GUID,
    table: Vec<usize>,
    type_info: *mut c_void,
    constructor: Construct,
    release: ReleaseInstance,
    owner: u32,
    single_use: bool,
}
impl Drop for Class {
    fn drop(&mut self) {
        // SAFETY: registrations and instances have gone before their type owners.
        unsafe {
            release_raw(self.type_info);
        }
    }
}
#[repr(C)]
struct Object {
    table: *const usize,
    /// Native wrappers load this second word before entering the compiled body.
    instance: *mut c_void,
    references: AtomicU32,
    class: *const Class,
}
#[repr(C)]
struct ClassFactory {
    table: *const FactoryTable,
    references: AtomicU32,
    class: *const Class,
}
#[repr(C)]
struct FactoryTable {
    query: unsafe extern "system" fn(*mut ClassFactory, *const GUID, *mut *mut c_void) -> i32,
    add: unsafe extern "system" fn(*mut ClassFactory) -> u32,
    release: unsafe extern "system" fn(*mut ClassFactory) -> u32,
    create: unsafe extern "system" fn(
        *mut ClassFactory,
        *mut c_void,
        *const GUID,
        *mut *mut c_void,
    ) -> i32,
    lock: unsafe extern "system" fn(*mut ClassFactory, i32) -> i32,
}
static FACTORY_TABLE: FactoryTable = FactoryTable {
    query: factory_query,
    add: factory_add,
    release: factory_release,
    create: factory_create,
    lock: factory_lock,
};

unsafe extern "system" fn query(this: *mut Object, iid: *const GUID, out: *mut *mut c_void) -> i32 {
    if out.is_null() {
        return E_POINTER;
    }
    // SAFETY: COM passes a live self and writable out; null IID is rejected.
    unsafe {
        out.write(ptr::null_mut());
        if iid.is_null() {
            return E_POINTER;
        }
        if equal(&*iid, &UNKNOWN) || equal(&*iid, &DISPATCH) || equal(&*iid, &(*(*this).class).iid)
        {
            add(this);
            out.write(this.cast());
            0
        } else {
            E_NOINTERFACE
        }
    }
}
unsafe extern "system" fn add(this: *mut Object) -> u32 {
    // SAFETY: live COM self; reference counts alone are thread-safe.
    unsafe { (*this).references.fetch_add(1, Ordering::Relaxed) + 1 }
}
unsafe fn dispose(this: *mut Object) {
    // SAFETY: final COM release is routed to the owning apartment. The compiled
    // ByRef release body clears the project owner and drains Class_Terminate.
    unsafe {
        let mut object = Box::from_raw(this);
        let class = &*object.class;
        let status = (class.release)(&mut object.instance);
        if status < 0 {
            eprintln!("compiled object release failed: {status:#x}");
        }
        drop(object);
        if CoReleaseServerProcess() == 0 {
            PostQuitMessage(0);
        }
    }
}
unsafe extern "system" fn release(this: *mut Object) -> u32 {
    // SAFETY: every caller owns one reference. Foreign final releases enqueue
    // disposal; no VBA body or finalizer can run on that foreign thread.
    unsafe {
        let remaining = (*this).references.fetch_sub(1, Ordering::AcqRel) - 1;
        if remaining == 0 {
            let owner = (*(*this).class).owner;
            if GetCurrentThreadId() == owner {
                dispose(this);
            } else {
                PostThreadMessageW(owner, DISPOSE, this as usize, 0);
            }
        }
        remaining
    }
}
unsafe extern "system" fn type_count(_: *mut Object, out: *mut u32) -> i32 {
    if out.is_null() {
        return E_POINTER;
    }
    // SAFETY: COM supplied writable UINT storage.
    unsafe {
        out.write(1);
    }
    0
}
unsafe extern "system" fn type_info(
    this: *mut Object,
    index: u32,
    _: u32,
    out: *mut *mut c_void,
) -> i32 {
    if out.is_null() {
        return E_POINTER;
    }
    // SAFETY: live COM self, retained type information, writable out.
    unsafe {
        out.write(ptr::null_mut());
        if index != 0 {
            return 0x8002000bu32 as i32;
        }
        let info = (*(*this).class).type_info;
        add_ref_raw(info);
        out.write(info);
    }
    0
}
unsafe extern "system" fn names(
    this: *mut Object,
    _: *const GUID,
    names: *const *const u16,
    count: u32,
    _: u32,
    ids: *mut i32,
) -> i32 {
    // SAFETY: Automation validates the caller's names/count buffers; immutable
    // type information resolves the compiled vtable's member slots.
    unsafe { DispGetIDsOfNames((*(*this).class).type_info, names, count, ids) }
}
unsafe extern "system" fn invoke(
    this: *mut Object,
    member: i32,
    _: *const GUID,
    _: u32,
    flags: u16,
    params: *mut DISPPARAMS,
    result: *mut VARIANT,
    exception: *mut EXCEPINFO,
    argument_error: *mut u32,
) -> i32 {
    // SAFETY: COM marshals callers onto this registered STA; the additional
    // owner check precedes DispInvoke and every generated VBA entry.
    unsafe {
        let class = &*(*this).class;
        if GetCurrentThreadId() != class.owner {
            return E_WRONG_THREAD;
        }
        DispInvoke(
            this.cast(),
            class.type_info,
            member,
            flags,
            params,
            result,
            exception,
            argument_error,
        )
    }
}
unsafe extern "system" fn factory_query(
    this: *mut ClassFactory,
    iid: *const GUID,
    out: *mut *mut c_void,
) -> i32 {
    if out.is_null() {
        return E_POINTER;
    }
    // SAFETY: live factory self and writable COM out.
    unsafe {
        out.write(ptr::null_mut());
        if iid.is_null() {
            return E_POINTER;
        }
        if equal(&*iid, &UNKNOWN) || equal(&*iid, &FACTORY) {
            factory_add(this);
            out.write(this.cast());
            0
        } else {
            E_NOINTERFACE
        }
    }
}
unsafe extern "system" fn factory_add(this: *mut ClassFactory) -> u32 {
    // SAFETY: a live factory owns its atomic count.
    unsafe { (*this).references.fetch_add(1, Ordering::Relaxed) + 1 }
}
unsafe extern "system" fn factory_release(this: *mut ClassFactory) -> u32 {
    // SAFETY: caller transfers a live factory reference; final release owns it.
    unsafe {
        let remaining = (*this).references.fetch_sub(1, Ordering::AcqRel) - 1;
        if remaining == 0 {
            drop(Box::from_raw(this));
        }
        remaining
    }
}
unsafe extern "system" fn factory_create(
    this: *mut ClassFactory,
    outer: *mut c_void,
    iid: *const GUID,
    out: *mut *mut c_void,
) -> i32 {
    if out.is_null() {
        return E_POINTER;
    }
    // SAFETY: all pointers belong to the live factory/COM request; no compiled
    // constructor runs before the thread, aggregation and IID checks.
    unsafe {
        out.write(ptr::null_mut());
        let class = &*(*this).class;
        if GetCurrentThreadId() != class.owner {
            return E_WRONG_THREAD;
        }
        if !outer.is_null() {
            return CLASS_E_NOAGGREGATION;
        }
        if iid.is_null() {
            return E_POINTER;
        }
        if !equal(&*iid, &UNKNOWN) && !equal(&*iid, &DISPATCH) && !equal(&*iid, &class.iid) {
            return E_NOINTERFACE;
        }
        let mut instance = ptr::null_mut();
        let status = (class.constructor)(&mut instance);
        if status < 0 {
            return status;
        }
        if instance.is_null() {
            return 0x80004005u32 as i32;
        }
        CoAddRefServerProcess();
        let object = Box::new(Object {
            table: class.table.as_ptr(),
            instance,
            references: AtomicU32::new(1),
            class,
        });
        out.write(Box::into_raw(object).cast());
        0
    }
}
unsafe extern "system" fn factory_lock(this: *mut ClassFactory, lock: i32) -> i32 {
    // SAFETY: process lifetime counts are manipulated on this factory's STA.
    unsafe {
        if GetCurrentThreadId() != (*(*this).class).owner {
            return E_WRONG_THREAD;
        }
        if lock != 0 {
            CoAddRefServerProcess();
        } else if CoReleaseServerProcess() == 0 {
            PostQuitMessage(0);
        }
    }
    0
}
struct Apartment;
impl Drop for Apartment {
    fn drop(&mut self) {
        // SAFETY: run initialized this thread's apartment exactly once.
        unsafe {
            CoUninitialize();
        }
    }
}
struct Registrations(Vec<u32>);
impl Drop for Registrations {
    fn drop(&mut self) {
        // SAFETY: only cookies returned by this process's CoRegisterClassObject are revoked.
        unsafe {
            for cookie in self.0.drain(..) {
                CoRevokeClassObject(cookie);
            }
        }
    }
}
fn check(operation: &str, status: i32) -> Result<(), String> {
    if status < 0 {
        Err(format!("{operation}: {status:#x}"))
    } else {
        Ok(())
    }
}

pub(super) fn run() -> Result<(), String> {
    let executable = std::env::current_exe().map_err(|e| e.to_string())?;
    let descriptor: NativeComServer = serde_json::from_slice(
        &std::fs::read(executable.with_extension("native-com.json")).map_err(|e| e.to_string())?,
    )
    .map_err(|e| e.to_string())?;
    // SAFETY: the process main thread owns initialization, registration, message
    // dispatch, every VBA activation, revocation and final shutdown in this scope.
    unsafe {
        check(
            "CoInitializeEx",
            CoInitializeEx(ptr::null(), COINIT_APARTMENTTHREADED as u32),
        )?;
        let _apartment = Apartment;
        let owner = GetCurrentThreadId();
        let mut message: MSG = std::mem::zeroed();
        PeekMessageW(&mut message, ptr::null_mut(), 0, 0, PM_NOREMOVE);
        let path: Vec<_> = executable
            .with_extension("tlb")
            .to_string_lossy()
            .encode_utf16()
            .chain(Some(0))
            .collect();
        let mut library = ptr::null_mut();
        check(
            "LoadTypeLibEx",
            LoadTypeLibEx(path.as_ptr(), REGKIND_NONE, &mut library),
        )?;
        let mut classes = Vec::new();
        let load_info = || -> Result<Vec<Box<Class>>, String> {
            let get_info: unsafe extern "system" fn(
                *mut c_void,
                *const GUID,
                *mut *mut c_void,
            ) -> i32 = std::mem::transmute(*(*(library.cast::<*const usize>())).add(6));
            let mut result = Vec::new();
            for metadata in descriptor.classes {
                let iid = guid(&metadata.iid)?;
                let mut type_info = ptr::null_mut();
                check(
                    "ITypeLib::GetTypeInfoOfGuid",
                    get_info(library, &iid, &mut type_info),
                )?;
                let mut table = vec![
                    query as *const () as usize,
                    add as *const () as usize,
                    release as *const () as usize,
                    type_count as *const () as usize,
                    self::type_info as *const () as usize,
                    names as *const () as usize,
                    invoke as *const () as usize,
                ];
                table.extend(
                    metadata
                        .entries
                        .iter()
                        .map(|name| entry(name))
                        .collect::<Result<Vec<_>, _>>()?,
                );
                result.push(Box::new(Class {
                    clsid: guid(&metadata.clsid)?,
                    iid,
                    table,
                    type_info,
                    constructor: std::mem::transmute::<usize, Construct>(entry(
                        &metadata.constructor,
                    )?),
                    release: std::mem::transmute::<usize, ReleaseInstance>(entry(
                        &metadata.release,
                    )?),
                    owner,
                    single_use: metadata.single_use,
                }));
            }
            Ok(result)
        };
        let loaded = load_info();
        release_raw(library);
        classes.extend(loaded?);
        let mut registrations = Registrations(Vec::new());
        for class in &classes {
            let factory = Box::into_raw(Box::new(ClassFactory {
                table: &FACTORY_TABLE,
                references: AtomicU32::new(1),
                class: &**class,
            }));
            let mut cookie = 0;
            let status = CoRegisterClassObject(
                &class.clsid,
                factory.cast(),
                CLSCTX_LOCAL_SERVER,
                ((if class.single_use {
                    REGCLS_SINGLEUSE
                } else {
                    REGCLS_MULTIPLEUSE
                }) | REGCLS_SUSPENDED) as u32,
                &mut cookie,
            );
            factory_release(factory);
            check("CoRegisterClassObject", status)?;
            registrations.0.push(cookie);
        }
        check("CoResumeClassObjects", CoResumeClassObjects())?;
        eprintln!(
            "CalcVBA COM server ready: pid={} thread={owner}",
            std::process::id()
        );
        loop {
            let status = GetMessageW(&mut message, ptr::null_mut(), 0, 0);
            if status == -1 {
                return Err("GetMessageW failed".into());
            }
            if status == 0 {
                break;
            }
            if message.message == DISPOSE {
                dispose(message.wParam as *mut Object);
            } else {
                TranslateMessage(&message);
                DispatchMessageW(&message);
            }
        }
        drop(registrations);
        let shutdown: unsafe extern "system" fn() -> i32 =
            std::mem::transmute(entry("CalcVbaShutdown")?);
        check("compiled shutdown", shutdown())?;
        eprintln!(
            "CalcVBA COM server stopped: pid={} thread={owner}",
            std::process::id()
        );
    }
    Ok(())
}
