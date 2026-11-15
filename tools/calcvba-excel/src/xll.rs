//! The XLL loader: the same DLL opened by `RegisterXLL`. `xlAutoOpen` registers
//! `CALCVBA.RUN(source, module, entry, backend)`, which compiles `source` as the standard
//! module `module` (with Excel's type library referenced), registers the hosting Excel's
//! `Application` as the instance of `Excel.Global`, and returns `module.entry()` as text.
//!
//! The hosting Excel is found from `xlGetHwnd`: its `XLMAIN` window, the `EXCEL7` child, and
//! `AccessibleObjectFromWindow(OBJID_NATIVEOM)` (the bd-p5bb identity rule), never the
//! Running Object Table. No VBA project is involved, so this checks CalcVBA alone in Excel.

use super::{Dispatch, Entry, failure, register, run};
use calcvba_engine::{Engine, HostConfig};
use calcvba_hal::model::HostPolicy;
use calcvba_symbol::manifest as sym;
use std::cell::RefCell;
use windows_sys::Win32::Foundation::HWND;
use windows_sys::Win32::System::LibraryLoader::{GetModuleHandleW, GetProcAddress};
use windows_sys::Win32::System::Threading::GetCurrentThreadId;
use windows_sys::Win32::UI::Accessibility::AccessibleObjectFromWindow;
use windows_sys::Win32::UI::WindowsAndMessaging::{
    EnumChildWindows, EnumThreadWindows, GetClassNameW,
};
use windows_sys::core::GUID;

/// `XLOPER12`: a 24-byte value union and its type tag.
#[repr(C)]
#[derive(Clone, Copy)]
pub struct Xloper12 {
    val: [u64; 3],
    xltype: u32,
}

const XLTYPE_NUM: u32 = 0x0001;
const XLTYPE_STR: u32 = 0x0002;
const XLTYPE_INT: u32 = 0x0800;
const XL_SPECIAL: i32 = 0x4000;
const XL_FREE: i32 = XL_SPECIAL;
const XL_GET_NAME: i32 = 9 | XL_SPECIAL;
const XL_GET_HWND: i32 = 10 | XL_SPECIAL;
const XLF_REGISTER: i32 = 149;
const OBJID_NATIVEOM: i32 = 0xFFFF_FFF0_u32 as i32;
const EXCEL_GLOBAL: &str = "{00020812-0000-0000-C000-000000000046}";
const EXCEL_LIBID: &str = "{00020813-0000-0000-C000-000000000046}";
const IID_IDISPATCH: GUID = GUID::from_u128(0x0002_0400_0000_0000_c000_0000_0000_0046);

type MdCallBack12 = unsafe extern "system" fn(i32, i32, *mut *mut Xloper12, *mut Xloper12) -> i32;

/// `Excel12v` over the host's `MdCallBack12`.
fn excel(function: i32, args: &mut [&mut Xloper12]) -> Option<Xloper12> {
    // SAFETY: the hosting EXCEL.EXE exports MdCallBack12 with this signature; the argument
    // and result XLOPER12s are live for the call.
    unsafe {
        let callback = GetProcAddress(
            GetModuleHandleW(core::ptr::null()),
            c"MdCallBack12".as_ptr().cast(),
        )?;
        let callback: MdCallBack12 = core::mem::transmute(callback);
        let mut pointers: Vec<*mut Xloper12> = args
            .iter_mut()
            .map(|arg| core::ptr::from_mut(*arg))
            .collect();
        let mut result = Xloper12 {
            val: [0; 3],
            xltype: 0,
        };
        let status = callback(
            function,
            pointers.len() as i32,
            pointers.as_mut_ptr(),
            &mut result,
        );
        (status == 0).then_some(result)
    }
}

/// A length-prefixed UTF-16 string the XLOPER12 points into.
struct XlString(Vec<u16>);

impl XlString {
    fn new(text: &str) -> Self {
        let mut units: Vec<u16> = text.encode_utf16().take(32767).collect();
        units.insert(0, units.len() as u16);
        Self(units)
    }

    fn oper(&mut self) -> Xloper12 {
        Xloper12 {
            val: [self.0.as_mut_ptr() as u64, 0, 0],
            xltype: XLTYPE_STR,
        }
    }
}

/// The text of a string XLOPER12 (anything else is "").
///
/// # Safety
/// `oper` is null or a live XLOPER12 from Excel.
unsafe fn oper_text(oper: *const Xloper12) -> String {
    // SAFETY: a string XLOPER12 points at a length-prefixed UTF-16 buffer.
    unsafe {
        if oper.is_null() || (*oper).xltype & 0x0FFF != XLTYPE_STR {
            return String::new();
        }
        let text = (*oper).val[0] as *const u16;
        let len = usize::from(*text);
        String::from_utf16_lossy(std::slice::from_raw_parts(text.add(1), len))
    }
}

/// The number in a numeric XLOPER12.
///
/// # Safety
/// `oper` is null or a live XLOPER12 from Excel.
unsafe fn oper_number(oper: *const Xloper12) -> Option<f64> {
    // SAFETY: the tag says which union member is live.
    unsafe {
        if oper.is_null() {
            return None;
        }
        match (*oper).xltype & 0x0FFF {
            XLTYPE_NUM => Some(f64::from_bits((*oper).val[0])),
            XLTYPE_INT => Some(f64::from((*oper).val[0] as i32)),
            _ => None,
        }
    }
}

/// Registers `CALCVBA.RUN`.
#[unsafe(no_mangle)]
pub extern "system" fn xlAutoOpen() -> i32 {
    let Some(mut name) = excel(XL_GET_NAME, &mut []) else {
        return 0;
    };
    let mut procedure = XlString::new("calcvba_xll_run");
    // Q result, four Q arguments; `#` (a macro-sheet function) may use the C API and the
    // object model while it runs.
    let mut types = XlString::new("QQQQQ#");
    let mut function = XlString::new("CALCVBA.RUN");
    let mut arguments = XlString::new("source,module,entry,backend");
    let mut kind = Xloper12 {
        val: [1f64.to_bits(), 0, 0],
        xltype: XLTYPE_NUM,
    };
    let (mut procedure, mut types, mut function, mut arguments) = (
        procedure.oper(),
        types.oper(),
        function.oper(),
        arguments.oper(),
    );
    let registered = excel(
        XLF_REGISTER,
        &mut [
            &mut name,
            &mut procedure,
            &mut types,
            &mut function,
            &mut arguments,
            &mut kind,
        ],
    )
    .is_some();
    let _ = excel(XL_FREE, &mut [&mut name]);
    i32::from(registered)
}

/// Releases nothing the add-in keeps: every `CALCVBA.RUN` closes its project.
#[unsafe(no_mangle)]
pub extern "system" fn xlAutoClose() -> i32 {
    1
}

/// The hosting Excel's `Application`, from its main window (one owned reference).
fn host_application() -> Result<Dispatch, (i32, String)> {
    let Some(oper) = excel(XL_GET_HWND, &mut []) else {
        return Err(failure("xlGetHwnd failed"));
    };
    let mut main = (oper.val[0] as u32 as usize) as HWND;
    if main.is_null() {
        // The XLMAIN window of this thread: Excel's main thread runs the add-in.
        unsafe extern "system" fn find_main(window: HWND, found: isize) -> i32 {
            let mut class = [0u16; 16];
            // SAFETY: a live window of this thread and a bounded buffer.
            let len = unsafe { GetClassNameW(window, class.as_mut_ptr(), class.len() as i32) };
            if String::from_utf16_lossy(&class[..len.max(0) as usize]) == "XLMAIN" {
                // SAFETY: `found` is the caller's HWND slot for the enumeration.
                unsafe { *(found as *mut HWND) = window };
                return 0;
            }
            1
        }
        // SAFETY: enumerates this thread's own top-level windows into a local slot.
        unsafe {
            EnumThreadWindows(
                GetCurrentThreadId(),
                Some(find_main),
                core::ptr::from_mut(&mut main) as isize,
            )
        };
        if main.is_null() {
            return Err(failure(format!(
                "xlGetHwnd gave no window (xltype {:#x}) and this thread has no XLMAIN",
                oper.xltype
            )));
        }
    }
    // SAFETY: plain window lookups and an accessibility query on Excel's own windows; the
    // returned IDispatch carries one reference, which the Dispatch takes over.
    unsafe {
        // The first EXCEL7 (workbook window) anywhere under the main window.
        unsafe extern "system" fn find(window: HWND, found: isize) -> i32 {
            let mut class = [0u16; 16];
            // SAFETY: a live child window and a bounded buffer.
            let len = unsafe { GetClassNameW(window, class.as_mut_ptr(), class.len() as i32) };
            if String::from_utf16_lossy(&class[..len.max(0) as usize]) == "EXCEL7" {
                // SAFETY: `found` is the caller's HWND slot for the enumeration.
                unsafe { *(found as *mut HWND) = window };
                return 0;
            }
            1
        }
        let mut book: HWND = core::ptr::null_mut();
        EnumChildWindows(main, Some(find), core::ptr::from_mut(&mut book) as isize);
        if book.is_null() {
            let mut class = [0u16; 32];
            let len = GetClassNameW(main, class.as_mut_ptr(), class.len() as i32);
            return Err(failure(format!(
                "no EXCEL7 window under {:?} ({}): open a workbook window first",
                main,
                String::from_utf16_lossy(&class[..len.max(0) as usize])
            )));
        }
        let mut window: *mut core::ffi::c_void = core::ptr::null_mut();
        if AccessibleObjectFromWindow(book, OBJID_NATIVEOM as u32, &IID_IDISPATCH, &mut window) < 0
            || window.is_null()
        {
            return Err(failure("AccessibleObjectFromWindow failed"));
        }
        let window = Dispatch(window);
        property(&window, "Application")
    }
}

fn w(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(Some(0)).collect()
}

/// `object.name` (a property get returning an object), with one owned reference.
fn property(object: &Dispatch, name: &str) -> Result<Dispatch, (i32, String)> {
    type GetIdsOfNames = unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const GUID,
        *const *const u16,
        u32,
        u32,
        *mut i32,
    ) -> i32;
    type Invoke = unsafe extern "system" fn(
        *mut core::ffi::c_void,
        i32,
        *const GUID,
        u32,
        u16,
        *mut [usize; 4],
        *mut [u8; 24],
        *mut core::ffi::c_void,
        *mut u32,
    ) -> i32;
    let wide = w(name);
    let names = [wide.as_ptr()];
    let mut id = 0;
    let null_iid = GUID::from_u128(0);
    // SAFETY: IDispatch slots 5 (GetIDsOfNames) and 6 (Invoke) of a live interface; DISPPARAMS
    // is empty and the result VARIANT is ours, moved into a Dispatch when it holds an object.
    unsafe {
        let vtable = *(object.0 as *const *const usize);
        let get_ids: GetIdsOfNames = core::mem::transmute(*vtable.add(5));
        let invoke: Invoke = core::mem::transmute(*vtable.add(6));
        if get_ids(object.0, &null_iid, names.as_ptr(), 1, 0, &mut id) < 0 {
            return Err(failure(format!("no `{name}` member")));
        }
        let mut params = [0usize; 4];
        let mut result = [0u8; 24];
        let hr = invoke(
            object.0,
            id,
            &null_iid,
            0,
            2, // DISPATCH_PROPERTYGET
            &mut params,
            &mut result,
            core::ptr::null_mut(),
            core::ptr::null_mut(),
        );
        let vt = u16::from_le_bytes([result[0], result[1]]);
        let pointer = usize::from_le_bytes(result[8..16].try_into().expect("8 bytes"));
        if hr < 0 || vt != 9 || pointer == 0 {
            return Err(failure(format!("`{name}` returned no object (hr {hr:#x})")));
        }
        Ok(Dispatch(pointer as *mut core::ffi::c_void))
    }
}

thread_local! {
    /// The text of the last result, which Excel copies before the next call.
    static RESULT: RefCell<XlString> = RefCell::new(XlString::new(""));
    static RESULT_OPER: RefCell<Xloper12> = const { RefCell::new(Xloper12 { val: [0; 3], xltype: XLTYPE_STR }) };
}

/// `CALCVBA.RUN(source, module, entry, backend)`: the result text, or `#ERROR: …`.
///
/// # Safety
/// The arguments are live XLOPER12s from Excel.
#[unsafe(no_mangle)]
pub unsafe extern "system" fn calcvba_xll_run(
    source: *const Xloper12,
    module: *const Xloper12,
    entry: *const Xloper12,
    backend: *const Xloper12,
) -> *mut Xloper12 {
    // SAFETY: Excel's live arguments.
    let (source, module, entry_name, backend) = unsafe {
        (
            oper_text(source),
            oper_text(module),
            oper_text(entry),
            oper_number(backend).unwrap_or(0.0),
        )
    };
    let text = match run_source(&source, &module, &entry_name, backend == 1.0) {
        Ok(text) => text,
        Err((number, description)) => format!("#ERROR {number}: {description}"),
    };
    RESULT.with(|result| {
        let mut result = result.borrow_mut();
        *result = XlString::new(&text);
        let oper = result.oper();
        RESULT_OPER.with(|slot| {
            *slot.borrow_mut() = oper;
            slot.as_ptr()
        })
    })
}

fn run_source(
    source: &str,
    module: &str,
    entry_name: &str,
    jit: bool,
) -> Result<String, (i32, String)> {
    let mut engine = Engine::new(if jit {
        HostConfig::jit()
    } else {
        HostConfig::vm3()
    });
    engine.set_host_policy(HostPolicy::native());
    let application = host_application()?;
    register(&engine, EXCEL_GLOBAL, &application)?;
    let entry = Entry {
        modules: vec![sym::ModuleUnit {
            module_name: module.to_string(),
            module_kind: sym::ModuleKind::Procedural,
            attributes: sym::ModuleAttributes::named(module),
            source: source.to_string(),
        }],
        references: vec![sym::ProjectReference::TypeLibrary {
            name: "Excel".into(),
            guid: Some(EXCEL_LIBID.into()),
            version_major: Some(1),
            version_minor: Some(9),
            lcid: Some(0),
            import_lib: None,
        }],
        app_objects: vec![(EXCEL_GLOBAL.into(), application)],
        engine: Some((engine, jit)),
        error: (0, String::new()),
    };
    let value = run(&entry, module, entry_name)?;
    Ok(value
        .as_bstr()
        .map(|text| text.to_string())
        .unwrap_or_else(|| format!("{value:?}")))
}
