//! The in-process COM host (calcvba-comhost) driven as an Automation client drives it:
//! the built DLL is loaded and asked for its class factory directly, so nothing is
//! registered and no Office application is involved.
#![cfg(target_os = "windows")]
#![allow(unsafe_op_in_unsafe_fn)]

use std::ffi::c_void;
use std::path::PathBuf;
use std::process::Command;
use std::sync::OnceLock;

use windows_sys::Win32::Foundation::SysFreeString;
use windows_sys::Win32::System::Com::{
    COINIT_APARTMENTTHREADED, CoInitializeEx, DISPATCH_METHOD, DISPPARAMS, EXCEPINFO,
};
use windows_sys::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};
use windows_sys::Win32::System::Variant::{VARIANT, VT_BYREF, VT_I4, VariantClear};
use windows_sys::core::GUID;

const IID_ICLASSFACTORY: GUID = GUID {
    data1: 0x0000_0001,
    data2: 0,
    data3: 0,
    data4: [0xC0, 0, 0, 0, 0, 0, 0, 0x46],
};
const IID_IUNKNOWN: GUID = GUID {
    data1: 0,
    data2: 0,
    data3: 0,
    data4: [0xC0, 0, 0, 0, 0, 0, 0, 0x46],
};
const IID_IDISPATCH: GUID = GUID {
    data1: 0x0002_0400,
    data2: 0,
    data3: 0,
    data4: [0xC0, 0, 0, 0, 0, 0, 0, 0x46],
};
const IID_NULL: GUID = GUID {
    data1: 0,
    data2: 0,
    data3: 0,
    data4: [0; 8],
};
const DISP_E_UNKNOWNNAME: i32 = 0x8002_0006u32 as i32;
const DISP_E_PARAMNOTFOUND: i32 = 0x8002_0004u32 as i32;

#[repr(C)]
struct ClassFactoryVtbl {
    query_interface: usize,
    add_ref: usize,
    release: unsafe extern "system" fn(*mut c_void) -> u32,
    create_instance:
        unsafe extern "system" fn(*mut c_void, *mut c_void, *const GUID, *mut *mut c_void) -> i32,
    lock_server: usize,
}

#[repr(C)]
struct DispatchVtbl {
    query_interface: unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> i32,
    add_ref: usize,
    release: unsafe extern "system" fn(*mut c_void) -> u32,
    get_type_info_count: usize,
    get_type_info: usize,
    get_ids_of_names: unsafe extern "system" fn(
        *mut c_void,
        *const GUID,
        *const *const u16,
        u32,
        u32,
        *mut i32,
    ) -> i32,
    invoke: unsafe extern "system" fn(
        *mut c_void,
        i32,
        *const GUID,
        u32,
        u16,
        *mut DISPPARAMS,
        *mut VARIANT,
        *mut EXCEPINFO,
        *mut u32,
    ) -> i32,
}

struct Dispatch(*mut c_void);

impl Dispatch {
    unsafe fn vtbl(&self) -> &DispatchVtbl {
        &**self.0.cast::<*const DispatchVtbl>()
    }

    /// GetIDsOfNames for a member and its parameter names: the HRESULT and the DISPIDs.
    unsafe fn ids(&self, names: &[&str]) -> (i32, Vec<i32>) {
        let wide: Vec<Vec<u16>> = names
            .iter()
            .map(|name| name.encode_utf16().chain([0]).collect())
            .collect();
        let pointers: Vec<*const u16> = wide.iter().map(|name| name.as_ptr()).collect();
        let mut dispids = vec![0; names.len()];
        let hr = (self.vtbl().get_ids_of_names)(
            self.0,
            &IID_NULL,
            pointers.as_ptr(),
            names.len() as u32,
            0,
            dispids.as_mut_ptr(),
        );
        (hr, dispids)
    }

    /// Invoke a method with `rgvarg` as given (last argument first, named ones first of all).
    unsafe fn call(&self, member: &str, rgvarg: &mut [VARIANT], named: &[i32]) -> (i32, VARIANT) {
        let (hr, ids) = self.ids(&[member]);
        assert_eq!(hr, 0, "GetIDsOfNames({member})");
        let mut named = named.to_vec();
        let mut params = DISPPARAMS {
            rgvarg: rgvarg.as_mut_ptr(),
            rgdispidNamedArgs: named.as_mut_ptr(),
            cArgs: rgvarg.len() as u32,
            cNamedArgs: named.len() as u32,
        };
        let mut result: VARIANT = std::mem::zeroed();
        let mut excep: EXCEPINFO = std::mem::zeroed();
        let mut arg_err = 0;
        let hr = (self.vtbl().invoke)(
            self.0,
            ids[0],
            &IID_NULL,
            0,
            DISPATCH_METHOD,
            &mut params,
            &mut result,
            &mut excep,
            &mut arg_err,
        );
        if hr != 0 && !excep.bstrDescription.is_null() {
            let len = windows_sys::Win32::Foundation::SysStringLen(excep.bstrDescription) as usize;
            eprintln!(
                "{member}: {}",
                String::from_utf16_lossy(std::slice::from_raw_parts(excep.bstrDescription, len))
            );
        }
        SysFreeString(excep.bstrSource);
        SysFreeString(excep.bstrDescription);
        SysFreeString(excep.bstrHelpFile);
        (hr, result)
    }

    unsafe fn call_long(&self, member: &str, rgvarg: &mut [VARIANT], named: &[i32]) -> i32 {
        let (hr, result) = self.call(member, rgvarg, named);
        assert_eq!(hr, 0, "Invoke({member}) HRESULT 0x{:08X}", hr as u32);
        assert_eq!(
            result.Anonymous.Anonymous.vt, VT_I4,
            "{member} returns a Long"
        );
        result.Anonymous.Anonymous.Anonymous.lVal
    }

    unsafe fn identity(&self, value: &VARIANT) -> *mut c_void {
        let object = value.Anonymous.Anonymous.Anonymous.pdispVal;
        assert!(!object.is_null(), "an object result");
        let vtbl = &**object.cast::<*const DispatchVtbl>();
        let mut unknown = std::ptr::null_mut();
        assert_eq!(
            (vtbl.query_interface)(object, &IID_IUNKNOWN, &mut unknown),
            0
        );
        (vtbl.release)(unknown);
        unknown
    }
}

impl Drop for Dispatch {
    fn drop(&mut self) {
        // SAFETY: this owns one reference on a live IDispatch.
        unsafe {
            (self.vtbl().release)(self.0);
        }
    }
}

fn long(value: i32) -> VARIANT {
    // SAFETY: an all-zero VARIANT is VT_EMPTY; a VT_I4 sets only lVal.
    unsafe {
        let mut variant: VARIANT = std::mem::zeroed();
        variant.Anonymous.Anonymous.vt = VT_I4;
        variant.Anonymous.Anonymous.Anonymous.lVal = value;
        variant
    }
}

fn long_ref(target: &mut i32) -> VARIANT {
    // SAFETY: a VT_BYREF | VT_I4 cell refers to the caller's Long.
    unsafe {
        let mut variant: VARIANT = std::mem::zeroed();
        variant.Anonymous.Anonymous.vt = VT_BYREF | VT_I4;
        variant.Anonymous.Anonymous.Anonymous.plVal = target;
        variant
    }
}

fn comhost_dll() -> PathBuf {
    static PREBUILT: OnceLock<PathBuf> = OnceLock::new();
    PREBUILT
        .get_or_init(|| {
            let status = Command::new("cargo")
                .args(["build", "-p", "calcvba-comhost", "--release"])
                .status()
                .expect("cargo build -p calcvba-comhost should run");
            assert!(status.success(), "cargo build -p calcvba-comhost failed");
            PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                .join("../../target/release/calcvba_comhost.dll")
        })
        .clone()
}

const CLASS: &str = r#"Attribute VB_Name = "Mixer"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Public Function Mix(ByVal a As Long, Optional ByVal b As Long = 10, Optional ByVal c As Long = 100) As Long
    Mix = a * 10000 + b * 100 + c
End Function

Public Sub Bump(ByRef target As Long, ByVal by As Long)
    target = target + by
End Sub

Public Function Self() As Object
    Set Self = Me
End Function
"#;

const PROJECT: &str = r#"<Project Sdk="CalcVBA.Sdk/0.1.0">
  <PropertyGroup>
    <OutputType>ComServer</OutputType>
    <BuildTarget>WrappedComServer</BuildTarget>
    <ProjectName>InprocServer</ProjectName>RUNTIME
  </PropertyGroup>
  <ItemGroup>
    <ClassModule Include="Mixer.cls">
      <ProgId>InprocServer.Mixer</ProgId>
    </ClassModule>
  </ItemGroup>
</Project>
"#;

#[test]
fn comhost_takes_named_arguments_and_keeps_one_wrapper_per_object() {
    check_server("vm3", "");
}

/// The same server with RuntimeFlavor=Jit: comhost runs the image on the JIT (dc-yqq0.10).
#[test]
fn comhost_on_the_jit_takes_named_arguments_and_keeps_one_wrapper_per_object() {
    check_server("jit", "\n    <RuntimeFlavor>Jit</RuntimeFlavor>");
}

fn check_server(name: &str, runtime: &str) {
    let dir = std::env::temp_dir().join(format!(
        "calcvba_comhost_inproc_{name}_{}",
        std::process::id()
    ));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    std::fs::write(dir.join("Mixer.cls"), CLASS).unwrap();
    std::fs::write(
        dir.join("Inproc.basproj"),
        PROJECT.replace("RUNTIME", runtime),
    )
    .unwrap();
    let output =
        calcvba_build::build_wrapped_com_server(&calcvba_build::WrappedComServerBuildOptions {
            project_path: dir.join("Inproc.basproj"),
            out_dir: dir.join("out"),
            compile_dll: true,
            comhost_dll_path: Some(comhost_dll()),
        })
        .expect("WrappedComServer build");
    let descriptor: calcvba_build::ComServerDescriptor =
        serde_json::from_str(&std::fs::read_to_string(&output.descriptor_path).unwrap()).unwrap();
    let clsid = parse_guid(&descriptor.classes[0].clsid);

    // SAFETY: the test owns the loaded module and every interface it obtains, on one
    // apartment-threaded thread.
    unsafe {
        CoInitializeEx(std::ptr::null(), COINIT_APARTMENTTHREADED as u32);
        let path: Vec<u16> = output
            .dll_target_path
            .as_os_str()
            .to_string_lossy()
            .encode_utf16()
            .chain([0])
            .collect();
        let module = LoadLibraryW(path.as_ptr());
        assert!(
            !module.is_null(),
            "load {}",
            output.dll_target_path.display()
        );
        let get_class_object: unsafe extern "system" fn(
            *const GUID,
            *const GUID,
            *mut *mut c_void,
        ) -> i32 = std::mem::transmute(
            GetProcAddress(module, c"DllGetClassObject".as_ptr().cast()).expect("export"),
        );
        let mut factory = std::ptr::null_mut();
        assert_eq!(
            get_class_object(&clsid, &IID_ICLASSFACTORY, &mut factory),
            0
        );
        let factory_vtbl = &**factory.cast::<*const ClassFactoryVtbl>();
        let mut instance = std::ptr::null_mut();
        assert_eq!(
            (factory_vtbl.create_instance)(
                factory,
                std::ptr::null_mut(),
                &IID_IDISPATCH,
                &mut instance
            ),
            0
        );
        (factory_vtbl.release)(factory);
        let mixer = Dispatch(instance);

        // Parameter DISPIDs are their positions; an unknown name is DISPID_UNKNOWN.
        let (hr, ids) = mixer.ids(&["Mix", "c", "a"]);
        assert_eq!((hr, &ids[1..]), (0, &[2, 0][..]));
        let (hr, ids) = mixer.ids(&["Mix", "a", "nope"]);
        assert_eq!((hr, &ids[1..]), (DISP_E_UNKNOWNNAME, &[0, -1][..]));

        // Mix(c:=3, a:=1): b takes its default.
        assert_eq!(
            mixer.call_long("Mix", &mut [long(3), long(1)], &[2, 0]),
            11003
        );
        // Mix(1, c:=5): positional a, named c.
        assert_eq!(mixer.call_long("Mix", &mut [long(5), long(1)], &[2]), 11005);
        // Mix(1, 2, 3), all positional.
        assert_eq!(
            mixer.call_long("Mix", &mut [long(3), long(2), long(1)], &[]),
            10203
        );
        // A parameter named twice, or one that does not exist.
        let (hr, _) = mixer.call("Mix", &mut [long(1), long(1)], &[0, 0]);
        assert_eq!(hr, DISP_E_PARAMNOTFOUND);
        let (hr, _) = mixer.call("Mix", &mut [long(1)], &[7]);
        assert_eq!(hr, DISP_E_PARAMNOTFOUND);

        // Bump(by:=5, target:=x) writes back through the named ByRef argument.
        let mut x = 2;
        let (hr, _) = mixer.call("Bump", &mut [long(5), long_ref(&mut x)], &[1, 0]);
        assert_eq!(hr, 0);
        assert_eq!(x, 7);

        // The same project object comes back as the same COM object, so Is holds.
        let (hr, mut first) = mixer.call("Self", &mut [], &[]);
        assert_eq!(hr, 0);
        let (hr, mut second) = mixer.call("Self", &mut [], &[]);
        assert_eq!(hr, 0);
        assert_eq!(mixer.identity(&first), mixer.identity(&second));
        VariantClear(&mut first);
        VariantClear(&mut second);

        // The module stays loaded: the host keeps its session for the process.
        drop(mixer);
    }
}

fn parse_guid(text: &str) -> GUID {
    let hex: String = text.chars().filter(char::is_ascii_hexdigit).collect();
    let byte = |i: usize| u8::from_str_radix(&hex[i * 2..i * 2 + 2], 16).unwrap();
    GUID {
        data1: u32::from_str_radix(&hex[0..8], 16).unwrap(),
        data2: u16::from_str_radix(&hex[8..12], 16).unwrap(),
        data3: u16::from_str_radix(&hex[12..16], 16).unwrap(),
        data4: [
            byte(8),
            byte(9),
            byte(10),
            byte(11),
            byte(12),
            byte(13),
            byte(14),
            byte(15),
        ],
    }
}
