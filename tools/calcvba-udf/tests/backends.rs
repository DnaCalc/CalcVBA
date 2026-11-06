#![cfg(windows)]
use calcvba_engine::{Engine, HostConfig};
use std::{
    ffi::{CString, c_void},
    fs,
    path::PathBuf,
    process::Command,
};

#[repr(C)]
#[derive(Clone, Copy)]
struct Oper {
    payload: [u64; 3],
    kind: u32,
    padding: u32,
}
impl Oper {
    fn number(n: f64) -> Self {
        Self {
            payload: [n.to_bits(), 0, 0],
            kind: 1,
            padding: 0,
        }
    }
    fn missing() -> Self {
        Self {
            payload: [0; 3],
            kind: 128,
            padding: 0,
        }
    }
}
#[link(name = "kernel32")]
unsafe extern "system" {
    fn LoadLibraryW(path: *const u16) -> *mut c_void;
    fn GetProcAddress(module: *mut c_void, name: *const u8) -> *mut c_void;
    fn FreeLibrary(module: *mut c_void) -> i32;
}
struct Dll(*mut c_void);
impl Dll {
    fn load(path: &std::path::Path) -> Self {
        let path = path
            .to_string_lossy()
            .encode_utf16()
            .chain(Some(0))
            .collect::<Vec<_>>();
        // SAFETY: terminated path; this test retains the image until Drop.
        let module = unsafe { LoadLibraryW(path.as_ptr()) };
        assert!(!module.is_null());
        Self(module)
    }
    unsafe fn entry(&self, name: &str) -> *mut c_void {
        let name = CString::new(name).unwrap();
        // SAFETY: the image and terminated export name remain retained.
        let entry = unsafe { GetProcAddress(self.0, name.as_ptr().cast()) };
        assert!(!entry.is_null(), "missing {name:?}");
        entry
    }
    fn call(&self, name: &str, args: &[Oper]) -> String {
        // SAFETY: generated exports use only LongPtr inputs/results. Every
        // argument points to a retained XLOPER12, and the owned result is freed
        // by the same image on this calling thread after copying its contents.
        unsafe {
            let entry = self.entry(name);
            let p = args
                .iter()
                .map(|v| std::ptr::from_ref(v) as usize)
                .collect::<Vec<_>>();
            let pointer = match p.as_slice() {
                [] => std::mem::transmute::<*mut c_void, extern "C" fn() -> usize>(entry)(),
                [a] => std::mem::transmute::<*mut c_void, extern "C" fn(usize) -> usize>(entry)(*a),
                [a, b] => std::mem::transmute::<*mut c_void, extern "C" fn(usize, usize) -> usize>(
                    entry,
                )(*a, *b),
                [a, b, c] => std::mem::transmute::<
                    *mut c_void,
                    extern "C" fn(usize, usize, usize) -> usize,
                >(entry)(*a, *b, *c),
                [a, b, c, d, e] => std::mem::transmute::<
                    *mut c_void,
                    extern "C" fn(usize, usize, usize, usize, usize) -> usize,
                >(entry)(*a, *b, *c, *d, *e),
                _ => panic!("test call arity"),
            };
            assert_ne!(pointer, 0, "{name} returned null");
            let text = render(&*(pointer as *const Oper));
            let free: extern "C" fn(usize) = std::mem::transmute(self.entry("xlAutoFree12"));
            free(pointer);
            text
        }
    }
}
impl Drop for Dll {
    fn drop(&mut self) {
        // SAFETY: all output buffers were freed; cleanup precedes image unload
        // and remains on this image's first calling thread.
        unsafe {
            let shutdown: extern "C" fn() -> i32 =
                std::mem::transmute(self.entry("CalcVbaShutdown"));
            assert_eq!(shutdown(), 0);
            FreeLibrary(self.0);
        }
    }
}
unsafe fn render(value: &Oper) -> String {
    match value.kind & 4095 {
        1 => f64::from_bits(value.payload[0]).to_string(),
        2 => {
            let p = value.payload[0] as *const u16;
            // SAFETY: a returned string owns a counted UTF-16 buffer through render.
            unsafe {
                String::from_utf16_lossy(std::slice::from_raw_parts(p.add(1), p.read() as usize))
            }
        }
        4 => if value.payload[0] == 0 {
            "False"
        } else {
            "True"
        }
        .into(),
        16 => format!("E{}", value.payload[0]),
        64 => {
            let count = (value.payload[1] as u32 as usize) * (value.payload[1] >> 32) as usize;
            // SAFETY: the result owns its exact rows*columns contiguous cell vector.
            let cells =
                unsafe { std::slice::from_raw_parts(value.payload[0] as *const Oper, count) };
            // SAFETY: every child is retained by that vector through this recursion.
            cells
                .iter()
                .map(|v| unsafe { render(v) })
                .collect::<Vec<_>>()
                .join(":")
        }
        kind => panic!("unexpected Excel kind {kind}"),
    }
}
#[test]
fn real_com_udf_bodies_and_generated_native_wrappers_agree() {
    let directory =
        std::env::temp_dir().join(format!("calcvba-udf-backends-{}", std::process::id()));
    fs::create_dir(&directory).unwrap();
    assert_eq!(
        directory.canonicalize().unwrap().parent(),
        Some(std::env::temp_dir().canonicalize().unwrap().as_path())
    );
    let fixture =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/projects/Functions.basproj");
    let output = directory.join("generated");
    let generated = Command::new(env!("CARGO_BIN_EXE_calcvba-udf"))
        .args([
            "--project",
            fixture.to_str().unwrap(),
            "--out",
            output.to_str().unwrap(),
            "--prefix",
            "Native_",
            "--paramarray-slots",
            "4",
        ])
        .output()
        .unwrap();
    assert!(
        generated.status.success(),
        "{}",
        String::from_utf8_lossy(&generated.stderr)
    );
    let mut manifest = calcvba_project::load_project_closure(&fixture)
        .unwrap()
        .pop()
        .unwrap();
    manifest.modules.push(calcvba_symbol::manifest::ModuleUnit {
        module_name:"BackendProbe".into(),module_kind:calcvba_symbol::manifest::ModuleKind::Procedural,
        attributes:calcvba_symbol::manifest::ModuleAttributes::named("BackendProbe"),
        source:"Public Sub Main()\nDim result As String, matrixInput(1 To 2,1 To 2) As Variant, matrixResult As Variant\nmatrixInput(1,1)=1\nmatrixResult=Functions.Matrix(matrixInput)\nresult=CStr(Functions.Scalar(2#,3&)) & \"|\" & Functions.Defaults() & \"|\" & CStr(Functions.SumItems(1#,2#,3#,4#)) & \"|\" & CStr(Functions.SumItems(1#)) & \"|\" & CStr(matrixResult(1,1)) & \":\" & CStr(matrixResult(1,2)) & \":\" & CStr(matrixResult(2,1)) & \":\" & CStr(IsError(matrixResult(2,2)))\nEnd Sub\nPublic Property Get Application() As Object\nEnd Property\n".into(),
    });
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(calcvba_hal::HostPolicy::native());
        let values = engine
            .execute_manifest_with_variant_snapshot(&manifest)
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|v| v.as_bstr())
                .any(|v| v.to_string() == "5|7:True|10|1|1:matrix:True:True"),
            "{values:?}"
        );
    }
    let table = fs::read_to_string(output.join("registration.tsv")).unwrap();
    let export = |source: &str| {
        table
            .lines()
            .skip(1)
            .find_map(|line| {
                let fields = line.split('\t').collect::<Vec<_>>();
                (fields[4] == source).then_some(fields[0])
            })
            .unwrap()
    };
    let dll = Dll::load(&output.join("FunctionsUdf.xll"));
    assert_eq!(
        dll.call(
            export("Functions.Scalar"),
            &[Oper::number(2.0), Oper::number(3.0)]
        ),
        "5"
    );
    assert_eq!(
        dll.call(
            export("Functions.Defaults"),
            &[Oper::missing(), Oper::missing()]
        ),
        "7:True"
    );
    assert_eq!(
        dll.call(
            export("Functions.SumItems"),
            &[
                Oper::number(1.0),
                Oper::number(2.0),
                Oper::number(3.0),
                Oper::number(4.0),
                Oper::missing()
            ]
        ),
        "10"
    );
    assert_eq!(
        dll.call(
            export("Functions.SumItems"),
            &[
                Oper::number(1.0),
                Oper::missing(),
                Oper::missing(),
                Oper::missing(),
                Oper::missing()
            ]
        ),
        "1"
    );
    let cells = [
        Oper::number(1.0),
        Oper::number(2.0),
        Oper::number(3.0),
        Oper::number(4.0),
    ];
    let matrix = Oper {
        payload: [cells.as_ptr() as u64, 2 | (2u64 << 32), 0],
        kind: 64,
        padding: 0,
    };
    assert_eq!(
        dll.call(export("Functions.Matrix"), &[matrix]),
        "1:matrix:True:E42"
    );
    assert_eq!(dll.call(export("Functions.ArrayArgument"), &[matrix]), "5");
    assert_eq!(dll.call(export("Functions.ErrorValue"), &[]), "E42");
    assert_eq!(dll.call(export("Functions.Unhandled"), &[]), "E15");
    assert_eq!(
        dll.call(export("Functions.OptionalReference"), &[Oper::missing()]),
        "none"
    );
    assert_eq!(
        dll.call(
            export("Functions.OptionalData"),
            &[Oper::missing(), Oper::missing()]
        ),
        "2:123.4567"
    );
    let text = [4u16, 84, 114, 117, 101];
    let string = Oper {
        payload: [text.as_ptr() as u64, 0, 0],
        kind: 2,
        padding: 0,
    };
    assert_eq!(
        dll.call(
            export("Functions.ScalarKinds"),
            &[string, Oper::number(3.5), Oper::number(9.0)]
        ),
        "True:4:9"
    );
    drop(dll);
    fs::remove_dir_all(&directory).unwrap();
}
