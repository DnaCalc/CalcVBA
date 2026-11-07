#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_build::{NativeLibraryBuildOptions, build_native_library};
use calcvba_engine::{Engine, HostConfig};
use std::{
    path::PathBuf,
    sync::atomic::{AtomicUsize, Ordering},
};
use windows_sys::Win32::{
    Foundation::{FreeLibrary, HMODULE, SysFreeString, SysStringLen},
    System::LibraryLoader::{GetProcAddress, LoadLibraryW},
};

struct Directory(PathBuf);
impl Directory {
    fn new() -> Self {
        static NEXT: AtomicUsize = AtomicUsize::new(0);
        let path = std::env::temp_dir().join(format!(
            "calcvba-native-members-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        std::fs::create_dir_all(&path).unwrap();
        assert_eq!(
            path.canonicalize().unwrap().parent(),
            Some(std::env::temp_dir().canonicalize().unwrap().as_path())
        );
        Self(path)
    }
}
impl Drop for Directory {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}
struct Dll(HMODULE);
impl Dll {
    fn load(path: &std::path::Path) -> Self {
        let name: Vec<_> = path
            .to_string_lossy()
            .encode_utf16()
            .chain(Some(0))
            .collect();
        // SAFETY: terminated absolute path; this test owns the DLL until Drop.
        let module = unsafe { LoadLibraryW(name.as_ptr()) };
        assert!(!module.is_null());
        Self(module)
    }
    fn call(&self, name: &str) -> String {
        let name = std::ffi::CString::new(name).unwrap();
        // SAFETY: test exports are zero-argument BSTR Functions; owning exports
        // and the error accessor are called on this DLL's first calling thread.
        unsafe {
            let entry: unsafe extern "C" fn() -> *mut u16 =
                std::mem::transmute(GetProcAddress(self.0, name.as_ptr().cast()).unwrap());
            let error: unsafe extern "C" fn(*mut usize, *mut usize) -> i32 = std::mem::transmute(
                GetProcAddress(self.0, c"CalcVbaGetLastError".as_ptr().cast()).unwrap(),
            );
            let result = entry();
            let mut source = 0;
            let mut description = 0;
            let number = error(&mut source, &mut description);
            let take = |pointer: *mut u16| {
                let text = if pointer.is_null() {
                    String::new()
                } else {
                    String::from_utf16_lossy(std::slice::from_raw_parts(
                        pointer,
                        SysStringLen(pointer) as usize,
                    ))
                };
                SysFreeString(pointer);
                text
            };
            let result = take(result);
            let source = take(source as *mut u16);
            let description = take(description as *mut u16);
            assert_eq!(number, 0, "{name:?}: {source}: {description}");
            result
        }
    }
}

#[test]
fn variant_array_and_project_default_indexing_share_compiled_routes() {
    let directory = Directory::new();
    let support = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/support/native_members");
    for name in ["Indexing.bas", "Indexed.cls", "Keyed.cls"] {
        std::fs::copy(support.join(name), directory.0.join(name)).unwrap();
    }
    let project = directory.0.join("Indexing.basproj");
    std::fs::write(&project, r#"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>Indexing</ProjectName></PropertyGroup><ItemGroup><Module Include="Indexing.bas"/><ClassModule Include="Indexed.cls"/><ClassModule Include="Keyed.cls"/><NativeExport Include="IndexSuite"><Module>Indexing</Module><Procedure>IndexSuite</Procedure></NativeExport></ItemGroup></Project>"#).unwrap();
    let closure = calcvba_project::load_project_closure(&project).unwrap();
    let expected = "blue:green|1:changed|True:True|key:stored|entry|dictionary|13";
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(calcvba_hal::HostPolicy::native());
        let values = engine
            .execute_manifest_with_variant_snapshot(closure.last().unwrap())
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|v| v.as_bstr())
                .any(|v| v.to_string() == expected),
            "{values:?}"
        );
    }
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: directory.0.join("out"),
        template_path: std::env::var_os("CALCVBA_NATIVE_TEMPLATE_TEST").map(PathBuf::from),
    })
    .unwrap();
    let dll = Dll::load(&output.dll_path);
    assert_eq!(dll.call("IndexSuite"), expected);
}
impl Drop for Dll {
    fn drop(&mut self) {
        // SAFETY: same-thread cleanup runs before the owned image is unloaded.
        unsafe {
            let shutdown: unsafe extern "C" fn() -> i32 = std::mem::transmute(
                GetProcAddress(self.0, c"CalcVbaShutdown".as_ptr().cast()).unwrap(),
            );
            let _ = shutdown();
            FreeLibrary(self.0);
        }
    }
}

#[test]
fn collection_interfaces_late_calls_and_failed_construction_agree() {
    let directory = Directory::new();
    let support = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/support/native_members");
    for entry in std::fs::read_dir(support).unwrap() {
        let entry = entry.unwrap();
        std::fs::copy(entry.path(), directory.0.join(entry.file_name())).unwrap();
    }
    let project = directory.0.join("Members.basproj");
    std::fs::write(&project,r#"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>NativeMembers</ProjectName></PropertyGroup><ItemGroup><Module Include="Members.bas"/><ClassModule Include="ICounter.cls"/><ClassModule Include="Worker.cls"/><ClassModule Include="OtherWorker.cls"/><ClassModule Include="BadCtor.cls"/><NativeExport Include="MemberSuite"><Module>Members</Module><Procedure>MemberSuite</Procedure></NativeExport><NativeExport Include="ConstructorFaultSuite"><Module>Members</Module><Procedure>ConstructorFaultSuite</Procedure></NativeExport><NativeExport Include="NothingSuite"><Module>Members</Module><Procedure>NothingSuite</Procedure></NativeExport><NativeExport Include="ProjectCallByNameSuite"><Module>Members</Module><Procedure>ProjectCallByNameSuite</Procedure></NativeExport></ItemGroup></Project>"#).unwrap();
    let closure = calcvba_project::load_project_closure(&project).unwrap();
    // Captured by VBE Compile and execution in the reused owned Excel process.
    let expected =
        "15:5|35:7:True|Collection:3:a:a23:2:0|z:3|1:9|438:3#513:Nothing:1:0#91#named:worker";
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(calcvba_hal::HostPolicy::native());
        let values = engine
            .execute_manifest_with_variant_snapshot(closure.last().unwrap())
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|value| value.as_bstr())
                .any(|value| value.to_string() == expected),
            "{values:?}"
        );
    }
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: directory.0.join("out"),
        template_path: std::env::var_os("CALCVBA_NATIVE_TEMPLATE_TEST").map(PathBuf::from),
    })
    .unwrap();
    let dll = Dll::load(&output.dll_path);
    assert_eq!(
        format!(
            "{}#{}#{}#{}",
            dll.call("MemberSuite"),
            dll.call("ConstructorFaultSuite"),
            dll.call("NothingSuite"),
            dll.call("ProjectCallByNameSuite")
        ),
        expected
    );
}

#[test]
fn imported_early_dictionary_members_agree() {
    let directory = Directory::new();
    let support = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/support/native_members/NativeCom.bas");
    std::fs::copy(support, directory.0.join("NativeCom.bas")).unwrap();
    let project = directory.0.join("Early.basproj");
    std::fs::write(&project,r#"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>NativeEarly</ProjectName></PropertyGroup><ItemGroup><Module Include="NativeCom.bas"/><COMReference Include="Scripting"><Guid>{420B2830-E718-11CF-893D-00A0C9054228}</Guid><VersionMajor>1</VersionMajor><VersionMinor>0</VersionMinor><Lcid>0</Lcid><ImportLib>scrrun.dll</ImportLib></COMReference><NativeExport Include="EarlyOwnDictionarySuite"><Module>NativeCom</Module><Procedure>EarlyOwnDictionarySuite</Procedure></NativeExport></ItemGroup></Project>"#).unwrap();
    let closure = calcvba_project::load_project_closure(&project).unwrap();
    let expected = "2:7:11:True:True:False|5:Invalid procedure call or argument";
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(calcvba_hal::HostPolicy::native());
        let values = engine
            .execute_manifest_with_variant_snapshot(closure.last().unwrap())
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|value| value.as_bstr())
                .any(|value| value.to_string() == expected),
            "{values:?}"
        );
    }
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: directory.0.join("out"),
        template_path: std::env::var_os("CALCVBA_NATIVE_TEMPLATE_TEST").map(PathBuf::from),
    })
    .unwrap();
    let dll = Dll::load(&output.dll_path);
    assert_eq!(dll.call("EarlyOwnDictionarySuite"), expected);
}

#[test]
fn paramarray_storage_optional_and_gosub_bodies_agree() {
    let directory = Directory::new();
    std::fs::copy(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests/support/native_members/Parameters.bas"),
        directory.0.join("Parameters.bas"),
    )
    .unwrap();
    let project = directory.0.join("Parameters.basproj");
    // VBE Compile and execution in the same recorded owned Excel session.
    let cases = [
        ("ParameterSuite", "13:15:b:0:4:11:13|0:-1"),
        ("FixedSuite", "3:[\0\0\0]:abc:[x  ]:4:[xy  ]"),
        ("GoSubSuite", "OIROIR|3"),
        ("OptionalSuite", "7:True:s|7:False:t|2:True:q"),
        ("NumericWidthSuite", "-12:-3:7:Long"),
        ("TypedParameterSuite", "9:[zzz]:6:Byte:1:[zzzz]:Byte:9:7"),
    ];
    let exports=cases.iter().map(|(name,_)|format!("<NativeExport Include=\"{name}\"><Module>Parameters</Module><Procedure>{name}</Procedure></NativeExport>")).collect::<String>();
    std::fs::write(&project,format!("<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>NativeParameters</ProjectName></PropertyGroup><ItemGroup><Module Include=\"Parameters.bas\"/>{exports}</ItemGroup></Project>")).unwrap();
    let closure = calcvba_project::load_project_closure(&project).unwrap();
    let expected = cases
        .iter()
        .map(|(_, expected)| *expected)
        .collect::<Vec<_>>()
        .join("#");
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(calcvba_hal::HostPolicy::native());
        let values = engine
            .execute_manifest_with_variant_snapshot(closure.last().unwrap())
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|value| value.as_bstr())
                .any(|value| value.to_string() == expected),
            "{values:?}"
        );
    }
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: directory.0.join("out"),
        template_path: std::env::var_os("CALCVBA_NATIVE_TEMPLATE_TEST").map(PathBuf::from),
    })
    .unwrap();
    let dll = Dll::load(&output.dll_path);
    for (name, expected) in cases {
        assert_eq!(dll.call(name), expected, "{name}");
    }
}

#[test]
fn dynamic_project_names_have_source_refusals() {
    let directory = Directory::new();
    let support = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/support/native_members");
    for name in ["ICounter.cls", "Worker.cls"] {
        std::fs::copy(support.join(name), directory.0.join(name)).unwrap();
    }
    let cases = [(
        "Public Function Probe() As String\nDim receiver As Worker, member As String\nSet receiver=New Worker\nmember=\"Echo\"\nProbe=CallByName(receiver,member,vbMethod)\nEnd Function\n",
        "dynamic project CallByName",
    )];
    for (index, (source, expected)) in cases.iter().enumerate() {
        std::fs::write(
            directory.0.join("Code.bas"),
            format!("Attribute VB_Name = \"Code\"\n{source}"),
        )
        .unwrap();
        let project = directory.0.join("Refused.basproj");
        std::fs::write(&project,"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>Refused</ProjectName></PropertyGroup><ItemGroup><Module Include=\"Code.bas\"/><ClassModule Include=\"ICounter.cls\"/><ClassModule Include=\"Worker.cls\"/><NativeExport Include=\"Probe\"><Module>Code</Module><Procedure>Probe</Procedure></NativeExport></ItemGroup></Project>").unwrap();
        let error = build_native_library(&NativeLibraryBuildOptions {
            project_path: project,
            out_dir: directory.0.join(format!("out{index}")),
            template_path: std::env::var_os("CALCVBA_NATIVE_TEMPLATE_TEST").map(PathBuf::from),
        })
        .unwrap_err()
        .to_string();
        assert!(error.contains(expected), "{error}");
        assert!(error.contains("Code:"), "{error}");
    }
}

#[test]
fn temporary_arrays_character_input_and_nullable_fields_agree() {
    let directory = Directory::new();
    let support = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/support/native_members");
    let input = directory.0.join("input.dat");
    std::fs::write(&input, [97u8, 44, 98, 13, 10, 34, 99, 34, 0, 233]).unwrap();
    let control = directory.0.join("control.dat");
    std::fs::write(&control, [65u8, 26, 66, 0, 67]).unwrap();
    let buffered = directory.0.join("buffered.dat");
    let mut buffered_data = vec![66u8; 5000];
    buffered_data[0] = 65;
    buffered_data[1] = 26;
    std::fs::write(&buffered, buffered_data).unwrap();
    let source = std::fs::read_to_string(support.join("Expressions.bas"))
        .unwrap()
        .replace("@INPUT_PATH@", &input.to_string_lossy())
        .replace("@CONTROL_PATH@", &control.to_string_lossy())
        .replace("@BUFFERED_PATH@", &buffered.to_string_lossy());
    std::fs::write(directory.0.join("Expressions.bas"), source).unwrap();
    std::fs::copy(
        support.join("Nullable.cls"),
        directory.0.join("Nullable.cls"),
    )
    .unwrap();
    // The array/Input results were captured with visible VBE Compile in the
    // reused owned Excel. Text input skips NUL; binary input retains it.
    let cases = [
        ("ArraySuite", "2:0:b:left:1:IF:1:FD"),
        (
            "InputSuite",
            "97@2,44@3,98@4,13@5,10@6,34@7,99@8,34@9,233@11,E62@11|zero=0:0|negative=5|handle=52|over=62:11:0#97@2,44@3,98@4,13@5,10@6,34@7,99@8,34@9,0@10,233@11,|zero=0:0|negative=5|handle=52|over=0:11:10",
        ),
        ("ObjectIndexSuite", "collection:8"),
        ("ScopeSuite", "9:True:True|107:107|438:438"),
        ("ByValFieldSuite", "8:field!:4294967297"),
        (
            "ControlSuite",
            "65@2:True,E62@8:True,E62@8:True,E62@8:True,",
        ),
        (
            "BufferedControlSuite",
            "65@2:True,E62@515:True,E62@515:True,E62@515:True,",
        ),
    ];
    let exports = cases.iter().map(|(name,_)| format!("<NativeExport Include=\"{name}\"><Module>Expressions</Module><Procedure>{name}</Procedure></NativeExport>")).collect::<String>();
    let project = directory.0.join("Expressions.basproj");
    std::fs::write(&project, format!("<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>Expressions</ProjectName></PropertyGroup><ItemGroup><Module Include=\"Expressions.bas\"/><ClassModule Include=\"Nullable.cls\"/>{exports}</ItemGroup></Project>")).unwrap();
    let closure = calcvba_project::load_project_closure(&project).unwrap();
    let expected = cases
        .iter()
        .map(|(_, value)| *value)
        .collect::<Vec<_>>()
        .join("#");
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let mut engine = Engine::new(config);
        engine.set_host_policy(calcvba_hal::HostPolicy::native());
        let values = engine
            .execute_manifest_with_variant_snapshot(closure.last().unwrap())
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|v| v.as_bstr())
                .any(|v| v.to_string() == expected),
            "{values:?}"
        );
    }
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: directory.0.join("out"),
        template_path: std::env::var_os("CALCVBA_NATIVE_TEMPLATE_TEST").map(PathBuf::from),
    })
    .unwrap();
    let dll = Dll::load(&output.dll_path);
    for (name, expected) in cases {
        assert_eq!(dll.call(name), expected, "{name}");
    }
}

#[test]
fn chained_project_receivers_and_inline_cases_agree() {
    let directory = Directory::new();
    let support = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/support/native_members");
    for name in [
        "Chains.bas",
        "IChain.cls",
        "ChainCore.cls",
        "ChainControl.cls",
    ] {
        std::fs::copy(support.join(name), directory.0.join(name)).unwrap();
    }
    let project = directory.0.join("Chains.basproj");
    std::fs::write(&project, r#"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>Chains</ProjectName></PropertyGroup><ItemGroup><Module Include="Chains.bas"/><ClassModule Include="IChain.cls"/><ClassModule Include="ChainCore.cls"/><ClassModule Include="ChainControl.cls"/><NativeExport Include="ChainSuite"><Module>Chains</Module><Procedure>ChainSuite</Procedure></NativeExport><NativeExport Include="CaseSuite"><Module>Chains</Module><Procedure>CaseSuite</Procedure></NativeExport></ItemGroup></Project>"#).unwrap();
    let closure = calcvba_project::load_project_closure(&project).unwrap();
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let engine = Engine::new(config);
        let values = engine
            .execute_manifest_with_variant_snapshot(closure.last().unwrap())
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|v| v.as_bstr())
                .any(|v| v.to_string() == "42:42:438:438:438"),
            "{values:?}"
        );
        let mut cases = closure.last().unwrap().clone();
        cases.modules[0].source = cases.modules[0]
            .source
            .replace("observed = ChainSuite()", "observed = CaseSuite()");
        let values = engine
            .execute_manifest_with_variant_snapshot(&cases)
            .unwrap();
        assert!(
            values
                .iter()
                .filter_map(|v| v.as_bstr())
                .any(|v| v.to_string() == "4:8:8:12:16:"),
            "{values:?}"
        );
    }
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: directory.0.join("out"),
        template_path: None,
    })
    .unwrap();
    let dll = Dll::load(&output.dll_path);
    assert_eq!(dll.call("ChainSuite"), "42:42:438:438:438");
    assert_eq!(dll.call("CaseSuite"), "4:8:8:12:16:");
}

#[test]
fn execution_census_integer_array_and_copyout_regressions_agree() {
    let directory = Directory::new();
    let source = include_str!("support/native_execution_regressions.bas");
    std::fs::write(directory.0.join("ExecutionRegressions.bas"), source).unwrap();
    let project = directory.0.join("ExecutionRegressions.basproj");
    let cases = [
        ("ExecutionSuite", "2:3:11:13:13:91:91:91:91"),
        (
            "ArrayOrderSuite",
            "13:23:14:24:|135:235:145:245:136:236:146:246:",
        ),
        ("RandomAritySuite", "True:True:True:True"),
        ("CounterSuite", "2:5:3:4"),
        ("NestedCallSuite", "2:10:3:10"),
        (
            "DisagreementSuite",
            "424:424:91:9:9:5:99:2:6:6:6:True:False",
        ),
        ("ObjectLetSuite", "424:424:424"),
        ("ArrayPtrSuite", "2:66"),
        ("ArrayPtrShapeSuite", "2:1:3:8209:Byte()"),
    ];
    let exports = cases.iter().map(|(name, _)| format!("<NativeExport Include=\"{name}\"><Module>ExecutionRegressions</Module><Procedure>{name}</Procedure></NativeExport>")).collect::<String>();
    std::fs::write(&project, format!(r#"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>ExecutionRegressions</ProjectName></PropertyGroup><ItemGroup><Module Include="ExecutionRegressions.bas"/>{exports}</ItemGroup></Project>"#)).unwrap();
    for (name, expected) in cases {
        let source = source.replace(
            "observed = ExecutionSuite()",
            &format!("observed = {name}()"),
        );
        for config in [HostConfig::vm3(), HostConfig::jit()] {
            let mut engine = Engine::new(config);
            engine.set_host_policy(calcvba_hal::HostPolicy::native());
            let values = engine
                .execute_source_with_variant_snapshot_clean(&source)
                .unwrap();
            assert!(
                values
                    .iter()
                    .filter_map(|v| v.as_bstr())
                    .any(|v| v.to_string() == expected),
                "{name}: {values:?}"
            );
        }
    }
    let output = build_native_library(&NativeLibraryBuildOptions {
        project_path: project,
        out_dir: directory.0.join("out"),
        template_path: None,
    })
    .unwrap();
    let dll = Dll::load(&output.dll_path);
    for (name, expected) in cases {
        assert_eq!(dll.call(name), expected, "{name}");
    }
}
