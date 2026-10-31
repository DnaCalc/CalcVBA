//! Bounded Windows x64 scalar COM/Declare execution, not release certification.
#![cfg(all(windows, target_arch = "x86_64"))]

use oxvba_differential::{Canon, canon};
use oxvba_hal::model::{ComInvocationStrategy, HostPolicy};
use oxvba_host::{Engine, FinalErr, HostConfig, SnapshotOutcome};
use oxvba_runtime::{Variant, current_thread_live_handle_counts};
use oxvba_symbol::manifest as sym;
use std::{path::PathBuf, process::Command};

#[derive(Debug, PartialEq)]
struct Observation {
    values: Vec<Canon>,
    err: FinalErr,
    raised: bool,
    transport: (u64, u64),
}

fn run(source: &str, jit: bool) -> Observation {
    let before = current_thread_live_handle_counts();
    let observation = {
        let mut engine = Engine::new(if jit {
            HostConfig::jit()
        } else {
            HostConfig::vm3()
        });
        // Deterministic HAL tokens are not native-call evidence.
        let mut policy = HostPolicy::interactive_dev();
        policy.com_invocation_strategy = ComInvocationStrategy::DispatchOnly;
        engine.set_host_policy(policy);
        let manifest = sym::SymbolProjectManifest {
            project_name: "Main".into(),
            project_kind: sym::ProjectKind::Source,
            modules: vec![sym::ModuleUnit {
                module_name: "Main".into(),
                module_kind: sym::ModuleKind::Procedural,
                attributes: sym::ModuleAttributes::named("Main"),
                source: source.into(),
            }],
            references: vec![],
            reference_projects: vec![],
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
        };
        let outcome = if jit {
            engine.execute_manifest_snapshot_with_err_jit(&manifest)
        } else {
            engine.execute_manifest_snapshot_with_err_vm3(&manifest)
        };
        let transport = engine.host_services().com().com_dispatch_transport_counts();
        match outcome {
            SnapshotOutcome::Completed { values, err } => Observation {
                values: values.iter().map(canon).collect(),
                err,
                raised: false,
                transport,
            },
            SnapshotOutcome::Raised { err } => Observation {
                values: vec![],
                err,
                raised: true,
                transport,
            },
            other => panic!(
                "{} must execute, got {other:?}",
                if jit { "JIT" } else { "VM3" }
            ),
        }
    };
    let balance = before.balance_to(current_thread_live_handle_counts());
    assert!(
        balance.is_zero(),
        "owned value handles leaked, jit={jit}: {balance:?}"
    );
    observation
}

fn compare(source: &str) -> Observation {
    let vm = run(source, false);
    let jit = run(source, true);
    assert_eq!(jit, vm, "native VM3/JIT observable differs");
    vm
}

#[test]
fn scalar_late_com_methods_properties_and_errors_execute() {
    let observed = compare(
        r#"
Public result As Long
Public ping As Long
Public count As Long
Sub Main()
    Dim obj As Object
    Set obj = CreateObject("OxVba.TestDispatch")
    obj.SetValue = 42
    result = obj.Value
    ping = obj.Ping()
    count = obj.Count
    Set obj = Nothing
End Sub
"#,
    );
    assert_eq!(
        &observed.values[..3],
        &[42, 123, 7].map(|n| canon(&Variant::from_i32(n)))
    );
    assert!(!observed.raised);
    assert_eq!(observed.err.number, 0);
    assert_eq!(
        observed.transport,
        (0, 4),
        "must use actual IDispatch calls"
    );

    let failure = compare(
        r#"
Sub Main()
    Dim obj As Object
    Set obj = CreateObject("OxVba.TestDispatch")
    obj.RaiseException
End Sub
"#,
    );
    assert!(failure.raised);
    assert_ne!(failure.err.number, 0);
    assert!(
        failure
            .err
            .description
            .contains("controlled dispatch exception")
    );
    // The existing transport counter counts successful invokes only.
    assert_eq!(failure.transport, (0, 0));
}

#[test]
fn late_com_activates_registered_class_from_dynamic_progid() {
    let observed = compare(
        r#"
Public result As Long
Public count As Long
Sub Main()
    Dim progId As String
    Dim obj As Object
    progId = "Scripting.Dictionary"
    Set obj = CreateObject(progId)
    obj.Add "answer", 19
    obj.Item("answer") = 42
    result = obj.Item("answer")
    count = obj.Count
    Set obj = Nothing
End Sub
"#,
    );
    assert_eq!(
        &observed.values[..2],
        &[42, 1].map(|n| canon(&Variant::from_i32(n)))
    );
    assert!(!observed.raised);
    assert_eq!(observed.transport, (0, 4));
}

/// Test-owned deterministic DLL. No registration or global PATH changes.
struct NativeFixture(PathBuf);

impl NativeFixture {
    fn build() -> Self {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root =
            std::env::temp_dir().join(format!("oxvba-scalar-{}-{nonce}", std::process::id()));
        std::fs::create_dir(&root).unwrap();
        let fixture = Self(root);
        let output = Command::new(std::env::var_os("RUSTC").unwrap_or_else(|| "rustc".into()))
            .args([
                "--edition=2024",
                "--crate-type=cdylib",
                "--crate-name=oxvba_scalar_fixture",
            ])
            .arg(PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/support/native_scalar.rs"))
            .arg("-o")
            .arg(fixture.path())
            .output()
            .expect("start fixture rustc");
        assert!(
            output.status.success(),
            "fixture compilation failed: {}",
            String::from_utf8_lossy(&output.stderr)
        );
        fixture
    }

    fn path(&self) -> PathBuf {
        self.0.join("scalar.dll")
    }
}

impl Drop for NativeFixture {
    fn drop(&mut self) {
        std::fs::remove_dir_all(&self.0).expect("remove owned scalar fixture after unloading DLL");
    }
}

#[test]
fn scalar_declare_function_sub_byref_and_last_dll_error_execute() {
    let fixture = NativeFixture::build();
    let path = fixture.path().display().to_string();
    let observed = compare(&format!(
        r#"
Public result As Long
Public changed As Long
Public nativeError As Long
Declare PtrSafe Function ScalarAdd Lib "{path}" (ByVal a As Long, ByVal b As Long) As Long
Declare PtrSafe Sub ScalarWrite Lib "{path}" (ByRef value As Long)
Sub Main()
    result = ScalarAdd(19, 23)
    ScalarWrite changed
    nativeError = Err.LastDllError
End Sub
"#
    ));
    assert_eq!(
        &observed.values[..3],
        &[42, 73, 1234].map(|n| canon(&Variant::from_i32(n)))
    );
    assert!(!observed.raised);
    assert_eq!(observed.err.number, 0);
    assert_eq!(observed.err.last_dll_error, 1234);
    assert_eq!(
        observed.transport,
        (0, 0),
        "Declare uses Win64 ABI, not COM"
    );

    let missing = compare(&format!(
        r#"
Declare PtrSafe Function Absent Lib "{path}" () As Long
Sub Main()
    Dim value As Long
    value = Absent()
End Sub
"#
    ));
    assert!(missing.raised, "missing export is a runtime error");
    assert_eq!(
        missing.err.number, 453,
        "a found DLL with a missing named export is VBA error 453"
    );
}
