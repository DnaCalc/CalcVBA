//! Bounded Windows x64 scalar COM/Declare execution, not release certification.
#![cfg(all(windows, target_arch = "x86_64"))]

use calcvba_differential::{Canon, canon};
use calcvba_engine::{Engine, FinalErr, HostConfig, SnapshotOutcome};
use calcvba_hal::model::{ComInvocationStrategy, HostPolicy};
use calcvba_runtime::{Variant, current_thread_live_handle_counts};
use calcvba_symbol::manifest as sym;
use std::{path::PathBuf, process::Command};

#[path = "support/early_com.rs"]
mod early_com;

#[derive(Debug, PartialEq)]
struct Observation {
    values: Vec<Canon>,
    err: FinalErr,
    raised: bool,
    transport: (u64, u64),
}

fn run(source: &str, jit: bool) -> Observation {
    run_referenced(source, jit, Vec::new())
}

fn run_referenced(source: &str, jit: bool, references: Vec<sym::ProjectReference>) -> Observation {
    let before = current_thread_live_handle_counts();
    let observation = {
        let mut engine = Engine::new(if jit {
            HostConfig::jit()
        } else {
            HostConfig::vm3()
        });
        // Deterministic HAL tokens are not native-call evidence.
        let mut policy = HostPolicy::native();
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
            references,
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
            std::env::temp_dir().join(format!("calcvba-scalar-{}-{nonce}", std::process::id()));
        std::fs::create_dir(&root).unwrap();
        let fixture = Self(root);
        let output = Command::new(std::env::var_os("RUSTC").unwrap_or_else(|| "rustc".into()))
            .args([
                "--edition=2024",
                "--crate-type=cdylib",
                "--crate-name=calcvba_scalar_fixture",
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
Declare PtrSafe Sub NativeSetError Lib "kernel32" Alias "SetLastError" (ByVal value As Long)
Sub Main()
    Dim value As Long
    NativeSetError 1234
    value = Absent()
End Sub
"#
    ));
    assert!(missing.raised, "missing export is a runtime error");
    assert_eq!(
        missing.err.number, 453,
        "a found DLL with a missing named export is VBA error 453"
    );
    assert_eq!(
        missing.err.description,
        format!("Can't find DLL entry point Absent in {path}")
    );
    assert_eq!(
        missing.err.last_dll_error, 1234,
        "resolution failure preserves the last invoked DLL error"
    );

    for alias in ["#32767", "#00032767", "#&H7FFF", "#&O77777", "#&77777"] {
        let ordinal = compare(&format!(
            r##"
Declare PtrSafe Function AbsentOrdinal Lib "{path}" Alias "{alias}" () As Long
Declare PtrSafe Sub NativeSetError Lib "kernel32" Alias "SetLastError" (ByVal value As Long)
Sub Main()
    Dim value As Long
    NativeSetError 1234
    value = AbsentOrdinal()
End Sub
"##
        ));
        assert!(ordinal.raised);
        assert_eq!(ordinal.err.number, 452);
        assert_eq!(
            ordinal.err.description,
            format!("Can't find DLL entry point 32767 in {path}")
        );
        assert_eq!(ordinal.err.last_dll_error, 1234);
    }
}

#[test]
fn native_scalar_variable_aliases_match_excel_including_forwarding_and_copies() {
    let fixture = NativeFixture::build();
    let path = fixture.path().display().to_string();
    let observed = compare(&format!(
        r#"
Public result As String
Public nativeError As Long
Public callCount As Long
Public g As Long
Private Declare PtrSafe Function NativeAlias Lib "{path}" Alias "AliasLong" (ByRef first As Long, ByRef second As Long) As Long
Private Declare PtrSafe Function NativeCount Lib "{path}" Alias "AliasCallCount" () As Long
Private Function Forward(ByRef first As Long, ByRef second As Long) As Long
    Forward = NativeAlias(first, second)
End Function
Private Function GlobalMix(ByRef argument As Long) As Long
    GlobalMix = NativeAlias(argument, g)
End Function
Sub Main()
    Dim x As Long, y As Long, r As Long
    x = 1
    r = NativeAlias(x, x)
    result = CStr(r) & "|" & CStr(x)
    x = 1: y = 1
    r = NativeAlias(x, y)
    result = result & "|" & CStr(r) & "|" & CStr(x) & "|" & CStr(y)
    x = 1
    r = Forward(x, x)
    result = result & "|" & CStr(r) & "|" & CStr(x)
    x = 1
    r = NativeAlias((x), x)
    result = result & "|" & CStr(r) & "|" & CStr(x)
    g = 1
    r = GlobalMix(g)
    result = result & "|" & CStr(r) & "|" & CStr(g)
    callCount = NativeCount()
    nativeError = Err.LastDllError
End Sub
"#
    ));
    assert_eq!(
        &observed.values[..3],
        &[
            canon(&Variant::from_string("42|42|41|41|2|42|42|41|2|42|42")),
            canon(&Variant::from_i32(4321)),
            canon(&Variant::from_i32(5))
        ]
    );
    assert!(!observed.raised);
    assert_eq!(
        observed.err,
        FinalErr {
            last_dll_error: 4321,
            ..Default::default()
        }
    );
    assert_eq!(observed.transport, (0, 0));
}

#[test]
fn native_scalar_alias_storage_preserves_widths_and_vba_types() {
    let fixture = NativeFixture::build();
    let path = fixture.path().display().to_string();
    let observed = compare(&format!(
        r#"
Public i As Integer
Public b As Byte
Public truth As Boolean
Public dbl As Double
Public s As Single
Public ll As LongLong
Public p As LongPtr
Public c As Currency
Public d As Date
Public calls As Long
Declare PtrSafe Function AI Lib "{path}" Alias "AliasInteger" (ByRef a As Integer, ByRef b As Integer) As Integer
Declare PtrSafe Function AB Lib "{path}" Alias "AliasByte" (ByRef a As Byte, ByRef b As Byte) As Byte
Declare PtrSafe Function AT Lib "{path}" Alias "AliasBoolean" (ByRef a As Boolean, ByRef b As Boolean) As Boolean
Declare PtrSafe Function ADbl Lib "{path}" Alias "AliasDouble" (ByRef a As Double, ByRef b As Double) As Double
Declare PtrSafe Function ASng Lib "{path}" Alias "AliasSingle" (ByRef a As Single, ByRef b As Single) As Single
Declare PtrSafe Function ALL Lib "{path}" Alias "AliasI64" (ByRef a As LongLong, ByRef b As LongLong) As LongLong
Declare PtrSafe Function AP Lib "{path}" Alias "AliasI64" (ByRef a As LongPtr, ByRef b As LongPtr) As LongPtr
Declare PtrSafe Function AC Lib "{path}" Alias "AliasCurrency" (ByRef a As Currency, ByRef b As Currency) As Currency
Declare PtrSafe Function AD Lib "{path}" Alias "AliasDate" (ByRef a As Date, ByRef b As Date) As Date
Declare PtrSafe Function NativeCount Lib "{path}" Alias "AliasCallCount" () As Long
Sub Main()
    i = AI(i, i): Call AB(b, b): truth = AT(truth, truth)
    dbl = ADbl(dbl, dbl): s = ASng(s, s): ll = ALL(ll, ll): p = AP(p, p)
    c = AC(c, c): d = AD(d, d): calls = NativeCount()
End Sub
"#
    ));
    let expected = [
        Variant::from_i16(-32766),
        Variant::from_u8(251),
        Variant::from_bool(false),
        Variant::from_f64(123456789.75),
        Variant::from_f32(123.75),
        Variant::from_i64(4_294_967_338),
        Variant::from_i64(4_294_967_338),
        Variant::from_currency_scaled_i64(1_234_567_900_000),
        Variant::from_date_f64(45123.75),
        Variant::from_i32(9),
    ];
    assert_eq!(
        &observed.values[..expected.len()],
        &expected.iter().map(canon).collect::<Vec<_>>()
    );
    assert!(!observed.raised);
    assert_eq!(
        observed.err,
        FinalErr {
            last_dll_error: 4321,
            ..Default::default()
        }
    );
}

#[test]
fn native_scalar_return_classes_and_mixed_register_stack_arguments_match_excel() {
    let fixture = NativeFixture::build();
    let path = fixture.path().display().to_string();
    let observed = compare(&format!(
        r#"
Public result As String
Declare PtrSafe Function NativeCurrency Lib "{path}" Alias "CurrencyValue" () As Currency
Declare PtrSafe Function NativeDate Lib "{path}" Alias "DateValue" () As Date
Declare PtrSafe Function NativeByte Lib "{path}" Alias "ByteValue" () As Byte
Declare PtrSafe Function NativeMixed Lib "{path}" Alias "Mixed" (ByVal a As Long, ByVal b As Double, ByVal c As Single, ByVal d As Currency, ByVal e As Integer, ByVal f As Date, ByVal g As LongPtr, ByVal h As Boolean, ByVal i As Byte, ByVal j As LongLong) As Long
Sub Main()
    Dim d As Date, c As Currency, p As LongPtr, ll As LongLong
    d = CDate(45123.75): c = 1234.5@: p = CLngPtr(4294967338#): ll = CLngLng(-9876543210987#)
    result = CStr(NativeCurrency() = 123456789.0123@) & "|" & CStr(CDbl(NativeDate()) = 45123.75) & "|" & TypeName(NativeByte()) & "|" & CStr(NativeMixed(19, 23.5, 12.5!, c, -2, d, p, True, 251, ll))
End Sub
"#
    ));
    assert_eq!(
        observed.values[0],
        canon(&Variant::from_string("True|True|Byte|42"))
    );
    assert!(!observed.raised);
    assert_eq!(
        observed.err,
        FinalErr {
            last_dll_error: 4321,
            ..Default::default()
        }
    );
}

#[test]
fn unsupported_alias_carriers_fail_before_entering_native_code() {
    let fixture = NativeFixture::build();
    let path = fixture.path().display().to_string();
    let observed = compare(&format!(
        r#"
Public result As Long
Public observedError As Long
Public calls As Long
Declare PtrSafe Function Probe Lib "{path}" Alias "PointerProbe" (ByRef a As Any, ByRef b As Any) As Long
Declare PtrSafe Function NativeCount Lib "{path}" Alias "AliasCallCount" () As Long
Sub Main()
    Dim text As String
    text = "same variable"
    On Error Resume Next
    result = Probe(text, text)
    observedError = Err.Number
    Err.Clear
    calls = NativeCount()
End Sub
"#
    ));
    assert_eq!(
        &observed.values[..3],
        &[0, 5, 0].map(|n| canon(&Variant::from_i32(n)))
    );
    assert!(!observed.raised);
    assert_eq!(
        observed.err,
        FinalErr {
            last_dll_error: 4321,
            ..Default::default()
        }
    );
}

#[test]
fn native_variant_alias_cells_preserve_payload_changes_and_copy_boundaries() {
    let fixture = NativeFixture::build();
    let path = fixture.path().display().to_string();
    for (initial, call, expected) in [
        ("Empty", "Probe(value, value)", 42),
        ("Null", "Probe(value, value)", 42),
        ("CInt(7)", "Probe(value, value)", 42),
        ("CLng(7)", "Probe(value, value)", 42),
        ("CByte(7)", "Probe(value, value)", 42),
        ("CSng(7.5)", "Probe(value, value)", 42),
        ("42.5", "Probe(value, value)", 42),
        ("True", "Probe(value, value)", 42),
        ("CCur(7)", "Probe(value, value)", 42),
        ("CDate(45123.5)", "Probe(value, value)", 42),
        ("CDec(7)", "Probe(value, value)", 42),
        ("CLngLng(7)", "Probe(value, value)", 42),
        ("CVErr(2042)", "Probe(value, value)", 42),
        ("\"owned string\"", "Probe(value, value)", 42),
        ("\"alpha\" & Chr(0) & \"beta\"", "Probe(value, value)", 42),
        ("\"owned string\"", "Forward(value, value)", 42),
        ("\"owned string\"", "Probe((value), value)", -1),
    ] {
        let observed = compare(&format!(
            r#"
Public result As String
Public calls As Long
Declare PtrSafe Function Probe Lib "{path}" Alias "AliasVariant" (ByRef first As Variant, ByRef second As Variant) As Long
Declare PtrSafe Function NativeCount Lib "{path}" Alias "AliasCallCount" () As Long
Private Function Forward(ByRef first As Variant, ByRef second As Variant) As Long
    Forward = Probe(first, second)
End Function
Sub Main()
    Dim value As Variant
    Dim answer As Long, savedError As Long, beforeCalls As Long
    value = {initial}
    beforeCalls = NativeCount()
    answer = {call}
    savedError = Err.LastDllError
    calls = NativeCount() - beforeCalls
    result = CStr(answer) & "|" & TypeName(value) & "|" & CStr(value = 42.5) & "|" & CStr(savedError)
End Sub
"#
        ));
        assert!(!observed.raised, "{initial}: {observed:?}");
        assert_eq!(
            observed.values[0],
            canon(&Variant::from_string(format!(
                "{expected}|Double|True|4321"
            )))
        );
        assert_eq!(observed.values[1], canon(&Variant::from_i32(1)));
        assert_eq!(
            observed.err,
            FinalErr {
                last_dll_error: 4321,
                ..Default::default()
            }
        );
    }
}

#[test]
fn native_variant_alias_readback_owns_returned_bstr_bytes_and_distinct_cells() {
    let fixture = NativeFixture::build();
    let path = fixture.path().display().to_string();
    for (second, expected) in [
        ("first", "42|13|final:payload|4321"),
        ("second", "-1|13|final:payload|4321"),
    ] {
        let observed = compare(&format!(
            r#"
Public result As String
Public firstValue As Variant
Declare PtrSafe Function Probe Lib "{path}" Alias "AliasVariantString" (ByRef first As Variant, ByRef second As Variant) As Long
Sub Main()
    Dim first As Variant, second As Variant
    Dim answer As Long, savedError As Long
    first = "old owned string"
    second = first
    answer = Probe(first, {second})
    savedError = Err.LastDllError
    firstValue = first
    result = CStr(answer) & "|" & CStr(Len({second})) & "|" & Replace({second}, Chr(0), ":") & "|" & CStr(savedError)
End Sub
"#
        ));
        assert!(!observed.raised, "{observed:?}");
        assert_eq!(observed.values[0], canon(&Variant::from_string(expected)));
        let expected_first = if second == "first" {
            "final\0payload"
        } else {
            "first"
        };
        assert_eq!(
            observed.values[1],
            canon(&Variant::from_string(expected_first))
        );
        assert_eq!(
            observed.err,
            FinalErr {
                last_dll_error: 4321,
                ..Default::default()
            }
        );
    }
}

#[test]
fn native_variant_unproved_shapes_still_refuse_aliases_before_callee_execution() {
    let fixture = NativeFixture::build();
    let path = fixture.path().display().to_string();
    // These are explicit remaining gaps, not VBA compatibility expectations.
    for (initial, first_decl, second_decl) in [
        ("value = Array(1, 2)", "Variant", "Variant"),
        (
            "Set value = CreateObject(\"Scripting.Dictionary\")",
            "Variant",
            "Variant",
        ),
        ("value = CInt(7)", "Variant", "Any"),
        ("value = CInt(7)", "Any", "Variant"),
        ("value = CInt(7)", "Any", "Any"),
        ("value = \"owned string\"", "Variant", "Any"),
    ] {
        let observed = compare(&format!(
            r#"
Public result As Long
Public observedError As Long
Public calls As Long
Declare PtrSafe Function Probe Lib "{path}" Alias "PointerProbe" (ByRef first As {first_decl}, ByRef second As {second_decl}) As Long
Declare PtrSafe Function NativeCount Lib "{path}" Alias "AliasCallCount" () As Long
Sub Main()
    Dim value As Variant, beforeCalls As Long
    {initial}
    beforeCalls = NativeCount()
    On Error Resume Next
    result = Probe(value, value)
    observedError = Err.Number
    Err.Clear
    calls = NativeCount() - beforeCalls
End Sub
"#
        ));
        assert!(!observed.raised);
        assert_eq!(
            &observed.values[..3],
            &[0, 5, 0].map(|value| canon(&Variant::from_i32(value)))
        );
    }
}

#[test]
fn getobject_empty_path_activates_a_registered_class_on_both_backends() {
    let observed = compare(
        r#"
Public result As Long
Public count As Long
Sub Main()
    Dim className As String, obj As Object
    className = "Scripting.Dictionary"
    Set obj = GetObject("", className)
    obj.Add "answer", 42
    result = obj.Item("answer")
    count = obj.Count
End Sub
"#,
    );
    assert_eq!(
        &observed.values[..2],
        &[42, 1].map(|value| canon(&Variant::from_i32(value)))
    );
    assert!(!observed.raised);
    assert_eq!(observed.err, FinalErr::default());
    assert_eq!(observed.transport, (0, 3));
}

/// The caller supplies a workbook held by the PID-scoped oracle supervisor.
/// A live owned PID/HWND is required before binding, and checked again in VBA
/// before writes. The supervisor must hold that workbook until this test ends.
#[test]
#[ignore = "requires OXVBA_OWNED_EXCEL_WORKBOOK, OXVBA_OWNED_EXCEL_HWND and OXVBA_OWNED_EXCEL_PID from a live owned oracle run"]
fn getobject_reads_owned_excel_workbook_values_on_both_backends() {
    let path = std::env::var("OXVBA_OWNED_EXCEL_WORKBOOK")
        .expect("explicitly owned workbook path is required");
    assert!(std::path::Path::new(&path).is_file());
    let hwnd: i32 = std::env::var("OXVBA_OWNED_EXCEL_HWND")
        .expect("owned Excel HWND is required")
        .parse()
        .expect("owned Excel HWND must be a VBA Long");
    assert_ne!(hwnd, 0);
    let expected_pid: u32 = std::env::var("OXVBA_OWNED_EXCEL_PID")
        .expect("live supervisor-owned Excel PID is required")
        .parse()
        .expect("owned Excel PID must be numeric");
    assert_ne!(expected_pid, 0);
    let mut actual_pid = 0;
    // SAFETY: the HWND is only queried; actual_pid is a live writable out-slot.
    let window_thread = unsafe {
        windows_sys::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId(
            hwnd as u32 as usize as _,
            &mut actual_pid,
        )
    };
    assert_ne!(
        window_thread, 0,
        "owned window must be live before GetObject"
    );
    assert_eq!(actual_pid, expected_pid, "owned window PID changed");
    let path = path.replace('"', "\"\"");
    let observed = compare(&format!(
        r#"
Public initialValue As Variant
Public changedValue As Variant
Public textValue As Variant
Public truthValue As Variant
Public errorValue As Variant
Public rowLower As Long
Public rowUpper As Long
Public columnLower As Long
Public columnUpper As Long
Public address As String
Public version As String
Public nativeWindow As Long
Sub Main()
    Dim wb As Object, app As Object, sheet As Object, cell As Object
    Dim matrix As Variant
    Set wb = GetObject("{path}")
    Set app = wb.Application
    nativeWindow = app.Hwnd
    If nativeWindow <> {hwnd} Then Err.Raise 5
    Set sheet = wb.Worksheets.Item(1)
    Set cell = sheet.Range("A1")
    initialValue = cell.Value2
    cell.Value2 = 19
    changedValue = cell.Value2
    cell.Value2 = 42
    address = cell.Address(False, False)
    Set cell = sheet.Range("A1:B2")
    matrix = cell.Value2
    textValue = matrix(1, 2)
    truthValue = matrix(2, 1)
    errorValue = matrix(2, 2)
    Set cell = sheet.Range("B2")
    cell.Value2 = errorValue
    errorValue = cell.Value2
    rowLower = LBound(matrix, 1): rowUpper = UBound(matrix, 1)
    columnLower = LBound(matrix, 2): columnUpper = UBound(matrix, 2)
    version = app.Version
End Sub
"#
    ));
    let expected = [
        Variant::from_f64(42.0),
        Variant::from_f64(19.0),
        Variant::from_string("alpha"),
        Variant::from_bool(true),
        Variant::from_error_code(2042),
        Variant::from_i32(1),
        Variant::from_i32(2),
        Variant::from_i32(1),
        Variant::from_i32(2),
        Variant::from_string("A1"),
        Variant::from_string("16.0"),
        Variant::from_i32(hwnd),
    ];
    assert_eq!(
        &observed.values[..expected.len()],
        &expected.iter().map(canon).collect::<Vec<_>>()
    );
    assert!(!observed.raised);
    assert_eq!(observed.err, FinalErr::default());
    assert_eq!(observed.transport.0, 0);
    assert!(
        observed.transport.1 >= 13,
        "actual IDispatch calls required"
    );
}

#[test]
fn native_string_aliases_use_argument_temporaries_and_ordered_copyback() {
    let fixture = NativeFixture::build();
    let path = fixture.path().display().to_string();
    for (mode, export, call) in [
        ("ByRef", "ObserveAnsiRef", "Probe(value, value)"),
        ("ByRef", "ObserveAnsiRef", "Probe((value), value)"),
        ("ByRef", "ObserveAnsiRef", "Forward(value, value)"),
        ("ByVal", "ObserveAnsiVal", "Probe(value, value)"),
    ] {
        let observed = compare(&format!(
            r#"
Public result As String
Declare PtrSafe Function Probe Lib "{path}" Alias "{export}" ({mode} first As String, {mode} second As String) As Long
Private Function Forward(ByRef first As String, ByRef second As String) As Long
    Forward = Probe(first, second)
End Function
Sub Main()
    Dim value As String
    Dim answer As Long, savedError As Long
    value = "abcd"
    answer = {call}
    savedError = Err.LastDllError
    result = CStr(answer) & "|" & value & "|" & CStr(savedError)
End Sub
"#
        ));
        assert!(!observed.raised, "{mode}/{call}: {observed:?}");
        assert_eq!(
            observed.values[0],
            canon(&Variant::from_string("0|aYcd|4321"))
        );
        assert_eq!(
            observed.err,
            FinalErr {
                last_dll_error: 4321,
                ..Default::default()
            }
        );
    }
}

#[test]
fn native_string_alias_copyback_preserves_distinct_values_and_embedded_nul() {
    let fixture = NativeFixture::build();
    let path = fixture.path().display().to_string();
    for (body, expected, raw_expected) in [
        (
            r#"
    first = "abcd"
    second = first
    answer = Probe(first, second)
    savedError = Err.LastDllError
    bytesValue = second
    result = CStr(answer) & "|" & first & "|" & second & "|" & CStr(savedError)
"#,
            "0|Xbcd|aYcd|4321",
            "aYcd",
        ),
        (
            r#"
    value = "ab" & Chr(0) & "d"
    answer = Probe(value, value)
    savedError = Err.LastDllError
    bytesValue = value
    result = CStr(answer) & "|" & CStr(Len(value)) & "|" & Replace(value, Chr(0), ":") & "|" & CStr(savedError)
"#,
            "0|4|aY:d|4321",
            "aY\0d",
        ),
    ] {
        let observed = compare(&format!(
            r#"
Public result As String
Public bytesValue As String
Declare PtrSafe Function Probe Lib "{path}" Alias "ObserveAnsiRef" (ByRef first As String, ByRef second As String) As Long
Sub Main()
    Dim first As String, second As String, value As String
    Dim answer As Long, savedError As Long
    {body}
End Sub
"#
        ));
        assert!(!observed.raised, "{observed:?}");
        assert_eq!(observed.values[0], canon(&Variant::from_string(expected)));
        assert_eq!(
            observed.values[1],
            canon(&Variant::from_string(raw_expected))
        );
        assert_eq!(
            observed.err,
            FinalErr {
                last_dll_error: 4321,
                ..Default::default()
            }
        );
    }
}

#[test]
fn mixed_native_string_any_aliases_still_refuse_before_entry() {
    let fixture = NativeFixture::build();
    let path = fixture.path().display().to_string();
    // Protective refusals remain compatibility gaps with explicit owners.
    for (first_decl, second_decl) in [("String", "Any"), ("Any", "String")] {
        let observed = compare(&format!(
            r#"
Public result As Long
Public observedError As Long
Public calls As Long
Declare PtrSafe Function Probe Lib "{path}" Alias "PointerProbe" (ByRef first As {first_decl}, ByRef second As {second_decl}) As Long
Declare PtrSafe Function NativeCount Lib "{path}" Alias "AliasCallCount" () As Long
Sub Main()
    Dim value As String, beforeCalls As Long
    value = "abcd"
    beforeCalls = NativeCount()
    On Error Resume Next
    result = Probe(value, value)
    observedError = Err.Number
    Err.Clear
    calls = NativeCount() - beforeCalls
End Sub
"#
        ));
        assert!(!observed.raised);
        assert_eq!(
            &observed.values[..3],
            &[0, 5, 0].map(|value| canon(&Variant::from_i32(value)))
        );
    }
}
