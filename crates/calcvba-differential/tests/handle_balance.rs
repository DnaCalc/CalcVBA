use calcvba_differential::{Executor, RunOutcome, canon, run, run_modules, run_with_project};
use calcvba_runtime::{Variant, live_handle_counts};
use calcvba_symbol::manifest::ModuleKind::{Class, Procedural};
use std::sync::Mutex;
static RUN: Mutex<()> = Mutex::new(());
const POLICY_ERROR_BALANCE_FIXTURE: &str = "host-policy-error";
const ALL_BALANCE_FIXTURES: &[&str] = &[
    "carrier-string",
    "carrier-array",
    "carrier-object",
    "carrier-record",
    POLICY_ERROR_BALANCE_FIXTURE,
];
fn assert_balanced(outcome: &RunOutcome) {
    assert!(outcome.unsupported.is_none(), "{outcome:?}");
    assert!(
        outcome
            .handle_balance
            .is_some_and(|balance| balance.is_zero()),
        "{outcome:?}"
    );
}
#[test]
fn carrier_results_and_ownership_match_on_both_runtimes() {
    let _run = RUN.lock().unwrap();
    let before = live_handle_counts();
    for (fixture, expected) in [
        ("carrier-string", Variant::from_string("alpha-beta")),
        ("carrier-array", Variant::from_i32(60)),
        ("carrier-object", Variant::from_i32(42)),
        ("carrier-record", Variant::from_i32(41)),
    ] {
        let expected = canon(&expected);
        let vm3 = run_named_fixture(Executor::Vm3, fixture).unwrap();
        let jit = run_named_fixture(Executor::Jit, fixture).unwrap();
        for outcome in [&vm3, &jit] {
            assert_balanced(outcome);
            assert!(!outcome.raised, "{fixture}: {outcome:?}");
            assert_eq!(outcome.err, Default::default());
            assert!(
                outcome.result.as_ref().unwrap().contains(&expected),
                "{fixture}: {outcome:?}"
            );
        }
        assert_eq!(jit.result, vm3.result, "{fixture}");
        assert_eq!(jit.err, vm3.err, "{fixture}");
    }
    assert!(before.balance_to(live_handle_counts()).is_zero());
}
#[test]
fn vm3_denied_pointer_call_preserves_full_err_and_releases_strings() {
    let _run = RUN.lock().unwrap();
    let before = live_handle_counts();
    for _ in 0..8 {
        let outcome = run_named_fixture(Executor::Vm3, POLICY_ERROR_BALANCE_FIXTURE).unwrap();
        assert_denied(&outcome);
    }
    assert!(before.balance_to(live_handle_counts()).is_zero());
}
#[test]
fn denied_scalar_native_call_matches_vm3_and_jit() {
    let _run = RUN.lock().unwrap();
    let before = live_handle_counts();
    let source = r#"
Private Declare PtrSafe Function VarR8FromI4 Lib "oleaut32" (ByVal inputValue As Long, ByRef outValue As Double) As Long
Sub Main()
    Dim retained As String
    Dim numericOut As Double
    Dim status As Long
    retained = "alpha"
    status = VarR8FromI4(123, numericOut)
End Sub
"#;
    for _ in 0..4 {
        let vm3 = run_with_project(Executor::Vm3, source, "VBAProject");
        let jit = run_with_project(Executor::Jit, source, "VBAProject");
        assert_denied(&vm3);
        assert_denied(&jit);
        assert_eq!(jit.err, vm3.err);
        assert_eq!(jit.result, vm3.result);
    }
    assert!(before.balance_to(live_handle_counts()).is_zero());
}
fn assert_denied(outcome: &RunOutcome) {
    assert_balanced(outcome);
    assert!(outcome.raised, "{outcome:?}");
    assert_eq!(outcome.err.number, 5);
    assert_eq!(outcome.err.source, "VBAProject");
    assert_eq!(outcome.err.description, "operation blocked by host policy");
    assert_eq!(outcome.err.last_dll_error, 0);
    assert_eq!(outcome.result, Err("VBA error 5".into()));
}

fn run_named_fixture(executor: Executor, fixture: &str) -> Result<RunOutcome, String> {
    match fixture {
        "carrier-string" => Ok(run(
            executor,
            r#"
Public gResult As String

Sub Main()
    Dim value As String
    value = "alpha" & "-beta"
    gResult = value
End Sub
"#,
        )),
        "carrier-array" => Ok(run(
            executor,
            r#"
Public gResult As Long

Sub Main()
    Dim values() As Long
    ReDim values(1 To 3)
    values(1) = 10
    values(2) = 20
    values(3) = 30
    gResult = values(1) + values(2) + values(3)
End Sub
"#,
        )),
        "carrier-object" => Ok(run_modules(
            executor,
            &[
                (
                    "Main",
                    Procedural,
                    r#"
Public gResult As Long

Sub Main()
    Dim value As Box
    Set value = New Box
    value.Number = 41
    gResult = value.Number + 1
    Set value = Nothing
End Sub
"#,
                ),
                (
                    "Box",
                    Class,
                    r#"
Public Number As Long
"#,
                ),
            ],
            "BalanceFixture",
        )),
        "carrier-record" => Ok(run(
            executor,
            r#"
Private Type Pair
    Number As Long
    Text As String
End Type

Public gResult As Long

Sub Main()
    Dim value As Pair
    value.Number = 35
    value.Text = "record"
    gResult = value.Number + Len(value.Text)
End Sub
"#,
        )),
        POLICY_ERROR_BALANCE_FIXTURE => Ok(run_with_project(
            executor,
            include_str!(
                "../../../conformance/jit_v2/tracer_bullets/\
                 tb08_native_declare_shared_abi.bas"
            ),
            "VBAProject",
        )),
        _ => Err(format!(
            "unknown fixture `{fixture}`; known fixtures: {}",
            ALL_BALANCE_FIXTURES.join(", ")
        )),
    }
}
