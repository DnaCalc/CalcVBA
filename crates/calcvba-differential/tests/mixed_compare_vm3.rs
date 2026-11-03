//! VBA comparison depends on declared types as well as Variant subtypes.
//! Numeric Variants sort before String Variants. Numeric comparison tries conversion
//! and otherwise retains subtype order; assignment coercion remains strict. The exact former cases are retained in the owned
//! comparisons.xlsm probe; see Microsoft's VBA comparison-operators reference.

use calcvba_differential::{Canon, Executor, RunOutcome, canon, run};
use calcvba_runtime::Variant;

fn run_main(executor: Executor, body: &str) -> RunOutcome {
    let source = format!("Public r As Variant\nSub Main()\n{body}End Sub\n");
    run(executor, &source)
}

fn assert_raises_13(body: &str) {
    for executor in [Executor::Vm3, Executor::Jit] {
        let outcome = run_main(executor, body);
        assert!(
            outcome.unsupported.is_none(),
            "unsupported: {:?}",
            outcome.unsupported
        );
        assert!(
            outcome.result.is_err(),
            "expected Type mismatch, completed: {:?}",
            outcome.result
        );
        assert_eq!(outcome.err.number, 13, "err={:?}", outcome.err);
    }
}

fn assert_first(body: &str, expected: &Canon) {
    for executor in [Executor::Vm3, Executor::Jit] {
        let outcome = run_main(executor, body);
        assert!(
            outcome.unsupported.is_none(),
            "unsupported: {:?}",
            outcome.unsupported
        );
        let snap = outcome.result.unwrap_or_else(|e| panic!("run failed: {e}"));
        assert_eq!(snap.first(), Some(expected), "{snap:?}");
        assert_eq!(outcome.err.number, 0);
    }
}

#[test]
fn variant_string_and_number_equality_is_false() {
    assert_first(
        "    Dim s As Variant\n    Dim n As Variant\n    s = \"abc\"\n    n = 5\n    r = (s = n)\n",
        &canon(&Variant::from_bool(false)),
    );
}

#[test]
fn variant_numeric_string_sorts_after_variant_number() {
    assert_first(
        "    Dim s As Variant\n    Dim n As Variant\n    s = \"5\"\n    n = 5\n    r = (s < n)\n",
        &canon(&Variant::from_bool(false)),
    );
}

#[test]
fn declared_number_assignment_rejects_nonnumeric_string_variant() {
    assert_raises_13(
        "    Dim s As Variant\n    Dim n As Long\n    s = \"abc\"\n    n = 5\n    n = s\n",
    );
}

#[test]
fn empty_vs_number_is_exempt() {
    // Empty adapts to 0, so `Empty = 0` is True (no Type mismatch).
    assert_first(
        "    Dim e As Variant\n    Dim n As Variant\n    n = 0\n    r = (e = n)\n",
        &canon(&Variant::from_bool(true)),
    );
}

#[test]
fn same_kind_comparisons_still_work() {
    // String vs String compares lexically; number vs number compares numerically.
    assert_first(
        "    Dim a As Variant\n    Dim b As Variant\n    a = \"abc\"\n    b = \"abc\"\n    r = (a = b)\n",
        &canon(&Variant::from_bool(true)),
    );
    assert_first(
        "    Dim a As Variant\n    Dim b As Variant\n    a = 5\n    b = 5\n    r = (a = b)\n",
        &canon(&Variant::from_bool(true)),
    );
}

#[test]
fn declared_number_comparison_keeps_nonnumeric_variant_order() {
    assert_first(
        "    Dim s As Variant\n    Dim n As Long\n    s = \"!\"\n    n = 5\n    r = (n < s)\n",
        &canon(&Variant::from_bool(true)),
    );
}
