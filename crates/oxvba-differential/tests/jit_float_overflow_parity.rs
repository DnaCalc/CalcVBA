//! Floating arithmetic regressions and scoped typed-callee error transport.
//! Generic Double fault/store ordering remains CORE-3 bd-59co.2.4.3.

use oxvba_differential::{Executor, run};

fn assert_jit_matches_vm3(source: &str) {
    let vm3 = run(Executor::Vm3, source);
    assert!(
        vm3.unsupported.is_none(),
        "vm3 declined: {:?}\n{source}",
        vm3.unsupported
    );
    let jit = run(Executor::Jit, source);
    assert!(
        jit.unsupported.is_none(),
        "jit declined: {:?}\n{source}",
        jit.unsupported
    );
    assert_eq!(jit.raised, vm3.raised, "raised mismatch\n{source}");
    assert_eq!(jit.err, vm3.err, "Err mismatch\n{source}");
    assert_eq!(jit.result, vm3.result, "snapshot mismatch\n{source}");
}

/// Both executors raise VBA Overflow (error 6) for `source`.
fn assert_both_overflow(source: &str) {
    for exec in [Executor::Vm3, Executor::Jit] {
        let outcome = run(exec, source);
        assert!(
            outcome.unsupported.is_none(),
            "{exec:?} declined: {:?}\n{source}",
            outcome.unsupported
        );
        assert!(outcome.raised, "{exec:?} did not raise\n{source}");
        assert_eq!(outcome.err.number, 6, "{exec:?} error != 6\n{source}");
    }
}

fn single(body: &str) -> String {
    format!(
        "Public r As Variant\nSub Main()\n    Dim a As Single\n    Dim b As Single\n{body}    r = b\nEnd Sub\n"
    )
}

fn double(body: &str) -> String {
    format!(
        "Public r As Variant\nSub Main()\n    Dim a As Double\n    Dim b As Double\n{body}    r = b\nEnd Sub\n"
    )
}

#[test]
fn jit_single_overflow_matches_vm3() {
    for src in [
        single("    a = 2E38\n    b = a + a\n"),
        single("    a = -2E38\n    b = a - 2E38\n"),
        single("    a = 2E19\n    b = a * a\n"),
    ] {
        assert_jit_matches_vm3(&src);
        assert_both_overflow(&src);
    }
}

#[test]
fn jit_double_overflow_matches_vm3() {
    // This compares the existing generic route only. Excel raises error 6 and
    // preserves the overflowed assignment under Resume Next; the full store/
    // fault-order correction remains CORE-3 bd-59co.2.4.3. The admitted typed
    // scalar family separately proves unhandled overflow on both backends.
    for src in [
        double("    a = 1E308\n    b = a + a\n"),
        double("    a = 1E308\n    b = a * 10\n"),
    ] {
        assert_jit_matches_vm3(&src);
    }
}

#[test]
fn jit_finite_float_arithmetic_still_matches_vm3() {
    // Regression: non-overflowing arithmetic is unchanged and produces no error.
    assert_jit_matches_vm3(&single("    a = 1.5\n    b = a + 2.5\n"));
    assert_jit_matches_vm3(&single("    a = 100\n    b = a * 3\n"));
    assert_jit_matches_vm3(&double("    a = 1E100\n    b = a + a\n"));
    assert_jit_matches_vm3(&double("    a = 12.5\n    b = a - 4.5\n"));
}

#[test]
fn admitted_scalar_callee_overflow_reaches_caller_handler() {
    let source = "Public r As Variant\nFunction Multiply(ByVal a As Double, ByVal b As Double) As Double\nMultiply = a * b\nEnd Function\nSub Main()\nOn Error Resume Next\nr = Multiply(1E308, 2#)\nr = Err.Number\nEnd Sub\n";
    assert_jit_matches_vm3(source);
    for executor in [Executor::Vm3, Executor::Jit] {
        let outcome = run(executor, source);
        assert!(!outcome.raised, "caller must handle overflow: {outcome:?}");
        assert_eq!(
            outcome.err.number, 6,
            "{executor:?} must retain callee error"
        );
    }
}
