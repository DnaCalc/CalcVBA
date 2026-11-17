//! `Round(Null)` is `Null` and `Sgn(Null)` raises "Invalid use of Null" (94), as Excel answers.
//! `Sgn` coerces its argument to a number (and coercing `Null` is error 94); `Round` keeps the
//! Variant `Null` like `Abs`/`Int`/`Fix` (see `abs_int_fix_sgn_vm3`).

use calcvba_differential::{Executor, canon, run};
use calcvba_runtime::Variant;

fn error_number(expr: &str) -> i32 {
    let source = format!("Public r As Variant\nSub Main()\n    r = {expr}\nEnd Sub\n");
    let outcome = run(Executor::Vm3, &source);
    assert!(
        outcome.unsupported.is_none(),
        "unsupported: {:?}",
        outcome.unsupported
    );
    assert!(
        outcome.raised,
        "expected a VBA error for {expr}, got {:?}",
        outcome.result
    );
    outcome.err.number
}

#[test]
fn round_of_null_is_null_and_sgn_of_null_raises_94() {
    // Excel: `Round(Null)` is Null (the argument keeps its Variant subtype); `Sgn(Null)` is 94.
    for expr in ["Round(Null)", "Round(Null, 2)"] {
        let source = format!("Public r As Variant\nSub Main()\n    r = {expr}\nEnd Sub\n");
        let outcome = run(Executor::Vm3, &source);
        assert!(outcome.unsupported.is_none(), "{expr}");
        let values = outcome.result.expect("run");
        assert_eq!(values[0], canon(&Variant::null()), "{expr}");
    }
    assert_eq!(error_number("Sgn(Null)"), 94);
}
