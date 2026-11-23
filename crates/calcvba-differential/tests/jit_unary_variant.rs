use calcvba_differential::{Executor, run};

#[test]
fn jit_unary_variant_library_calls_and_noninteger_not_match_vm3() {
    let mut differences = Vec::new();
    for expression in [
        "Abs(value)",
        "CVErr(value)",
        "Hex(value)",
        "Oct(value)",
        "Str(value)",
        "Not 1.5",
        "Not \"2\"",
        "Not 1.5@",
    ] {
        let source = format!(
            "Public result As Variant\nSub Main()\n    Dim value As Variant\n    value = 17\n    result = {expression}\nEnd Sub\n"
        );
        let vm3 = run(Executor::Vm3, &source);
        let expected = vm3
            .result
            .unwrap_or_else(|err| panic!("VM3 could not run {expression}: {err}"));
        assert!(!expected.is_empty(), "VM3 did not run {expression}");
        let jit = run(Executor::Jit, &source);
        if let Some(reason) = jit.unsupported {
            differences.push(format!("{expression}: JIT declined: {reason}"));
        } else if jit.raised || jit.err != vm3.err || jit.result.as_ref() != Ok(&expected) {
            differences.push(format!("{expression}: VM3 {expected:?}, JIT {jit:?}"));
        }
    }
    assert!(differences.is_empty(), "{differences:#?}");
}
