//! Compile-time acceptance of numeric assignments and calls, against Excel's VBE Debug > Compile
//! (`conformance/excel-measurements/numeric-mode-compile.excel.txt`, Excel 16.0 build 20527).
//!
//! Office rejects every implicit LongLong -> Byte/Integer/Long/Currency store and argument as a
//! compile-time type mismatch, and ByRef arguments of a different numeric type as a ByRef
//! argument type mismatch. Cases the engines still decide differently are listed in `KNOWN`.

use calcvba_differential::{Executor, run};

const EXCEL: &str =
    include_str!("../../../conformance/excel-measurements/numeric-mode-compile.excel.txt");

/// Case ids the engine still decides differently from Excel.
const KNOWN: &[&str] = &[
    // Abs(LongLong) and CLngLng(..) are typed LongLong in Office; the engine types them Variant.
    "asgq.Byte.abs",
    "asgq.Byte.clnglng",
    "asgq.Currency.abs",
    "asgq.Currency.clnglng",
    "asgq.Integer.abs",
    "asgq.Integer.clnglng",
    "asgq.Long.abs",
    "asgq.Long.clnglng",
];

struct Case {
    id: String,
    expected_compile_error: bool,
    source: String,
}

fn cases() -> Vec<Case> {
    let mut out: Vec<Case> = Vec::new();
    for line in EXCEL.lines() {
        if let Some(head) = line.strip_prefix("### ") {
            let mut parts = head.split('\t');
            let id = parts.next().unwrap().to_string();
            let status = parts.next().unwrap();
            out.push(Case {
                id,
                expected_compile_error: status == "compile-error",
                source: String::new(),
            });
        } else if !line.starts_with('#')
            && let Some(case) = out.last_mut()
        {
            case.source.push_str(line);
            case.source.push('\n');
        }
    }
    out
}

fn engine_rejects(executor: Executor, case: &Case) -> Option<bool> {
    let source = format!(
        "Public result As String\n{}\nSub Main()\n  result = Run()\nEnd Sub\n",
        case.source
    );
    let outcome = run(executor, &source);
    if outcome.unsupported.is_some() {
        return None;
    }
    Some(outcome.result.is_err() && !outcome.raised)
}

fn check(executor: Executor) {
    let mut wrong = Vec::new();
    let mut checked = 0;
    for case in cases() {
        let Some(rejected) = engine_rejects(executor, &case) else {
            continue;
        };
        checked += 1;
        if rejected != case.expected_compile_error {
            wrong.push(case.id.clone());
        }
    }
    assert!(checked > 80, "{executor:?} decided only {checked} cases");
    let mut known: Vec<String> = KNOWN.iter().map(|s| s.to_string()).collect();
    known.sort();
    wrong.sort();
    assert_eq!(
        wrong, known,
        "{executor:?}: cases decided differently from Excel (left) versus the listed ones (right)"
    );
}

#[test]
fn numeric_compile_decisions_vm3_match_excel() {
    check(Executor::Vm3);
}

#[test]
fn numeric_compile_decisions_jit_match_excel() {
    check(Executor::Jit);
}
