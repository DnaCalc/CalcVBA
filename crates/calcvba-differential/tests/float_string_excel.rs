//! Double and Single to String (`CStr`, `Str$`, implicit `&`) against Excel VBA
//! (`conformance/excel-measurements/double-string*.excel.txt`, Excel 16.0 build 20527).
//!
//! Office prints at most 15 (Double) or 7 (Single) significant digits, plain when the integer part
//! and the fraction each fit in that many digits and as `d.dddE+XX` otherwise; negative zero keeps
//! its sign.

use calcvba_differential::{Canon, Executor, run};
use std::collections::BTreeMap;

fn lines(text: &str) -> BTreeMap<String, String> {
    text.lines()
        .filter_map(|l| l.split_once('='))
        .map(|(k, v)| (k.to_string(), v.to_string()))
        .collect()
}

fn compare(executor: Executor, module: &str, excel: &str) {
    let source = format!(
        "Public result As String\r\n{module}\r\nSub Main()\r\n  result = RunAll()\r\nEnd Sub\r\n"
    );
    let outcome = run(executor, &source);
    if outcome.unsupported.is_some() {
        return;
    }
    let values = outcome
        .result
        .unwrap_or_else(|e| panic!("{executor:?} failed: {e}"));
    let text = values
        .into_iter()
        .find_map(|c| match c {
            Canon::Str(s) if s.contains('=') => Some(s),
            _ => None,
        })
        .expect("probe output");
    let got = lines(&text.replace("\r\n", "\n"));
    let want = lines(excel);
    let wrong: Vec<String> = want
        .iter()
        .filter(|(k, v)| got.get(*k) != Some(*v))
        .map(|(k, v)| format!("{k}: excel {v:?} engine {:?}", got.get(k)))
        .collect();
    assert!(
        wrong.is_empty(),
        "{executor:?}: {} of {} lines differ from Excel:\n{}",
        wrong.len(),
        want.len(),
        wrong.join("\n")
    );
}

const LIST: &str = include_str!("../../../conformance/excel-measurements/double-string.vba");
const LIST_EXCEL: &str =
    include_str!("../../../conformance/excel-measurements/double-string.excel.txt");
const BOUNDS: &str =
    include_str!("../../../conformance/excel-measurements/double-string-bounds.vba");
const BOUNDS_EXCEL: &str =
    include_str!("../../../conformance/excel-measurements/double-string-bounds.excel.txt");

#[test]
fn float_string_list_vm3_matches_excel() {
    compare(Executor::Vm3, LIST, LIST_EXCEL);
}

#[test]
fn float_string_bounds_vm3_matches_excel() {
    compare(Executor::Vm3, BOUNDS, BOUNDS_EXCEL);
}

#[test]
fn float_string_list_jit_matches_excel() {
    compare(Executor::Jit, LIST, LIST_EXCEL);
}

#[test]
fn float_string_bounds_jit_matches_excel() {
    compare(Executor::Jit, BOUNDS, BOUNDS_EXCEL);
}
