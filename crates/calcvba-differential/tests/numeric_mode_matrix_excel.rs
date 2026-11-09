//! Compares the engines with the committed Excel VBA NumericMode decision matrix
//! (`conformance/excel-measurements/numeric-mode-matrix.excel.txt`, Excel 16.0 build 20527).
//!
//! The probe module (`numeric-mode-matrix.vba`, produced by `gen_numeric_mode_matrix.py`) records one
//! `id=TypeName:value` or `id=ERR<n>` line per case. Cases the engines still answer differently are
//! listed in `numeric-mode-matrix.known-differences.txt` (one `id` per line, per executor); the test
//! fails on any new difference and on any listed difference that is now fixed.

use calcvba_differential::{Canon, Executor, run};
use std::collections::{BTreeMap, BTreeSet};

const MODULE: &str =
    include_str!("../../../conformance/excel-measurements/numeric-mode-matrix.vba");
const EXCEL: &str =
    include_str!("../../../conformance/excel-measurements/numeric-mode-matrix.excel.txt");
const KNOWN: &str = include_str!(
    "../../../conformance/excel-measurements/numeric-mode-matrix.known-differences.txt"
);

fn lines(text: &str) -> BTreeMap<String, String> {
    text.lines()
        .filter_map(|l| l.split_once('='))
        .map(|(k, v)| (k.to_string(), v.to_string()))
        .collect()
}

fn engine_output(executor: Executor) -> Option<String> {
    let source = format!(
        "Public result As String\r\n{MODULE}\r\nSub Main()\r\n  result = RunAll()\r\nEnd Sub\r\n"
    );
    let outcome = run(executor, &source);
    if let Some(why) = &outcome.unsupported {
        eprintln!("{executor:?} declined the matrix: {why}");
        return None;
    }
    let values = outcome
        .result
        .unwrap_or_else(|e| panic!("{executor:?} failed to run the matrix: {e}"));
    values.into_iter().find_map(|c| match c {
        Canon::Str(s) if s.contains("bin.Byte0.add.Byte0=") => Some(s),
        _ => None,
    })
}

fn differences(executor: Executor) -> Option<BTreeSet<String>> {
    let got = lines(&engine_output(executor)?.replace("\r\n", "\n"));
    let want = lines(EXCEL);
    assert_eq!(
        got.keys().collect::<Vec<_>>(),
        want.keys().collect::<Vec<_>>(),
        "{executor:?} produced a different set of case ids"
    );
    Some(
        want.iter()
            .filter(|(k, v)| got[*k] != **v)
            .map(|(k, v)| format!("{k}\texcel={v}\tengine={}", got[k]))
            .collect(),
    )
}

fn known(executor: Executor) -> BTreeSet<String> {
    let tag = format!("{executor:?}");
    KNOWN
        .lines()
        .filter(|l| !l.trim().is_empty() && !l.starts_with('#'))
        .filter_map(|l| l.split_once('\t'))
        .filter(|(t, _)| *t == tag)
        .map(|(_, id)| id.to_string())
        .collect()
}

fn check(executor: Executor) {
    let Some(diffs) = differences(executor) else {
        return;
    };
    if let Ok(path) = std::env::var("NUMERIC_MODE_MATRIX_DUMP") {
        let body: String = diffs
            .iter()
            .map(|d| format!("{executor:?}\t{}\n", d.split('\t').next().unwrap()))
            .collect();
        let detail: String = diffs
            .iter()
            .map(|d| format!("{executor:?}\t{d}\n"))
            .collect();
        let _ = std::fs::write(format!("{path}.{executor:?}.ids"), body);
        let _ = std::fs::write(format!("{path}.{executor:?}.detail"), detail);
    }
    let ids: BTreeSet<String> = diffs
        .iter()
        .map(|d| d.split('\t').next().unwrap().to_string())
        .collect();
    let known = known(executor);
    let new: Vec<_> = diffs
        .iter()
        .filter(|d| !known.contains(d.split('\t').next().unwrap()))
        .take(40)
        .collect();
    let fixed: Vec<_> = known.difference(&ids).take(40).collect();
    assert!(
        new.is_empty() && fixed.is_empty(),
        "{executor:?}: {} differences from Excel ({} listed)\nnew: {new:#?}\nlisted but now fixed: {fixed:#?}",
        ids.len(),
        known.len()
    );
}

#[test]
fn numeric_mode_matrix_vm3_matches_excel() {
    check(Executor::Vm3);
}

#[test]
fn numeric_mode_matrix_jit_matches_excel() {
    check(Executor::Jit);
}
