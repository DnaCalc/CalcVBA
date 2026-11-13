//! Automation VARIANT layout and ByRef Variant binding against Excel VBA (x64), for dc-yqq0.1.
//!
//! The probe module `conformance/excel-measurements/variant-layout.vba` records one `id=value`
//! line per case: `VarPtr` strides, cell bytes per subtype, `SAFEARRAY` headers, `CopyMemory`
//! round-trips, writes through `VarPtr`, UDT sizes and offsets, and how a ByRef Variant parameter
//! binds each kind of argument (`VT_BYREF` cells). `variant-layout.excel.txt` is Excel's answer,
//! captured with `scripts/run-excel-probe-session.ps1` on 2026-11-13.
//! Cases an engine still answers differently are listed in `variant-layout.known-differences.txt`
//! (`<executor>\t<id>`; `DECLINED` when the engine refuses the whole module). The test fails on
//! any new difference and on any listed difference that is now fixed.
#![cfg(all(windows, target_pointer_width = "64"))]

use calcvba_differential::{Canon, Executor, run};
use std::collections::{BTreeMap, BTreeSet};

const MODULE: &str = include_str!("../../../conformance/excel-measurements/variant-layout.vba");
const EXCEL: &str =
    include_str!("../../../conformance/excel-measurements/variant-layout.excel.txt");
const KNOWN: &str =
    include_str!("../../../conformance/excel-measurements/variant-layout.known-differences.txt");

fn lines(text: &str) -> BTreeMap<String, String> {
    text.lines()
        .filter_map(|l| l.split_once('='))
        .map(|(k, v)| (k.to_string(), v.to_string()))
        .collect()
}

/// Every difference id for this executor; `DECLINED` when the engine does not run the module.
fn differences(executor: Executor) -> BTreeMap<String, String> {
    let source = format!(
        "Public result As String\r\n{MODULE}\r\nSub Main()\r\n  result = RunAll()\r\nEnd Sub\r\n"
    );
    let outcome = run(executor, &source);
    let declined = |why: String| BTreeMap::from([("DECLINED".to_string(), why)]);
    if let Some(why) = outcome.unsupported {
        return declined(format!("unsupported: {why}"));
    }
    let values = match outcome.result {
        Ok(values) => values,
        Err(e) => return declined(format!("failed: {e}")),
    };
    let Some(got) = values.into_iter().find_map(|c| match c {
        Canon::Str(s) if s.contains("STRIDE.Variant=") => Some(s),
        _ => None,
    }) else {
        return declined("no result string".to_string());
    };
    let got = lines(&got.replace("\r\n", "\n"));
    lines(EXCEL)
        .into_iter()
        .filter(|(k, v)| got.get(k) != Some(v))
        .map(|(k, v)| {
            let engine = got.get(&k).cloned().unwrap_or_else(|| "<missing>".into());
            (k, format!("excel={v}\tengine={engine}"))
        })
        .collect()
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
    let diffs = differences(executor);
    if let Ok(path) = std::env::var("VARIANT_LAYOUT_DUMP") {
        let ids: String = diffs
            .keys()
            .map(|k| format!("{executor:?}\t{k}\n"))
            .collect();
        let detail: String = diffs
            .iter()
            .map(|(k, d)| format!("{executor:?}\t{k}\t{d}\n"))
            .collect();
        let _ = std::fs::write(format!("{path}.{executor:?}.ids"), ids);
        let _ = std::fs::write(format!("{path}.{executor:?}.detail"), detail);
    }
    let ids: BTreeSet<String> = diffs.keys().cloned().collect();
    let known = known(executor);
    let new: Vec<_> = diffs.iter().filter(|(k, _)| !known.contains(*k)).collect();
    let fixed: Vec<_> = known.difference(&ids).collect();
    assert!(
        new.is_empty() && fixed.is_empty(),
        "{executor:?}: {} differences from Excel ({} listed)\nnew: {new:#?}\nlisted but now fixed: {fixed:#?}",
        ids.len(),
        known.len()
    );
}

#[test]
fn variant_layout_vm3_matches_excel() {
    check(Executor::Vm3);
}

#[test]
fn variant_layout_jit_matches_excel() {
    check(Executor::Jit);
}
