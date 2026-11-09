//! Late-bound conversions against Excel VBA (16.0 build 20527, owned session):
//! `Null` to a ByVal scalar parameter is error 94 when early bound and error 13 when late bound
//! (`Object` call or `CallByName`), for every scalar type including Boolean; and `CallByName`
//! with `VbLet` converts text, Integer and Double values to every numeric property type.

use calcvba_differential::{Canon, Executor, run_modules};
use calcvba_symbol::manifest::ModuleKind;
use std::collections::BTreeMap;

fn lines(text: &str) -> BTreeMap<String, String> {
    text.lines()
        .filter_map(|l| l.split_once('='))
        .map(|(k, v)| (k.to_string(), v.to_string()))
        .collect()
}

fn check(modules: &[(&str, ModuleKind, &str)], excel: &str) {
    for executor in [Executor::Vm3, Executor::Jit] {
        let outcome = run_modules(executor, modules, "VBAProject");
        if outcome.unsupported.is_some() {
            continue;
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
        // Date results print in the owner's regional format and Excel rejects text for a Date
        // property; those rows are measurements of the machine locale, not compared here.
        let wrong: Vec<String> = lines(excel)
            .iter()
            .filter(|(k, _)| !k.starts_with("Dt."))
            .filter(|(k, v)| got.get(*k) != Some(*v))
            .map(|(k, v)| format!("{k}: excel {v:?} engine {:?}", got.get(k)))
            .collect();
        assert!(wrong.is_empty(), "{executor:?}:\n{}", wrong.join("\n"));
    }
}

#[test]
fn null_to_byval_parameters_matches_excel() {
    let main = format!(
        "Public result As String\nSub Main()\n  result = Run()\nEnd Sub\n{}",
        include_str!("../../../conformance/excel-measurements/null-byval.vba")
    );
    check(
        &[
            ("Main", ModuleKind::Procedural, &main),
            (
                "Target",
                ModuleKind::Class,
                include_str!("../../../conformance/excel-measurements/null-byval-target.clsvba"),
            ),
        ],
        include_str!("../../../conformance/excel-measurements/null-byval.excel.txt"),
    );
}

#[test]
fn callbyname_vblet_conversions_match_excel() {
    let main = format!(
        "Public result As String\nSub Main()\n  result = Run()\nEnd Sub\n{}",
        include_str!("../../../conformance/excel-measurements/callbyname-vblet.vba")
    );
    check(
        &[
            ("Main", ModuleKind::Procedural, &main),
            (
                "Holder",
                ModuleKind::Class,
                include_str!("../../../conformance/excel-measurements/callbyname-holder.clsvba"),
            ),
        ],
        include_str!("../../../conformance/excel-measurements/callbyname-vblet.excel.txt"),
    );
}
