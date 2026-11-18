//! MS-VBAL conformance pass (dc-ez3rd): the probe modules in
//! `conformance/excel-measurements/spec-pass`, written from a section-by-section reading of
//! MS-VBAL against the CalcVBA code, run in Excel VBA (16.0, `<Module>.excel.txt`, captured with
//! `scripts/run-inproc-probe.ps1 -PerModule`) and here on VM3 and the JIT.
//!
//! A module's Run returns `key=value` lines. A `.excel.txt` that starts with `compile error:`
//! records a module VBE refused to compile; CalcVBA must refuse it too. Lines CalcVBA still answers
//! differently are listed in `known-differences.txt` as `Module: line` (prefixed `vm3 ` or `jit `
//! when only one executor differs; `Module: run fails` when the module does not run at all).

use calcvba_differential::{Canon, Executor, run_modules};
use calcvba_symbol::manifest::ModuleKind::{Class, Procedural};

const SPECK: &str = include_str!("../../../conformance/excel-measurements/spec-pass/SpecK.cls");
const KNOWN: &str =
    include_str!("../../../conformance/excel-measurements/spec-pass/known-differences.txt");

macro_rules! cases {
    ($($module:literal),* $(,)?) => {
        &[$((
            $module,
            include_str!(concat!("../../../conformance/excel-measurements/spec-pass/", $module, ".bas")),
            include_str!(concat!("../../../conformance/excel-measurements/spec-pass/", $module, ".excel.txt")),
        )),*]
    };
}

/// (module, source, Excel's answer)
const CASES: &[(&str, &str, &str)] = cases!(
    "SpecConst",
    "SpecExpr",
    "SpecLib1",
    "SpecLib2",
    "SpecOps",
    "SpecOpsText",
    "SpecStmt",
    "SpecTypes",
    "CmpByRefVariantElem",
    "CmpCallByVal",
    "CmpCaret",
    "CmpCatBytes",
    "CmpCcCont",
    "CmpCcDup",
    "CmpCcEndIf",
    "CmpCcIfCont",
    "CmpCcString",
    "CmpCcSuffix",
    "CmpCcVba6",
    "CmpCommentCont",
    "CmpDate13PM",
    "CmpDate2Part",
    "CmpDate3PM",
    "CmpDateAttachedAM",
    "CmpDateDash",
    "CmpDateMonYear",
    "CmpDateNoonPM",
    "CmpDateYMD2",
    "CmpDateYY",
    "CmpDateYY29",
    "CmpEnumDbl",
    "CmpEnumFwd",
    "CmpEnumOver",
    "CmpGotoFor",
    "CmpNamedThenPos",
    "CmpNumBig",
    "CmpNumD",
    "CmpNumDot",
    "CmpNumOct",
    "CmpOptBase1Amp",
    "CmpOptBase2",
    "CmpOptCompareFoo",
    "CmpOptExplicit2",
    "CmpPropMismatch",
    "CmpPropOptional",
);

/// SpecStmt cases that touch the file system. `run_modules` runs under the restricted host policy
/// (no real file system), so these are checked by `spec_pass_laned_vm3.rs` on a native-policy
/// engine (temp directory, Linux and Windows) instead.
const FILE_KEYS: &[&str] = &[
    "openMissing",
    "openTwice",
    "openHash0",
    "openHash600",
    "seekZero",
    "putOverLen",
    "eofBad",
    "printEmpty",
    "printLeadComma",
    "writeEmpty",
    "writeTrailSemi",
    "printSpcTab",
    "printWidth",
    "printDbl",
    "writeDbl",
    "printBig",
    "randomDefLen",
    "randomLenBad",
    "killMissing",
    "mkdirExisting",
    "rmdirNonEmpty",
    "chdirMissing",
    "fileLenMissing",
    "dirAttrs",
    "dirFolder",
    "fileAttrMode",
    "eofBinary",
    "getAttrType",
    "curDirDollar",
];

/// Modules CalcVBA does not finish (an endless loop); not run, and counted as differing.
const HANGS: &[&str] = &[];

/// Known differences for `executor`: lines prefixed `vm3 ` or `jit ` hold for that executor only.
fn known(executor: Executor) -> std::collections::BTreeSet<&'static str> {
    let own = match executor {
        Executor::Vm3 => "vm3 ",
        Executor::Jit => "jit ",
    };
    KNOWN
        .lines()
        .map(str::trim_end)
        .filter(|l| !l.trim().is_empty() && !l.starts_with('#'))
        .filter_map(|l| {
            if let Some(rest) = l.strip_prefix(own) {
                Some(rest)
            } else if l.starts_with("vm3 ") || l.starts_with("jit ") {
                None
            } else {
                Some(l)
            }
        })
        .collect()
}

fn skipped(module: &str, line: &str) -> bool {
    module == "SpecStmt" && FILE_KEYS.contains(&line.split('=').next().unwrap_or(""))
}

/// The lines (as `Module: line`) where `executor` differs from Excel, and any known-differences
/// entries that now match.
fn compare(executor: Executor) -> (Vec<String>, Vec<String>) {
    let known = known(executor);
    let mut wrong = Vec::new();
    let mut fixed = Vec::new();
    for (module, source, expected) in CASES {
        if HANGS.contains(module) {
            let entry = format!("{module}: hangs");
            if !known.contains(entry.as_str()) {
                wrong.push(entry);
            }
            continue;
        }
        let main =
            format!("Public result As String\nSub Main()\n    result = {module}.Run()\nEnd Sub\n");
        let mut modules = vec![
            ("Main", Procedural, main.as_str()),
            (*module, Procedural, *source),
        ];
        if *module == "SpecExpr" {
            modules.push(("SpecK", Class, SPECK));
        }
        let outcome = run_modules(executor, &modules, "VBAProject");
        let text = match (&outcome.unsupported, &outcome.result) {
            (Some(why), _) => Err(format!("declined: {why:?}")),
            (None, Err(e)) => Err(format!("failed: {e}")),
            (None, Ok(values)) => values
                .iter()
                .find_map(|v| match v {
                    Canon::Str(s) if s.contains('=') => Some(s.clone()),
                    _ => None,
                })
                .ok_or_else(|| format!("no result in {values:?}")),
        };
        let expected = expected.replace("\r\n", "\n");
        if let Some(dialog) = expected.strip_prefix("compile error:") {
            let entry = format!(
                "{module}: compile error:{}",
                dialog.lines().next().unwrap_or("")
            );
            match (&text, known.contains(entry.as_str())) {
                (Ok(got), false) => wrong.push(format!("{entry} -- {executor:?} ran: {got:?}")),
                (Err(_), true) => fixed.push(entry),
                _ => {}
            }
            continue;
        }
        let got = match text {
            Ok(got) => got,
            Err(why) => {
                let entry = format!("{module}: {why}");
                if !known.contains(format!("{module}: run fails").as_str()) {
                    wrong.push(entry);
                }
                continue;
            }
        };
        if known.contains(format!("{module}: run fails").as_str()) {
            fixed.push(format!("{module}: run fails"));
        }
        let got: Vec<&str> = got.lines().collect();
        for (i, want) in expected.lines().enumerate() {
            if skipped(module, want) {
                continue;
            }
            let entry = format!("{module}: {want}");
            let have = got.get(i).copied().unwrap_or("<missing>");
            match (want == have, known.contains(entry.as_str())) {
                (false, false) => wrong.push(format!("{entry} -- {executor:?} {have:?}")),
                (true, true) => fixed.push(entry),
                _ => {}
            }
        }
    }
    (wrong, fixed)
}

fn check(executor: Executor) {
    let (wrong, fixed) = compare(executor);
    assert!(
        wrong.is_empty() && fixed.is_empty(),
        "{executor:?} against Excel ({} wrong, {} fixed but still in known-differences.txt):\n{}\n\
         fixed:\n{}",
        wrong.len(),
        fixed.len(),
        wrong.join("\n"),
        fixed.join("\n")
    );
}

#[test]
fn vm3_matches_excel() {
    check(Executor::Vm3);
}

#[test]
fn jit_matches_excel() {
    check(Executor::Jit);
}
