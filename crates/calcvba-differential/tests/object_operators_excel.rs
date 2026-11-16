//! Operators, conversions and logical result types over VBA class objects held in Variants,
//! against Excel VBA (16.0 build 20527; probe temp/variant-layout-probe/ops13, the COM-free
//! lines kept in `conformance/excel-measurements/object-operators`).
//!
//! MS-VBAL defines the operators as "simple data operators" (each operand evaluated as a simple
//! data value, an object's default member) and gives the logical operators' declared types; VBE7
//! follows that except for a binary operator over a VBA class object in a Variant, which answers
//! that object's default member (the right operand's when it is one and not a call result, else
//! the left's) without performing the operation. The simple-data-value rules hold on VM3 and the
//! JIT; that VBE7 behaviour is not copied yet (`known-differences.txt`).

use calcvba_differential::{Canon, Executor, run_modules};
use calcvba_symbol::manifest::ModuleKind::{Class, Procedural};

const DIR: &str = "../../conformance/excel-measurements/object-operators";
const DFLT: &str =
    include_str!("../../../conformance/excel-measurements/object-operators/Dflt.cls");
const SEVEN: &str =
    include_str!("../../../conformance/excel-measurements/object-operators/Seven.cls");
/// Excel's lines CalcVBA still answers differently (the class-object operator behaviour).
const KNOWN: &str =
    include_str!("../../../conformance/excel-measurements/object-operators/known-differences.txt");

fn known() -> std::collections::BTreeSet<&'static str> {
    KNOWN
        .lines()
        .filter(|l| !l.trim().is_empty() && !l.starts_with('#'))
        .collect()
}

fn check(executor: Executor, module: &str, source: &str, expected: &str) {
    let main =
        format!("Public result As String\nSub Main()\n    result = {module}.Run()\nEnd Sub\n");
    let modules = [
        ("Main", Procedural, main.as_str()),
        (module, Procedural, source),
        ("Dflt", Class, DFLT),
        ("Seven", Class, SEVEN),
    ];
    let outcome = run_modules(executor, &modules, "VBAProject");
    assert!(
        outcome.unsupported.is_none(),
        "{executor:?} declined {module}: {:?}",
        outcome.unsupported
    );
    let values = outcome
        .result
        .unwrap_or_else(|e| panic!("{executor:?} {module} failed: {e}"));
    let text = values
        .iter()
        .find_map(|v| match v {
            Canon::Str(s) if s.contains('=') => Some(s.clone()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("{executor:?} {module}: no result in {values:?}"));
    let got: Vec<&str> = text.lines().collect();
    let want: Vec<&str> = expected.lines().collect();
    let known = known();
    let wrong: Vec<String> = want
        .iter()
        .zip(&got)
        .filter(|(w, g)| w != g && !known.contains(**w))
        .map(|(w, g)| format!("VBA {w:?}, {executor:?} {g:?}"))
        .collect();
    let fixed: Vec<&str> = want
        .iter()
        .zip(&got)
        .filter(|(w, g)| w == g && known.contains(**w))
        .map(|(w, _)| *w)
        .collect();
    assert_eq!(
        got.len(),
        want.len(),
        "{executor:?} {module}: {text} (in {DIR})"
    );
    assert!(wrong.is_empty(), "{module}:\n{}", wrong.join("\n"));
    assert!(
        fixed.is_empty(),
        "{module}: listed in known-differences.txt but now matching Excel: {fixed:#?}"
    );
}

macro_rules! probe {
    ($name:ident, $module:literal) => {
        mod $name {
            use super::*;
            const SOURCE: &str = include_str!(concat!(
                "../../../conformance/excel-measurements/object-operators/",
                $module,
                ".bas"
            ));
            const EXPECTED: &str = include_str!(concat!(
                "../../../conformance/excel-measurements/object-operators/",
                $module,
                ".excel.txt"
            ));
            #[test]
            fn vm3_matches_excel() {
                check(Executor::Vm3, $module, SOURCE, EXPECTED);
            }
            #[test]
            fn jit_matches_excel() {
                check(Executor::Jit, $module, SOURCE, EXPECTED);
            }
        }
    };
}

probe!(operand_kinds, "Ops2Probe");
probe!(operand_kinds_more, "Ops3Probe");
probe!(logical_result_types, "LogProbe");
