//! VBA `Rate`, bit for bit (dc-bjt7v): `conformance/excel-measurements/rate-shape/RateShape.bas`
//! run in Excel VBA (`RateShape.excel.txt`) and here on VM3 and the JIT. Each line is
//! `key=<IEEE bits>|<CStr>` or `key=ERRn` for one `Rate` call.

use calcvba_differential::{Canon, Executor, run_modules};
use calcvba_symbol::manifest::ModuleKind::Procedural;

const SOURCE: &str =
    include_str!("../../../conformance/excel-measurements/rate-shape/RateShape.bas");
const EXCEL: &str =
    include_str!("../../../conformance/excel-measurements/rate-shape/RateShape.excel.txt");

fn check(executor: Executor) {
    let main = "Public result As String\nSub Main()\n    result = RateShape.Run()\nEnd Sub\n";
    let outcome = run_modules(
        executor,
        &[
            ("Main", Procedural, main),
            ("RateShape", Procedural, SOURCE),
        ],
        "VBAProject",
    );
    assert!(outcome.unsupported.is_none(), "{:?}", outcome.unsupported);
    let values = outcome.result.expect("RateShape runs");
    let got = values
        .iter()
        .find_map(|v| match v {
            Canon::Str(s) if s.contains('=') => Some(s.clone()),
            _ => None,
        })
        .expect("RateShape result");
    let got: Vec<&str> = got.lines().collect();
    let expected = EXCEL.replace("\r\n", "\n");
    let mut wrong = Vec::new();
    for (i, want) in expected.lines().filter(|l| !l.is_empty()).enumerate() {
        let have = got.get(i).copied().unwrap_or("<missing>");
        if want != have {
            wrong.push(format!("{want} -- {executor:?} {have:?}"));
        }
    }
    assert!(
        wrong.is_empty(),
        "{executor:?} differs from Excel on {} RateShape lines:\n{}",
        wrong.len(),
        wrong.join("\n")
    );
}

#[test]
fn vm3_rate_matches_excel() {
    check(Executor::Vm3);
}

#[test]
fn jit_rate_matches_excel() {
    check(Executor::Jit);
}
