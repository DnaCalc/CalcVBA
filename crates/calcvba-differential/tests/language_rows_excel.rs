//! The 192 basic-language conformance fixtures (`conformance/tests/*.bas`, suite `basic-language`)
//! against Excel's VBA (16.0 build 20527): value, VarType, Err.Number and text form recorded in
//! `conformance/excel-measurements/language-rows.excel.txt` (`gen_language_rows.py` prepares the
//! fixtures, `vba-job.ps1 -Action language` in DnaCalc runs them, `collect` writes the answers).
//!
//! A row matches when the engine finishes with the same typed values as Excel (module variables
//! first, then the locals of `Main`, each by VarType and value; floating point by bit pattern;
//! strings by UTF-16 units), or raises the same `Err.Number`, or both refuse to compile. An
//! array or object variable is compared by kind only, and a user-defined-type local is not
//! recorded: such rows are listed as `container` rows in the report.
//!
//! Rows the engine still answers differently are listed in
//! `language-rows.known-differences.txt` (`<Executor>\t<file>\t<reason>`); the test fails on any
//! new difference and on any listed difference that is now fixed.

use calcvba_differential::{Canon, Executor, RunOutcome, canon, run};
use calcvba_runtime::Variant;
use std::collections::BTreeSet;
use std::path::PathBuf;

const EXCEL: &str = include_str!("../../../conformance/excel-measurements/language-rows.excel.txt");
const KNOWN: &str =
    include_str!("../../../conformance/excel-measurements/language-rows.known-differences.txt");

fn conformance(file: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../conformance/tests")
        .join(file)
}

/// What Excel recorded for one variable.
enum Recorded {
    /// A scalar with its exact canonical form.
    Scalar(Canon),
    /// An array or an object: compared by kind.
    Container,
    /// A Decimal: compared by VarType only.
    Decimal,
    /// Not recorded (a user-defined-type local).
    Unrecorded,
}

fn hex_units(text: &str) -> Vec<u16> {
    (0..text.len() / 4)
        .map(|i| u16::from_str_radix(&text[i * 4..i * 4 + 4], 16).expect("hex unit"))
        .collect()
}

fn currency_scaled(text: &str) -> i64 {
    let negative = text.starts_with('-');
    let digits = text.trim_start_matches(['-', '+']);
    let (whole, fraction) = digits.split_once('.').unwrap_or((digits, ""));
    let whole: i64 = if whole.is_empty() {
        0
    } else {
        whole.parse().unwrap()
    };
    let mut fraction = fraction.to_string();
    while fraction.len() < 4 {
        fraction.push('0');
    }
    let scaled = whole * 10_000 + fraction[..4].parse::<i64>().unwrap();
    if negative { -scaled } else { scaled }
}

fn recorded(record: &str) -> Recorded {
    let (_, value) = record.split_once('=').expect("name=value");
    let (vt, payload) = value.split_once(':').expect("vt:payload");
    let vt: u32 = vt.parse().expect("vt");
    if vt >= 8192 || vt == 9 || vt == 13 {
        return Recorded::Container;
    }
    let bits = |p: &str| f64::from_bits(u64::from_str_radix(p, 16).expect("double bits"));
    let variant = match vt {
        0 => Variant::empty(),
        1 => Variant::null(),
        2 => Variant::from_i16(payload.parse().unwrap()),
        3 => Variant::from_i32(payload.parse().unwrap()),
        17 => Variant::from_u8(payload.parse().unwrap()),
        20 => Variant::from_i64(payload.parse().unwrap()),
        4 => Variant::from_f32(bits(payload) as f32),
        5 => Variant::from_f64(bits(payload)),
        7 => Variant::from_date_f64(bits(payload)),
        6 => Variant::from_currency_scaled_i64(currency_scaled(payload)),
        14 => return Recorded::Decimal,
        8 => Variant::from_utf16_units(&hex_units(payload)),
        10 => Variant::from_error_code(payload.parse().unwrap()),
        11 => Variant::from_bool(payload == "true"),
        other => panic!("unexpected VarType {other}"),
    };
    Recorded::Scalar(canon(&variant))
}

enum Answer {
    Values(Vec<(String, Recorded)>),
    Error(i32),
    Compile,
    Other(String),
}

fn answers() -> Vec<(String, Answer)> {
    EXCEL
        .lines()
        .filter(|l| !l.trim().is_empty())
        .map(|line| {
            let mut fields = line.split('\t');
            let file = fields.next().unwrap().to_string();
            let kind = fields.next().unwrap_or("");
            let answer = match kind {
                "OK" => Answer::Values(
                    fields
                        .filter(|f| !f.is_empty() && !f.starts_with('#'))
                        .map(|f| {
                            let name = f.split('=').next().unwrap().to_string();
                            let value = if f.contains("=udt:") {
                                Recorded::Unrecorded
                            } else {
                                recorded(f)
                            };
                            (name, value)
                        })
                        .collect(),
                ),
                "ERR" => Answer::Error(fields.next().unwrap().parse().unwrap()),
                "COMPILE" => Answer::Compile,
                other => Answer::Other(other.to_string()),
            };
            (file, answer)
        })
        .collect()
}

#[derive(Default)]
struct Tally {
    matched: usize,
    containers: usize,
    different: BTreeSet<String>,
}

fn compare(outcome: &RunOutcome, answer: &Answer) -> Result<bool, String> {
    match answer {
        Answer::Values(expected) => {
            let Ok(values) = &outcome.result else {
                return Err(format!("engine failed: {:?}", outcome.result));
            };
            if values.len() != expected.len() {
                return Err(format!(
                    "engine has {} values, Excel {}",
                    values.len(),
                    expected.len()
                ));
            }
            let mut container = false;
            for (got, (name, want)) in values.iter().zip(expected) {
                match want {
                    Recorded::Scalar(want) => {
                        if got != want {
                            return Err(format!("{name}: engine {got:?}, Excel {want:?}"));
                        }
                    }
                    Recorded::Container => {
                        container = true;
                        if !matches!(got, Canon::Opaque { .. }) {
                            return Err(format!("{name}: engine {got:?}, Excel a container"));
                        }
                    }
                    Recorded::Decimal => {
                        container = true;
                        if !matches!(got, Canon::Raw { tag: 14, .. }) {
                            return Err(format!("{name}: engine {got:?}, Excel a Decimal"));
                        }
                    }
                    Recorded::Unrecorded => container = true,
                }
            }
            Ok(container)
        }
        Answer::Error(number) => {
            if !outcome.raised {
                return Err(format!(
                    "Excel raised {number}; engine {:?}",
                    outcome.result.as_ref().map(|_| "completed")
                ));
            }
            if outcome.err.number != *number {
                return Err(format!(
                    "Excel raised {number}; engine raised {}",
                    outcome.err.number
                ));
            }
            Ok(false)
        }
        Answer::Compile => {
            if outcome.raised || outcome.result.is_ok() {
                return Err("Excel refused to compile; engine ran it".into());
            }
            Ok(false)
        }
        Answer::Other(kind) => Err(format!("Excel answer {kind}")),
    }
}

fn known(executor: Executor) -> BTreeSet<String> {
    let tag = format!("{executor:?}");
    KNOWN
        .lines()
        .filter(|l| !l.trim().is_empty() && !l.starts_with('#'))
        .filter_map(|l| {
            let mut fields = l.split('\t');
            (fields.next() == Some(tag.as_str())).then(|| fields.next().unwrap().to_string())
        })
        .collect()
}

fn survey(executor: Executor) -> (Tally, Vec<String>) {
    let mut tally = Tally::default();
    let mut details = Vec::new();
    for (file, answer) in answers() {
        let source = std::fs::read_to_string(conformance(&file))
            .or_else(|_| {
                std::fs::read(conformance(&file))
                    .map(|bytes| bytes.iter().map(|b| *b as char).collect())
            })
            .unwrap();
        let outcome = run(executor, &source);
        if let Some(why) = &outcome.unsupported {
            tally.different.insert(file.clone());
            details.push(format!("{file}\tunsupported: {why}"));
            continue;
        }
        match compare(&outcome, &answer) {
            Ok(container) => {
                tally.matched += 1;
                tally.containers += usize::from(container);
            }
            Err(reason) => {
                tally.different.insert(file.clone());
                details.push(format!("{file}\t{reason}"));
            }
        }
    }
    (tally, details)
}

fn check(executor: Executor) {
    let (tally, details) = survey(executor);
    eprintln!(
        "{executor:?}: {} of {} language rows match Excel ({} of them hold a container or Decimal compared by kind); {} differ",
        tally.matched,
        tally.matched + tally.different.len(),
        tally.containers,
        tally.different.len()
    );
    if let Ok(path) = std::env::var("LANGUAGE_ROWS_DUMP") {
        let body: String = details
            .iter()
            .map(|d| format!("{executor:?}\t{d}\n"))
            .collect();
        let mut old = std::fs::read_to_string(&path).unwrap_or_default();
        old.push_str(&body);
        std::fs::write(path, old).unwrap();
    }
    let listed = known(executor);
    let new: Vec<_> = tally.different.difference(&listed).collect();
    let fixed: Vec<_> = listed.difference(&tally.different).collect();
    assert!(
        new.is_empty(),
        "{executor:?}: new differences from Excel: {new:?}"
    );
    assert!(
        fixed.is_empty(),
        "{executor:?}: listed differences now match Excel, remove them: {fixed:?}"
    );
}

#[test]
fn vm3_language_rows_against_excel() {
    check(Executor::Vm3);
}

#[test]
fn jit_language_rows_against_excel() {
    check(Executor::Jit);
}
