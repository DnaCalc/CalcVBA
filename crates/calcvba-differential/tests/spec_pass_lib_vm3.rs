//! dc-ez3rd lane C: the VBA library points the spec-pass probes measured in Excel, plus the
//! neighbours the probes do not cover. Each case runs on VM3 and on the JIT.

use calcvba_differential::{Canon, Executor, run};
use calcvba_runtime::regional::{RegionalSettings, with_settings};

/// `result = <expr>` as a String, on both executors.
fn eval(expr: &str) -> Vec<String> {
    let source = format!("Public result As String\nSub Main()\n    result = {expr}\nEnd Sub\n");
    [Executor::Vm3, Executor::Jit]
        .into_iter()
        .map(|executor| {
            let outcome = with_settings(RegionalSettings::en_us(), || run(executor, &source));
            assert!(
                outcome.unsupported.is_none(),
                "{executor:?} declined {expr}: {:?}",
                outcome.unsupported
            );
            match outcome.result {
                Ok(values) => values
                    .iter()
                    .find_map(|v| match v {
                        Canon::Str(s) => Some(s.clone()),
                        _ => None,
                    })
                    .unwrap_or_else(|| format!("no string in {values:?}")),
                Err(error) => format!("failed: {error}"),
            }
        })
        .collect()
}

fn check(expr: &str, expected: &str) {
    for (got, executor) in eval(expr).into_iter().zip(["vm3", "jit"]) {
        assert!(
            got == expected || got.contains(expected),
            "{executor}: {expr} -> {got:?}, want {expected:?}"
        );
    }
}

#[test]
fn byte_functions_count_bytes() {
    check("CStr(AscB(\"A\"))", "65");
    check("CStr(LenB(ChrB(65)))", "1");
    check("CStr(InStrB(\"abc\", \"b\"))", "3");
    check("CStr(InStrB(1, \"abc\", \"c\"))", "5");
    check("CStr(AscB(MidB(\"abc\", 3, 2)))", "98");
    check("CStr(LenB(MidB(\"abc\", 2)))", "5");
}

#[test]
fn like_matches_utf16_units_and_rejects_bad_patterns() {
    check("CStr(\"ab\" Like \"a[]b\")", "True");
    check("CStr(\"[\" Like \"[\")", "93");
    check("CStr(\"b\" Like \"[z-a]\")", "93");
    check("CStr(\"b\" Like \"[a-c]\")", "True");
    check("CStr(\"d\" Like \"[!a-c]\")", "True");
    check("CStr(ChrW(&HD83D) & ChrW(&HDE00) Like \"??\")", "True");
    check("CStr(\"a3\" Like \"a#\")", "True");
}

#[test]
fn week_numbers_follow_first_day_and_first_week() {
    // 1/1/2005 is a Saturday.
    check("CStr(DatePart(\"ww\", #1/1/2005#))", "1");
    check(
        "CStr(DatePart(\"ww\", #1/1/2005#, vbSunday, vbFirstFullWeek))",
        "52",
    );
    check(
        "CStr(DatePart(\"ww\", #1/1/2005#, vbMonday, vbFirstFourDays))",
        "53",
    );
    check(
        "CStr(DatePart(\"ww\", #12/29/2003#, vbMonday, vbFirstFourDays))",
        "1",
    );
    check(
        "Format(#1/1/2005#, \"ww\", vbMonday, vbFirstFourDays)",
        "53",
    );
    check("Format(#1/5/2000#, \"w\", vbMonday)", "3");
}

#[test]
fn dateserial_windows_years_and_carries_months() {
    check("CStr(Year(DateSerial(29, 1, 1)))", "2029");
    check("CStr(Year(DateSerial(30, 1, 1)))", "1930");
    check(
        "Format(DateSerial(2000, 15, 1), \"yyyy-mm-dd\")",
        "2001-03-01",
    );
    check(
        "Format(DateSerial(2000, -10, 1), \"yyyy-mm-dd\")",
        "1999-02-01",
    );
    check(
        "Format(DateSerial(2000, 1, 0), \"yyyy-mm-dd\")",
        "1999-12-31",
    );
    check("CStr(TypeName(Day(#1/2/2000#)))", "Integer");
    check(
        "CStr(TypeName(DateDiff(\"d\", #1/1/2000#, #1/2/2000#)))",
        "Long",
    );
}

#[test]
fn datediff_counts_boundaries_crossed() {
    check("CStr(DateDiff(\"q\", #3/31/2000#, #4/1/2000#))", "1");
    check("CStr(DateDiff(\"yyyy\", #12/31/1999#, #1/1/2000#))", "1");
    check("CStr(DateDiff(\"m\", #1/31/2000#, #2/1/2000#))", "1");
    check(
        "CStr(DateDiff(\"h\", #1/1/2000 10:59:00 PM#, #1/1/2000 11:01:00 PM#))",
        "1",
    );
    check(
        "CStr(DateDiff(\"n\", #1/1/2000 10:59:30 PM#, #1/1/2000 11:00:10 PM#))",
        "1",
    );
    check("CStr(DateDiff(\"w\", #1/9/2023#, #1/8/2023#))", "0");
}

#[test]
fn strings_numbers_and_null_follow_the_argument_rules() {
    check("CStr(IsNull(Day(Null)))", "True");
    check("CStr(Day(\"1/15/2000\"))", "15");
    check("CStr(Choose(1.9, \"a\", \"b\"))", "a");
    check("CStr(IIf(\"True\", 1, 2))", "1");
    check("TypeName(Split(\"a b\"))", "String()");
    check("CStr(UBound(Split(\"\")))", "-1");
    check("CStr(InStr(\"abc\", \"\"))", "1");
    check("CStr(InStr(\"\", \"a\"))", "0");
    check("CStr(InStrRev(\"abc\", \"c\", 10))", "0");
    check("StrConv(\"a\" & vbTab & \"b c\", vbProperCase)", "A\tB C");
    check("CStr(LenB(StrConv(\"abc\", vbFromUnicode)))", "3");
    check("CStr(Len(StrConv(\"abc\", vbUnicode)))", "6");
}

#[test]
fn round_keeps_the_argument_type() {
    check("TypeName(Round(CCur(2.5)))", "Currency");
    check("TypeName(Round(CInt(2)))", "Integer");
    check("TypeName(Round(CLng(7)))", "Long");
    check("CStr(Round(CDec(\"2.55\"), 1))", "2.6");
    check("CStr(Round(2.5))", "2");
    check("CStr(Round(0.125, 2))", "0.12");
}

#[test]
fn financial_arguments_are_checked() {
    check("CStr(Pmt(0.1, 0, 1000))", "5");
    check("CStr(DDB(1000, 100, 5, 6))", "5");
    check("CStr(SLN(1, 0, 0))", "5");
    check("CStr(Round(FV(0.1, 10, -100, 0, 2), 2))", "1753.12");
    check("CStr(IRR(Array(-100, 230, -132)))", "0.1");
}

#[test]
fn string_masks_format_text() {
    check("Format(\"abc\", \"@@@@@\")", "  abc");
    check("Format(\"abc\", \">\")", "ABC");
    check("Format(\"ABC\", \"<\")", "abc");
    check("Format(\"abc\", \"0.00\")", "abc");
    check("Format(5, \"\"\"x\"\"0\")", "x5");
    check("Format(Null, \"0;0;0;\"\"nil\"\"\")", "nil");
}

#[test]
fn collection_errors_and_keys() {
    let body = |stmt: &str| {
        format!(
            "Public result As String\nSub Main()\n    Dim c As New Collection\n    On Error Resume Next\n    c.Add 1, \"a\"\n    {stmt}\n    result = CStr(Err.Number)\nEnd Sub\n"
        )
    };
    for (stmt, want) in [
        ("Dim v: v = c(\"zz\")", "5"),
        ("c.Remove \"zz\"", "5"),
        ("Dim v: v = c(5)", "9"),
        ("c.Add 2, 5", "13"),
        ("c.Add 2, \"A\"", "457"),
        ("c.Add 2, ChrW(196): c.Add 3, ChrW(228)", "457"),
    ] {
        for executor in [Executor::Vm3, Executor::Jit] {
            let outcome = run(executor, &body(stmt));
            let got = outcome
                .result
                .ok()
                .and_then(|values| {
                    values.iter().find_map(|v| match v {
                        Canon::Str(s) => Some(s.clone()),
                        _ => None,
                    })
                })
                .unwrap_or_default();
            assert_eq!(got, want, "{executor:?}: {stmt}");
        }
    }
}

#[test]
fn rnd_and_randomize_seed_like_vba() {
    check("CStr(Rnd(-1.1))", "7.187074E-02");
    let source = "Public result As String\nSub Main()\n    Dim x As Single\n    x = Rnd(-1)\n    Randomize 2\n    result = CStr(Rnd)\nEnd Sub\n";
    for executor in [Executor::Vm3, Executor::Jit] {
        let outcome = run(executor, source);
        let got = outcome.result.ok().and_then(|values| {
            values.iter().find_map(|v| match v {
                Canon::Str(s) => Some(s.clone()),
                _ => None,
            })
        });
        assert_eq!(got.as_deref(), Some("0.5828429"), "{executor:?}");
    }
}

#[test]
fn constants_added_from_the_spec() {
    check("TypeName(vbKeyReturn) & CStr(vbKeyReturn)", "Integer13");
    check(
        "CStr(vbKeyA) & \" \" & CStr(vbKeyF16) & \" \" & CStr(vbKeyNumpad9)",
        "65 127 105",
    );
    check("CStr(vbWindowText)", "-2147483640");
    check(
        "CStr(vbMinimizedNoFocus) & CStr(vbFormMDIForm) & CStr(vbCalHijri)",
        "641",
    );
}
