//! dc-ez3rd lane B: the SpecTypes/SpecOps probe lines about types, coercion, operators and
//! Decimal arithmetic, checked against Excel's answers on VM3 and the JIT.

use calcvba_differential::{Canon, Executor, run_modules};
use calcvba_symbol::manifest::ModuleKind::Procedural;

const DIR: &str = "../../conformance/excel-measurements/spec-pass/";

/// Lines owned by other lanes (Like, Set/member access, regional date text).
const OTHER_LANES: &[&str] = &[
    "likeOpenBr",
    "likeRevRange",
    "likeEmptyBr",
    "setObjInt",
    "setVarInt",
    "memberOnInt",
];

fn run(executor: Executor, module: &str, source: &str) -> Result<String, String> {
    let main =
        format!("Public result As String\nSub Main()\n    result = {module}.Run()\nEnd Sub\n");
    let modules = vec![
        ("Main", Procedural, main.as_str()),
        (module, Procedural, source),
    ];
    let outcome = run_modules(executor, &modules, "VBAProject");
    match (&outcome.unsupported, &outcome.result) {
        (Some(why), _) => Err(format!("declined: {why:?}")),
        (None, Err(e)) => Err(format!("failed: {e}")),
        (None, Ok(values)) => values
            .iter()
            .find_map(|v| match v {
                Canon::Str(s) if s.contains('=') => Some(s.clone()),
                _ => None,
            })
            .ok_or_else(|| format!("no result in {values:?}")),
    }
}

fn read(name: &str) -> String {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join(DIR)
        .join(name);
    std::fs::read_to_string(&path).unwrap_or_else(|e| panic!("{}: {e}", path.display()))
}

fn check_module(module: &str) {
    let source = read(&format!("{module}.bas"));
    let expected = read(&format!("{module}.excel.txt")).replace("\r\n", "\n");
    let mut wrong = Vec::new();
    for executor in [Executor::Vm3, Executor::Jit] {
        let got = run(executor, module, &source).unwrap_or_else(|e| panic!("{module}: {e}"));
        let got: Vec<&str> = got.lines().collect();
        for (i, want) in expected.lines().enumerate() {
            let key = want.split('=').next().unwrap_or("");
            if OTHER_LANES.contains(&key) {
                continue;
            }
            let have = got.get(i).copied().unwrap_or("<missing>");
            if want != have {
                wrong.push(format!("{executor:?} {module}: {want} -- got {have}"));
            }
        }
    }
    assert!(wrong.is_empty(), "{}", wrong.join("\n"));
}

#[test]
fn spec_types_lines_match_excel() {
    check_module("SpecTypes");
}

#[test]
fn spec_ops_lines_match_excel() {
    check_module("SpecOps");
}

#[test]
fn concatenating_a_byte_array_is_a_compile_error() {
    let source = read("CmpCatBytes.bas");
    for executor in [Executor::Vm3, Executor::Jit] {
        let got = run(executor, "CmpCatBytes", &source);
        assert!(got.is_err(), "{executor:?} compiled `bb & \"!\"`: {got:?}");
    }
}

/// Runs `body` (which sets `Run`) as a probe module on both executors.
fn probe(body: &str) -> Vec<(Executor, String)> {
    let source = format!("Attribute VB_Name = \"LaneB\"\n{body}");
    [Executor::Vm3, Executor::Jit]
        .into_iter()
        .map(|executor| {
            let got = run(executor, "LaneB", &source).unwrap_or_else(|e| panic!("{e}"));
            (executor, got)
        })
        .collect()
}

#[test]
fn error_value_into_a_string_parameter_is_a_type_mismatch() {
    let body = r#"
Private Function Take(ByVal s As String) As String
    Take = s
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "arg=" & Take(CVErr(5))
    Exit Function
fail:
    Run = "arg=ERR" & Err.Number
End Function
"#;
    for (executor, got) in probe(body) {
        assert_eq!(got, "arg=ERR13", "{executor:?}");
    }
}

#[test]
fn decimal_and_logical_edge_cases() {
    let body = r#"
Private Function T(ByVal k As String) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case "decHalfEven": v = CDec("0.0000000000000000000000000025") * CDec("0.1")
        Case "decTwoThirds": v = CDec(2) / 3
        Case "decPlusDbl": v = CDec("0.1") + 0.2
        Case "decCLng": v = CLng(CDec("2.5"))
        Case "decNegZero": v = CDec("-1.5") + CDec("1.5")
        Case "falseImpNull": v = False Imp Null
        Case "zeroOrNull": v = 0 Or Null
        Case "notByte": v = Not CByte(1)
        Case "emptyMinus5": v = Empty - "5"
    End Select
    T = k & "=" & TypeName(v) & ":" & CStr(v)
    Exit Function
fail:
    T = k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Variant
    For Each k In Split("decHalfEven decTwoThirds decPlusDbl decCLng decNegZero falseImpNull zeroOrNull notByte emptyMinus5", " ")
        Run = Run & T(k) & vbLf
    Next
End Function
"#;
    let want = "decHalfEven=Decimal:0.0000000000000000000000000002\n\
                decTwoThirds=Decimal:0.6666666666666666666666666667\n\
                decPlusDbl=Decimal:0.3\n\
                decCLng=Long:2\n\
                decNegZero=Decimal:0\n\
                falseImpNull=Boolean:True\n\
                zeroOrNull=ERR94\n\
                notByte=Byte:254\n\
                emptyMinus5=Double:-5\n";
    for (executor, got) in probe(body) {
        assert_eq!(got, want, "{executor:?}");
    }
}
