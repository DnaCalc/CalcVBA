//! MS-VBAL lexical rules, conditional compilation, Option directives and enums (dc-ez3rd):
//! each case runs on VM3 and the JIT and matches Excel's answer.

use calcvba_differential::{Canon, Executor, run};

/// The `result` string of `Public result As String` after `Main`, or `Err` when the module
/// does not compile or run.
fn eval(executor: Executor, source: &str) -> Result<String, String> {
    let outcome = run(executor, source);
    if let Some(why) = outcome.unsupported {
        return Err(format!("declined: {why:?}"));
    }
    outcome.result.map_err(|e| e.to_string()).and_then(|vals| {
        vals.iter()
            .find_map(|v| match v {
                Canon::Str(s) => Some(s.clone()),
                _ => None,
            })
            .ok_or_else(|| format!("no string in {vals:?}"))
    })
}

fn both(source: &str) -> String {
    let vm3 = eval(Executor::Vm3, source).unwrap_or_else(|e| panic!("vm3: {e}\n{source}"));
    let jit = eval(Executor::Jit, source).unwrap_or_else(|e| panic!("jit: {e}\n{source}"));
    assert_eq!(vm3, jit, "vm3 and jit differ for\n{source}");
    vm3
}

fn rejected(source: &str) {
    for executor in [Executor::Vm3, Executor::Jit] {
        assert!(
            eval(executor, source).is_err(),
            "{executor:?} accepted\n{source}"
        );
    }
}

fn main_with(body: &str) -> String {
    format!("Public result As String\nSub Main()\n{body}\nEnd Sub\n")
}

#[test]
fn comment_continuation_swallows_the_next_line() {
    let src = main_with("    result = \"a\" ' note _\n    result = \"b\"");
    assert_eq!(both(&src), "a");
}

#[test]
fn caret_declares_longlong_but_still_exponentiates() {
    let src = "Dim x^\nPublic result As String\nSub Main()\n    x = 5\n    result = TypeName(x) & \":\" & x\nEnd Sub\n";
    assert_eq!(both(src), "LongLong:5");
    let src = main_with("    Dim y As Long: y = 3\n    result = CStr(y^2) & \" \" & CStr(y ^ 2)");
    assert_eq!(both(&src), "9 9");
}

#[test]
fn number_token_forms() {
    let src = main_with(
        "    result = TypeName(1D3) & CStr(1D3) & \" \" & TypeName(.5) & CStr(.5) & \" \" & \
         TypeName(&17) & CStr(&17) & \" \" & TypeName(99999999999999999999) & \" \" & \
         CStr(0.00015@) & \" \" & CStr(0.00025@) & \" \" & CStr(0.00005@)",
    );
    assert_eq!(
        both(&src),
        "Double1000 Double0.5 Integer15 Double 0.0002 0.0002 0"
    );
}

#[test]
fn conditional_compilation_rules() {
    let cont = "#Const A = 1 _\n    + 1\n#If A = 2 _\n  Then\nPublic result As String\n#End If\nSub Main()\n    result = \"two\"\nEnd Sub\n";
    assert_eq!(both(cont), "two");
    let vba6 = "#If VBA6 And Not Win16 And Not Mac Then\nPublic result As String\n#End If\nSub Main()\n    result = \"v\"\nEnd Sub\n";
    assert_eq!(both(vba6), "v");
    let endif =
        "#If True Then\nPublic result As String\n#EndIf\nSub Main()\n    result = \"e\"\nEnd Sub\n";
    assert_eq!(both(endif), "e");
    let suffix = "#Const X% = 1\n#If X Then\nPublic result As String\n#End If\nSub Main()\n    result = \"s\"\nEnd Sub\n";
    assert_eq!(both(suffix), "s");
    rejected(
        "#If Win64 Then\n#Const P = 8\n#Else\n#Const P = 4\n#End If\nPublic result As String\nSub Main()\nEnd Sub\n",
    );
    rejected(
        "#Const S = \"x\"\n#If S Then\n#End If\nPublic result As String\nSub Main()\nEnd Sub\n",
    );
}

#[test]
fn option_directives() {
    let body = "Public result As String\nSub Main()\n    Dim a(3)\n    result = CStr(LBound(a))\nEnd Sub\n";
    assert_eq!(both(&format!("Option Base 1&\n{body}")), "1");
    assert_eq!(both(&format!("Option Base 0\n{body}")), "0");
    rejected(&format!("Option Base 2\n{body}"));
    rejected(&format!("Option Compare Foo\n{body}"));
    rejected(&format!("Option Explicit\nOption Explicit\n{body}"));
}

#[test]
fn enum_initialiser_rounds_to_long() {
    let src = "Enum E1\n    A = 1.5\n    B\n    C = 2.5\nEnd Enum\nPublic result As String\nSub Main()\n    result = A & \" \" & B & \" \" & C\nEnd Sub\n";
    assert_eq!(both(src), "2 3 2");
}

#[test]
fn property_get_let_passing_modes_must_agree() {
    let prop = |let_mode: &str| {
        format!(
            "Private m As Long\nProperty Get P(ByVal i As Long) As Long\n    P = m + i\nEnd Property\n\
             Property Let P({let_mode} i As Long, ByVal v As Long)\n    m = v\nEnd Property\n\
             Public result As String\nSub Main()\n    P(1) = 4\n    result = CStr(P(1))\nEnd Sub\n"
        )
    };
    assert_eq!(both(&prop("ByVal")), "5");
    rejected(&prop("ByRef"));
}
