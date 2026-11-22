//! A `Let` from a Variant or Object source takes the default member of the object it holds,
//! once, with no arguments (`CoreValue::LetValue`). Expected lines are Excel VBA's (probe
//! temp/variant-layout-probe/let12, Excel 16.0 build 20527), on VM3 and the JIT.

use calcvba_differential::{Canon, Executor, run_modules};
use calcvba_symbol::manifest::ModuleKind::{Class, Procedural};

const DFLT: &str = "Public Property Get Value() As Variant\n\
Attribute Value.VB_UserMemId = 0\n\
    Value = 42\n\
End Property\n";
const CHAIN: &str = "Public Property Get Value() As Variant\n\
Attribute Value.VB_UserMemId = 0\n\
    Set Value = New Dflt\n\
End Property\n";
const NO_DFLT: &str = "Public Function Foo() As Long\n    Foo = 1\nEnd Function\n";
const INDEXED: &str = "Public Property Get Item(ByVal key As Variant) As Variant\n\
Attribute Item.VB_UserMemId = 0\n\
    If key = \"object\" Then Set Item = New Indexed Else Item = 42\n\
End Property\n";

/// `(case, statement, Excel VBA's line)`.
const CASES: &[(&str, &str, &str)] = &[
    ("var", "Set w = New Dflt: v = w", "Integer:42"),
    ("obj", "Set o = New Dflt: v = o", "Integer:42"),
    ("fn_variant", "v = MakeV()", "Integer:42"),
    ("fn_object", "v = MakeO()", "Integer:42"),
    ("long", "Set w = New Dflt: n = w: v = n", "Long:42"),
    ("string", "Set w = New Dflt: s = w: v = s", "String:42"),
    ("chain", "Set w = New Chain: v = w", "obj Dflt"),
    (
        "nodflt",
        "Set w = New NoDflt: v = w",
        "raised 438 Object doesn't support this property or method",
    ),
    (
        "collection",
        "Set w = New Collection: w.Add 5: v = w",
        "raised 450 Wrong number of arguments or invalid property assignment",
    ),
    (
        "nothing",
        "Set w = Nothing: v = w",
        "raised 91 Object variable or With block variable not set",
    ),
    ("empty", "w = Empty: v = w", "Empty:"),
    (
        "indexed_object",
        "Set o = New Indexed: v = o(\"object\")",
        "raised 449 Argument not optional",
    ),
    (
        "indexed_number",
        "Set o = New Indexed: v = o(\"number\")",
        "Integer:42",
    ),
    ("return", "v = RetLet()", "Integer:42"),
    (
        "array_elem",
        "Set w = New Dflt: a(0) = w: v = a(0)",
        "Integer:42",
    ),
    (
        "byval_variant",
        "Set w = New Dflt: v = TypeOfArg(w)",
        "String:Dflt",
    ),
    (
        "object_copy",
        "Set w = New Dflt: Set o2 = w: v = o2",
        "Integer:42",
    ),
];

fn main_module() -> String {
    let mut src = String::from(
        "Public result As String\n\
Function MakeV() As Variant\n    Set MakeV = New Dflt\nEnd Function\n\
Function MakeO() As Object\n    Set MakeO = New Dflt\nEnd Function\n\
Function RetLet() As Variant\n    Dim w As Variant\n    Set w = New Dflt\n    RetLet = w\nEnd Function\n\
Function TypeOfArg(ByVal x As Variant) As String\n    TypeOfArg = TypeName(x)\nEnd Function\n\
Private Function Try(ByVal what As String) As String\n\
    Dim w As Variant, v As Variant, o As Object, o2 As Object, n As Long, s As String\n\
    Dim a(0) As Variant\n\
    On Error GoTo fail\n\
    Select Case what\n",
    );
    for (case, statement, _) in CASES {
        src.push_str(&format!("        Case \"{case}\": {statement}\n"));
    }
    src.push_str(
        "    End Select\n\
    If IsObject(v) Then Try = \"obj \" & TypeName(v) Else Try = TypeName(v) & \":\" & CStr(v)\n\
    Exit Function\n\
fail:\n\
    Try = \"raised \" & Err.Number & \" \" & Err.Description\n\
End Function\n\
Sub Main()\n",
    );
    for (case, _, _) in CASES {
        src.push_str(&format!(
            "    result = result & \"{case}=\" & Try(\"{case}\") & vbLf\n"
        ));
    }
    src.push_str("End Sub\n");
    src
}

fn check(executor: Executor) {
    let main = main_module();
    let modules = [
        ("Main", Procedural, main.as_str()),
        ("Dflt", Class, DFLT),
        ("Chain", Class, CHAIN),
        ("NoDflt", Class, NO_DFLT),
        ("Indexed", Class, INDEXED),
    ];
    let outcome = run_modules(executor, &modules, "VBAProject");
    assert!(
        outcome.unsupported.is_none(),
        "{executor:?} declined: {:?}",
        outcome.unsupported
    );
    let values = outcome
        .result
        .unwrap_or_else(|e| panic!("{executor:?} failed: {e}"));
    let text = values
        .iter()
        .find_map(|v| match v {
            Canon::Str(s) if s.contains("var=") => Some(s.clone()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("{executor:?}: no result in {values:?}"));
    let got: Vec<&str> = text.lines().collect();
    let wrong: Vec<String> = CASES
        .iter()
        .zip(&got)
        .filter_map(|((case, _, want), line)| {
            let value = line.strip_prefix(&format!("{case}=")).unwrap_or(line);
            (value != *want).then(|| format!("{case}: VBA {want:?}, {executor:?} {value:?}"))
        })
        .collect();
    assert_eq!(got.len(), CASES.len(), "{executor:?}: {text}");
    assert!(wrong.is_empty(), "{}", wrong.join("\n"));
}

#[test]
fn let_takes_the_default_member_as_vba_does_on_vm3() {
    check(Executor::Vm3);
}

#[test]
fn let_takes_the_default_member_as_vba_does_on_the_jit() {
    check(Executor::Jit);
}
