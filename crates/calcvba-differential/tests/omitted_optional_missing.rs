//! An omitted `Optional ... As Variant` parameter without a default is Missing however the
//! call is made: late bound or early bound, with or without an argument list, and through a
//! default member (dc-mxsg8). One with a default, or of a declared type, is never Missing: a
//! declared default reaches the callee on every call shape, coerced to the parameter type
//! (dc-wmjmb, measured in Excel 2026-11-16).

use calcvba_differential::{Canon, Executor, run_modules};
use calcvba_symbol::manifest::ModuleKind::{Class, Procedural};

const OPT: &str = "Public Function Item(Optional ByVal i As Variant) As Variant\n\
Attribute Item.VB_UserMemId = 0\n\
    If IsMissing(i) Then Item = \"missing\" Else Item = \"got \" & TypeName(i)\n\
End Function\n\
Public Function WithDefault(Optional ByVal i As Variant = 3) As Variant\n\
    If IsMissing(i) Then WithDefault = \"missing\" Else WithDefault = \"got \" & CStr(i)\n\
End Function\n\
Public Function Typed(Optional ByVal n As Long) As Variant\n\
    Typed = \"got \" & TypeName(n) & \" \" & CStr(n)\n\
End Function\n\
Public Function TypedDefault(Optional ByVal n As Long = 5) As Variant\n\
    TypedDefault = \"got \" & TypeName(n) & \" \" & CStr(n)\n\
End Function\n";

const MAIN: &str = "Public result As String\n\
Sub Main()\n\
    Dim o As Object, e As Opt, v As Variant\n\
    Set o = New Opt: Set e = o\n\
    result = result & \"late=\" & o.Item & vbLf\n\
    result = result & \"late_parens=\" & o.Item() & vbLf\n\
    result = result & \"early=\" & e.Item & vbLf\n\
    result = result & \"early_parens=\" & e.Item() & vbLf\n\
    result = result & \"default_call=\" & e() & vbLf\n\
    v = o\n\
    result = result & \"let_default=\" & v & vbLf\n\
    result = result & \"early_with_default=\" & e.WithDefault() & vbLf\n\
    result = result & \"early_typed=\" & e.Typed & vbLf\n\
    result = result & \"late_typed=\" & o.Typed & vbLf\n\
    result = result & \"early_default_noparens=\" & e.WithDefault & vbLf\n\
    result = result & \"late_default=\" & o.WithDefault & vbLf\n\
    result = result & \"callbyname_default=\" & CallByName(o, \"WithDefault\", VbMethod) & vbLf\n\
    result = result & \"early_typed_default=\" & e.TypedDefault & vbLf\n\
    result = result & \"late_typed_default=\" & o.TypedDefault & vbLf\n\
    result = result & \"early_typed_default_parens=\" & e.TypedDefault() & vbLf\n\
End Sub\n";

const EXPECTED: &[&str] = &[
    "late=missing",
    "late_parens=missing",
    "early=missing",
    "early_parens=missing",
    "default_call=missing",
    "let_default=missing",
    "early_with_default=got 3",
    "early_typed=got Long 0",
    "late_typed=got Long 0",
    "early_default_noparens=got 3",
    "late_default=got 3",
    "callbyname_default=got 3",
    "early_typed_default=got Long 5",
    "late_typed_default=got Long 5",
    "early_typed_default_parens=got Long 5",
];

fn check(executor: Executor) {
    let modules = [("Main", Procedural, MAIN), ("Opt", Class, OPT)];
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
            Canon::Str(s) if s.contains("late=") => Some(s.clone()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("{executor:?}: no result in {values:?}"));
    assert_eq!(text.lines().collect::<Vec<_>>(), EXPECTED, "{executor:?}");
}

#[test]
fn omitted_variant_optionals_are_missing_on_vm3() {
    check(Executor::Vm3);
}

#[test]
fn omitted_variant_optionals_are_missing_on_the_jit() {
    check(Executor::Jit);
}
