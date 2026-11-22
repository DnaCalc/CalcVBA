//! dc-jrthq lane H, measured in Excel VBA (`conformance/excel-measurements/spec-pass/SpecExpr`):
//! a named argument in a late-bound call passes ByRef like a positional one, and an omitted
//! ParamArray element is Missing (vbError DISP_E_PARAMNOTFOUND), not Empty.

use calcvba_differential::{Canon, Executor, run_modules};
use calcvba_symbol::manifest::ModuleKind::{Class, Procedural};

const K: &str = "Option Explicit\n\
Public Sub Bump(n As Long)\n    n = n + 1\nEnd Sub\n\
Public Property Get Number() As Variant\n    Number = 41\nEnd Property\n\
Public Sub Bump2(a As Long, b As Variant)\n    a = a + 10\n    b = b & \"!\"\nEnd Sub\n";

fn result(executor: Executor, body: &str) -> String {
    let main = format!("Option Explicit\nPublic result As String\n{body}");
    let outcome = run_modules(
        executor,
        &[("Main", Procedural, main.as_str()), ("K", Class, K)],
        "VBAProject",
    );
    assert!(
        outcome.unsupported.is_none(),
        "{executor:?} declined: {:?}",
        outcome.unsupported
    );
    let values = outcome
        .result
        .unwrap_or_else(|e| panic!("{executor:?} failed: {e}"));
    values
        .iter()
        .find_map(|v| match v {
            Canon::Str(s) => Some(s.clone()),
            _ => None,
        })
        .unwrap_or_else(|| panic!("{executor:?}: no string in {values:?}"))
}

fn both(body: &str, expected: &str) {
    for executor in [Executor::Vm3, Executor::Jit] {
        assert_eq!(result(executor, body), expected, "{executor:?}");
    }
}

#[test]
fn late_bound_named_argument_passes_byref() {
    both(
        "Sub Main()\n    Dim o As Object, lk As Long\n    Set o = New K\n    lk = 1\n    \
         o.Bump n:=lk\n    result = CStr(lk)\nEnd Sub\n",
        "2",
    );
}

#[test]
fn late_bound_named_arguments_out_of_order_pass_byref() {
    both(
        "Sub Main()\n    Dim o As Object, x As Long, s As Variant\n    Set o = New K\n    \
         x = 1: s = \"a\"\n    o.Bump2 b:=s, a:=x\n    result = x & \" \" & s\nEnd Sub\n",
        "11 a!",
    );
}

#[test]
fn late_bound_named_argument_expression_is_byval() {
    both(
        "Sub Main()\n    Dim o As Object, lk As Long\n    Set o = New K\n    lk = 1\n    \
         o.Bump n:=(lk)\n    result = CStr(lk)\nEnd Sub\n",
        "1",
    );
}

#[test]
fn omitted_paramarray_element_is_missing() {
    both(
        "Private Function PA(ParamArray r() As Variant) As String\n    \
         PA = VarType(r(1)) & \" \" & IsMissing(r(1)) & \" \" & IsError(r(1)) & \" \" & \
         IsMissing(r(0))\nEnd Function\n\
         Sub Main()\n    result = PA(1, , 3)\nEnd Sub\n",
        "10 True True False",
    );
}

#[test]
fn omitted_optional_variant_is_still_missing() {
    both(
        "Private Function F(Optional v As Variant) As String\n    \
         F = VarType(v) & \" \" & IsMissing(v)\nEnd Function\n\
         Sub Main()\n    result = F()\nEnd Sub\n",
        "10 True",
    );
}

#[test]
fn late_bound_value_to_byref_long_parameter_is_a_private_copy() {
    both(
        "Sub Main()\n    Dim o As Object, lk As Long\n    Set o = New K\n    lk = 1\n    \
         o.Bump (lk)\n    o.Bump n:=5\n    result = CStr(lk)\nEnd Sub\n",
        "1",
    );
}

#[test]
fn property_value_to_static_byref_long_parameter_is_a_private_copy() {
    both(
        "Private Function AddOne(n As Long) As Long\n    n = n + 1\n    AddOne = n\nEnd Function\n\
         Sub Main()\n    Dim o As Object\n    Set o = New K\n    \
         result = AddOne(o.Number) & \"|\" & o.Number\nEnd Sub\n",
        "42|41",
    );
}
