//! dc-ez3rd lane E: expressions, binding, arguments, Set and object typing. Each case is a
//! piece of `conformance/excel-measurements/spec-pass/SpecExpr.bas` (with the class SpecK) run
//! alone on VM3 and the JIT; the expected value is Excel's answer.

use calcvba_differential::{Canon, Executor, run_modules};
use calcvba_symbol::manifest::ModuleKind::{Class, Procedural};

const SPECK: &str = include_str!("../../../conformance/excel-measurements/spec-pass/SpecK.cls");
const SPEC_EXPR: &str =
    include_str!("../../../conformance/excel-measurements/spec-pass/SpecExpr.bas");

/// SpecExpr with the `Select Case` cut down to the one case `key`.
fn module_for(key: &str) -> String {
    let mut out = String::new();
    for line in SPEC_EXPR.lines() {
        let t = line.trim_start();
        if t.starts_with("Case \"") && !t.starts_with(&format!("Case \"{key}\"")) {
            continue;
        }
        out.push_str(line);
        out.push('\n');
    }
    out
}

fn answer(executor: Executor, key: &str) -> String {
    let main = format!(
        "Public result As String\nSub Main()\n    result = SpecExpr.T(\"{key}\")\nEnd Sub\n"
    );
    let source = module_for(key).replace("Private Function T(", "Public Function T(");
    let modules = [
        ("Main", Procedural, main.as_str()),
        ("SpecExpr", Procedural, source.as_str()),
        ("SpecK", Class, SPECK),
    ];
    let outcome = run_modules(executor, &modules, "VBAProject");
    if let Some(why) = outcome.unsupported {
        return format!("declined: {why:?}");
    }
    match outcome.result {
        Err(e) => format!("failed: {e}"),
        Ok(values) => values
            .iter()
            .find_map(|v| match v {
                Canon::Str(s) if s.contains('=') => Some(s.clone()),
                _ => None,
            })
            .unwrap_or_else(|| format!("no result in {values:?}")),
    }
}

fn check(key: &str, want: &str) {
    for executor in [Executor::Vm3, Executor::Jit] {
        assert_eq!(
            answer(executor, key),
            format!("{key}={want}"),
            "{executor:?}"
        );
    }
}

#[test]
fn bang_collection() {
    check("bangColl", "Integer:7");
}

#[test]
fn bang_inside_with() {
    check("bangWith", "Integer:8");
}

#[test]
fn index_zero_param_method() {
    check("idxZeroParamFn", "Integer:20");
}

#[test]
fn index_zero_param_property() {
    check("idxZeroParamProp", "String:[second]");
}

#[test]
fn index_zero_param_module_function() {
    check("idxZeroParamMod", "Integer:8");
}

#[test]
fn parameter_shadows_module_const() {
    check("paramVsConst", "Long:9");
}

#[test]
fn parenthesised_object_argument_reads_default_member() {
    check("parenObjArg", "ERR438");
}

#[test]
fn typeof_project_object_is_object() {
    check("typeOfObject", "Boolean:True");
}

#[test]
fn typeof_collection_is_object() {
    check("typeOfObjectColl", "Boolean:True");
}

#[test]
fn typeof_empty_is_false() {
    check("typeOfEmpty", "Boolean:False");
}

#[test]
fn typeof_nothing_is_error_91() {
    check("typeOfNothing", "ERR91");
}

#[test]
fn with_udt_element_subscript_evaluated_once() {
    check("withUdtIndex", "String:[5,0]");
}

#[test]
fn left_exception_calls_library_left() {
    check("leftException", "String:[ab]");
}

#[test]
fn empty_index_list_passes_whole_array() {
    check("emptyIndexArr", "Long:102");
}

#[test]
fn late_positional_argument_is_byref() {
    check("latePosByRef", "Long:2");
}

#[test]
fn class_field_passed_byref_is_a_copy() {
    check("byRefField", "Long:1");
}

#[test]
fn me_as_argument() {
    check("passMe", "String:[SpecK]");
}

#[test]
fn set_collection_into_class_variable_is_13() {
    check("setCollToClass", "ERR13");
}

#[test]
fn set_class_into_collection_variable_is_13() {
    check("setClassToColl", "ERR13");
}

#[test]
fn set_class_into_object_variable() {
    check("setClassToObj", "String:[SpecK]");
}

/// SpecTypes' `Set` of a non-object: 424 into an Object variable, 13 into a Variant; and a
/// member call on an Integer is 424.
#[test]
fn set_of_a_non_object() {
    let source = "Public result As String\n\
        Function T(ByVal k As String) As String\n\
        Dim vv As Variant, o As Object, w As Variant, v As Variant\n\
        On Error GoTo fail\n\
        vv = 5\n\
        Select Case k\n\
        Case \"obj\": Set o = vv\n\
        Case \"var\": Set w = vv\n\
        Case \"member\": v = vv.Foo\n\
        End Select\n\
        T = \"ok\": Exit Function\n\
        fail:\n\
        T = \"ERR\" & Err.Number\n\
        End Function\n\
        Sub Main()\n\
        result = T(\"obj\") & \" \" & T(\"var\") & \" \" & T(\"member\")\n\
        End Sub\n";
    for executor in [Executor::Vm3, Executor::Jit] {
        let outcome = run_modules(executor, &[("Main", Procedural, source)], "VBAProject");
        let values = outcome
            .result
            .unwrap_or_else(|e| panic!("{executor:?}: {e}"));
        assert_eq!(
            values.first(),
            Some(&Canon::Str("ERR424 ERR13 ERR424".into())),
            "{executor:?}"
        );
    }
}
