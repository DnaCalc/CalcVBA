//! Compile decisions measured through Excel's visible VBE Compile command on 2026-11-18.
//! Len's typed expressions and Variant expressions have different syntax rules.

use calcvba_bind::{BindError, bind_program};
use calcvba_symbol::{CatalogTypeLibResolver, manifest as sym};

fn compile(body: &str, declarations: &str) -> Result<(), BindError> {
    let source = format!(
        "Option Explicit\nPublic Function Run() As String\n{body}\nEnd Function\n{declarations}"
    );
    let manifest = sym::SymbolProjectManifest {
        project_name: "VBAProject".into(),
        project_kind: sym::ProjectKind::Source,
        modules: vec![sym::ModuleUnit {
            module_name: "M".into(),
            module_kind: sym::ModuleKind::Procedural,
            attributes: sym::ModuleAttributes::named("M"),
            source,
        }],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
        codepage: None,
    };
    bind_program(&manifest, &CatalogTypeLibResolver).map(|_| ())
}

#[test]
fn len_numeric_expressions_require_a_variable() {
    for expression in [
        "123",
        "1.5",
        "True",
        "1 + 2",
        "CInt(5)",
        "#1/2/2020#",
        "Sqr(4)",
        "Abs(-1)",
        "Abs(CInt(5))",
        "Abs(Abs(-1))",
    ] {
        for name in ["Len", "LenB"] {
            let body = format!("Run = CStr({name}({expression}))");
            assert!(
                matches!(compile(&body, ""), Err(BindError::VariableRequired)),
                "{body}"
            );
        }
    }
    assert!(matches!(
        compile("Const n As Long = 5\nRun = CStr(Len(n))", ""),
        Err(BindError::VariableRequired)
    ));
    assert!(matches!(
        compile(
            "Run = CStr(Len(F()))",
            "Private Function F() As Long\nF = 5\nEnd Function\n"
        ),
        Err(BindError::VariableRequired)
    ));
}

#[test]
fn variant_len_and_array_conversions_remain_valid() {
    for body in [
        "Run = CStr(Len(CVar(5)))",
        "Run = CStr(Len(CVDate(1)))",
        "Run = CStr(Len([CVDate](1)))",
        "Run = CStr(Len(VBA.CVDate(1)))",
        "Dim v As Variant\nv = -1\nRun = CStr(Len(Abs(v)))",
        "Run = CStr(Len(Abs(CVar(-1))))",
        "Dim v As Variant\nv = #1/2/2020#\nRun = CStr(Len(Year(v)))",
        "Run = CStr(Len(Year(#1/2/2020#)))",
        "Run = CStr(Len(IIf(True, 1, 2)))",
        "Dim v As Variant\nv = 5\nRun = CStr(Len(v + 1))",
        "Dim a(0 To 1) As Long, v As Variant\nv = CVar(a)\nRun = TypeName(v)",
        "Dim a(0 To 1) As Long, v As Variant\nv = CVar(a())\nRun = TypeName(v)",
        "Dim a As Variant, v As Variant\na = Array(1, 2)\nv = CVar(a)\nRun = TypeName(v)",
        "Run = CStr(LBound(Array(1, 2)))",
    ] {
        compile(body, "").unwrap_or_else(|err| panic!("{body}: {err}"));
    }
    compile(
        "Run = CStr(Len(F()))",
        "Private Function F() As Variant\nF = 5\nEnd Function\n",
    )
    .unwrap();
}

#[test]
fn bound_conversions_and_call_do_events_are_syntax_errors() {
    for body in [
        "Dim a(0 To 1) As Long\nRun = CStr(LBound(CVar(a)))",
        "Dim a As Variant\na = Array(1, 2)\nRun = CStr(LBound(CVar(a)))",
        "Run = CStr(LBound(CVar(Array(1, 2))))",
        "Call DoEvents\nRun = \"ok\"",
        "Call DoEvents()\nRun = \"ok\"",
    ] {
        assert!(
            matches!(compile(body, ""), Err(BindError::SyntaxError)),
            "{body}"
        );
    }
}

#[test]
fn a_trailing_omission_is_not_an_optional_argument() {
    for (body, declarations) in [
        ("Run = CStr(DDB(2400, 300, 10, 1, ))", ""),
        (
            "Run = CStr(F(1, ))",
            "Private Function F(ByVal x As Long, Optional ByVal y As Long = 5) As Long\nF = x + y\nEnd Function\n",
        ),
    ] {
        let err = compile(body, declarations).expect_err(body);
        assert!(
            err.to_string().contains("expected argument after `,`"),
            "{body}: {err}"
        );
    }
}
