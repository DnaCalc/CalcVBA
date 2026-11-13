//! Shapes the Riff and Wasabi libraries need to load and run (dc-tey4.5).

use calcvba_engine::{Engine, HostConfig};
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};

fn run_with(config: HostConfig, main_src: &str) -> Result<Vec<calcvba_runtime::Variant>, String> {
    let manifest = SymbolProjectManifest {
        project_name: "Proj".into(),
        project_kind: ProjectKind::Source,
        modules: vec![ModuleUnit {
            module_name: "Main".into(),
            module_kind: ModuleKind::Procedural,
            attributes: ModuleAttributes::named("Main"),
            source: main_src.into(),
        }],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    Engine::new(config)
        .execute_manifest_with_variant_snapshot(&manifest)
        .map_err(|d| format!("{:?}: {}", d.phase(), d.message()))
}

fn both(src: &str, check: impl Fn(&[calcvba_runtime::Variant]) -> bool) {
    for (name, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let v = run_with(config, src).unwrap_or_else(|e| panic!("{name}: {e}"));
        assert!(check(&v), "{name}: {v:?}");
    }
}

#[test]
fn call_statement_with_a_with_member_argument() {
    both(
        "Private Type Rec\n    pass As String\nEnd Type\nPublic verdict As Long\nPrivate Sub Zero(ByRef s As String)\n    If Len(s) > 0 Then s = vbNullString\nEnd Sub\nSub Main()\n    Dim r As Rec\n    r.pass = \"abc\"\n    With r\n        Zero .pass\n    End With\n    verdict = Len(r.pass) + 10\nEnd Sub\n",
        |v| v.iter().any(|x| x.as_i32() == Some(10)),
    );
}

/// Survey, not a check: `VBA_FILES=Name:class|std:path;...` loads those sources as one project on both
/// executors and prints the full diagnostic (the corpus runner keeps only its first 240 characters).
#[test]
#[ignore]
fn survey_project_from_files() {
    let Ok(files) = std::env::var("VBA_FILES") else {
        return;
    };
    let modules: Vec<ModuleUnit> = files
        .split(';')
        .map(|spec| {
            let mut parts = spec.splitn(3, ':');
            let (name, kind, path) = (
                parts.next().unwrap(),
                parts.next().unwrap(),
                parts.next().unwrap(),
            );
            let bytes = std::fs::read(path).unwrap();
            let source: String = match String::from_utf8(bytes.clone()) {
                Ok(s) => s,
                Err(_) => bytes.iter().map(|b| char::from(*b)).collect(),
            };
            ModuleUnit {
                module_name: name.into(),
                module_kind: if kind == "class" {
                    ModuleKind::Class
                } else {
                    ModuleKind::Procedural
                },
                attributes: ModuleAttributes::named(name),
                source,
            }
        })
        .collect();
    for (label, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let manifest = SymbolProjectManifest {
            project_name: "Proj".into(),
            project_kind: ProjectKind::Source,
            modules: modules.clone(),
            references: Vec::new(),
            reference_projects: Vec::new(),
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
        };
        let out = Engine::new(config).execute_manifest_with_variant_snapshot(&manifest);
        match out {
            Ok(_) => println!("{label}: ran"),
            Err(d) => println!("{label}: {:?}: {}", d.phase(), d.message()),
        }
    }
}

#[test]
fn open_statement_with_a_with_member_path() {
    both(
        "Private Type Rec\n    path As String\nEnd Type\nPublic verdict As Long\nSub Main()\n    Dim r As Rec\n    Dim n As Integer\n    r.path = \"/nonexistent-dir-p10f/none.bin\"\n    n = FreeFile\n    On Error Resume Next\n    With r\n        Open .path For Binary Access Read As #n\n    End With\n    verdict = Err.Number\n    Close #n\nEnd Sub\n",
        |v| v.iter().any(|x| matches!(x.as_i32(), Some(5 | 53 | 76))),
    );
}

const CONN: &str = "Private Type Conn\n    buf() As Byte\n    n As Long\nEnd Type\nPrivate m(0 To 3) As Conn\nPublic verdict As Long\n";

#[test]
fn with_array_element_udt_indexes_an_array_field() {
    for (name, body, expected) in [
        (
            "varptr_index",
            "With m(h)\n  verdict = VarPtr(.buf(0)) <> 0\nEnd With",
            -1,
        ),
        (
            "ubound",
            "With m(h)\n  verdict = UBound(.buf) + 1\nEnd With",
            5,
        ),
        (
            "index_read",
            "With m(h)\n  verdict = .buf(1) + 21\nEnd With",
            21,
        ),
        (
            "index_write",
            "With m(h)\n  .buf(1) = 7\n  verdict = .buf(1)\nEnd With",
            7,
        ),
        (
            "redim_field",
            "With m(h)\n  ReDim .buf(0 To 9)\n  verdict = UBound(.buf)\nEnd With",
            9,
        ),
    ] {
        let src = format!(
            "{CONN}Sub Main()\n  Dim h As Long\n  h = 2\n  ReDim m(h).buf(0 To 4)\n{body}\nEnd Sub\n"
        );
        for (label, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
            let v = run_with(config, &src).unwrap_or_else(|e| panic!("{name} {label}: {e}"));
            assert!(
                v.iter().any(|x| x.as_i32() == Some(expected)),
                "{name} {label}: {v:?}"
            );
        }
    }
}

#[test]
fn record_array_field_indexed_by_another_record_array_field() {
    // Riff: `busVol = rCtx.Buses(rVoices(i).busID)`
    let src = "Private Type Ctx\n    Buses(0 To 7) As Single\n    n As Long\nEnd Type\nPrivate Type Voice\n    busID As Long\nEnd Type\nPrivate rCtx As Ctx\nPrivate rVoices(0 To 3) As Voice\nPublic verdict As Long\nSub Main()\n    Dim i As Long\n    Dim busVol As Single\n    rCtx.Buses(5) = 2.5!\n    rVoices(1).busID = 5\n    i = 1\n    busVol = rCtx.Buses(rVoices(i).busID)\n    verdict = CLng(busVol * 2)\nEnd Sub\n";
    both(src, |v| v.iter().any(|x| x.as_i32() == Some(5)));
}

const ITEMS: &str = "Private Type Conn\n    Items As Collection\nEnd Type\nPrivate m(0 To 3) As Conn\nPublic verdict As Long\nPrivate Sub Fill(ByVal h As Long)\n    Set m(h).Items = New Collection\n    m(h).Items.Add New Collection\n    m(h).Items.Add New Collection\nEnd Sub\n";

#[test]
fn collection_field_of_an_array_element_udt_adds_and_counts() {
    let src = format!("{ITEMS}Sub Main()\n    Fill 1\n    verdict = m(1).Items.Count\nEnd Sub\n");
    both(&src, |v| v.iter().any(|x| x.as_i32() == Some(2)));
}

#[test]
fn for_each_over_a_collection_field_of_an_array_element_udt() {
    // Wasabi: `For Each mw In m_Connections(handle).Middlewares` with `Dim mw As Object`
    let src = format!(
        "{ITEMS}Sub Main()\n    Dim mw As Object\n    Dim h As Long\n    h = 1\n    Fill h\n    For Each mw In m(h).Items\n        verdict = verdict + 1\n    Next mw\nEnd Sub\n"
    );
    both(&src, |v| v.iter().any(|x| x.as_i32() == Some(2)));
}

#[test]
fn module_level_enum_variable_is_a_long() {
    let src = "Option Explicit\nOption Private Module\nPublic Enum Err2\n    E_NONE = 0\n    E_BAD = 7\nEnd Enum\nPrivate m_Last As Err2\nPublic verdict As Long\nPrivate Sub SetIt(ByVal e As Err2)\n    m_Last = e\nEnd Sub\nSub Main()\n    m_Last = E_NONE\n    SetIt E_BAD\n    verdict = m_Last\nEnd Sub\n";
    both(src, |v| v.iter().any(|x| x.as_i32() == Some(7)));
}

const MIDCLS: &str = "Private s As String\nPublic Property Get Mid(ByVal Start As Long, Optional ByVal Length As Variant) As String\n    If IsMissing(Length) Then Length = Len(s) - Start + 1\n    Mid = VBA.Mid$(s, Start, Length)\nEnd Property\nPublic Property Let Mid(ByVal Start As Long, Optional ByVal Length As Variant, ByVal RHS As String)\n    If IsMissing(Length) Then Length = Len(RHS)\n    Mid$(s, Start, Length) = RHS\nEnd Property\nPublic Sub Init(ByVal v As String)\n    s = v\nEnd Sub\nPublic Function Peek() As String\n    Peek = s\nEnd Function\n";

#[test]
fn property_let_with_an_optional_middle_parameter() {
    // vba-stringbuilder: `sb.Mid(3) = "ZZ"` calls Property Let Mid(Start, Optional Length, RHS)
    let class_src = MIDCLS;
    for (label, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let manifest = SymbolProjectManifest {
            project_name: "Proj".into(),
            project_kind: ProjectKind::Source,
            modules: vec![
                ModuleUnit {
                    module_name: "Main".into(),
                    module_kind: ModuleKind::Procedural,
                    attributes: ModuleAttributes::named("Main"),
                    source: "Public verdict As Long\nSub Main()\n    Dim c As New Cls\n    c.Init \"hello world\"\n    c.Mid(1) = \"J\"\n    c.Mid(7, 2) = \"WORLD\"\n    If c.Peek = \"Jello WOrld\" Then verdict = 1\nEnd Sub\n".into(),
                },
                ModuleUnit {
                    module_name: "Cls".into(),
                    module_kind: ModuleKind::Class,
                    attributes: ModuleAttributes::named("Cls"),
                    source: class_src.into(),
                },
            ],
            references: Vec::new(),
            reference_projects: Vec::new(),
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
        };
        let v = Engine::new(config)
            .execute_manifest_with_variant_snapshot(&manifest)
            .unwrap_or_else(|d| panic!("{label}: {:?}: {}", d.phase(), d.message()));
        assert!(v.iter().any(|x| x.as_i32() == Some(1)), "{label}: {v:?}");
    }
}

#[test]
fn variable_named_compare_and_a_call_continued_over_lines() {
    // VbaRegex: `compare = StrComp( _ identifier, _ Mid$(...), _ vbBinaryCompare _ )` and `If compare < 0`
    let src = "Public verdict As Long\nSub Main()\n    Dim compare As Long\n    Dim identifier As String\n    identifier = \"abc\"\n    compare = StrComp( _\n        identifier, _\n        Mid$(\"xabcx\", 2, 3), _\n        vbBinaryCompare _\n    )\n    If compare < 0 Then\n        verdict = 1\n    ElseIf compare > 0 Then\n        verdict = 2\n    Else\n        verdict = 3\n    End If\nEnd Sub\n";
    both(src, |v| v.iter().any(|x| x.as_i32() == Some(3)));
}

fn run_two(main_src: &str, class_src: &str) -> Result<Vec<calcvba_runtime::Variant>, String> {
    let manifest = SymbolProjectManifest {
        project_name: "Proj".into(),
        project_kind: ProjectKind::Source,
        modules: vec![
            ModuleUnit {
                module_name: "Main".into(),
                module_kind: ModuleKind::Procedural,
                attributes: ModuleAttributes::named("Main"),
                source: main_src.into(),
            },
            ModuleUnit {
                module_name: "Cls".into(),
                module_kind: ModuleKind::Class,
                attributes: ModuleAttributes::named("Cls"),
                source: class_src.into(),
            },
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    Engine::new(HostConfig::vm3())
        .execute_manifest_with_variant_snapshot(&manifest)
        .map_err(|d| format!("{:?}: {}", d.phase(), d.message()))
}

#[test]
fn public_enum_of_a_class_module_is_visible_to_other_modules() {
    // vba-test: `Public Enum TestResultType` lives in the TestSuite class; other modules write
    // `TestResultType.Pass` (and a bare member name).
    let cls = "Public Enum Kind\n    Alpha = 3\n    Beta = 8\nEnd Enum\nPublic Function Pick() As Kind\n    Pick = Beta\nEnd Function\n";
    let main = "Public verdict As Long\nSub Main()\n    Dim c As New Cls\n    Dim k As Kind\n    k = c.Pick\n    If k = Kind.Beta Then verdict = Kind.Alpha + 1\n    If k = Beta Then verdict = verdict + 100\nEnd Sub\n";
    let v = run_two(main, cls).expect("runs");
    assert!(v.iter().any(|x| x.as_i32() == Some(104)), "{v:?}");
}

#[test]
fn library_constant_named_through_its_enum() {
    // VBA-KeyedCollection: `CompareMode = VBA.VbCompareMethod.vbTextCompare`
    let src = "Public verdict As Long\nSub Main()\n    verdict = VBA.VbCompareMethod.vbTextCompare + 10 * VBA.vbBinaryCompare\nEnd Sub\n";
    both(src, |v| v.iter().any(|x| x.as_i32() == Some(1)));
}

#[test]
fn optional_parameter_default_continued_on_the_next_line() {
    // vba-test: `Optional Message As String = _ <newline> "Expected {0} to equal {1}"`
    let src = "Public verdict As Long\nPrivate Function Msg(ByVal a As Long, Optional Message As String = _\n    \"Expected {0}\") As String\n    Msg = Message\nEnd Function\nSub Main()\n    verdict = Len(Msg(1))\nEnd Sub\n";
    both(src, |v| v.iter().any(|x| x.as_i32() == Some(12)));
}

#[test]
fn call_statement_whose_first_argument_is_a_comparison() {
    // vba-test: `Check Err.Number = 0, Message, Text` is a call, not an assignment to `Err.Number`
    let src = "Public verdict As Long\nPrivate Sub Check(ByVal ok As Boolean, ByVal msg As String)\n    If ok Then verdict = verdict + Len(msg)\nEnd Sub\nSub Main()\n    Dim n As Long\n    n = 0\n    Check n = 0, \"abc\"\n    Check n = 1, \"xyz\"\n    Check Len(\"q\") = 1, \"de\"\nEnd Sub\n";
    both(src, |v| v.iter().any(|x| x.as_i32() == Some(5)));
}

#[test]
fn single_line_with_block() {
    // VbaRegex: `With outEntireMatch: .start = 0: .Length = 0: End With`
    let src = "Private Type Span\n    start As Long\n    Length As Long\nEnd Type\nPublic verdict As Long\nSub Main()\n    Dim m As Span\n    m.start = 5\n    With m: .start = 3: .Length = 4: End With\n    verdict = m.start * 10 + m.Length\nEnd Sub\n";
    both(src, |v| v.iter().any(|x| x.as_i32() == Some(34)));
}

#[test]
fn bracketed_hidden_enum_members_are_referenced_by_name() {
    // VBA-ArrayTools: `[_opMin] = opEqual` ... `For i = [_opMin] To [_opMax]`
    let src = "Public Enum Op\n    opNone = 0\n    opEqual = 1\n    [_opMin] = opEqual\n    opLess = 2\n    opLike = 3\n    [_opMax] = opLike\nEnd Enum\nPublic verdict As Long\nSub Main()\n    Dim i As Long\n    For i = [_opMin] To [_opMax]\n        verdict = verdict + i\n    Next i\nEnd Sub\n";
    both(src, |v| v.iter().any(|x| x.as_i32() == Some(6)));
}

#[test]
fn vba_library_module_named_without_the_vba_prefix() {
    // VBA-ArrayTools: `Math.Abs(diff)`, `Math.Sgn(diff)`
    let src = "Option Explicit\nPublic verdict As Long\nSub Main()\n    Dim d As Long\n    d = -7\n    verdict = Math.Abs(d) * 10 + Math.Sgn(d) + Len(Strings.Left(\"abcdef\", 3))\nEnd Sub\n";
    both(src, |v| v.iter().any(|x| x.as_i32() == Some(72)));
}

fn run_modules(modules: &[(&str, &str)]) -> Result<Vec<calcvba_runtime::Variant>, String> {
    let manifest = SymbolProjectManifest {
        project_name: "Proj".into(),
        project_kind: ProjectKind::Source,
        modules: modules
            .iter()
            .map(|(name, source)| ModuleUnit {
                module_name: (*name).into(),
                module_kind: ModuleKind::Procedural,
                attributes: ModuleAttributes::named(*name),
                source: (*source).into(),
            })
            .collect(),
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    Engine::new(HostConfig::vm3())
        .execute_manifest_with_variant_snapshot(&manifest)
        .map_err(|d| format!("{:?}: {}", d.phase(), d.message()))
}

#[test]
fn enum_initialisers_name_members_of_modules_declared_later() {
    // VbaRegex: `PSF_NONE = RegexNumericConstants.LONG_MIN` in a module that precedes it.
    let consts = "Public Enum NumericConstantLong\n    LONG_FIRST_BIT = &H80000000\n    LONG_ALL_BUT_FIRST_BIT = Not LONG_FIRST_BIT\n    LONG_MIN = &H80000000\n    LONG_MAX = &H7FFFFFFF\n    LONG_MAX_DIV_10 = LONG_MAX \\ 10\nEnd Enum\n";
    let ast = "Public Const MAX_AST_CODE As Long = 40\n";
    let main = "Public verdict As Long\nSub Main()\n    verdict = Go()\nEnd Sub\n";
    for (name, user, expected) in [
        (
            "long_min",
            "Private Enum E\n    PSF_NONE = Consts.LONG_MIN\nEnd Enum\nPublic Function Go() As Long\n    Go = PSF_NONE\nEnd Function\n",
            i32::MIN,
        ),
        (
            "long_max",
            "Private Enum E\n    H = Consts.LONG_MAX\nEnd Enum\nPublic Function Go() As Long\n    Go = H\nEnd Function\n",
            i32::MAX,
        ),
        (
            "neg_const",
            "Private Enum E\n    A = -(Ast.MAX_AST_CODE + 1)\n    B = A\nEnd Enum\nPublic Function Go() As Long\n    Go = B\nEnd Function\n",
            -41,
        ),
        (
            "not_member",
            "Private Enum E\n    X = Consts.LONG_ALL_BUT_FIRST_BIT\nEnd Enum\nPublic Function Go() As Long\n    Go = X\nEnd Function\n",
            i32::MAX,
        ),
    ] {
        for order in [0, 1] {
            let modules: Vec<(&str, &str)> = if order == 0 {
                vec![
                    ("Main", main),
                    ("User", user),
                    ("Consts", consts),
                    ("Ast", ast),
                ]
            } else {
                vec![
                    ("Main", main),
                    ("Consts", consts),
                    ("Ast", ast),
                    ("User", user),
                ]
            };
            let v = run_modules(&modules).unwrap_or_else(|e| panic!("{name} order {order}: {e}"));
            assert!(
                v.iter().any(|x| x.as_i32() == Some(expected)),
                "{name} order {order}: {v:?}"
            );
        }
    }
}

#[test]
fn library_enum_constant_named_without_the_vba_prefix() {
    // VBA-ArrayTools: `VarType(x) = VbVarType.vbString`
    let src = "Option Explicit\nPublic verdict As Long\nSub Main()\n    Dim s As Variant\n    s = \"abc\"\n    If VarType(s) = VbVarType.vbString Then verdict = 1\n    If VbCompareMethod.vbTextCompare = 1 Then verdict = verdict + 10\nEnd Sub\n";
    both(src, |v| v.iter().any(|x| x.as_i32() == Some(11)));
}

#[test]
fn event_and_property_of_one_name_in_a_class() {
    // vba-test's TestSuite: `Public Event Result(...)` and `Public Property Get Result()`
    let cls = "Public Event Result(ByVal n As Long)\nPublic Property Get Result() As Long\n    Result = 7\nEnd Property\nPublic Sub Fire()\n    RaiseEvent Result(1)\nEnd Sub\n";
    let main = "Public verdict As Long\nSub Main()\n    Dim c As New Cls\n    verdict = c.Result + 1\n    c.Fire\nEnd Sub\n";
    let v = run_two(main, cls).expect("runs");
    assert!(v.iter().any(|x| x.as_i32() == Some(8)), "{v:?}");
}

#[test]
fn property_get_fills_its_result_through_a_byref_property_let() {
    // VBA-KeyedCollection: `LetSet(Item) = m_collItems(indexOrKey)` inside `Property Get Item`
    let cls = "Public Property Get Item(ByVal k As Variant) As Variant\n    LetSet(Item) = k * 2\nEnd Property\nPrivate Property Let LetSet(ByRef result As Variant, ByRef v As Variant)\n    If IsObject(v) Then Set result = v Else result = v\nEnd Property\n";
    let main = "Public verdict As Long\nSub Main()\n    Dim c As New Cls\n    verdict = c.Item(21)\nEnd Sub\n";
    let v = run_two(main, cls).expect("runs");
    assert!(v.iter().any(|x| x.as_i32() == Some(42)), "{v:?}");
}
