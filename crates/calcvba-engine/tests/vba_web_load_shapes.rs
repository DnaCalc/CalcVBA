//! A Private Enum of a class module is reachable as `EnumName.Member` inside the class
//! (VBA-Web's `web_WinHttpRequestOption.web_WinHttpRequestOption_...`, dc-tey4.4).

use calcvba_engine::{Engine, HostConfig};
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};

fn run(class_src: &str, main_src: &str) -> Result<Vec<calcvba_runtime::Variant>, String> {
    run_with(HostConfig::vm3(), class_src, main_src)
}

fn run_with(
    config: HostConfig,
    class_src: &str,
    main_src: &str,
) -> Result<Vec<calcvba_runtime::Variant>, String> {
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
    Engine::new(config)
        .execute_manifest_with_variant_snapshot(&manifest)
        .map_err(|d| format!("{:?}: {}", d.phase(), d.message()))
}

const MAIN: &str = "Public verdict As Long\nSub Main()\nDim c As New Cls\nverdict = c.Go()\nEnd Sub\n";

#[test]
fn qualified_private_enum_member_in_value_position() {
    let v = run(
        "Private Enum Opt\n    Opt_A = 0\n    Opt_B = 18\nEnd Enum\nPublic Function Go() As Long\n    Go = Opt.Opt_B\nEnd Function\n",
        MAIN,
    )
    .expect("runs");
    assert!(v.iter().any(|x| x.as_i32() == Some(18)));
}

#[test]
fn qualified_private_enum_member_as_late_bound_index() {
    let v = run(
        "Private Enum Opt\n    Opt_A = 0\n    Opt_B = 18\nEnd Enum\nPublic Function Go() As Long\n    Dim d As Object\n    Set d = CreateObject(\"Scripting.Dictionary\")\n    d.Item(Opt.Opt_B) = False\n    Go = d.Count\nEnd Function\n",
        MAIN,
    )
    .expect("runs");
    assert!(v.iter().any(|x| x.as_i32() == Some(1)));
}

#[test]
fn class_typed_variable_to_byref_object_parameter() {
    let v = run(
        "Public Function Go() As Long\n    Dim c As Cls\n    Set c = Me\n    Go = Take(c)\nEnd Function\nPrivate Function Take(o As Object) As Long\n    Take = 7\nEnd Function\n",
        MAIN,
    )
    .expect("runs");
    assert!(v.iter().any(|x| x.as_i32() == Some(7)));
}

#[test]
fn string_stored_into_byte_array_copies_utf16_bytes() {
    let v = run(
        "Public Function Go() As Long\n    Dim b() As Byte\n    b = \"Ab\"\n    Go = (UBound(b) + 1) * 1000 + b(0) + b(1) + b(2) + b(3)\nEnd Function\n",
        MAIN,
    )
    .expect("runs");
    // 4 bytes: 65, 0, 98, 0
    assert!(v.iter().any(|x| x.as_i32() == Some(4000 + 163)), "{v:?}");
}

#[test]
fn strconv_result_stored_into_byte_array() {
    let v = run(
        "Public Function Go() As Long\n    Dim b() As Byte\n    b = StrConv(\"Ab\", vbFromUnicode)\n    Go = (UBound(b) + 1) * 1000 + b(0) + b(1)\nEnd Function\n",
        MAIN,
    )
    .expect("runs");
    assert!(v.iter().any(|x| x.as_i32() == Some(2000 + 163)), "{v:?}");
}

const REDIM_CLASS: &str = "Public Function Go() As Long\n    Dim b() As Byte\n    Dim a() As Byte\n    b = \"Ab\"\n    ReDim a(Int(UBound(b) / 2))\n    Go = UBound(a) + 1\nEnd Function\n";

#[test]
fn redim_with_int_of_ubound_vm3() {
    let v = run(REDIM_CLASS, MAIN).expect("runs");
    assert!(v.iter().any(|x| x.as_i32() == Some(2)), "{v:?}");
}

#[test]
fn redim_with_int_of_ubound_jit() {
    let v = run_with(HostConfig::jit(), REDIM_CLASS, MAIN).expect("runs");
    assert!(v.iter().any(|x| x.as_i32() == Some(2)), "{v:?}");
}

const HEX_CLASS: &str = "Public Function Go() As String\n    Dim b() As Byte\n    Dim i As Long\n    b = StrConv(\"Ab\", vbFromUnicode)\n    For i = LBound(b) To UBound(b)\n        Go = Go & Hex$(b(i))\n    Next i\nEnd Function\n";

#[test]
fn hex_of_byte_array_element_jit() {
    let main = "Public verdict As String\nSub Main()\nDim c As New Cls\nverdict = c.Go()\nEnd Sub\n";
    let v = run_with(HostConfig::jit(), HEX_CLASS, main).expect("runs");
    assert!(v.iter().any(|x| x.as_bstr().map(|s| s.as_str()) == Some("4162".into())), "{v:?}");
}

const PROP_CLASS: &str = "Public Enum Fmt\n    A = 0\n    B = 1\nEnd Enum\nPrivate pF As Fmt\nPrivate pC As String\nPublic Property Get F() As Fmt\n    F = pF\nEnd Property\nPublic Property Get C() As String\n    C = pC\nEnd Property\nPublic Property Get Content() As String\n    Content = Name2(Me.F, Me.C)\nEnd Property\nPublic Function Name2(Format As Fmt, Optional Custom As String) As String\n    Select Case Format\n    Case Fmt.B\n        Name2 = \"b\"\n    Case Else\n        Name2 = \"a\" & Custom\n    End Select\nEnd Function\nPublic Function Go() As String\n    pF = B\n    Go = Me.Content\nEnd Function\n";

#[test]
fn enum_property_and_string_property_as_byref_arguments_jit() {
    let main = "Public verdict As String\nSub Main()\nDim c As New Cls\nverdict = c.Go()\nEnd Sub\n";
    let v = run_with(HostConfig::jit(), &PROP_CLASS.replace("\\\n", "\n"), main).expect("runs");
    assert!(v.iter().any(|x| x.as_bstr().map(|s| s.as_str()) == Some("b".into())), "{v:?}");
}

const FIELD_DEFAULT_CLASS: &str = "Public Items As Collection\nPublic Function Go() As Long\n    Dim d As Object\n    Set d = CreateObject(\"Scripting.Dictionary\")\n    Set Items = New Collection\n    Items.Add 41\n    Go = Me.Items(1) + 1\nEnd Function\n";

#[test]
fn object_field_default_member_index_jit() {
    let main = "Public verdict As Long\nSub Main()\nDim c As New Cls\nverdict = c.Go()\nEnd Sub\n";
    let v = run_with(HostConfig::jit(), FIELD_DEFAULT_CLASS, main).expect("runs");
    assert!(v.iter().any(|x| x.as_i32() == Some(42)), "{v:?}");
}

const VARIANT_INDEX_CLASS: &str = "Private pBody As Variant\nPublic Function Go() As Long\n    Dim k As Variant\n    k = \"a\"\n    Set pBody = CreateObject(\"Scripting.Dictionary\")\n    pBody(k) = 41\n    pBody(\"b\") = 1\n    Dim v As Variant\n    Dim arr(3) As Long\n    v = arr\n    v(2) = 5\n    Go = pBody.Item(\"a\") + pBody.Item(\"b\") + pBody.Count + v(2)\nEnd Function\n";

#[test]
fn variant_holding_object_or_array_indexed_store_both_executors() {
    let main = "Public verdict As Long\nSub Main()\nDim c As New Cls\nverdict = c.Go()\nEnd Sub\n";
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let v = run_with(config, VARIANT_INDEX_CLASS, main).expect("runs");
        assert!(v.iter().any(|x| x.as_i32() == Some(41 + 1 + 2 + 5)), "{v:?}");
    }
}

const BYREF_TEMP_CLASS: &str = "Private pContent As String\nPrivate pBody As Variant\nPublic Property Get Content() As String\n    Content = pContent\nEnd Property\nPublic Property Let Content(Value As String)\n    pContent = Value\nEnd Property\nPublic Property Get Body() As Variant\n    Body = pBody\nEnd Property\nPublic Property Let Body(Value As Variant)\n    pBody = Value\nEnd Property\nPublic Function Go() As Long\n    Dim h As Object\n    Set h = CreateObject(\"Scripting.Dictionary\")\n    Me.Content = \"abc\"\n    Me.Body = 4\n    Go = Load(h.Count, Me.Content, Me.Body)\nEnd Function\nPrivate Function Load(Headers As String, Content As String, Body As Variant) As Long\n    Load = Len(Content) + Body\nEnd Function\n";

#[test]
fn byref_temporaries_of_property_results_jit() {
    let main = "Public verdict As Long\nSub Main()\nDim c As New Cls\nverdict = c.Go()\nEnd Sub\n";
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let v = run_with(config, BYREF_TEMP_CLASS, main).expect("runs");
        assert!(v.iter().any(|x| x.as_i32() == Some(7)), "{v:?}");
    }
}

#[test]
fn dateadd_with_long_local_count_jit() {
    let cls = "Public Function Go() As Long\n    Dim s As Long\n    s = 90\n    Dim t As Date\n    t = DateAdd(\"s\", s, #1/1/2020 12:00:00 AM#)\n    Go = Minute(t) * 100 + Second(t)\nEnd Function\n";
    let main = "Public verdict As Long\nSub Main()\nDim c As New Cls\nverdict = c.Go()\nEnd Sub\n";
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let v = run_with(config, cls, main).expect("runs");
        assert!(v.iter().any(|x| x.as_i32() == Some(130)), "{v:?}");
    }
}
