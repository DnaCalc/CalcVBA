//! Shapes the vba-regex, vba-arraytools and vba-test samples need to compile and run
//! (dc-tey4.8, dc-tey4.9, dc-tey4.10). Excel's VBE compiles and runs each of them.

use calcvba_engine::{Engine, HostConfig};
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};

fn module(name: &str, kind: ModuleKind, source: &str) -> ModuleUnit {
    ModuleUnit {
        module_name: name.into(),
        module_kind: kind,
        attributes: ModuleAttributes::named(name),
        source: source.into(),
    }
}

fn std_module(name: &str, source: &str) -> ModuleUnit {
    module(name, ModuleKind::Procedural, source)
}

fn class_module(name: &str, source: &str) -> ModuleUnit {
    module(name, ModuleKind::Class, source)
}

fn run_with(
    config: HostConfig,
    modules: &[ModuleUnit],
) -> Result<Vec<calcvba_runtime::Variant>, String> {
    let manifest = SymbolProjectManifest {
        project_name: "Proj".into(),
        project_kind: ProjectKind::Source,
        modules: modules.to_vec(),
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    Engine::new(config)
        .execute_manifest_with_variant_snapshot(&manifest)
        .map_err(|d| format!("{:?}: {}", d.phase(), d.message()))
}

/// Run `modules` (Main.Main is the entry) on VM3 and the JIT; each must leave a global equal
/// to `verdict`.
fn both(modules: &[ModuleUnit], verdict: i32) {
    for (name, config) in [("vm3", HostConfig::vm3()), ("jit", HostConfig::jit())] {
        let values = run_with(config, modules).unwrap_or_else(|e| panic!("{name}: {e}"));
        assert!(
            values.iter().any(|v| v.as_i32() == Some(verdict)),
            "{name}: expected {verdict} in {values:?}"
        );
    }
}

const ARRAY_BUFFER: &str = "Public Type Ty\n    Buffer() As Long\n    Capacity As Long\n    Length As Long\nEnd Type\nPublic Sub AppendLong(ByRef lab As Ty, ByVal v As Long)\n    With lab\n        If .Length + 1 > .Capacity Then Grow lab, .Length + 1\n        .Buffer(.Length) = v\n        .Length = .Length + 1\n    End With\nEnd Sub\nPrivate Sub Grow(ByRef lab As Ty, ByVal n As Long)\n    lab.Capacity = n * 2\n    ReDim Preserve lab.Buffer(0 To lab.Capacity - 1) As Long\nEnd Sub\n";

const STRING_BUILDER: &str = "Public Type Ty\n    Text As String\n    Count As Long\nEnd Type\nPublic Sub Add(ByRef sb As Ty, ByVal s As String)\n    sb.Text = sb.Text & s\n    sb.Count = sb.Count + 1\nEnd Sub\n";

/// vba-regex: ArrayBuffer, RegexLexer and StaticStringBuilder each declare `Public Type Ty`.
/// A module's own `Ty` shadows the others; other modules qualify it (`ArrayBuffer.Ty`).
#[test]
fn same_named_public_types_in_several_modules_stay_distinct() {
    both(
        &[
            std_module("ArrayBuffer", ARRAY_BUFFER),
            std_module("StaticStringBuilder", STRING_BUILDER),
            std_module(
                "Main",
                "Private Type Holder\n    buf As ArrayBuffer.Ty\n    sb As StaticStringBuilder.Ty\nEnd Type\nPublic verdict As Long\nSub Main()\n    Dim a As ArrayBuffer.Ty, s As StaticStringBuilder.Ty, h As Holder\n    ArrayBuffer.AppendLong a, 7\n    ArrayBuffer.AppendLong a, 9\n    StaticStringBuilder.Add s, \"x\"\n    StaticStringBuilder.Add s, \"yz\"\n    ArrayBuffer.AppendLong h.buf, 5\n    StaticStringBuilder.Add h.sb, \"q\"\n    verdict = 1000 * a.Length + 100 * a.Buffer(1) + 10 * Len(s.Text) + h.buf.Buffer(0) + h.sb.Count\nEnd Sub\n",
            ),
        ],
        2000 + 900 + 30 + 5 + 1,
    );
}

/// An unqualified `Ty` in a module that declares none stays ambiguous between two public owners.
#[test]
fn unqualified_type_with_two_public_owners_outside_them_is_ambiguous() {
    let error = run_with(
        HostConfig::vm3(),
        &[
            std_module("ArrayBuffer", ARRAY_BUFFER),
            std_module("StaticStringBuilder", STRING_BUILDER),
            std_module("Main", "Sub Main()\n    Dim a As Ty\nEnd Sub\n"),
        ],
    )
    .expect_err("ambiguous");
    assert!(error.contains("ambiguous name detected: Ty"), "{error}");
}

/// vba-regex: `RegexUnicodeSupport.StaticData(i)` indexes another module's fixed array.
#[test]
fn module_qualified_array_variable_is_indexed() {
    both(
        &[
            std_module(
                "Store",
                "Public Enum C\n    N = 10\nEnd Enum\nPublic Data(0 To N - 1) As Long\nPublic Sub Fill(ByRef a() As Long)\n    a(3) = 42\nEnd Sub\n",
            ),
            std_module(
                "Main",
                "Public verdict As Long\nSub Main()\n    Dim x As Long, v As Variant\n    Store.Fill Store.Data\n    x = Store.Data(3)\n    Store.Data(4) = 8\n    v = 4\n    verdict = 100 * x + Store.Data(v) + Data(4)\nEnd Sub\n",
            ),
        ],
        4216,
    );
}

/// vba-regex: `ReDim .table(0 To .mask) As Entry` inside `With` on a ByRef UDT parameter, and
/// a record field of a record-array element passed ByRef to a UDT parameter.
#[test]
fn record_array_fields_resize_and_pass_record_fields_by_reference() {
    both(
        &[std_module(
            "Main",
            "Private Type Pair\n    a As Long\n    b As Long\nEnd Type\nPrivate Type Node\n    ref As Pair\nEnd Type\nPrivate Type Tree\n    mask As Long\n    buf() As Node\n    ptrs() As Long\nEnd Type\nPublic verdict As Long\nPrivate Sub Init(ByRef t As Tree)\n    With t\n        ReDim .buf(0 To .mask) As Node\n        ReDim .ptrs(0 To .mask + 1) As Long\n    End With\nEnd Sub\nPrivate Function Cmp(ByRef p As Pair, ByRef q As Pair) As Long\n    Cmp = p.a - q.a\n    q.b = 99\nEnd Function\nSub Main()\n    Dim t As Tree, x As Pair\n    t.mask = 3\n    Init t\n    t.buf(1).ref.a = 5\n    x.a = 8\n    verdict = 10000 * Cmp(x, t.buf(1).ref) + 100 * t.buf(1).ref.b + 10 * UBound(t.buf) + UBound(t.ptrs)\nEnd Sub\n",
        )],
        30000 + 9900 + 30 + 4,
    );
}

/// vba-arraytools: a `Variant` UDT field holds an array (`ReDim d.items(n)`, `d.items(i) = v`),
/// `Module.Enum.Member`, a field of a UDT function result and a library-enum field type.
#[test]
fn variant_fields_enum_qualifiers_record_results_and_library_enum_fields() {
    both(
        &[
            std_module(
                "Lib",
                "Public Enum OPS\n    opLess = 1\n    opBigger = 2\nEnd Enum\nPrivate Type Desc\n    count As Long\n    items As Variant\n    method As VbCompareMethod\nEnd Type\nPrivate Function Make(ByVal n As Long) As Desc\n    Dim d As Desc, i As Long\n    d.count = n\n    d.method = IIf(n > 2, vbTextCompare, vbBinaryCompare)\n    ReDim d.items(0 To n - 1)\n    For i = 0 To n - 1\n        d.items(i) = i * 2\n    Next i\n    Make = d\nEnd Function\nPublic Function Count(ByVal n As Long) As Long\n    Count = Make(n).count * 10 + Make(n).method\nEnd Function\nPublic Function Last(ByVal n As Long) As Long\n    Dim d As Desc\n    d = Make(n)\n    Last = d.items(n - 1)\nEnd Function\n",
            ),
            std_module(
                "Main",
                "Public verdict As Long\nSub Main()\n    verdict = 1000 * Lib.OPS.opBigger + 10 * Lib.Count(3) + Lib.Last(4)\nEnd Sub\n",
            ),
        ],
        2000 + 310 + 6,
    );
}

/// vba-arraytools: `IsIterable(c)` with `Dim c As New Collection` creates it when it is passed
/// to a ByRef Variant; `vector(p)` passed ByRef where `vector` holds a Collection passes the
/// default member's value.
#[test]
fn as_new_and_default_member_values_reach_byref_variant_parameters() {
    both(
        &[std_module(
            "Main",
            "Public verdict As Long\nPrivate Function Iterable(ByRef vList As Variant) As Long\n    Dim v As Variant\n    On Error Resume Next\n    For Each v In vList\n        Exit For\n    Next v\n    If Err.Number = 0 And TypeName(vList) = \"Collection\" Then Iterable = 1\nEnd Function\nPrivate Function Twice(ByRef v As Variant) As Long\n    Twice = 2 * v\nEnd Function\nPrivate Function At(ByRef vector As Variant, ByVal p As Long) As Long\n    At = Twice(vector(p))\nEnd Function\nSub Main()\n    Dim c As New Collection, d As New Collection\n    d.Add 5: d.Add 7\n    verdict = 1000 * Iterable(c) + 10 * At(d, 2) + At(Array(1, 2, 3), 1)\nEnd Sub\n",
        )],
        1000 + 140 + 4,
    );
}

/// vba-test: `If Not pSuite Is Nothing` reads an unset WithEvents field as Nothing.
#[test]
fn unset_withevents_field_reads_as_nothing() {
    both(
        &[
            class_module(
                "Suite",
                "Public Event Done()\nPublic Sub Fire()\n    RaiseEvent Done\nEnd Sub\n",
            ),
            class_module(
                "Rep",
                "Private WithEvents pSuite As Suite\nPublic Seen As Long\nPublic Sub ListenTo(s As Suite)\n    If pSuite Is Nothing Then Seen = Seen + 100\n    Set pSuite = s\n    If Not pSuite Is Nothing Then Seen = Seen + 10\nEnd Sub\nPrivate Sub pSuite_Done()\n    Seen = Seen + 1\nEnd Sub\n",
            ),
            std_module(
                "Main",
                "Public verdict As Long\nSub Main()\n    Dim s As New Suite, r As New Rep\n    r.ListenTo s\n    s.Fire\n    verdict = r.Seen + 1000\nEnd Sub\n",
            ),
        ],
        1111,
    );
}

/// vba-arraytools `SwapCollectionValues`: `coll.Add Item:=..., Before:=i` binds by name.
#[test]
fn collection_add_named_arguments_bind_by_name() {
    both(
        &[std_module(
            "Main",
            "Public verdict As Long\nSub Main()\n    Dim c As New Collection, v As Variant, s As Long\n    c.Add 1: c.Add 2: c.Add 3\n    c.Add Item:=9, Before:=2\n    c.Add Item:=8, After:=1\n    c.Add 7, Key:=\"k\", Before:=1\n    For Each v In c\n        s = s * 10 + v\n    Next v\n    verdict = s + 0 * c(\"k\")\nEnd Sub\n",
        )],
        718923,
    );
}

/// dc-yqq0.11: string functions keep an unpaired surrogate (VBA strings are UTF-16 code units).
#[test]
fn string_functions_keep_unpaired_surrogates() {
    both(
        &[std_module(
            "Main",
            "Public verdict As Long\nSub Main()\n    Dim s As String, v As Variant, n As Long\n    s = ChrW(&HD800) & \"ab\"\n    v = s\n    If AscW(UCase(s)) = -10240 Then n = n + 1\n    If AscW(LCase(s)) = -10240 Then n = n + 2\n    If AscW(Trim(\" \" & s)) = -10240 Then n = n + 4\n    If AscW(Replace(s, \"a\", \"x\")) = -10240 Then n = n + 8\n    If AscW(Split(s, \"a\")(0)) = -10240 Then n = n + 16\n    If AscW(CStr(v)) = -10240 Then n = n + 32\n    If Len(Replace(s, \"b\", \"\", 2)) = 1 Then n = n + 64\n    verdict = 1000 + n\nEnd Sub\n",
        )],
        1127,
    );
}
