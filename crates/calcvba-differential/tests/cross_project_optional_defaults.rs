//! A call into a referenced project that leaves out an `Optional` parameter with a declared default
//! passes that default, on every call shape, as a call inside one project does (dc-wmjmb): a
//! standard-module function with and without an argument list, and a referenced class read early
//! bound, late bound and through `CallByName`. Both ways of referencing the project are covered:
//! its source in the closure, and only its compiled program and export surface.

use std::collections::BTreeMap;

use calcvba_differential::{Canon, Executor, RunOutcome, run_ox_programs, run_project_closure};
use calcvba_oxir::{OxImage, OxProgram};
use calcvba_symbol::CatalogTypeLibResolver;
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, ProjectReference,
    ReferencedProjectManifest, SymbolProjectManifest,
};
use calcvba_symbol::surface::{ProjectExportSurface, synthesize_export_surface_from_core_program};

const TOOLS: &str = "Public Function Pick(Optional ByVal i As Variant = 3) As Variant\n\
    If IsMissing(i) Then Pick = \"missing\" Else Pick = \"got \" & TypeName(i) & \" \" & CStr(i)\n\
End Function\n\
Public Function PickLong(Optional ByVal n As Long = 5) As Variant\n\
    PickLong = \"got \" & TypeName(n) & \" \" & CStr(n)\n\
End Function\n\
Public Function PickNone(Optional ByVal i As Variant) As Variant\n\
    If IsMissing(i) Then PickNone = \"missing\" Else PickNone = \"got \" & TypeName(i)\n\
End Function\n";

const OPT: &str = "Public Function WithDefault(Optional ByVal i As Variant = 3) As Variant\n\
    If IsMissing(i) Then WithDefault = \"missing\" Else WithDefault = \"got \" & TypeName(i) & \" \" & CStr(i)\n\
End Function\n\
Public Function TypedDefault(Optional ByVal n As Long = 5) As Variant\n\
    TypedDefault = \"got \" & TypeName(n) & \" \" & CStr(n)\n\
End Function\n\
Public Function Pair(Optional ByVal a As Long = 1, Optional ByVal b As String = \"b\") As Variant\n\
    Pair = CStr(a) & \"/\" & b\n\
End Function\n\
Public Function ImplicitPair(Optional ByRef a As Long, Optional ByVal b As String, Optional ByVal suffix As String = \"\") As String\n\
    ImplicitPair = TypeName(a) & \" \" & CStr(a) & \"/\" & b & suffix\n\
    a = 99\n\
End Function\n";

const MAIN: &str = "Public result As String\n\
Sub Main()\n\
    Dim o As Object, e As Lib.Opt\n\
    Set e = New Lib.Opt: Set o = e\n\
    result = result & \"pick_parens=\" & Lib.Tools.Pick() & vbLf\n\
    result = result & \"pick=\" & Lib.Tools.Pick & vbLf\n\
    result = result & \"pick_unqualified=\" & Pick() & vbLf\n\
    result = result & \"pick_long=\" & Lib.Tools.PickLong & vbLf\n\
    result = result & \"pick_long_parens=\" & PickLong() & vbLf\n\
    result = result & \"pick_none=\" & Lib.Tools.PickNone() & vbLf\n\
    result = result & \"early_parens=\" & e.WithDefault() & vbLf\n\
    result = result & \"early=\" & e.WithDefault & vbLf\n\
    result = result & \"late=\" & o.WithDefault & vbLf\n\
    result = result & \"callbyname=\" & CallByName(o, \"WithDefault\", VbMethod) & vbLf\n\
    result = result & \"early_typed=\" & e.TypedDefault & vbLf\n\
    result = result & \"early_typed_parens=\" & e.TypedDefault() & vbLf\n\
    result = result & \"late_typed=\" & o.TypedDefault & vbLf\n\
    result = result & \"early_skip_first=\" & e.Pair(, \"x\") & vbLf\n\
    result = result & \"late_skip_first=\" & o.Pair(, \"x\") & vbLf\n\
    result = result & \"late_named=\" & o.Pair(b:=\"y\") & vbLf\n\
    result = result & \"early_named=\" & e.Pair(b:=\"y\") & vbLf\n\
    result = result & \"early_named_reordered=\" & e.Pair(b:=\"y\", a:=7) & vbLf\n\
    result = result & \"late_implicit_long=\" & o.ImplicitPair(, \"x\") & vbLf\n\
    result = result & \"late_implicit_string=\" & o.ImplicitPair(4, , \"!\") & vbLf\n\
End Sub\n";

const EXPECTED: &[&str] = &[
    "pick_parens=got Integer 3",
    "pick=got Integer 3",
    "pick_unqualified=got Integer 3",
    "pick_long=got Long 5",
    "pick_long_parens=got Long 5",
    "pick_none=missing",
    "early_parens=got Integer 3",
    "early=got Integer 3",
    "late=got Integer 3",
    "callbyname=got Integer 3",
    "early_typed=got Long 5",
    "early_typed_parens=got Long 5",
    "late_typed=got Long 5",
    "early_skip_first=1/x",
    "late_skip_first=1/x",
    "late_named=1/y",
    "early_named=1/y",
    "early_named_reordered=7/y",
    "late_implicit_long=Long 0/x",
    "late_implicit_string=Long 4/!",
];

fn module(name: &str, kind: ModuleKind, source: &str) -> ModuleUnit {
    let mut module = ModuleUnit {
        module_name: name.to_string(),
        module_kind: kind,
        attributes: ModuleAttributes::named(name),
        source: source.to_string(),
    };
    if kind == ModuleKind::Class {
        module.attributes.vb_exposed = true;
        module.attributes.vb_creatable = true;
    }
    module
}

fn lib_modules() -> Vec<ModuleUnit> {
    vec![
        module("Tools", ModuleKind::Procedural, TOOLS),
        module("Opt", ModuleKind::Class, OPT),
    ]
}

fn manifest(
    name: &str,
    modules: Vec<ModuleUnit>,
    references: Vec<ProjectReference>,
    reference_projects: Vec<ReferencedProjectManifest>,
) -> SymbolProjectManifest {
    SymbolProjectManifest {
        project_name: name.to_string(),
        project_kind: ProjectKind::Source,
        modules,
        references,
        reference_projects,
        conditional_constants: BTreeMap::new(),
        conditional_compilation_target: Default::default(),
    }
}

fn lib_reference() -> Vec<ProjectReference> {
    vec![ProjectReference::Project {
        referenced_project_name: "Lib".to_string(),
    }]
}

fn result_lines(label: &str, outcome: RunOutcome) -> Vec<String> {
    assert!(
        outcome.unsupported.is_none(),
        "{label} declined: {:?}",
        outcome.unsupported
    );
    let values = outcome
        .result
        .unwrap_or_else(|e| panic!("{label} failed: {e}"));
    values
        .iter()
        .find_map(|v| match v {
            Canon::Str(s) if s.contains("pick_parens=") => {
                Some(s.lines().map(str::to_string).collect())
            }
            _ => None,
        })
        .unwrap_or_else(|| panic!("{label}: no result in {values:?}"))
}

fn source_closure(executor: Executor) {
    let lib = manifest("Lib", lib_modules(), Vec::new(), Vec::new());
    let app = manifest(
        "App",
        vec![module("Main", ModuleKind::Procedural, MAIN)],
        lib_reference(),
        vec![ReferencedProjectManifest {
            project_name: "Lib".to_string(),
            conditional_compilation_target: None,
            project_kind: ProjectKind::Source,
            modules: lib_modules(),
        }],
    );
    let label = format!("{executor:?} source closure");
    let lines = result_lines(&label, run_project_closure(executor, &[lib, app]));
    assert_eq!(lines, EXPECTED, "{label}");
}

fn compiled_lib() -> (ProjectExportSurface, OxProgram) {
    let lib = manifest("Lib", lib_modules(), Vec::new(), Vec::new());
    let core = calcvba_bind::bind_program(&lib, &CatalogTypeLibResolver).expect("Lib binds");
    let surface = synthesize_export_surface_from_core_program(&core);
    let ox = calcvba_oxir::elaborate::elaborate(&core).expect("Lib elaborates");
    // Through the image bytes, as a compiled reference is loaded.
    let bytes = OxImage::new(vec![ox]).to_bytes().expect("image serializes");
    let mut image = OxImage::from_bytes(&bytes).expect("image deserializes");
    (surface, image.programs.pop().expect("one program"))
}

fn bundle_only(executor: Executor) {
    let (surface, lib_program) = compiled_lib();
    let app = manifest(
        "App",
        vec![module("Main", ModuleKind::Procedural, MAIN)],
        lib_reference(),
        Vec::new(),
    );
    let core = calcvba_bind::bind_program_with_project_surfaces(
        &app,
        &CatalogTypeLibResolver,
        std::slice::from_ref(&surface),
    )
    .expect("App binds against the Lib surface");
    let app_program = calcvba_oxir::elaborate::elaborate(&core).expect("App elaborates");
    let label = format!("{executor:?} bundle only");
    let lines = result_lines(
        &label,
        run_ox_programs(executor, &[lib_program, app_program]),
    );
    assert_eq!(lines, EXPECTED, "{label}");
}

#[test]
fn referenced_source_defaults_on_vm3() {
    source_closure(Executor::Vm3);
}

#[test]
fn referenced_source_defaults_on_the_jit() {
    source_closure(Executor::Jit);
}

#[test]
fn compiled_reference_defaults_on_vm3() {
    bundle_only(Executor::Vm3);
}

#[test]
fn compiled_reference_defaults_on_the_jit() {
    bundle_only(Executor::Jit);
}

#[test]
fn packaged_raw_typed_omission_uses_its_default() {
    use calcvba_oxir::{OxArg, OxInst};

    for (passing_mode, default, expected) in [
        ("ByVal", " = 7", 7),
        ("ByRef", " = 7", 7),
        ("ByVal", "", 0),
        ("ByRef", "", 0),
    ] {
        let source = format!(
            "Public Result As Long\n\
             Sub Main()\n\
                 Result = UseDefault()\n\
             End Sub\n\
             Private Function UseDefault(Optional {passing_mode} n As Long{default}) As Long\n\
                 UseDefault = n\n\
                 n = 99\n\
             End Function\n"
        );
        let manifest = manifest(
            "RawOmission",
            vec![module("Entry", ModuleKind::Procedural, &source)],
            vec![],
            vec![],
        );
        let core = calcvba_bind::bind_program(&manifest, &CatalogTypeLibResolver).unwrap();
        let mut program = calcvba_oxir::elaborate::elaborate(&core).unwrap();
        let entry = program.entry.unwrap();
        let args = program.funcs[entry.0]
            .blocks
            .iter_mut()
            .flat_map(|block| block.instrs.iter_mut())
            .find_map(|inst| match inst {
                OxInst::CallProc { args, .. } => Some(args),
                _ => None,
            })
            .expect("compiled call");
        assert_eq!(args.len(), 1);
        args[0] = OxArg::Omitted;

        let bytes = OxImage::new(vec![program]).to_bytes().unwrap();
        let image = OxImage::from_bytes(&bytes).unwrap();
        let expected = calcvba_differential::canon(&calcvba_runtime::Variant::from_i32(expected));
        for executor in [Executor::Vm3, Executor::Jit] {
            let run = run_ox_programs(executor, &image.programs);
            assert!(
                run.unsupported.is_none(),
                "{passing_mode} {executor:?}: {run:?}"
            );
            assert_eq!(
                run.result.as_deref().expect("completed call"),
                std::slice::from_ref(&expected)
            );
            assert_eq!(run.err.number, 0);
        }
    }
}
