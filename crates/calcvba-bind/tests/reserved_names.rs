use calcvba_bind::bind_program;
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};
use std::collections::BTreeMap;

fn manifest(source: &str) -> SymbolProjectManifest {
    SymbolProjectManifest {
        project_name: "Proj".into(),
        project_kind: ProjectKind::Source,
        modules: vec![ModuleUnit {
            module_name: "Mod1".into(),
            module_kind: ModuleKind::Procedural,
            attributes: ModuleAttributes::named("Mod1"),
            source: source.into(),
        }],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: BTreeMap::new(),
        conditional_compilation_target: Default::default(),
    }
}

#[test]
fn local_is_reserved_in_declarations_but_imported_is_legal() {
    for name in ["local", "LoCaL", "[local]"] {
        let source =
            format!("Sub Main()\nDim count As Long, {name} As Object, value As Long\nEnd Sub\n");
        let error = bind_program(&manifest(&source), &calcvba_symbol::CatalogTypeLibResolver)
            .unwrap_err()
            .to_string();
        assert!(error.contains("reserved identifier `local`"), "{error}");
    }
    bind_program(&manifest("Sub Main()\nDim count As Long, imported As Object, ownToken As Object, value As Long\nEnd Sub\n"), &calcvba_symbol::CatalogTypeLibResolver).unwrap();
}
