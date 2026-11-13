//! Survey, not a check: `VBA_FILES=Name:class|std:path;...` loads those sources as one project through
//! `Module::load_project` and prints the full diagnostic (dc-tey4.5, Wasabi/Riff).

use calcvba_host::Module;
use calcvba_symbol::manifest as sym;
use calcvba_symbol::CatalogTypeLibResolver;

#[test]
#[ignore]
fn survey_project_from_files() {
    let Ok(files) = std::env::var("VBA_FILES") else { return };
    let modules: Vec<sym::ModuleUnit> = files
        .split(';')
        .map(|spec| {
            let mut parts = spec.splitn(3, ':');
            let (name, kind, path) = (parts.next().unwrap(), parts.next().unwrap(), parts.next().unwrap());
            let bytes = std::fs::read(path).unwrap();
            let source: String = String::from_utf8(bytes.clone())
                .unwrap_or_else(|_| bytes.iter().map(|b| char::from(*b)).collect());
            sym::ModuleUnit {
                module_name: name.into(),
                module_kind: if kind == "class" { sym::ModuleKind::Class } else { sym::ModuleKind::Procedural },
                attributes: sym::ModuleAttributes::named(name),
                source,
            }
        })
        .collect();
    let manifest = sym::SymbolProjectManifest {
        project_name: "HostProject".into(),
        project_kind: sym::ProjectKind::Library,
        modules,
        references: vec![],
        reference_projects: vec![],
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    match Module::load_project(&manifest, &CatalogTypeLibResolver) {
        Ok(_) => println!("SURVEY loaded"),
        Err(e) => println!("SURVEY {e}"),
    }
}
