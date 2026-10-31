//! Serialized source-free execution must preserve the artifact's entry and faults.
use oxvba_host::{Engine, HostConfig};
use oxvba_oxir::{OxImage, OxProgram};
use oxvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};

fn program(name: &str, source: &str) -> OxProgram {
    let manifest = SymbolProjectManifest {
        project_name: name.into(),
        project_kind: ProjectKind::Source,
        modules: vec![ModuleUnit {
            module_name: "Main".into(),
            module_kind: ModuleKind::Procedural,
            attributes: ModuleAttributes::named("Main"),
            source: source.into(),
        }],
        references: vec![],
        reference_projects: vec![],
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let core = oxvba_bind::bind_program(&manifest, &oxvba_symbol::CatalogTypeLibResolver).unwrap();
    oxvba_oxir::elaborate::elaborate(&core).unwrap()
}

#[test]
fn source_free_image_honors_nonfinal_entry_and_preserves_runtime_error_number() {
    let mut image = OxImage::new(vec![
        program(
            "Entry",
            "Public result As Long\nSub Main()\nresult = 42\nEnd Sub\n",
        ),
        program(
            "Dependency",
            "Public result As Long\nSub Main()\nresult = 99\nEnd Sub\n",
        ),
    ]);
    image.entry = 0;
    let bytes = image.to_bytes().unwrap();
    let fault = OxImage::new(vec![program(
        "Fault",
        "Sub Main()\nErr.Raise 13\nEnd Sub\n",
    )])
    .to_bytes()
    .unwrap();
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let engine = Engine::new(config);
        let values = engine
            .execute_image_bytes_with_variant_snapshot(&bytes)
            .unwrap();
        assert_eq!(values.len(), 1);
        assert_eq!(values[0].as_i32(), Some(42));
        let error = engine
            .execute_image_bytes_with_variant_snapshot(&fault)
            .unwrap_err();
        assert_eq!(error.vba_code(), Some(13));
    }
}

#[test]
fn dependency_initializer_runs_before_the_entry_and_can_fail() {
    let mut dependency = program("Dependency", "Sub Main()\nErr.Raise 13\nEnd Sub\n");
    dependency.global_initializer = dependency.entry.take();
    let image = OxImage::new(vec![dependency, program("Entry", "Sub Main()\nEnd Sub\n")]);
    for config in [HostConfig::vm3(), HostConfig::jit()] {
        let error = Engine::new(config)
            .execute_image_with_variant_snapshot(&image)
            .unwrap_err();
        assert_eq!(error.vba_code(), Some(13));
    }
}
