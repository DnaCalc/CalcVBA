#![cfg(windows)]
use calcvba_host::{Backend, Module, ProcedureKind};
use calcvba_office::OfficeFile;
use std::path::Path;

#[test]
fn every_saved_workbook_compiles_and_binds_on_vm3_and_jit() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../calcvba-office/tests/fixtures");
    let mut count = 0;
    for case in std::fs::read_dir(root).unwrap() {
        let path = case.unwrap().path();
        if !path.is_dir() {
            continue;
        }
        let office = OfficeFile::open(path.join("workbook.xlsm")).unwrap();
        let manifest = office.project.to_manifest().unwrap();
        let snapshot: serde_json::Value =
            serde_json::from_slice(&std::fs::read(path.join("project.json")).unwrap()).unwrap();
        let module = Module::load_project(&manifest, &calcvba_symbol::CatalogTypeLibResolver)
            .unwrap_or_else(|e| panic!("{}: {e}", path.display()));
        for backend in [Backend::Interpreter, Backend::Jit]
            .into_iter()
            .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
        {
            let session = module
                .clone()
                .into_session(backend)
                .unwrap_or_else(|e| panic!("{} / {backend:?}: {e}", path.display()));
            let bound = session
                .bind(
                    snapshot["entry_module"].as_str().unwrap(),
                    "RunProbe",
                    ProcedureKind::Function,
                )
                .unwrap();
            session
                .bind_values(&bound, &[], calcvba_host_api::ValueType::String)
                .unwrap();
        }
        count += 1;
    }
    assert_eq!(count, 33);
}
