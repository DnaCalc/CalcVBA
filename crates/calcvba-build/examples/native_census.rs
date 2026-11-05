//! Compile every procedure in each individual conformance .bas module. Class
//! fixture suffixes retain their module kind. Missing project/reference context
//! stays a source diagnostic; these counts do not claim complete project execution.
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};
use std::{
    collections::BTreeMap,
    path::{Path, PathBuf},
};

fn collect(path: &Path, files: &mut Vec<PathBuf>) -> std::io::Result<()> {
    for entry in std::fs::read_dir(path)? {
        let path = entry?.path();
        if path.is_dir() {
            collect(&path, files)?;
        } else if path
            .extension()
            .is_some_and(|extension| extension.to_string_lossy().eq_ignore_ascii_case("bas"))
        {
            files.push(path);
        }
    }
    Ok(())
}

fn quoted(value: &str) -> String {
    format!(
        "\"{}\"",
        value.replace('"', "\"\"").replace(['\r', '\n'], " ")
    )
}

fn construct(detail: &str) -> String {
    for marker in ["typed instruction ", "unsupported terminator "] {
        if let Some(rest) = detail.split(marker).nth(1) {
            return rest
                .split([' ', '{', '('])
                .next()
                .unwrap_or(rest)
                .to_string();
        }
    }

    if let Some(rest) = detail
        .split("unsupported block ")
        .nth(1)
        .and_then(|text| text.split_once(": ").map(|(_, rest)| rest))
    {
        return rest.split([' ', '{']).next().unwrap_or(rest).to_string();
    }
    for (needle, name) in [
        ("native library call", "VBA library"),
        (
            "unlinked typed source project",
            "unresolved project/library reference",
        ),
        ("ParamArray", "ParamArray"),
        ("optional signature", "optional parameters"),
        ("project classes", "class modules"),
        ("Variant array receiver", "Variant array access"),
        ("native typed COM", "ComCallEarly signature"),
        ("typed entry terminator", "error/control terminator"),
        ("dynamic integer arithmetic", "Variant integer arithmetic"),
        ("requires the dynamic runtime", "carrier/instruction layout"),
    ] {
        if detail.contains(needle) {
            return name.into();
        }
    }
    detail
        .split("procedure ")
        .nth(1)
        .and_then(|text| text.split_once(": ").map(|(_, rest)| rest))
        .unwrap_or(detail)
        .split(['{', '\n'])
        .next()
        .unwrap_or(detail)
        .trim()
        .to_string()
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let root = PathBuf::from(args.next().unwrap_or_else(|| "conformance".into()));
    let output = PathBuf::from(
        args.next()
            .unwrap_or_else(|| "temp/native-census.csv".into()),
    );
    let mut files = Vec::new();
    collect(&root, &mut files)?;
    files.sort();
    let mut rows = String::from("file,status,construct,detail\n");
    let mut totals = BTreeMap::<String, usize>::new();
    let mut refused = BTreeMap::<String, usize>::new();
    for path in &files {
        let source = std::fs::read_to_string(path)?;
        let stem = path.file_stem().unwrap().to_string_lossy();
        let module_kind = if stem.ends_with(".class") {
            ModuleKind::Class
        } else {
            ModuleKind::Procedural
        };
        let name = stem
            .strip_suffix(".class")
            .or_else(|| stem.strip_suffix(".proc"))
            .unwrap_or(&stem)
            .replace('-', "_");
        let manifest = SymbolProjectManifest {
            project_name: "VBAProject".into(),
            project_kind: ProjectKind::Source,
            modules: vec![ModuleUnit {
                module_name: name.clone(),
                module_kind,
                attributes: ModuleAttributes::named(&name),
                source,
            }],
            references: Vec::new(),
            reference_projects: Vec::new(),
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
        };
        let compiled =
            calcvba_bind::bind_program(&manifest, &calcvba_symbol::CatalogTypeLibResolver)
                .map_err(|error| format!("bind: {error}"))
                .and_then(|core| {
                    calcvba_oxir::elaborate::elaborate(&core)
                        .map_err(|error| format!("elaborate: {error}"))
                });
        let (status, category, detail) = match compiled {
            Err(detail) => ("source-refused", "source diagnostic".into(), detail),
            Ok(program) => {
                let roots: Vec<_> = (0..program.funcs.len()).map(calcvba_oxir::FuncId).collect();
                match calcvba_build::check_native_program(&program, &roots) {
                    Ok(()) => ("native", String::new(), String::new()),
                    Err(error) => {
                        let detail = error.to_string();
                        let category = construct(&detail);
                        *refused.entry(category.clone()).or_default() += 1;
                        ("native-refused", category, detail)
                    }
                }
            }
        };
        *totals.entry(status.into()).or_default() += 1;
        rows.push_str(
            &[
                quoted(&path.to_string_lossy()),
                quoted(status),
                quoted(&category),
                quoted(&detail),
            ]
            .join(","),
        );
        rows.push('\n');
    }
    if let Some(parent) = output.parent() {
        std::fs::create_dir_all(parent)?;
    }
    std::fs::write(&output, rows)?;
    println!(
        "{} individual .bas modules; class suffixes recognized; all procedures are roots",
        files.len()
    );
    for (status, count) in totals {
        println!("{status}: {count}");
    }
    println!("First native refusal per input:");
    for (category, count) in refused {
        println!("{category}: {count}");
    }
    println!("CSV: {}", output.display());
    Ok(())
}
