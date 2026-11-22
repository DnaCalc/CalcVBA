//! Inspect corpus sources without executing them, or compare explicitly listed reviewed fixtures.

use std::{
    collections::BTreeSet,
    env, fs,
    path::{Path, PathBuf},
};

use calcvba_differential::parity::{Observation, compare_entry};
use calcvba_oxir::OxProgram;
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};
use serde_json::{Value, json};

fn extension(path: &Path) -> String {
    path.extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase()
}

fn collect(path: &Path, inputs: &mut BTreeSet<PathBuf>) -> std::io::Result<()> {
    let metadata = fs::symlink_metadata(path)?;
    if metadata.is_dir() {
        for entry in fs::read_dir(path)? {
            collect(&entry?.path(), inputs)?;
        }
    } else if metadata.is_file()
        && matches!(
            extension(path).as_str(),
            "bas" | "cls" | "frm" | "basproj" | "vbp" | "xlsm" | "xlsb" | "xlam" | "xls"
        )
    {
        inputs.insert(path.to_owned());
    }
    Ok(())
}

fn load(path: &Path) -> Result<Vec<SymbolProjectManifest>, String> {
    match extension(path).as_str() {
        "bas" => {
            let source = fs::read_to_string(path).map_err(|err| err.to_string())?;
            let module_name = source
                .lines()
                .find_map(|line| {
                    let (key, value) = line.trim().split_once('=')?;
                    key.trim()
                        .eq_ignore_ascii_case("Attribute VB_Name")
                        .then(|| value.trim().trim_matches('"').to_owned())
                })
                .unwrap_or_else(|| "Main".into());
            Ok(vec![SymbolProjectManifest {
                project_name: "VBAProject".into(),
                project_kind: ProjectKind::Source,
                modules: vec![ModuleUnit {
                    attributes: ModuleAttributes::named(&module_name),
                    module_name,
                    module_kind: ModuleKind::Procedural,
                    source,
                }],
                references: Vec::new(),
                reference_projects: Vec::new(),
                conditional_constants: Default::default(),
                conditional_compilation_target: Default::default(),
            }])
        }
        "basproj" | "vbp" => {
            calcvba_project::load_project_closure(path).map_err(|err| err.to_string())
        }
        "xlsm" | "xlsb" | "xlam" | "xls" => calcvba_office::OfficeFile::open(path)
            .and_then(|file| file.project.to_manifest())
            .map(|manifest| vec![manifest])
            .map_err(|err| err.to_string()),
        _ => Err("pass the owning project for this module".into()),
    }
}

fn compile(closure: &[SymbolProjectManifest]) -> Result<Vec<OxProgram>, String> {
    let _scope =
        calcvba_com::ComPolicyScope::enter(|_, _| Err("native COM disabled for inspection".into()));
    calcvba_bind::bind_projects(closure, &calcvba_symbol::CatalogTypeLibResolver)
        .map_err(|err| format!("bind: {err}"))?
        .iter()
        .map(|program| {
            calcvba_oxir::elaborate::elaborate(program).map_err(|err| format!("elaborate: {err}"))
        })
        .collect()
}

fn observation(run: &Observation) -> Value {
    let outcome = &run.outcome;
    let err = &outcome.err;
    json!({
        "values": outcome.result.as_ref().ok().map(|values| {
            values.iter().map(|value| format!("{value:?}")).collect::<Vec<_>>()
        }),
        "unsupported": outcome.unsupported,
        "failure": outcome.result.as_ref().err(),
        "raised": outcome.raised,
        "err": {
            "number": err.number,
            "description": err.description,
            "source": err.source,
            "last_dll_error": err.last_dll_error,
        },
        "output": run.output.iter().map(|(channel, text)| {
            json!({ "channel": format!("{channel:?}"), "text": text })
        }).collect::<Vec<_>>(),
    })
}

fn row(path: &Path, compare: bool) -> Value {
    let reported_path = path.to_string_lossy();
    let error =
        |status, reason: String| json!({"path": reported_path, "status": status, "reason": reason});
    let closure = match load(path) {
        Ok(closure) => closure,
        Err(reason) => return error("input", reason),
    };
    let programs = match compile(&closure) {
        Ok(programs) => programs,
        Err(reason) => return error("compile", reason),
    };
    if !compare {
        let programs: Vec<_> = programs.iter().collect();
        return match calcvba_jit::JitEngine.compile_image(&programs) {
            Ok(_) => json!({"path": reported_path, "status": "jit-admitted"}),
            Err(calcvba_jit::JitError::Unsupported(reason)) => error("jit-refused", reason),
            Err(err) => error("jit-compile", err.to_string()),
        };
    }
    if !programs
        .last()
        .and_then(|program| program.entry.and_then(|id| program.funcs.get(id.0)))
        .is_some_and(|entry| entry.param_count == 0)
    {
        return error("no-entry", "a parameterless entry point is required".into());
    }
    let comparison = compare_entry(&programs);
    let vm3 = &comparison.vm3.outcome;
    let jit = &comparison.jit.outcome;
    let differences = comparison.differences();
    let status = if vm3.unsupported.is_some() {
        "vm3-refused"
    } else if vm3.result.is_err() && !vm3.raised {
        "vm3-failed"
    } else if jit.unsupported.is_some() {
        "jit-refused"
    } else if jit.result.is_err() && !jit.raised {
        "jit-failed"
    } else if !differences.is_empty() {
        "different"
    } else if comparison.has_opaque_values() {
        "matched-with-opaque-values"
    } else {
        "matched"
    };
    json!({
        "path": reported_path, "status": status, "differences": differences,
        "vm3": observation(&comparison.vm3), "jit": observation(&comparison.jit),
    })
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let usage =
        "usage: source_parity inspect <files or directories...> | compare <reviewed files...>";
    let mut args = env::args_os().skip(1);
    let compare = match args.next().as_deref().and_then(|arg| arg.to_str()) {
        Some("inspect") => false,
        Some("compare") => true,
        _ => return Err(usage.into()),
    };
    let mut inputs = BTreeSet::new();
    for path in args.map(PathBuf::from) {
        if compare {
            if !path.is_file() || !matches!(extension(&path).as_str(), "bas" | "basproj" | "vbp") {
                return Err(format!(
                    "compare requires an explicit reviewed source file or project: {}",
                    path.display()
                )
                .into());
            }
            inputs.insert(path);
        } else {
            collect(&path, &mut inputs)?;
        }
    }
    if inputs.is_empty() {
        return Err(usage.into());
    }
    for path in inputs {
        println!("{}", row(&path, compare));
    }
    Ok(())
}
