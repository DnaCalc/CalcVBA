//! Inspect corpus sources without executing them, or compare explicitly listed reviewed fixtures.

use std::{
    collections::BTreeSet,
    env, fs,
    path::{Path, PathBuf},
};

use calcvba_differential::{
    excel_stub::{ExcelStubTypeLibs, excel_project_reference},
    parity::{
        Observation, compare_entry, compare_entry_with_cells, compare_entry_with_excel,
        compare_entry_with_files,
    },
};
use calcvba_oxir::{FuncId, OxProgram};
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

fn compile(closure: &[SymbolProjectManifest], excel: bool) -> Result<Vec<OxProgram>, String> {
    let _scope =
        calcvba_com::ComPolicyScope::enter(|_, _| Err("native COM disabled for inspection".into()));
    let owned: Vec<SymbolProjectManifest>;
    let closure = if excel {
        owned = closure
            .iter()
            .cloned()
            .map(|mut manifest| {
                let has_excel = manifest.references.iter().any(|reference| {
                    matches!(
                        reference,
                        calcvba_symbol::manifest::ProjectReference::TypeLibrary { name, .. }
                            if name.eq_ignore_ascii_case("Excel")
                    )
                });
                if !has_excel {
                    manifest.references.push(excel_project_reference());
                }
                manifest
            })
            .collect();
        &owned
    } else {
        closure
    };
    let resolver: &dyn calcvba_symbol::TypeLibResolver = if excel {
        &ExcelStubTypeLibs
    } else {
        &calcvba_symbol::CatalogTypeLibResolver
    };
    calcvba_bind::bind_projects(closure, resolver)
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
        "files": run.files.iter().map(|(path, bytes)| {
            json!({ "path": path.to_string_lossy(), "bytes": bytes })
        }).collect::<Vec<_>>(),
        "cells": run.cells.iter().map(|write| {
            json!({ "row": write.row, "column": write.column, "value": format!("{:?}", write.value) })
        }).collect::<Vec<_>>(),
        "denied_host_calls": run.denied_host_calls.iter().map(|cap| cap.name()).collect::<Vec<_>>(),
    })
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Mode {
    Inspect,
    Entry,
    Procedures,
    Files,
    Cells,
    Excel,
}

fn rows(path: &Path, mode: Mode, report: &mut impl FnMut(Value)) {
    let reported_path = path.to_string_lossy();
    let error =
        |status, reason: String| json!({"path": reported_path, "status": status, "reason": reason});
    let closure = match load(path) {
        Ok(closure) => closure,
        Err(reason) => return report(error("input", reason)),
    };
    let mut programs = match compile(&closure, mode == Mode::Excel) {
        Ok(programs) => programs,
        Err(reason) => return report(error("compile", reason)),
    };
    if mode == Mode::Inspect {
        let programs: Vec<_> = programs.iter().collect();
        return report(match calcvba_jit::JitEngine.compile_image(&programs) {
            Ok(_) => json!({"path": reported_path, "status": "jit-admitted"}),
            Err(calcvba_jit::JitError::Unsupported(reason)) => error("jit-refused", reason),
            Err(err) => error("jit-compile", err.to_string()),
        });
    }
    let Some(program) = programs.last() else {
        return report(error("no-entry", "empty project closure".into()));
    };
    if mode == Mode::Entry {
        let Some(entry) = program.entry else {
            return report(error(
                "no-entry",
                "a parameterless entry point is required".into(),
            ));
        };
        return report(compare(&programs, entry, path, mode));
    }
    let procedures: Vec<_> = (0..program.funcs.len())
        .map(FuncId)
        .filter(|id| Some(*id) != program.global_initializer)
        .collect();
    if procedures.is_empty() {
        return report(error(
            "no-entry",
            "the project contains no procedures".into(),
        ));
    }
    for proc in procedures {
        programs.last_mut().expect("entry program").entry = Some(proc);
        report(compare(&programs, proc, path, mode));
    }
}

fn compare(programs: &[OxProgram], entry: FuncId, path: &Path, mode: Mode) -> Value {
    let program = programs.last().expect("entry program");
    let procedure = &program.funcs[entry.0];
    let path = path.to_string_lossy();
    if procedure.param_count != 0 {
        return json!({"path": path, "procedure": procedure.name, "status": "requires-arguments"});
    }
    let comparison = match mode {
        Mode::Files => match compare_entry_with_files(programs) {
            Ok(comparison) => comparison,
            Err(err) => {
                return json!({"path": path, "procedure": procedure.name,
                "status": "scratch-failed", "reason": err.to_string()});
            }
        },
        Mode::Cells => compare_entry_with_cells(programs),
        Mode::Excel => compare_entry_with_excel(programs),
        _ => compare_entry(programs),
    };
    let vm3 = &comparison.vm3.outcome;
    let jit = &comparison.jit.outcome;
    let differences = comparison.differences();
    let failed = |run: &Observation| run.outcome.result.is_err() && !run.outcome.raised;
    let status = match (vm3.unsupported.is_some(), jit.unsupported.is_some()) {
        (true, true) => "both-refused",
        (true, false) if !failed(&comparison.jit) => "vm3-refused",
        (false, true) if !failed(&comparison.vm3) => "jit-refused",
        _ if failed(&comparison.vm3) && failed(&comparison.jit) => "both-failed",
        _ if failed(&comparison.vm3) => "vm3-failed",
        _ if failed(&comparison.jit) => "jit-failed",
        _ if !differences.is_empty() => "different",
        _ if !comparison.vm3.denied_host_calls.is_empty() => "host-limited",
        _ if comparison.has_opaque_values() => "matched-with-opaque-values",
        _ => "matched",
    };
    let return_slot = procedure
        .return_local
        .map(|id| program.globals.len() + id.0);
    json!({
        "path": path, "procedure": procedure.name, "status": status, "differences": differences,
        "return_slot": return_slot,
        "vm3": observation(&comparison.vm3), "jit": observation(&comparison.jit),
    })
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let usage = "usage: source_parity inspect <files or directories...> | \
                 {compare|procedures|files|cells|excel} <reviewed files...>";
    let mut args = env::args_os().skip(1);
    let mode = match args.next().as_deref().and_then(|arg| arg.to_str()) {
        Some("inspect") => Mode::Inspect,
        Some("compare") => Mode::Entry,
        Some("procedures") => Mode::Procedures,
        Some("files") => Mode::Files,
        Some("cells") => Mode::Cells,
        Some("excel") => Mode::Excel,
        _ => return Err(usage.into()),
    };
    let mut inputs = BTreeSet::new();
    for path in args.map(PathBuf::from) {
        if mode != Mode::Inspect {
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
        rows(&path, mode, &mut |row| println!("{row}"));
    }
    Ok(())
}
