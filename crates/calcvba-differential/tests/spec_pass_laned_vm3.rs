//! dc-ez3rd lane D: the SpecStmt probe (statements, file system, Print/Write, On..GoTo, For)
//! run on a real file system with the Windows paths replaced by a temp directory, so it also
//! runs on Linux. Excel's answers are in `conformance/excel-measurements/spec-pass`.

use calcvba_differential::{Canon, Executor, canon};
use calcvba_engine::{Engine, HostConfig, RuntimeProfileId, SnapshotOutcome};
use calcvba_hal::model::HostPolicy;
use calcvba_symbol::manifest as sym;

/// Run the modules on a real-file-system engine and return the snapshot values.
fn run_native(executor: Executor, modules: &[(&str, &str)]) -> Result<Vec<Canon>, String> {
    let manifest = sym::SymbolProjectManifest {
        project_name: "VBAProject".to_string(),
        project_kind: sym::ProjectKind::Source,
        modules: modules
            .iter()
            .map(|(name, src)| sym::ModuleUnit {
                module_name: name.to_string(),
                module_kind: sym::ModuleKind::Procedural,
                attributes: sym::ModuleAttributes::named(*name),
                source: src.to_string(),
            })
            .collect(),
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: std::collections::BTreeMap::new(),
        conditional_compilation_target: Default::default(),
        codepage: None,
    };
    let config = match executor {
        Executor::Vm3 => HostConfig::vm3(),
        Executor::Jit => HostConfig::jit(),
    };
    let profile = if cfg!(windows) {
        RuntimeProfileId::WindowsHeadless
    } else {
        RuntimeProfileId::LinuxStdio
    };
    let mut engine = Engine::new(config).with_runtime_profile(profile);
    engine.set_host_policy(HostPolicy::native());
    let outcome = match executor {
        Executor::Vm3 => engine.execute_manifest_snapshot_with_err_vm3(&manifest),
        Executor::Jit => engine.execute_manifest_snapshot_with_err_jit(&manifest),
    };
    match outcome {
        SnapshotOutcome::Completed { values, .. } => Ok(values.iter().map(canon).collect()),
        SnapshotOutcome::Raised { err } => Err(format!("VBA error {}", err.number)),
        SnapshotOutcome::Unsupported(w) => Err(format!("unsupported {w}")),
        SnapshotOutcome::Failed(m) => Err(m),
    }
}

const SRC: &str = include_str!("../../../conformance/excel-measurements/spec-pass/SpecStmt.bas");
const WANT: &str =
    include_str!("../../../conformance/excel-measurements/spec-pass/SpecStmt.excel.txt");

fn run_probe(executor: Executor) -> Vec<String> {
    let dir = std::env::temp_dir().join(format!("cvp-spec-{executor:?}-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let root = dir.to_string_lossy().replace('\\', "/");
    let src = SRC
        .replace(
            "Root = Environ(\"TEMP\") & \"\\cvp-spec\"",
            &format!("Root = \"{root}\""),
        )
        .replace("\\a.txt", "/a.txt")
        .replace("\\b.txt", "/b.txt")
        .replace("\\none.txt", "/none.txt")
        .replace("\\nodir", "/nodir")
        .replace("\\d\\", "/d/")
        .replace("\\d", "/d")
        .replace("\\x.txt", "/x.txt")
        .replace("\\*", "/*")
        .replace("\\h.txt", "/h.txt")
        .replace("\\sub", "/sub")
        .replace("\\\"", "/\"");
    let main = "Public result As String\nSub Main()\n    result = SpecStmt.Run()\nEnd Sub\n";
    let values = run_native(executor, &[("Main", main), ("SpecStmt", src.as_str())])
        .unwrap_or_else(|e| panic!("failed: {e}"));
    let text = values
        .iter()
        .find_map(|v| match v {
            Canon::Str(s) if s.contains('=') => Some(s.clone()),
            _ => None,
        })
        .expect("no result");
    let _ = std::fs::remove_dir_all(&dir);
    text.lines().map(str::to_string).collect()
}

fn check(executor: Executor) {
    let got = run_probe(executor);
    let mut bad = Vec::new();
    for (i, want) in WANT.replace("\r\n", "\n").lines().enumerate() {
        // Number formatting inside Print/Write belongs to another lane, and a hidden file
        // cannot be made on Linux.
        let hidden_file = !cfg!(windows) && want.starts_with("dirAttrs");
        if hidden_file
            || ["printDbl", "writeDbl", "printBig"]
                .iter()
                .any(|k| want.starts_with(k))
        {
            continue;
        }
        let have = got.get(i).map(String::as_str).unwrap_or("<missing>");
        if have != want {
            bad.push(format!("{want} -- got {have}"));
        }
    }
    assert!(bad.is_empty(), "{executor:?}:\n{}", bad.join("\n"));
}

#[test]
fn vm3_spec_stmt() {
    check(Executor::Vm3);
}

#[test]
fn jit_spec_stmt() {
    check(Executor::Jit);
}
