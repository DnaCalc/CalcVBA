//! `calcvba-cli` — command-line driver for the clean execution stack.
//!
//! The run subcommands execute VBA through `calcvba_bind` →
//! `calcvba_oxir::elaborate` → `calcvba_vm3`:
//!   * `run <source.bas>` — execute a single source module.
//!   * `run-project [path] [--entry M.P]` — load a `.basproj`/`.vbp` (and its
//!     transitive project-reference graph) into a closure and execute it.
//!   * `build <project> --target WrappedComServer --out-dir <dir>` — emit the
//!     clean package and wrapper artifacts for an in-process COM DLL target.
//!
//! The run commands accept a platform host profile. The legacy compiler and host execution paths
//! remain removed; `build` is the clean wrapper-artifact lane.

mod office;
use std::path::{Path, PathBuf};
use std::{env, fs};

use calcvba_diagnostics::{
    Diagnostic as OxDiagnostic, DiagnosticPhase as OxDiagnosticPhase, DiagnosticReport,
};
use calcvba_engine::{
    Engine, ExecBackend, HostConfig, ResolvedRunnerBootstrap, RunnerBootstrapFallbacks,
    RunnerBootstrapOptions, resolve_runner_bootstrap, resolve_runner_bootstrap_with_fallbacks,
};
use calcvba_runtime::{VarType, Variant};

fn main() {
    let cli_args: Vec<String> = env::args().skip(1).collect();
    match cli_args.first().map(String::as_str) {
        Some(command @ ("inspect-vba" | "export-vba" | "write-vba")) => {
            if let Err(error) = office::run(command, &cli_args[1..]) {
                eprintln!("calcvba-cli: {error}");
                std::process::exit(1);
            }
        }
        Some("run-project") => run_project(cli_args),
        Some("build") => run_build(cli_args),
        Some("run-image") => run_execute(cli_args),
        Some("run") | None => run_execute(cli_args),
        Some("help") | Some("--help") | Some("-h") => print_usage(),
        Some(other) => {
            eprintln!("calcvba-cli: unknown subcommand `{other}`");
            print_usage();
            std::process::exit(2);
        }
    }
}
fn print_usage() {
    eprintln!(
        "usage:\n  calcvba-cli run <source.bas> [--backend vm3|jit] [--dump-values]\n  calcvba-cli run-project [path] [--entry Module.Procedure] [--backend vm3|jit]\n  calcvba-cli build <project> --target OxImage|NativeLibrary|Exe|ComExe|ComServer --out-dir <dir>\n  calcvba-cli run-image <project.oxi> [--backend vm3|jit]\n  calcvba-cli inspect-vba <workbook>\n  calcvba-cli export-vba <workbook> --out-dir <new-folder>\n  calcvba-cli write-vba <workbook> --project <project.basproj> --out <new-workbook>\n\n  --profile <id>  --dump-bootstrap  --diagnostic-format human|json"
    );
}

// ---------------------------------------------------------------------------
// run subcommand: execute a single VBA source module on the clean path
// ---------------------------------------------------------------------------

fn run_execute(mut cli_args: Vec<String>) {
    let image_mode = cli_args.first().is_some_and(|arg| arg == "run-image");
    if image_mode {
        cli_args[0] = "run".into();
    }
    let args = if cli_args.is_empty() {
        None
    } else {
        Some(parse_run_args_from(cli_args).unwrap_or_else(|| {
            eprintln!("usage: calcvba-cli run <source.bas> [--diagnostic-format <human|json>] [bootstrap options]");
            std::process::exit(2);
        }))
    };
    let diagnostic_format = args
        .as_ref()
        .map(|a| a.diagnostic_format)
        .unwrap_or_default();
    let config = HostConfig {
        backend: args.as_ref().map(|a| a.backend).unwrap_or(ExecBackend::Vm3),
    };
    let mut engine = Engine::new(config);
    if let Some(run_args) = args.as_ref() {
        let resolved = resolve_runner_bootstrap(&run_args.bootstrap, |key| env::var(key).ok())
            .unwrap_or_else(|err| {
                let diagnostic = OxDiagnostic::error(
                    "HOST-E-BOOTSTRAP",
                    OxDiagnosticPhase::Host,
                    format!("bootstrap failed: {err}"),
                );
                emit_diagnostic("calcvba-cli", &diagnostic, diagnostic_format);
                std::process::exit(2);
            });
        engine.set_runtime_profile(resolved.runtime_profile);
        engine.set_host_policy(resolved.policy.clone());
        if run_args.dump_bootstrap {
            println!("BOOTSTRAP:{}", resolved.fingerprint());
        }
    }
    let source = args
        .as_ref()
        .map(|a| a.source.clone())
        .unwrap_or_else(|| "Sub Main()\nEnd Sub".to_string());
    let dump_values = args.as_ref().map(|a| a.dump_values).unwrap_or(false);

    let outcome = if image_mode {
        engine.execute_image_bytes_with_variant_snapshot(source.as_bytes())
    } else {
        engine.execute_source_with_variant_snapshot_clean(&source)
    };
    match outcome {
        Ok(values) => {
            if dump_values {
                print_values(&values);
            }
        }
        Err(err) => {
            emit_diagnostic(
                "calcvba-cli: execution failed",
                err.diagnostic(),
                diagnostic_format,
            );
            std::process::exit(1);
        }
    }
}

// ---------------------------------------------------------------------------
// run-project subcommand: load a `.basproj`/`.vbp` closure and execute it
// ---------------------------------------------------------------------------

fn run_project(args: Vec<String>) {
    let parsed = parse_run_project_args_from(args).unwrap_or_else(|| {
        eprintln!(
            "usage: calcvba-cli run-project [path] [--entry <Module.Procedure>] [--diagnostic-format <human|json>] [bootstrap options]"
        );
        std::process::exit(2);
    });

    // The clean path runs a `.basproj`/`.vbp` and its reference graph as a closure.
    let project_file = match resolve_project_file(parsed.input_path.as_ref()) {
        Ok(project_file) => project_file,
        Err(err) => {
            let location = parsed
                .input_path
                .as_deref()
                .map(|p| p.display().to_string())
                .unwrap_or_else(|| ".".to_string());
            let diagnostic = match err {
                ResolveProjectError::Ambiguous(err) => OxDiagnostic::error(
                    "PROJ-E-PROJECT-FILE-AMBIGUOUS",
                    OxDiagnosticPhase::ProjectLoad,
                    format!("{err}"),
                )
                .with_help(
                    "Pass an explicit .basproj or .vbp path to select which project to run.",
                ),
                ResolveProjectError::NotFound => OxDiagnostic::error(
                    "PROJ-E-PROJECT-FILE-NOT-FOUND",
                    OxDiagnosticPhase::ProjectLoad,
                    format!("no .basproj or .vbp project file found at {location}"),
                )
                .with_help("Pass an explicit .basproj or .vbp path, or run from a directory containing exactly one project file."),
            };
            emit_diagnostic(
                "calcvba-cli run-project",
                &diagnostic,
                parsed.diagnostic_format,
            );
            std::process::exit(1);
        }
    };

    // The single-project view supplies the runner-bootstrap fallbacks (default
    // profile / policy declared in the project file).
    let loaded = load_project_for_bootstrap(&project_file).unwrap_or_else(|err| {
        emit_diagnostic(
            "calcvba-cli run-project",
            &err.to_diagnostic(),
            parsed.diagnostic_format,
        );
        std::process::exit(1);
    });

    let mut engine = Engine::new(HostConfig {
        backend: parsed.backend,
    });
    let resolved =
        resolve_project_runner_bootstrap(&loaded, &parsed.bootstrap, |key| env::var(key).ok())
            .unwrap_or_else(|err| {
                let diagnostic = OxDiagnostic::error(
                    "HOST-E-BOOTSTRAP",
                    OxDiagnosticPhase::Host,
                    format!("bootstrap failed: {err}"),
                );
                emit_diagnostic(
                    "calcvba-cli run-project",
                    &diagnostic,
                    parsed.diagnostic_format,
                );
                std::process::exit(2);
            });
    engine.set_runtime_profile(resolved.runtime_profile);
    engine.set_host_policy(resolved.policy.clone());
    if parsed.dump_bootstrap {
        println!("BOOTSTRAP:{}", resolved.fingerprint());
    }

    let closure = calcvba_project::load_project_closure_with_entry(
        &project_file,
        parsed.entry_point_override.as_deref(),
    )
    .unwrap_or_else(|err| {
        emit_diagnostic(
            "calcvba-cli run-project",
            &err.to_diagnostic(),
            parsed.diagnostic_format,
        );
        std::process::exit(1);
    });

    match engine.execute_project_closure_with_variant_snapshot(&closure) {
        Ok(values) => {
            if parsed.dump_values {
                print_values(&values);
            }
        }
        Err(err) => {
            emit_diagnostic(
                "calcvba-cli run-project",
                err.diagnostic(),
                parsed.diagnostic_format,
            );
            std::process::exit(1);
        }
    }
}

// ---------------------------------------------------------------------------
// build subcommand: emit clean wrapper artifacts
// ---------------------------------------------------------------------------

fn run_build(args: Vec<String>) {
    let parsed = parse_build_args_from(args).unwrap_or_else(|| {
        eprintln!(
            "usage: calcvba-cli build <project.basproj|project.vbp> --target <ComServer|OxImage|NativeLibrary|Exe|ComExe> --out-dir <dir> [--diagnostic-format <human|json>]"
        );
        std::process::exit(2);
    });

    if parsed.native_library || parsed.native_executable || parsed.com_executable {
        let options = calcvba_build::NativeLibraryBuildOptions {
            project_path: parsed.project_path,
            out_dir: parsed.out_dir,
            template_path: None,
        };
        let output = if parsed.com_executable {
            calcvba_build::build_native_com_executable(&options)
                .map(|out| ("COM EXE", out.exe_path, out.map_path))
        } else if parsed.native_executable {
            calcvba_build::build_native_executable(&options)
                .map(|out| ("EXE", out.exe_path, out.map_path))
        } else {
            calcvba_build::build_native_library(&options)
                .map(|out| ("DLL", out.dll_path, out.map_path))
        };
        match output {
            Ok((kind, path, map)) => {
                println!("{kind}:{}", path.display());
                println!("MAP:{}", map.display());
            }
            Err(error) => {
                emit_diagnostic(
                    "calcvba-cli build",
                    &OxDiagnostic::error(
                        "BUILD-E-NATIVE-EXPORT",
                        OxDiagnosticPhase::Host,
                        error.to_string(),
                    ),
                    parsed.diagnostic_format,
                );
                std::process::exit(1);
            }
        }
        return;
    }
    if parsed.oximage {
        match calcvba_build::build_oximage(&parsed.project_path, &parsed.out_dir) {
            Ok(path) => println!("OXI:{}", path.display()),
            Err(err) => {
                emit_diagnostic(
                    "calcvba-cli build",
                    &OxDiagnostic::error(
                        "BUILD-E-OXIMAGE",
                        OxDiagnosticPhase::Host,
                        err.to_string(),
                    ),
                    parsed.diagnostic_format,
                );
                std::process::exit(1);
            }
        }
        return;
    }

    match calcvba_build::build_wrapped_com_server(&calcvba_build::WrappedComServerBuildOptions {
        project_path: parsed.project_path,
        out_dir: parsed.out_dir,
        compile_dll: true,
        comhost_dll_path: None,
    }) {
        Ok(output) => {
            println!("OXI:{}", output.oxi_path.display());
            println!("COM_DESCRIPTOR:{}", output.descriptor_path.display());
            println!("IDL:{}", output.idl_path.display());
            if let Some(path) = output.comhost_source_path.as_ref() {
                println!("COMHOST_SOURCE:{}", path.display());
            }
            println!("DLL_TARGET:{}", output.dll_target_path.display());
            println!("TLB_TARGET:{}", output.tlb_target_path.display());
        }
        Err(err) => {
            let diagnostic = OxDiagnostic::error(
                "BUILD-E-WRAPPED-COM-SERVER",
                OxDiagnosticPhase::Host,
                format!("WrappedComServer build failed: {err}"),
            );
            emit_diagnostic("calcvba-cli build", &diagnostic, parsed.diagnostic_format);
            std::process::exit(1);
        }
    }
}

/// Why a run-project input could not be resolved to a single project file.
enum ResolveProjectError {
    /// No `.basproj`/`.vbp` at the path (or a non-project file argument).
    NotFound,
    /// A directory contains more than one candidate project file — reporting
    /// this as "not found" (the old behavior) hid the real cause.
    Ambiguous(calcvba_project::BasProjError),
}

/// Resolve a run-project input to a concrete `.basproj`/`.vbp` file. A directory
/// is searched for a unique project file; a directory with several candidates is
/// [`ResolveProjectError::Ambiguous`], and one with none (or a non-project file
/// argument) is [`ResolveProjectError::NotFound`].
fn resolve_project_file(input: Option<&PathBuf>) -> Result<PathBuf, ResolveProjectError> {
    let input = input.cloned().unwrap_or_else(|| PathBuf::from("."));
    if input.is_dir() {
        for ext in ["basproj", "vbp"] {
            match calcvba_project::discover_project_file_in_dir(&input, ext) {
                Ok(Some(p)) => return Ok(p),
                Ok(None) => {}
                Err(err) => return Err(ResolveProjectError::Ambiguous(err)),
            }
        }
        return Err(ResolveProjectError::NotFound);
    }
    match input.extension().and_then(|e| e.to_str()) {
        Some(ext) if ext.eq_ignore_ascii_case("basproj") || ext.eq_ignore_ascii_case("vbp") => {
            Ok(input)
        }
        _ => Err(ResolveProjectError::NotFound),
    }
}

fn load_project_for_bootstrap(
    file: &Path,
) -> Result<calcvba_project::LoadedProject, calcvba_project::BasProjError> {
    if file
        .extension()
        .and_then(|e| e.to_str())
        .is_some_and(|e| e.eq_ignore_ascii_case("vbp"))
    {
        calcvba_project::load_vbp(file)
    } else {
        calcvba_project::load_basproj(file)
    }
}

fn resolve_project_runner_bootstrap(
    loaded: &calcvba_project::LoadedProject,
    bootstrap: &RunnerBootstrapOptions,
    env_get: impl Fn(&str) -> Option<String>,
) -> Result<ResolvedRunnerBootstrap, String> {
    resolve_runner_bootstrap_with_fallbacks(
        bootstrap,
        &RunnerBootstrapFallbacks {
            profile: loaded.default_runtime_profile.clone(),
        },
        env_get,
    )
}

fn print_values(values: &[Variant]) {
    let payload = values
        .iter()
        .map(format_variant_value)
        .collect::<Vec<_>>()
        .join("|");
    println!("VALUES:{payload}");
}

fn emit_diagnostic(prefix: &str, diagnostic: &OxDiagnostic, format: DiagnosticFormat) {
    match format {
        DiagnosticFormat::Human => eprintln!("{prefix}: {}", diagnostic.render_human()),
        DiagnosticFormat::Json => {
            match DiagnosticReport::single(diagnostic.clone()).to_json_pretty() {
                Ok(json) => eprintln!("{json}"),
                Err(err) => eprintln!(
                    "{prefix}: failed to serialize diagnostic JSON: {err}; {}",
                    diagnostic.render_human()
                ),
            }
        }
    }
}

// ---------------------------------------------------------------------------
// argument parsing
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
enum DiagnosticFormat {
    #[default]
    Human,
    Json,
}

#[derive(Debug, Clone)]
struct RunArgs {
    source: String,
    dump_values: bool,
    dump_bootstrap: bool,
    backend: ExecBackend,
    diagnostic_format: DiagnosticFormat,
    bootstrap: RunnerBootstrapOptions,
}

#[derive(Debug, Clone)]
struct RunProjectArgs {
    input_path: Option<PathBuf>,
    backend: ExecBackend,
    dump_values: bool,
    dump_bootstrap: bool,
    diagnostic_format: DiagnosticFormat,
    bootstrap: RunnerBootstrapOptions,
    entry_point_override: Option<String>,
}

#[derive(Debug, Clone)]
struct BuildArgs {
    project_path: PathBuf,
    out_dir: PathBuf,
    diagnostic_format: DiagnosticFormat,
    oximage: bool,
    native_library: bool,
    native_executable: bool,
    com_executable: bool,
}

fn parse_run_args_from(args: Vec<String>) -> Option<RunArgs> {
    let mut iter = args.into_iter();
    if iter.next()? != "run" {
        return None;
    }
    let collected: Vec<String> = iter.collect();
    let mut path: Option<String> = None;
    let mut dump_values = false;
    let mut dump_bootstrap = false;
    let mut backend = ExecBackend::Vm3;
    let mut diagnostic_format = DiagnosticFormat::Human;
    let mut bootstrap = RunnerBootstrapOptions::default();

    let mut i = 0;
    while i < collected.len() {
        let arg = collected[i].as_str();
        match arg {
            "--dump-values" => dump_values = true,
            "--dump-bootstrap" => dump_bootstrap = true,
            "--jit" => backend = ExecBackend::Jit,
            "--backend" => {
                i += 1;
                backend = parse_exec_backend(collected.get(i)?)?;
            }
            "--diagnostic-format" => {
                i += 1;
                diagnostic_format = parse_diagnostic_format(collected.get(i)?)?;
            }
            _ => match consume_bootstrap_flag(&mut bootstrap, arg, collected.get(i + 1)) {
                Some(Ok(())) => i += 1,
                Some(Err(())) => return None,
                None if !arg.starts_with('-') && path.is_none() => path = Some(arg.to_string()),
                None => return None,
            },
        }
        i += 1;
    }

    let source = fs::read_to_string(path?).ok()?;
    Some(RunArgs {
        source,
        dump_values,
        dump_bootstrap,
        backend,
        diagnostic_format,
        bootstrap,
    })
}

fn parse_build_args_from(args: Vec<String>) -> Option<BuildArgs> {
    let mut iter = args.into_iter();
    if iter.next()? != "build" {
        return None;
    }
    let collected: Vec<String> = iter.collect();
    let mut project_path: Option<PathBuf> = None;
    let mut target: Option<String> = None;
    let mut out_dir: Option<PathBuf> = None;
    let mut diagnostic_format = DiagnosticFormat::Human;

    let mut i = 0;
    while i < collected.len() {
        let arg = collected[i].as_str();
        match arg {
            "--target" => {
                i += 1;
                target = Some(collected.get(i)?.clone());
            }
            "--out-dir" => {
                i += 1;
                out_dir = Some(PathBuf::from(collected.get(i)?));
            }
            "--diagnostic-format" => {
                i += 1;
                diagnostic_format = parse_diagnostic_format(collected.get(i)?)?;
            }
            _ if !arg.starts_with('-') && project_path.is_none() => {
                project_path = Some(PathBuf::from(arg));
            }
            _ => return None,
        }
        i += 1;
    }

    let target = target?;
    let oximage = target.eq_ignore_ascii_case("OxImage") || target.eq_ignore_ascii_case("Bundle");
    let native_library = target.eq_ignore_ascii_case("NativeLibrary");
    let native_executable = target.eq_ignore_ascii_case("Exe");
    let com_executable = target.eq_ignore_ascii_case("ComExe");
    if !oximage
        && !native_library
        && !native_executable
        && !com_executable
        && !target.eq_ignore_ascii_case("ComServer")
        && !target.eq_ignore_ascii_case("WrappedComServer")
    {
        return None;
    }

    Some(BuildArgs {
        project_path: project_path?,
        out_dir: out_dir?,
        diagnostic_format,
        oximage,
        native_library,
        native_executable,
        com_executable,
    })
}

fn parse_run_project_args_from(args: Vec<String>) -> Option<RunProjectArgs> {
    let mut iter = args.into_iter();
    if iter.next()? != "run-project" {
        return None;
    }
    let collected: Vec<String> = iter.collect();
    let mut input_path: Option<PathBuf> = None;
    let mut backend = ExecBackend::Vm3;
    let mut dump_values = false;
    let mut dump_bootstrap = false;
    let mut diagnostic_format = DiagnosticFormat::Human;
    let mut bootstrap = RunnerBootstrapOptions::default();
    let mut entry_point_override: Option<String> = None;

    let mut i = 0;
    while i < collected.len() {
        let arg = collected[i].as_str();
        match arg {
            "--jit" => backend = ExecBackend::Jit,
            "--backend" => {
                i += 1;
                backend = parse_exec_backend(collected.get(i)?)?;
            }
            "--dump-values" => dump_values = true,
            "--dump-bootstrap" => dump_bootstrap = true,
            "--diagnostic-format" => {
                i += 1;
                diagnostic_format = parse_diagnostic_format(collected.get(i)?)?;
            }
            "--entry" => {
                i += 1;
                entry_point_override = Some(collected.get(i)?.clone());
            }
            _ => match consume_bootstrap_flag(&mut bootstrap, arg, collected.get(i + 1)) {
                Some(Ok(())) => i += 1,
                Some(Err(())) => return None,
                None if !arg.starts_with('-') && input_path.is_none() => {
                    input_path = Some(PathBuf::from(arg));
                }
                None => return None,
            },
        }
        i += 1;
    }

    Some(RunProjectArgs {
        input_path,
        backend,
        dump_values,
        dump_bootstrap,
        diagnostic_format,
        bootstrap,
        entry_point_override,
    })
}
fn consume_bootstrap_flag(
    bootstrap: &mut RunnerBootstrapOptions,
    flag: &str,
    value: Option<&String>,
) -> Option<Result<(), ()>> {
    if flag != "--profile" {
        return None;
    }
    Some(
        value
            .map(|value| {
                bootstrap.profile = Some(value.clone());
            })
            .ok_or(()),
    )
}

// ---------------------------------------------------------------------------
// value formatting + scalar parsers
// ---------------------------------------------------------------------------

fn format_variant_value(value: &Variant) -> String {
    match value.vtype() {
        VarType::Empty => "empty".to_string(),
        VarType::Null => "null".to_string(),
        VarType::Error => format!("error:{}", value.as_error_code().unwrap_or(0)),
        VarType::Integer => format!("i16:{}", value.as_i16().unwrap_or(0)),
        VarType::Long => format!("i32:{}", value.as_i32().unwrap_or(0)),
        VarType::SignedByte => format!("i8:{}", value.as_i8().unwrap_or(0)),
        VarType::Byte => format!("u8:{}", value.as_u8().unwrap_or(0)),
        VarType::UnsignedInteger => format!("u16:{}", value.as_u16().unwrap_or(0)),
        VarType::UnsignedLong => format!("u32:{}", value.as_u32().unwrap_or(0)),
        VarType::UnsignedInt => format!("uint:{}", value.as_u32().unwrap_or(0)),
        VarType::LongLong => format!("i64:{}", value.as_i64().unwrap_or(0)),
        VarType::UnsignedLongLong => format!("u64:{}", value.as_u64().unwrap_or(0)),
        VarType::Single => format!("f64:{}", f64::from(value.as_f32().unwrap_or(0.0))),
        VarType::Double => format!("f64:{}", value.as_f64().unwrap_or(0.0)),
        VarType::Date => format!("f64:{}", value.as_date_f64().unwrap_or(0.0)),
        VarType::Decimal => match value.as_decimal96() {
            Some(value) => format!("decimal:{value}"),
            None => "decimal:<invalid>".to_string(),
        },
        VarType::Currency => format!(
            "currency:{}",
            calcvba_runtime::CurrencyValue::from_scaled_i64(
                value.as_currency_scaled_i64().unwrap_or(0)
            )
        ),
        VarType::Boolean => format!("bool:{}", value.as_bool().unwrap_or(false)),
        VarType::String => format!(
            "string:{:?}",
            value
                .as_bstr()
                .map(|text| text.as_str())
                .unwrap_or_default()
        ),
        VarType::ArrayVariant => match value.as_safearray() {
            Some(array) => format!("array:{array:?}"),
            None => "array:<invalid>".to_string(),
        },
        VarType::Object => match value.as_object_ref() {
            Some(handle) => format!("object:{handle}"),
            None => "object:<null>".to_string(),
        },
        VarType::Record => match value.as_com_record() {
            Some(record) => format!("record:{record:?}"),
            None => "record:<invalid>".to_string(),
        },
        VarType::ProcRef => match value.as_proc_ref() {
            Some(proc) => format!("proc:{proc}"),
            None => "proc:<invalid>".to_string(),
        },
    }
}
fn parse_diagnostic_format(value: &str) -> Option<DiagnosticFormat> {
    match value.trim().to_ascii_lowercase().as_str() {
        "human" => Some(DiagnosticFormat::Human),
        "json" => Some(DiagnosticFormat::Json),
        _ => None,
    }
}

fn parse_exec_backend(value: &str) -> Option<ExecBackend> {
    ExecBackend::parse(value)
}
#[cfg(test)]
mod tests {
    #[test]
    fn parse_com_server_target() {
        let parsed = super::parse_build_args_from(
            [
                "build",
                "Project.basproj",
                "--target",
                "ComServer",
                "--out-dir",
                "out",
            ]
            .into_iter()
            .map(str::to_owned)
            .collect(),
        )
        .unwrap();
        assert!(!parsed.oximage && !parsed.native_library);
    }

    use super::*;

    fn unique_temp_dir(tag: &str) -> PathBuf {
        use std::sync::atomic::{AtomicU32, Ordering};
        static COUNTER: AtomicU32 = AtomicU32::new(0);
        let n = COUNTER.fetch_add(1, Ordering::Relaxed);
        let dir =
            std::env::temp_dir().join(format!("calcvba_cli_test_{}_{n}_{tag}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("create temp dir");
        dir
    }

    #[test]
    fn resolve_project_file_reports_ambiguous_dir() {
        // Two candidate project files must be reported as ambiguous, not as
        // "no project file found" (the old behavior swallowed the Err).
        let dir = unique_temp_dir("ambig");
        std::fs::write(dir.join("a.basproj"), "").unwrap();
        std::fs::write(dir.join("b.basproj"), "").unwrap();
        assert!(
            matches!(
                resolve_project_file(Some(&dir)),
                Err(ResolveProjectError::Ambiguous(_))
            ),
            "two .basproj files must be Ambiguous"
        );
        std::fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn resolve_project_file_single_and_missing() {
        let dir = unique_temp_dir("single");
        std::fs::write(dir.join("only.basproj"), "").unwrap();
        assert!(
            matches!(resolve_project_file(Some(&dir)), Ok(p) if p.ends_with("only.basproj")),
            "a single project file resolves"
        );
        std::fs::remove_dir_all(&dir).ok();

        let empty = unique_temp_dir("empty");
        assert!(matches!(
            resolve_project_file(Some(&empty)),
            Err(ResolveProjectError::NotFound)
        ));
        std::fs::remove_dir_all(&empty).ok();
    }

    #[test]
    fn parse_run_args_with_flags() {
        let args = vec![
            "run".to_string(),
            "Cargo.toml".to_string(),
            "--dump-values".to_string(),
            "--jit".to_string(),
        ];
        let parsed = parse_run_args_from(args).expect("args should parse");
        assert!(parsed.dump_values);
        assert_eq!(parsed.backend, ExecBackend::Jit);
    }

    #[test]
    fn parse_run_args_with_backend_jit() {
        let args = vec![
            "run".to_string(),
            "Cargo.toml".to_string(),
            "--backend".to_string(),
            "jit".to_string(),
        ];
        let parsed = parse_run_args_from(args).expect("args should parse");
        assert_eq!(parsed.backend, ExecBackend::Jit);
    }

    #[test]
    fn parse_run_args_rejects_noncanonical_backend_alias() {
        let args = vec![
            "run".to_string(),
            "Cargo.toml".to_string(),
            "--backend".to_string(),
            "vm".to_string(),
        ];
        assert!(parse_run_args_from(args).is_none());
    }

    #[test]
    fn parse_runner_bootstrap_flags() {
        let args = vec![
            "run".to_string(),
            "Cargo.toml".to_string(),
            "--diagnostic-format".to_string(),
            "json".to_string(),
            "--profile".to_string(),
            "linux-stdio".to_string(),
            "--dump-bootstrap".to_string(),
        ];
        let parsed = parse_run_args_from(args).expect("args should parse");
        assert!(parsed.dump_bootstrap);
        assert_eq!(parsed.diagnostic_format, DiagnosticFormat::Json);
        assert_eq!(parsed.bootstrap.profile.as_deref(), Some("linux-stdio"));
    }

    #[test]
    fn reject_unknown_flags() {
        let args = vec![
            "run".to_string(),
            "Cargo.toml".to_string(),
            "--unknown".to_string(),
        ];
        assert!(parse_run_args_from(args).is_none());
    }

    #[test]
    fn reject_bootstrap_flag_missing_value() {
        let args = vec![
            "run".to_string(),
            "Cargo.toml".to_string(),
            "--profile".to_string(),
        ];
        assert!(parse_run_args_from(args).is_none());
    }

    #[test]
    fn format_variant_value_preserves_date_payloads() {
        let value = Variant::from_date_f64(46_081.0);
        assert_eq!(format_variant_value(&value), "f64:46081");
    }

    #[test]
    fn format_variant_value_preserves_single_payloads() {
        let value = Variant::from_f32(0.3335753);
        assert_eq!(format_variant_value(&value), "f64:0.3335753083229065");
    }

    #[test]
    fn parse_run_project_args_with_entry_override() {
        let args = vec![
            "run-project".to_string(),
            ".".to_string(),
            "--entry".to_string(),
            "Startup.Boot".to_string(),
            "--diagnostic-format".to_string(),
            "json".to_string(),
            "--profile".to_string(),
            "windows-stdio".to_string(),
            "--jit".to_string(),
        ];
        let parsed = parse_run_project_args_from(args).expect("args should parse");
        assert_eq!(parsed.entry_point_override.as_deref(), Some("Startup.Boot"));
        assert_eq!(parsed.bootstrap.profile.as_deref(), Some("windows-stdio"));
        assert_eq!(parsed.backend, ExecBackend::Jit);
        assert_eq!(parsed.diagnostic_format, DiagnosticFormat::Json);
        assert_eq!(parsed.input_path, Some(PathBuf::from(".")));
    }

    #[test]
    fn parse_native_executable_build_target() {
        let parsed = parse_build_args_from(
            [
                "build",
                "Hello.basproj",
                "--target",
                "Exe",
                "--out-dir",
                "out",
            ]
            .into_iter()
            .map(str::to_owned)
            .collect(),
        )
        .expect("Exe target");
        assert!(parsed.native_executable);
        assert!(!parsed.native_library && !parsed.oximage);
    }

    #[test]
    fn parse_build_args_for_wrapped_com_server() {
        let args = vec![
            "build".to_string(),
            "demo.basproj".to_string(),
            "--target".to_string(),
            "WrappedComServer".to_string(),
            "--out-dir".to_string(),
            "target/calcvba-build/demo".to_string(),
            "--diagnostic-format".to_string(),
            "json".to_string(),
        ];
        let parsed = parse_build_args_from(args).expect("args should parse");
        assert_eq!(parsed.project_path, PathBuf::from("demo.basproj"));
        assert_eq!(parsed.out_dir, PathBuf::from("target/calcvba-build/demo"));
        assert_eq!(parsed.diagnostic_format, DiagnosticFormat::Json);
    }

    #[test]
    fn accept_bundle_compatibility_spelling_as_oximage_target() {
        let args = vec![
            "build".to_string(),
            "demo.basproj".to_string(),
            "--target".to_string(),
            "Bundle".to_string(),
            "--out-dir".to_string(),
            "target/calcvba-build/demo".to_string(),
        ];
        assert!(
            parse_build_args_from(args)
                .expect("Bundle is an OxImage target")
                .oximage
        );
    }

    #[test]
    fn reject_unknown_diagnostic_format() {
        let args = vec![
            "run".to_string(),
            "Cargo.toml".to_string(),
            "--diagnostic-format".to_string(),
            "sarif".to_string(),
        ];
        assert!(parse_run_args_from(args).is_none());
    }

    #[test]
    fn format_scalar_variants() {
        assert_eq!(format_variant_value(&Variant::from_i32(42)), "i32:42");
        assert_eq!(format_variant_value(&Variant::from_bool(true)), "bool:true");
        assert_eq!(format_variant_value(&Variant::empty()), "empty");
    }
}
