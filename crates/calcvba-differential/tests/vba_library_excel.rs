//! The VBA standard library against Excel VBA: every probe module in
//! `conformance/excel-measurements/vba-library/probes` (one per library function, listed in
//! `functions.tsv`, from VBE7.DLL's type library and MS-VBAL section 6) runs on VM3 and the JIT
//! together with `ProbeFmt.bas`, and each `Function.k=value` line is compared with Excel's
//! answer in `<Module>.excel.txt`. Bit-exact is the only passing state.
//!
//! Lines that still differ are listed in `known-differences.txt` as `vm3 |jit ` + `Module: line`
//! (unprefixed when both differ), or `Module: run fails` when the module does not run at all.
//! With `VBA_LIBRARY_REPORT=<path>` the test also writes every difference, with CalcVBA's
//! answer, to that file.
//!
//! Both executors run in ordinary workspace checks. For a selected module use
//! `VBA_LIBRARY_ONLY=P_Left cargo test -p calcvba-differential --test vba_library_excel matches_excel`.
//! The cases are still isolated if a module cannot bind or run as a whole.
//! Set `VBA_LIBRARY_USE_HOST_LOCALE=1` when comparing answers freshly captured
//! with the host's current regional settings.

use calcvba_differential::{Canon, Executor, run_modules};
use calcvba_symbol::manifest::ModuleKind::Procedural;
use std::collections::BTreeSet;
use std::path::PathBuf;

/// Modules that touch the file system through Windows paths; checked on Windows only.
const WINDOWS_ONLY_PREFIX: &[&str] = &[
    "P_ChDir",
    "P_ChDrive",
    "P_CurDir",
    "P_Dir",
    "P_EOF",
    "P_FileAttr",
    "P_FileCopy",
    "P_FileDateTime",
    "P_FileLen",
    "P_FreeFile",
    "P_GetAttr",
    "P_Kill",
    "P_Loc",
    "P_LOF",
    "P_MkDir",
    "P_Reset",
    "P_RmDir",
    "P_Seek",
    "P_SetAttr",
    "P_Input",
    "P_InputB",
    "P_Width",
];

/// Cases too slow for a debug-build test run (module, case): blanked before running and not
/// compared. Their Excel answers stay recorded.
const HEAVY: &[(&str, u32)] = &[
    ("P_Rnd", 34), // 16,777,215 Rnd calls (the LCG period)
];

/// Replace dropped probe cases with one `Case Else: v = Empty`. Keeping an empty
/// branch per dropped case makes debug JIT compilation unnecessarily expensive.
fn blank(source: &str, drop: &[u32]) -> String {
    let mut out = Vec::new();
    let (mut in_select, mut depth, mut dropping, mut blanked) = (false, 0usize, false, false);
    for line in source.lines() {
        let t = line.trim_start();
        let lower = t.to_ascii_lowercase();
        if !in_select {
            out.push(line.to_string());
            in_select = lower.starts_with("select case k");
            blanked = false;
            continue;
        }
        if depth == 0 && lower.starts_with("end select") {
            if blanked {
                out.push("        Case Else: v = Empty".to_string());
            }
            in_select = false;
            dropping = false;
            out.push(line.to_string());
            continue;
        }
        if depth == 0 && lower.starts_with("case else") {
            if blanked {
                let labels = drop
                    .iter()
                    .map(u32::to_string)
                    .collect::<Vec<_>>()
                    .join(", ");
                out.push(format!("        Case {labels}: v = Empty"));
                blanked = false;
            }
            dropping = false;
            out.push(line.to_string());
            continue;
        }
        if depth == 0 && lower.starts_with("case ") {
            let number = t[5..]
                .split(|c: char| !c.is_ascii_digit())
                .next()
                .and_then(|n| n.parse::<u32>().ok());
            if let Some(n) = number {
                dropping = drop.contains(&n);
                if dropping {
                    blanked = true;
                    continue;
                }
            }
        }
        if lower.starts_with("select case") {
            depth += 1;
        } else if lower.starts_with("end select") {
            depth = depth.saturating_sub(1);
        }
        if !dropping {
            out.push(line.to_string());
        }
    }
    out.join("\n") + "\n"
}

#[test]
fn isolated_probe_cases_preserve_nested_selects_and_else_behavior() {
    let source = r#"
Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Select Case k
        Case 1
            Select Case 2
                Case 1: v = "wrong"
                Case Else: v = "nested"
            End Select
        Case 2
            Select Case 1
                Case Else: Err.Raise 5
            End Select
        Case 3: Err.Raise 6
        Case Else: v = "original else"
    End Select
    T = "C." & k & "=" & TypeName(v) & ":" & v
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To 4
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
"#;
    for executor in [Executor::Vm3, Executor::Jit] {
        let actual = run_text(executor, "", "P_Isolation", &blank(source, &[2, 3])).unwrap();
        assert_eq!(
            actual,
            "C.1=String:nested\nC.2=Empty:\nC.3=Empty:\nC.4=String:original else\n"
        );
        let source = source.replace("        Case Else: v = \"original else\"\n", "");
        let actual = run_text(executor, "", "P_Isolation", &blank(&source, &[2, 3])).unwrap();
        assert_eq!(
            actual,
            "C.1=String:nested\nC.2=Empty:\nC.3=Empty:\nC.4=Empty:\n"
        );
    }
}

fn dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../conformance/excel-measurements/vba-library")
}

/// A known line may start with a host (`linux `, `windows `, `macos `): it then applies only on
/// that host, because the host maths library or code page decides the answer (dc-q6brh.112).
fn host_line(line: &str) -> Option<&str> {
    for host in ["linux", "windows", "macos"] {
        if let Some(rest) = line.strip_prefix(host).and_then(|r| r.strip_prefix(' ')) {
            return (host == std::env::consts::OS).then_some(rest);
        }
    }
    Some(line)
}

fn known(executor: Executor) -> BTreeSet<String> {
    let own = match executor {
        Executor::Vm3 => "vm3 ",
        Executor::Jit => "jit ",
    };
    let text = std::fs::read_to_string(dir().join("known-differences.txt")).unwrap_or_default();
    text.lines()
        .map(str::trim_end)
        .filter(|l| !l.trim().is_empty() && !l.starts_with('#'))
        .filter_map(host_line)
        .filter_map(|l| {
            if let Some(rest) = l.strip_prefix(own) {
                Some(rest.to_string())
            } else if l.starts_with("vm3 ") || l.starts_with("jit ") {
                None
            } else {
                Some(l.to_string())
            }
        })
        .collect()
}

/// Every probe module that has an Excel answer: (module, source, answer).
fn cases() -> Vec<(String, String, String)> {
    let probes = dir().join("probes");
    let mut out = Vec::new();
    let mut entries: Vec<_> = std::fs::read_dir(&probes)
        .expect("probes dir")
        .filter_map(Result::ok)
        .map(|e| e.path())
        .filter(|p| {
            p.extension().is_some_and(|x| x == "bas")
                && p.file_stem()
                    .is_some_and(|s| s.to_string_lossy().starts_with("P_"))
        })
        .collect();
    entries.sort();
    for path in entries {
        let module = path.file_stem().unwrap().to_string_lossy().to_string();
        // `VBA_LIBRARY_ONLY=P_Dir,P_Kill` runs just those modules (comparison limited to them).
        if let Ok(only) = std::env::var("VBA_LIBRARY_ONLY")
            && !only.split(',').any(|m| m == module)
        {
            continue;
        }
        let Ok(answer) = std::fs::read_to_string(probes.join(format!("{module}.excel.txt"))) else {
            continue;
        };
        let source = std::fs::read_to_string(&path).expect("probe source");
        out.push((module, source, answer.replace("\r\n", "\n")));
    }
    out
}

/// Run the Windows file-system probes or Shell with native host services.
/// `run_modules` uses the deterministic policy.
fn run_native(executor: Executor, modules: &[(&str, &str)]) -> Result<String, String> {
    // These probes share process current-directory state and scratch paths.
    // Keep their runtime calls on the invoking thread and serialize the two tests.
    static NATIVE_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());
    let _guard = NATIVE_LOCK
        .lock()
        .unwrap_or_else(|error| error.into_inner());
    use calcvba_engine::{Engine, HostConfig, RuntimeProfileId, SnapshotOutcome};
    use calcvba_symbol::manifest as sym;
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
    };
    let config = match executor {
        Executor::Vm3 => HostConfig::vm3(),
        Executor::Jit => HostConfig::jit(),
    };
    let profile = if modules.iter().any(|(name, _)| *name == "P_Shell") {
        RuntimeProfileId::default_for_hal_profile(calcvba_hal::model::native_host_profile())
    } else {
        RuntimeProfileId::WindowsHeadless
    };
    let mut engine = Engine::new(config).with_runtime_profile(profile);
    engine.set_host_policy(calcvba_hal::model::HostPolicy::native());
    let outcome = match executor {
        Executor::Vm3 => engine.execute_manifest_snapshot_with_err_vm3(&manifest),
        Executor::Jit => engine.execute_manifest_snapshot_with_err_jit(&manifest),
    };
    match outcome {
        SnapshotOutcome::Completed { values, .. } => values
            .iter()
            .map(calcvba_differential::canon)
            .find_map(|v| match v {
                Canon::Str(s) if s.contains('=') => Some(s),
                _ => None,
            })
            .ok_or_else(|| "no result".to_string()),
        SnapshotOutcome::Raised { err } => Err(format!("VBA error {}", err.number)),
        SnapshotOutcome::Unsupported(w) => Err(format!("declined: {w}")),
        SnapshotOutcome::Failed(m) => Err(format!("failed: {m}")),
    }
}

/// Run `module` (with ProbeFmt and a Main that stores its `Run()` result) and return the text.
fn run_text(executor: Executor, fmt: &str, module: &str, source: &str) -> Result<String, String> {
    if WINDOWS_ONLY_PREFIX.contains(&module) || module == "P_Shell" {
        let main =
            format!("Public result As String\nSub Main()\n    result = {module}.Run()\nEnd Sub\n");
        let modules = [("Main", main.as_str()), ("ProbeFmt", fmt), (module, source)];
        return std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            run_native(executor, &modules)
        }))
        .unwrap_or_else(|_| Err("panic".to_string()));
    }
    let main =
        format!("Public result As String\nSub Main()\n    result = {module}.Run()\nEnd Sub\n");
    let modules = [
        ("Main", Procedural, main.as_str()),
        ("ProbeFmt", Procedural, fmt),
        (module, Procedural, source),
    ];
    // A panic inside CalcVBA is a defect like any other answer: record it for this module.
    let outcome = match std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        run_modules(executor, &modules, "VBAProject")
    })) {
        Ok(outcome) => outcome,
        Err(panic) => {
            let msg = panic
                .downcast_ref::<String>()
                .cloned()
                .or_else(|| panic.downcast_ref::<&str>().map(|s| s.to_string()))
                .unwrap_or_default();
            return Err(format!("panic: {msg}"));
        }
    };
    match (&outcome.unsupported, &outcome.result) {
        (Some(why), _) => Err(format!("declined: {why:?}")),
        (None, Err(e)) => Err(format!("failed: {e}")),
        (None, Ok(values)) => values
            .iter()
            .find_map(|v| match v {
                Canon::Str(s) if s.contains('=') => Some(s.clone()),
                _ => None,
            })
            .ok_or_else(|| format!("no result in {values:?}")),
    }
}

/// `source` with every case except `keep` (of `cases`) blanked, so a case CalcVBA cannot bind
/// only fails its own line.
fn isolate(source: &str, keep: u32, cases: &[u32]) -> String {
    let drop: Vec<u32> = cases.iter().copied().filter(|&j| j != keep).collect();
    // Blanking the Select Case arms still runs setup and cleanup for every
    // other case. The file probes do that work outside Select Case, so run
    // only the selected iteration of the standard probe driver.
    blank(source, &drop).replace("For k = 1 To N", &format!("For k = {keep} To {keep}"))
}

/// The case number of a probe answer line `Func.k=value`.
fn case_number(line: &str) -> Option<u32> {
    let key = line.split('=').next()?;
    key.rsplit('.').next()?.parse().ok()
}

fn compare(executor: Executor) -> (Vec<String>, Vec<String>, Vec<String>) {
    let known = known(executor);
    let fmt = std::fs::read_to_string(dir().join("probes/ProbeFmt.bas")).expect("ProbeFmt.bas");
    let (mut wrong, mut fixed, mut report) = (Vec::new(), Vec::new(), Vec::new());
    for (module, source, answer) in cases() {
        if !cfg!(windows) && WINDOWS_ONLY_PREFIX.contains(&module.as_str()) {
            continue;
        }
        let heavy: Vec<u32> = HEAVY
            .iter()
            .filter(|(m, _)| *m == module)
            .map(|(_, k)| *k)
            .collect();
        let source = if heavy.is_empty() {
            source
        } else {
            blank(&source, &heavy)
        };
        let numbers: Vec<u32> = answer
            .lines()
            .enumerate()
            .map(|(i, l)| case_number(l).unwrap_or(i as u32 + 1))
            .collect();
        eprintln!("{executor:?} {module}: starting");
        let started = std::time::Instant::now();
        // A module that does not bind or run is run once per case (the others blanked), so the
        // failure counts against the cases that cause it.
        let whole = run_text(executor, &fmt, &module, &source);
        let lines: Vec<String> = match whole {
            Ok(text) => text.lines().map(str::to_string).collect(),
            Err(_) => answer
                .lines()
                .enumerate()
                .map(|(i, want)| {
                    let k = case_number(want).unwrap_or(i as u32 + 1);
                    match run_text(executor, &fmt, &module, &isolate(&source, k, &numbers)) {
                        Ok(text) => text
                            .lines()
                            .find(|line| case_number(line) == Some(k))
                            .unwrap_or("<missing>")
                            .to_string(),
                        Err(why) => {
                            let why: String = why.chars().take(160).collect();
                            format!("<{why}>")
                        }
                    }
                })
                .collect(),
        };
        let elapsed = started.elapsed();
        if elapsed.as_secs() >= 5 {
            eprintln!("{executor:?} {module}: {elapsed:?}");
        }
        for (i, want) in answer.lines().enumerate() {
            if heavy.contains(&numbers[i]) {
                continue;
            }
            let entry = format!("{module}: {want}");
            let have = lines.get(i).map(String::as_str).unwrap_or("<missing>");
            if want != have {
                report.push(format!("{module}\t{want}\t{have}"));
            }
            match (want == have, known.contains(&entry)) {
                (false, false) => wrong.push(format!("{entry} -- {executor:?} {have:?}")),
                (true, true) => fixed.push(entry),
                _ => {}
            }
        }
    }
    (wrong, fixed, report)
}

fn check(executor: Executor) {
    // The recorded Excel answers use the South African regional profile.
    use calcvba_runtime::regional::{RegionalSettings, with_settings};
    let mut profile = RegionalSettings::en_za();
    // DateSerial.9 and .10 put years 29 and 30 in the 2000s. The capture
    // establishes 2030 as the earliest cutoff consistent with these rows.
    profile.two_digit_year_max = 2030;
    let (wrong, fixed, report) =
        if std::env::var("VBA_LIBRARY_USE_HOST_LOCALE").as_deref() == Ok("1") {
            compare(executor)
        } else {
            with_settings(profile, || compare(executor))
        };
    if let Some(path) = std::env::var_os("VBA_LIBRARY_REPORT") {
        let mut path = PathBuf::from(path);
        path.set_extension(format!("{executor:?}.tsv").to_lowercase());
        std::fs::write(&path, report.join("\n") + "\n").expect("write report");
    }
    assert!(
        wrong.is_empty() && fixed.is_empty(),
        "{executor:?} against Excel ({} wrong, {} fixed but still in known-differences.txt):\n{}\n\
         fixed:\n{}",
        wrong.len(),
        fixed.len(),
        wrong.join("\n"),
        fixed.join("\n")
    );
}

/// Lines Excel's VBE refuses to compile (`compile-refusals.tsv`; the probe keeps the case as
/// `v = "compile-refused"`): with the original line restored, CalcVBA must refuse the module too.
/// Returns the refusals CalcVBA accepts, minus those listed in known-differences.txt as
/// `refusal: <module> <source>`.
fn accepted_refusals() -> (Vec<String>, Vec<String>) {
    let known = known(Executor::Vm3);
    let Ok(log) = std::fs::read_to_string(dir().join("compile-refusals.tsv")) else {
        return (Vec::new(), Vec::new());
    };
    let fmt = std::fs::read_to_string(dir().join("probes/ProbeFmt.bas")).expect("ProbeFmt.bas");
    let (mut wrong, mut fixed) = (Vec::new(), Vec::new());
    let (mut refused_count, mut accepted_count) = (0, 0);
    for row in log.lines().skip(1).filter(|l| !l.trim().is_empty()) {
        let mut cols = row.splitn(3, '\t');
        let (module, message, original) = (
            cols.next().unwrap_or(""),
            cols.next().unwrap_or(""),
            cols.next().unwrap_or(""),
        );
        let Some(prefix) = original.split_once(':').map(|(case, _)| case.trim()) else {
            continue;
        };
        let Ok(source) = std::fs::read_to_string(dir().join(format!("probes/{module}.bas"))) else {
            continue;
        };
        let keep: u32 = prefix
            .strip_prefix("Case ")
            .expect("numbered refusal case")
            .parse()
            .expect("case number");
        let answer = std::fs::read_to_string(dir().join(format!("probes/{module}.excel.txt")))
            .expect("refusal probe answers");
        let numbers: Vec<_> = answer.lines().filter_map(case_number).collect();
        let source = isolate(&source, keep, &numbers);
        let restored: Vec<String> = source
            .lines()
            .map(|l| {
                let t = l.trim();
                if t.starts_with(&format!("{prefix}:")) && t.contains("\"compile-refused\"") {
                    format!("        {original}")
                } else {
                    l.to_string()
                }
            })
            .collect();
        let restored = restored.join("\n");
        let main =
            format!("Public result As String\nSub Main()\n    result = {module}.Run()\nEnd Sub\n");
        let modules = [
            ("Main", Procedural, main.as_str()),
            ("ProbeFmt", Procedural, fmt.as_str()),
            (module, Procedural, restored.as_str()),
        ];
        // Check this restored line's compilation, independently of unrelated
        // cases and runtime errors. Application execution cannot prove refusal.
        use calcvba_symbol::manifest as sym;
        let manifest = sym::SymbolProjectManifest {
            project_name: "VBAProject".into(),
            project_kind: sym::ProjectKind::Source,
            modules: modules
                .iter()
                .map(|(name, kind, source)| sym::ModuleUnit {
                    module_name: (*name).into(),
                    module_kind: *kind,
                    attributes: sym::ModuleAttributes::named(*name),
                    source: (*source).into(),
                })
                .collect(),
            references: Vec::new(),
            reference_projects: Vec::new(),
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
        };
        let refused =
            calcvba_bind::bind_program(&manifest, &calcvba_symbol::CatalogTypeLibResolver).is_err();
        if refused {
            refused_count += 1;
        } else {
            accepted_count += 1;
        }
        let entry = format!("refusal: {module} {original}");
        match (refused, known.contains(&entry)) {
            (false, false) => wrong.push(format!("{entry} -- Excel: {message}")),
            (true, true) => fixed.push(entry),
            _ => {}
        }
    }
    eprintln!("Excel compile refusals: {refused_count} refused, {accepted_count} accepted");
    (wrong, fixed)
}

#[test]
fn excel_compile_refusals_are_refused() {
    let (wrong, fixed) = accepted_refusals();
    assert!(
        wrong.is_empty() && fixed.is_empty(),
        "CalcVBA accepts {} lines Excel refuses ({} listed but now refused):\n{}\nfixed:\n{}",
        wrong.len(),
        fixed.len(),
        wrong.join("\n"),
        fixed.join("\n")
    );
}

/// Keep the 90 measured class cases in ordinary workspace regression runs.
#[test]
fn regexp_classes_match_excel() {
    let fmt = std::fs::read_to_string(dir().join("probes/ProbeFmt.bas")).unwrap();
    for executor in [Executor::Vm3, Executor::Jit] {
        for module in ["P_RegExp", "P_Match", "P_MatchCollection"] {
            let source =
                std::fs::read_to_string(dir().join(format!("probes/{module}.bas"))).unwrap();
            let answer =
                std::fs::read_to_string(dir().join(format!("probes/{module}.excel.txt"))).unwrap();
            let actual = run_text(executor, &fmt, module, &source)
                .unwrap_or_else(|e| panic!("{executor:?} {module}: {e}"));
            assert_eq!(
                actual.replace("\r\n", "\n"),
                answer.replace("\r\n", "\n"),
                "{executor:?} {module}"
            );
        }
    }
}

/// Keep the measured portable names and empty UserForms collection in workspace checks.
#[test]
fn missing_library_names_match_excel() {
    let fmt = std::fs::read_to_string(dir().join("probes/ProbeFmt.bas")).unwrap();
    for executor in [Executor::Vm3, Executor::Jit] {
        for module in [
            "P_Calendar",
            "P_IMEStatus",
            "P_MacID",
            "P_MacScript",
            "P_UserForms",
        ] {
            let source =
                std::fs::read_to_string(dir().join(format!("probes/{module}.bas"))).unwrap();
            let answer =
                std::fs::read_to_string(dir().join(format!("probes/{module}.excel.txt"))).unwrap();
            let actual = run_text(executor, &fmt, module, &source)
                .unwrap_or_else(|e| panic!("{executor:?} {module}: {e}"));
            assert_eq!(
                actual.replace("\r\n", "\n"),
                answer.replace("\r\n", "\n"),
                "{executor:?} {module}"
            );
        }
    }
}

/// The eleven captured uses that require Err to be a first-class object value.
#[test]
fn err_object_values_match_excel() {
    let fmt = std::fs::read_to_string(dir().join("probes/ProbeFmt.bas")).unwrap();
    for executor in [Executor::Vm3, Executor::Jit] {
        for (module, keep) in [
            ("P_Err", &[12, 13, 15][..]),
            ("P_ErrObject", &[34][..]),
            ("P_TypeName", &[59][..]),
            ("P_IsObject", &[10][..]),
            ("P_VarType", &[49][..]),
            ("P_CallByName", &[29, 30, 31][..]),
            ("P_Ptr", &[24][..]),
        ] {
            let source =
                std::fs::read_to_string(dir().join(format!("probes/{module}.bas"))).unwrap();
            let answer =
                std::fs::read_to_string(dir().join(format!("probes/{module}.excel.txt"))).unwrap();
            let drop: Vec<_> = answer
                .lines()
                .filter_map(case_number)
                .filter(|n| !keep.contains(n))
                .collect();
            let actual = run_text(executor, &fmt, module, &blank(&source, &drop))
                .unwrap_or_else(|e| panic!("{executor:?} {module}: {e}"));
            for (have, want) in actual.lines().zip(answer.lines()) {
                if case_number(want).is_some_and(|n| keep.contains(&n)) {
                    assert_eq!(have, want, "{executor:?} {module}");
                }
            }
            assert_eq!(
                actual.lines().count(),
                answer.lines().count(),
                "{executor:?} {module}"
            );
        }
    }
}

#[test]
fn vm3_matches_excel() {
    check(Executor::Vm3);
}

#[test]
fn jit_matches_excel() {
    check(Executor::Jit);
}
