//! Safe-mode capability policy, one row per gated operation, on VM3 and the JIT.
//!
//! Each program runs under `PolicyHost::safe` over the standard host with a temporary
//! directory. Every row must give the same `Err.Number` and `Err.Description` on both
//! backends, the same audit record, and leave the directory untouched.

use std::cell::RefCell;
use std::path::Path;

use calcvba_differential::{Canon, Executor, RunOutcome, run_with_engine_setup};
use calcvba_hal::model::HostPolicy;
use calcvba_hal::{AuditHandle, AuditRecord, Capability, Decision, PolicyHost};

struct Row {
    label: &'static str,
    capability: Capability,
    /// Module-level declarations.
    decls: &'static str,
    /// The statement run under `On Error Resume Next`; `{d}` is the temp directory.
    stmt: &'static str,
}

const ROWS: &[Row] = &[
    // fs.read
    Row {
        label: "Open For Input",
        capability: Capability::FsRead,
        decls: "",
        stmt: "Open \"{d}/existing.txt\" For Input As #1",
    },
    Row {
        label: "Dir",
        capability: Capability::FsRead,
        decls: "",
        stmt: "x = Dir(\"{d}/existing.txt\")",
    },
    Row {
        label: "FileLen",
        capability: Capability::FsRead,
        decls: "",
        stmt: "x = FileLen(\"{d}/existing.txt\")",
    },
    Row {
        label: "GetAttr",
        capability: Capability::FsRead,
        decls: "",
        stmt: "x = GetAttr(\"{d}/existing.txt\")",
    },
    Row {
        label: "FileDateTime",
        capability: Capability::FsRead,
        decls: "",
        stmt: "x = FileDateTime(\"{d}/existing.txt\")",
    },
    Row {
        label: "CurDir",
        capability: Capability::FsRead,
        decls: "",
        stmt: "x = CurDir",
    },
    // fs.write
    Row {
        label: "Open For Output",
        capability: Capability::FsWrite,
        decls: "",
        stmt: "Open \"{d}/out.txt\" For Output As #1",
    },
    Row {
        label: "Open For Append",
        capability: Capability::FsWrite,
        decls: "",
        stmt: "Open \"{d}/out.txt\" For Append As #1",
    },
    Row {
        label: "Open For Binary",
        capability: Capability::FsWrite,
        decls: "",
        stmt: "Open \"{d}/out.bin\" For Binary As #1",
    },
    Row {
        label: "Open For Random",
        capability: Capability::FsWrite,
        decls: "",
        stmt: "Open \"{d}/out.dat\" For Random As #1 Len = 8",
    },
    Row {
        label: "Kill",
        capability: Capability::FsWrite,
        decls: "",
        stmt: "Kill \"{d}/existing.txt\"",
    },
    Row {
        label: "MkDir",
        capability: Capability::FsWrite,
        decls: "",
        stmt: "MkDir \"{d}/newdir\"",
    },
    Row {
        label: "RmDir",
        capability: Capability::FsWrite,
        decls: "",
        stmt: "RmDir \"{d}\"",
    },
    Row {
        label: "Name",
        capability: Capability::FsWrite,
        decls: "",
        stmt: "Name \"{d}/existing.txt\" As \"{d}/renamed.txt\"",
    },
    Row {
        label: "FileCopy",
        capability: Capability::FsWrite,
        decls: "",
        stmt: "FileCopy \"{d}/existing.txt\", \"{d}/copy.txt\"",
    },
    Row {
        label: "SetAttr",
        capability: Capability::FsWrite,
        decls: "",
        stmt: "SetAttr \"{d}/existing.txt\", vbReadOnly",
    },
    Row {
        label: "ChDir",
        capability: Capability::FsWrite,
        decls: "",
        stmt: "ChDir \"{d}\"",
    },
    Row {
        label: "ChDrive",
        capability: Capability::FsWrite,
        decls: "",
        stmt: "ChDrive \"C\"",
    },
    // process
    Row {
        label: "Shell",
        capability: Capability::Process,
        decls: "",
        stmt: "x = Shell(\"touch {d}/shell-ran\")",
    },
    // env
    Row {
        label: "Environ",
        capability: Capability::Env,
        decls: "",
        stmt: "x = Environ(\"HOME\")",
    },
    Row {
        label: "Command",
        capability: Capability::Env,
        decls: "",
        stmt: "x = Command",
    },
    // settings
    Row {
        label: "GetSetting",
        capability: Capability::Settings,
        decls: "",
        stmt: "x = GetSetting(\"app\", \"sec\", \"key\", \"dflt\")",
    },
    Row {
        label: "GetAllSettings",
        capability: Capability::Settings,
        decls: "",
        stmt: "x = GetAllSettings(\"app\", \"sec\")",
    },
    Row {
        label: "SaveSetting",
        capability: Capability::Settings,
        decls: "",
        stmt: "SaveSetting \"app\", \"sec\", \"key\", \"v\"",
    },
    Row {
        label: "DeleteSetting",
        capability: Capability::Settings,
        decls: "",
        stmt: "DeleteSetting \"app\", \"sec\", \"key\"",
    },
    // ffi
    Row {
        label: "Declare call",
        capability: Capability::Ffi,
        decls: "Declare PtrSafe Function NativeGetTickCount Lib \"kernel32\" Alias \"GetTickCount\" () As Long\n",
        stmt: "x = NativeGetTickCount()",
    },
    // com.activate
    Row {
        label: "CreateObject",
        capability: Capability::ComActivate,
        decls: "",
        stmt: "Set o = CreateObject(\"WinHttp.WinHttpRequest.5.1\")",
    },
    Row {
        label: "GetObject",
        capability: Capability::ComActivate,
        decls: "",
        stmt: "Set o = GetObject(\"\", \"WinHttp.WinHttpRequest.5.1\")",
    },
    // ui.keys
    Row {
        label: "SendKeys",
        capability: Capability::UiKeys,
        decls: "",
        stmt: "SendKeys \"abc\", False",
    },
    Row {
        label: "AppActivate",
        capability: Capability::UiKeys,
        decls: "",
        stmt: "AppActivate \"no such window\"",
    },
];

fn source(row: &Row, dir: &str) -> String {
    format!(
        "Public n As String\nPublic d As String\n{decls}Sub Main()\n  Dim x As Variant\n  Dim o As Object\n  On Error Resume Next\n  {stmt}\n  n = CStr(Err.Number)\n  d = Err.Description\nEnd Sub\n",
        decls = row.decls,
        stmt = row.stmt.replace("{d}", dir),
    )
}

fn fixture_dir(tag: &str) -> tempdir::TestDir {
    tempdir::TestDir::new(tag)
}

mod tempdir {
    use std::path::{Path, PathBuf};

    pub struct TestDir(PathBuf);

    impl TestDir {
        pub fn new(tag: &str) -> Self {
            let path = std::env::temp_dir().join(format!(
                "calcvba-policy-{}-{}-{}",
                tag,
                std::process::id(),
                std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|d| d.as_nanos())
                    .unwrap_or(0)
            ));
            std::fs::create_dir_all(&path).expect("create fixture dir");
            std::fs::write(path.join("existing.txt"), b"hello").expect("fixture file");
            Self(path)
        }
        pub fn path(&self) -> &Path {
            &self.0
        }
        pub fn text(&self) -> String {
            self.0.to_string_lossy().replace('\\', "/")
        }
    }

    impl Drop for TestDir {
        fn drop(&mut self) {
            let _ = std::fs::remove_dir_all(&self.0);
        }
    }
}

fn listing(dir: &Path) -> Vec<(String, Vec<u8>)> {
    let mut entries: Vec<_> = std::fs::read_dir(dir)
        .expect("list fixture dir")
        .map(|entry| {
            let entry = entry.expect("entry");
            (
                entry.file_name().to_string_lossy().into_owned(),
                std::fs::read(entry.path()).unwrap_or_default(),
            )
        })
        .collect();
    entries.sort();
    entries
}

/// n and d are the first two globals of the snapshot.
fn err_pair(outcome: &RunOutcome) -> (i64, String) {
    let values = outcome.result.as_ref().expect("run completed");
    let n = match &values[0] {
        Canon::Str(n) => n.parse::<i64>().expect("n parses"),
        other => panic!("n is not a String: {other:?}"),
    };
    let d = match &values[1] {
        Canon::Str(s) => s.clone(),
        other => panic!("d is not a String: {other:?}"),
    };
    (n, d)
}

/// The HAL profile whose adapter really touches this machine (the engine default is the
/// Windows one, which only runs deterministically off Windows).
fn native_profile() -> calcvba_hal::HalProfileId {
    if cfg!(target_os = "windows") {
        calcvba_hal::HalProfileId::Windows
    } else if cfg!(target_os = "macos") {
        calcvba_hal::HalProfileId::MacOs
    } else {
        calcvba_hal::HalProfileId::Linux
    }
}

fn run_policy(
    executor: Executor,
    source: &str,
    make: &dyn Fn(std::sync::Arc<dyn calcvba_hal::HostServices>) -> PolicyHost,
    native: bool,
) -> (RunOutcome, Vec<AuditRecord>) {
    let handle: RefCell<Option<AuditHandle>> = RefCell::new(None);
    let outcome = run_with_engine_setup(executor, source, &|engine| {
        if native {
            engine.set_hal_profile(native_profile());
            engine.set_host_policy(HostPolicy::native());
        }
        engine.wrap_host_services(|inner| {
            let policy = make(inner);
            *handle.borrow_mut() = Some(policy.audit());
            policy.into_host()
        });
    });
    let records = handle
        .borrow()
        .as_ref()
        .expect("policy installed")
        .drain()
        .records;
    (outcome, records)
}

fn comparable(
    records: &[AuditRecord],
) -> Vec<(Capability, &'static str, String, Decision, Option<i32>)> {
    records
        .iter()
        .map(|r| {
            (
                r.capability,
                r.op,
                r.detail.clone(),
                r.decision,
                r.err_number,
            )
        })
        .collect()
}

#[test]
fn safe_mode_denies_every_gated_operation_the_same_on_both_backends() {
    let mut table = Vec::new();
    for row in ROWS {
        let dir = fixture_dir("safe");
        let before = listing(dir.path());
        let program = source(row, &dir.text());
        let (expected_number, expected_description) = row.capability.safe_error();
        let mut audits = Vec::new();
        for executor in [Executor::Vm3, Executor::Jit] {
            let (outcome, records) = run_policy(executor, &program, &PolicyHost::safe, false);
            assert!(
                outcome.unsupported.is_none(),
                "{} on {executor:?} unsupported: {:?}",
                row.label,
                outcome.unsupported
            );
            assert!(
                !outcome.raised,
                "{} on {executor:?} raised: {:?}",
                row.label, outcome.result
            );
            let (n, d) = err_pair(&outcome);
            assert_eq!(
                n,
                i64::from(expected_number),
                "{} on {executor:?}: Err.Number",
                row.label
            );
            assert_eq!(
                d, expected_description,
                "{} on {executor:?}: Err.Description",
                row.label
            );
            assert_eq!(
                records.len(),
                1,
                "{} on {executor:?}: records {records:?}",
                row.label
            );
            let record = &records[0];
            assert_eq!(
                record.capability, row.capability,
                "{} on {executor:?}",
                row.label
            );
            assert_eq!(record.decision, Decision::Denied);
            assert_eq!(record.err_number, Some(expected_number));
            audits.push(comparable(&records));
            assert_eq!(
                listing(dir.path()),
                before,
                "{} on {executor:?} changed the fixture directory",
                row.label
            );
        }
        assert_eq!(
            audits[0], audits[1],
            "{}: VM3 and JIT audit differ",
            row.label
        );
        assert!(!dir.path().join("shell-ran").exists());
        table.push(format!(
            "{:<16} {:<13} err {:>3}  {} (VM3 = JIT)",
            row.label,
            row.capability.name(),
            expected_number,
            audits[0][0].2.replace(&dir.text(), "<dir>"),
        ));
    }
    eprintln!(
        "operation        capability    result\n{}",
        table.join("\n")
    );
}

#[test]
fn a_denial_without_on_error_raises_the_policy_error_on_both_backends() {
    let program = "Public r As Long\nSub Main()\n  r = Shell(\"echo hi\")\nEnd Sub\n";
    for executor in [Executor::Vm3, Executor::Jit] {
        let (outcome, records) = run_policy(executor, program, &PolicyHost::safe, false);
        assert!(outcome.raised, "{executor:?}: {:?}", outcome.result);
        assert_eq!(outcome.err.number, 70);
        assert_eq!(outcome.err.description, "Permission denied");
        assert_eq!(records.len(), 1);
    }
}

#[test]
fn a_denial_carries_the_stable_code() {
    use calcvba_hal::traits::HostServices as _;
    use calcvba_hal::{HalErrorKind, POLICY_DENIED_CODE};
    let inner = calcvba_hal::adapters::builder::HostBuilder::new().build();
    let host = PolicyHost::safe(inner);
    let err = host
        .process()
        .environ_variant(calcvba_runtime::Variant::from_string("HOME"))
        .expect_err("denied");
    assert_eq!(err.kind, HalErrorKind::PolicyDenied);
    assert_eq!(err.stable_code, POLICY_DENIED_CODE);
    assert_eq!(err.host_error_code, Some(70));
    assert_eq!(
        err.host_error_info.expect("info").description.as_deref(),
        Some("Permission denied")
    );
}

#[test]
fn full_mode_logs_and_passes_through_on_both_backends() {
    let dir = fixture_dir("full");
    let program = format!(
        "Public n As String\nPublic d As String\nSub Main()\n  Dim x As Variant\n  On Error Resume Next\n  Open \"{d}/written.txt\" For Output As #1\n  Print #1, \"hi\"\n  Close #1\n  x = FileLen(\"{d}/written.txt\")\n  x = Environ(\"HOME\")\n  n = CStr(Err.Number)\n  d = Err.Description\nEnd Sub\n",
        d = dir.text()
    );
    let mut audits = Vec::new();
    for executor in [Executor::Vm3, Executor::Jit] {
        let _ = std::fs::remove_file(dir.path().join("written.txt"));
        let (outcome, records) = run_policy(executor, &program, &PolicyHost::full, true);
        let (n, d) = err_pair(&outcome);
        assert_eq!((n, d.as_str()), (0, ""), "{executor:?}");
        assert!(
            dir.path().join("written.txt").exists(),
            "{executor:?}: no file written"
        );
        assert!(records.iter().all(|r| r.decision == Decision::Allowed));
        let caps: Vec<_> = records.iter().map(|r| r.capability).collect();
        assert_eq!(
            caps,
            [Capability::FsWrite, Capability::FsRead, Capability::Env],
            "{executor:?}: {records:?}"
        );
        audits.push(comparable(&records));
    }
    assert_eq!(audits[0], audits[1]);
}

#[test]
fn pure_code_makes_no_policy_request() {
    let program = "Public n As String\nPublic d As String\nSub Main()\n  n = Len(UCase$(\"abc\")) + Abs(-3)\nEnd Sub\n";
    for executor in [Executor::Vm3, Executor::Jit] {
        let (outcome, records) = run_policy(executor, program, &PolicyHost::safe, false);
        assert!(!outcome.raised);
        assert!(records.is_empty(), "{executor:?}: {records:?}");
    }
}

const FIXTURE: &str = include_str!("fixtures/capability_triage/Fixture.bas");

/// Triage predicts the fixture's capabilities from bound OxIR; Safe runs on both backends must
/// observe every one of them (recall 100%), including the denials that `On Error Resume Next`
/// swallowed, and must not observe the function that is never called.
#[test]
fn triage_prediction_covers_every_observed_capability_on_both_backends() {
    let predicted = calcvba_engine::triage::predict_source(FIXTURE);
    assert_eq!(predicted.fallback_reason, None, "the fixture must bind");
    let p = &predicted.prediction;
    let predicted_set: Vec<&str> = p.capabilities.iter().map(|c| c.name()).collect();
    assert_eq!(
        predicted_set,
        [
            "fs.read",
            "fs.write",
            "process",
            "env",
            "settings",
            "ffi",
            "com.activate",
            "ui.keys"
        ]
    );
    let predicted_com = &p.details[&Capability::ComActivate];
    assert!(predicted_com.contains("Uncalled.Class"));

    for executor in [Executor::Vm3, Executor::Jit] {
        let handle: RefCell<Option<AuditHandle>> = RefCell::new(None);
        let outcome = run_with_engine_setup(executor, FIXTURE, &|engine| {
            engine.wrap_host_services(|inner| {
                let policy = PolicyHost::safe(inner);
                *handle.borrow_mut() = Some(policy.audit());
                policy.into_host()
            });
        });
        assert!(
            outcome.unsupported.is_none(),
            "{executor:?}: {:?}",
            outcome.unsupported
        );
        assert!(!outcome.raised, "{executor:?}: {:?}", outcome.result);
        let snapshot = handle.borrow().as_ref().expect("installed").drain();
        let observed = snapshot.requested();
        // Recall on exercised code: observed minus predicted is empty, and all eight are seen.
        let predicted_caps =
            calcvba_hal::CapabilitySet::of(&p.capabilities.iter().copied().collect::<Vec<_>>());
        assert!(
            observed.difference(predicted_caps).is_empty(),
            "{executor:?}: triage miss"
        );
        assert_eq!(
            observed,
            calcvba_hal::CapabilitySet::ALL,
            "{executor:?}: {observed:?}"
        );
        // Every call ran twice (plain and with Resume Next); every request was denied.
        assert_eq!(snapshot.denied, snapshot.total, "{executor:?}");
        assert_eq!(snapshot.total, 16, "{executor:?}: {:?}", snapshot.records);
        for count in &snapshot.counts {
            assert_eq!(count.denied, 2, "{executor:?}: {count:?}");
        }
        // The function Main never calls: predicted, never observed.
        let observed_com: Vec<_> = snapshot
            .counts
            .iter()
            .filter(|c| c.capability == Capability::ComActivate)
            .map(|c| c.detail.as_str())
            .collect();
        assert_eq!(observed_com, ["WinHttp.WinHttpRequest.5.1"], "{executor:?}");
        let unconfirmed: Vec<_> = predicted_com
            .iter()
            .filter(|d| !observed_com.contains(&d.as_str()))
            .collect();
        assert_eq!(unconfirmed, ["Uncalled.Class"]);
    }
}

/// Scripting.Dictionary is served by the engine itself: it works in Safe mode, asks the host
/// for nothing, and is not predicted as a COM activation.
#[test]
fn scripting_dictionary_needs_no_capability_in_safe_mode() {
    let program = "Public r As Variant\nSub Main()\n  Dim d As Object\n  Set d = CreateObject(\"Scripting.Dictionary\")\n  d.Add \"a\", 4\n  r = d(\"a\")\nEnd Sub\n";
    for executor in [Executor::Vm3, Executor::Jit] {
        let (outcome, records) = run_policy(executor, program, &PolicyHost::safe, false);
        assert!(
            outcome.unsupported.is_none() && !outcome.raised,
            "{executor:?}: {:?}",
            outcome.result
        );
        assert_eq!(
            outcome.result.as_ref().unwrap()[0],
            calcvba_differential::canon(&calcvba_runtime::Variant::from_i16(4)),
            "{executor:?}"
        );
        assert!(records.is_empty(), "{executor:?}: {records:?}");
    }
}
