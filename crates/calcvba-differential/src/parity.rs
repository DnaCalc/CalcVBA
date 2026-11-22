//! Compare the entry of a reviewed OxIR project on the two execution backends.

use std::{
    collections::BTreeMap,
    fs, io,
    path::{Path, PathBuf},
    sync::{
        Arc, Mutex,
        atomic::{AtomicU64, Ordering},
    },
};

use calcvba_com::{
    PortableComProjection, platform::portable::PortableDispatch,
    platform::portable::PortableObjectFactory,
};
use calcvba_hal::{
    Capability, CapabilitySet, HostOutputChannel, HostPolicy, HostServices, PolicyHost,
    adapters::builder::HostBuilder, callbacks::HostCallbacks, install_thread_output_tap,
};
use calcvba_oxir::OxProgram;
use calcvba_runtime::{Variant, variant::VarType};

use crate::{Canon, Executor, RunOutcome, canon, run_ox_programs_on_host, vm3_oracle_engine};

const STEPS: u64 = 10_000_000;

fn exact_value(value: &Variant) -> Canon {
    fn project(value: &Variant, depth: usize) -> Canon {
        let tag = value.core().raw_vt();
        let opaque = Canon::Opaque { tag };
        if depth == 64 {
            return opaque;
        }
        let nested = |value: Variant| project(&value, depth + 1);
        match value.vtype() {
            VarType::String => Canon::StringBytes(value.string_bytes().unwrap_or_default()),
            VarType::ArrayVariant => {
                let (bounds, length) = if let Some((bounds, length)) = value.safearray_bounds_len()
                {
                    (
                        Some(bounds.into_iter().map(|b| (b.lower, b.count)).collect()),
                        length,
                    )
                } else if value.raw_safearray_ptr().is_null() {
                    (None, 0)
                } else {
                    return opaque;
                };
                let elements = (0..length)
                    .map(|i| value.safearray_element(i).and_then(Result::ok).map(nested))
                    .collect::<Option<Vec<_>>>();
                elements.map_or(opaque, |elements| Canon::Array {
                    tag,
                    fixed: value.is_fixed_size_array(),
                    iid: value.array_element_iid(),
                    record_type: value
                        .record_array_layout()
                        .and_then(|l| l.identity().map(|id| format!("{} {:?}", id.name, id.guid))),
                    bounds,
                    elements,
                })
            }
            VarType::Record => {
                let Some(layout) = value.vba_record_layout() else {
                    return opaque;
                };
                let fields = (0..layout.fields().len())
                    .map(|i| value.read_record_field_variant(i).map(nested))
                    .collect::<Result<Vec<_>, _>>();
                fields.map_or(opaque, |fields| Canon::Record {
                    identity: layout
                        .identity()
                        .map(|id| format!("{} {:?}", id.name, id.guid)),
                    fields,
                })
            }
            VarType::Object if value.as_object_ref().is_none() => Canon::Nothing,
            VarType::Object | VarType::ProcRef | VarType::ByRef => opaque,
            VarType::Single | VarType::Double | VarType::Date | VarType::Boolean => Canon::Raw {
                tag,
                bytes: value.value_bytes(),
                reserved: value.value_reserved(),
            },
            _ => canon(value),
        }
    }
    project(value, 0)
}

fn opaque(value: &Canon) -> bool {
    match value {
        Canon::Opaque { .. } => true,
        Canon::Array { elements, .. } => elements.iter().any(opaque),
        Canon::Record { fields, .. } => fields.iter().any(opaque),
        _ => false,
    }
}

/// Final scratch tree, including empty directories and exact file bytes.
pub type Files = BTreeMap<PathBuf, Option<Vec<u8>>>;

struct Scratch(PathBuf);

impl Scratch {
    fn new() -> io::Result<Self> {
        static NEXT: AtomicU64 = AtomicU64::new(0);
        loop {
            let dir = std::env::temp_dir().join(format!(
                "calcvba-parity-{}-{}",
                std::process::id(),
                NEXT.fetch_add(1, Ordering::Relaxed)
            ));
            match fs::create_dir(&dir) {
                Ok(()) => return Ok(Self(dir)),
                Err(err) if err.kind() == io::ErrorKind::AlreadyExists => continue,
                Err(err) => return Err(err),
            }
        }
    }

    fn snapshot(&self) -> io::Result<Files> {
        fn walk(root: &Path, dir: &Path, files: &mut Files) -> io::Result<()> {
            for entry in fs::read_dir(root.join(dir))? {
                let entry = entry?;
                let kind = entry.file_type()?;
                let path = dir.join(entry.file_name());
                if kind.is_dir() {
                    walk(root, &path, files)?;
                    files.insert(path, None);
                } else if kind.is_file() {
                    files.insert(path, Some(fs::read(entry.path())?));
                } else {
                    return Err(io::Error::other("unexpected scratch entry type"));
                }
            }
            Ok(())
        }
        let mut files = BTreeMap::new();
        walk(&self.0, Path::new(""), &mut files)?;
        Ok(files)
    }
}

impl Drop for Scratch {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[derive(Debug)]
pub struct Observation {
    pub outcome: RunOutcome,
    pub output: Vec<(HostOutputChannel, String)>,
    pub files: Files,
    pub cells: Vec<CellWrite>,
    pub denied_host_calls: Vec<Capability>,
}

#[derive(Debug, PartialEq, Eq)]
pub struct CellWrite {
    pub row: i32,
    pub column: i32,
    pub value: Canon,
}

#[derive(Clone, Default)]
struct CellRecorder(Arc<Mutex<Vec<CellWrite>>>);

impl PortableObjectFactory for CellRecorder {
    fn create(&self) -> Box<dyn PortableDispatch> {
        Box::new(self.clone())
    }
}

impl PortableDispatch for CellRecorder {
    fn invoke(&self, member: &str, args: &[Variant]) -> Result<Variant, String> {
        if member.eq_ignore_ascii_case("WriteCell")
            && let [row, column, value] = args
            && let (Some(row), Some(column)) = (row.as_i32(), column.as_i32())
            && row > 0
            && column > 0
        {
            self.0.lock().unwrap().push(CellWrite {
                row,
                column,
                value: exact_value(value),
            });
            return Ok(Variant::empty());
        }
        Err("WriteCell requires positive Long row and column arguments and a value".into())
    }

    fn get(&self, _member: &str) -> Result<Variant, String> {
        Err("the parity sheet exposes only WriteCell".into())
    }

    fn put(&self, _member: &str, _value: Variant) -> Result<(), String> {
        Err("the parity sheet exposes only WriteCell".into())
    }

    fn member_names(&self) -> Vec<String> {
        vec!["WriteCell".into()]
    }
}

enum HostMode<'a> {
    Default,
    Files(&'a Scratch),
    Cells,
    Excel,
}

struct RecordedOutput;

impl HostCallbacks for RecordedOutput {
    fn on_msg_box(&self, _prompt: &str, style: i32) -> i32 {
        style.max(1)
    }

    fn on_input_box(&self, _prompt: &str, default: &str) -> String {
        default.to_owned()
    }

    fn on_status_bar(&self, _text: &str) {}

    fn on_console_print(&self, _text: &str) -> bool {
        true
    }

    fn on_debug_print(&self, _text: &str) {}
}

fn observe(
    executor: Executor,
    programs: &[OxProgram],
    mode: HostMode<'_>,
) -> io::Result<Observation> {
    let mut engine = vm3_oracle_engine();
    let mut policy = HostPolicy {
        allow_process_spawn: false,
        allow_com_activation: false,
        ..HostPolicy::default()
    };
    let callbacks = Arc::new(RecordedOutput);
    let cells = CellRecorder::default();
    let host = match mode {
        HostMode::Files(scratch) => {
            policy.deterministic_mode = false;
            policy.allow_filesystem_mutation = true;
            let inner = HostBuilder::new()
                .profile(calcvba_hal::HalProfileId::Linux)
                .policy(policy.with_windows_drive_root('C', &scratch.0))
                .callbacks(callbacks)
                .build();
            let files = CapabilitySet::of(&[Capability::FsRead, Capability::FsWrite]);
            PolicyHost::with_denied(inner, CapabilitySet::ALL.difference(files))
        }
        HostMode::Default | HostMode::Cells | HostMode::Excel => {
            let mut allowed = CapabilitySet::EMPTY;
            if matches!(mode, HostMode::Cells) {
                let projection = Arc::new(PortableComProjection::new());
                projection.register_object("CalcVBA.ParitySheet", Arc::new(cells.clone()));
                engine.set_portable_com_projection(Some(projection));
                policy.allow_com_activation = true;
                allowed.insert(Capability::ComActivate);
            } else if matches!(mode, HostMode::Excel) {
                let projection = Arc::new(PortableComProjection::new());
                crate::excel_stub::install_excel_stub(&projection);
                engine.set_portable_com_projection(Some(projection));
                policy.allow_com_activation = true;
                allowed.insert(Capability::ComActivate);
            }
            engine.set_host_policy(policy);
            engine.set_host_callbacks(Some(callbacks));
            PolicyHost::with_denied(
                engine.host_services(),
                CapabilitySet::ALL.difference(allowed),
            )
        }
    };
    let _policy = host.enter_com_policy();
    let output = Arc::new(Mutex::new(Vec::new()));
    let observer = output.clone();
    let tap = install_thread_output_tap(Arc::new(move |channel, text: &str| {
        observer.lock().unwrap().push((channel, text.to_owned()));
    }));
    let outcome =
        run_ox_programs_on_host(executor, programs, &host, true, exact_value, Some(STEPS));
    drop(tap);
    let output = std::mem::take(&mut *output.lock().unwrap());
    let denied_host_calls = host.audit().drain().denied_set().iter().collect();
    drop(host);
    let files = match mode {
        HostMode::Files(scratch) => scratch.snapshot()?,
        _ => Files::new(),
    };
    let cells = std::mem::take(&mut *cells.0.lock().unwrap());
    Ok(Observation {
        outcome,
        output,
        files,
        cells,
        denied_host_calls,
    })
}

#[derive(Debug)]
pub struct Comparison {
    pub vm3: Observation,
    pub jit: Observation,
}

impl Comparison {
    pub fn differences(&self) -> Vec<&'static str> {
        let vm3 = &self.vm3.outcome;
        let jit = &self.jit.outcome;
        if [vm3, jit]
            .iter()
            .any(|run| run.unsupported.is_some() || (!run.raised && run.result.is_err()))
        {
            return Vec::new();
        }
        let mut differences = Vec::new();
        if vm3.raised != jit.raised {
            differences.push("error propagation");
        }
        if vm3.err != jit.err {
            differences.push("Err");
        }
        if let (Ok(vm3_values), Ok(jit_values)) = (&vm3.result, &jit.result)
            && vm3_values != jit_values
        {
            differences.push("values");
        }
        if self.vm3.output != self.jit.output {
            differences.push("host output");
        }
        if self.vm3.files != self.jit.files {
            differences.push("scratch files");
        }
        if self.vm3.cells != self.jit.cells {
            differences.push("cell writes");
        }
        if self.vm3.denied_host_calls != self.jit.denied_host_calls {
            differences.push("host capabilities");
        }
        differences
    }

    pub fn has_opaque_values(&self) -> bool {
        [&self.vm3, &self.jit].iter().any(|run| {
            run.outcome
                .result
                .as_ref()
                .is_ok_and(|values| values.iter().any(opaque))
                || run.cells.iter().any(|write| opaque(&write.value))
        })
    }
}

/// Capture entry locals, project globals, Err and host output for a reviewed project.
pub fn compare_entry(programs: &[OxProgram]) -> Comparison {
    Comparison {
        vm3: observe(Executor::Vm3, programs, HostMode::Default).expect("in-memory observation"),
        jit: observe(Executor::Jit, programs, HostMode::Default).expect("in-memory observation"),
    }
}

/// Compare a reviewed fixture writing to the portable CalcVBA.ParitySheet host object.
pub fn compare_entry_with_cells(programs: &[OxProgram]) -> Comparison {
    Comparison {
        vm3: observe(Executor::Vm3, programs, HostMode::Cells).expect("in-memory observation"),
        jit: observe(Executor::Jit, programs, HostMode::Cells).expect("in-memory observation"),
    }
}

/// Compare a reviewed fixture against the portable in-memory Excel host stub.
pub fn compare_entry_with_excel(programs: &[OxProgram]) -> Comparison {
    Comparison {
        vm3: observe(Executor::Vm3, programs, HostMode::Excel).expect("in-memory observation"),
        jit: observe(Executor::Jit, programs, HostMode::Excel).expect("in-memory observation"),
    }
}

/// Compare a reviewed fixture with C: mapped to a fresh scratch folder per engine.
pub fn compare_entry_with_files(programs: &[OxProgram]) -> io::Result<Comparison> {
    if !cfg!(target_os = "linux") {
        return Err(io::Error::new(
            io::ErrorKind::Unsupported,
            "scratch drive mapping requires Linux",
        ));
    }
    Ok(Comparison {
        vm3: observe(Executor::Vm3, programs, HostMode::Files(&Scratch::new()?))?,
        jit: observe(Executor::Jit, programs, HostMode::Files(&Scratch::new()?))?,
    })
}
