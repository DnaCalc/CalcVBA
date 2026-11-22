//! Opt-in measurements for the shared Excel/CalcVBA workloads.
//! `conformance/perf/vba-bench/run-calc.py` runs this release test once per case.

use std::sync::Arc;
use std::time::Instant;

use calcvba_com::{
    PortableComProjection,
    platform::portable::{PortableDispatch, PortableObjectFactory},
};
use calcvba_hal::{
    adapters::builder::HostBuilder,
    model::{HostPolicy, native_host_profile},
};
use calcvba_host::{Backend, Module, ProcedureKind};
use calcvba_runtime::Variant;
use calcvba_symbol::{CatalogTypeLibResolver, manifest as sym};

const WORK: &str = include_str!("../../../conformance/perf/vba-bench/B_Work.bas");
const ACC: &str = include_str!("../../../conformance/perf/vba-bench/B_Acc.cls");
const CASES: &str = include_str!("../../../conformance/perf/vba-bench/cases.tsv");
const REPS: usize = 5;

struct WorksheetFunctions;

impl PortableObjectFactory for WorksheetFunctions {
    fn create(&self) -> Box<dyn PortableDispatch> {
        Box::new(Self)
    }
}

impl PortableDispatch for WorksheetFunctions {
    fn invoke(&self, member: &str, args: &[Variant]) -> Result<Variant, String> {
        if member.eq_ignore_ascii_case("Sum")
            && let [a, b] = args
            && let (Some(a), Some(b)) = (a.as_i32(), b.as_i32())
        {
            return Ok(Variant::from_f64(f64::from(a) + f64::from(b)));
        }
        Err("expected WorksheetFunction.Sum with two Longs".into())
    }

    fn get(&self, member: &str) -> Result<Variant, String> {
        Err(format!("unexpected property read {member}"))
    }

    fn put(&self, member: &str, _: Variant) -> Result<(), String> {
        Err(format!("unexpected property write {member}"))
    }

    fn member_names(&self) -> Vec<String> {
        vec!["Sum".into()]
    }
}

#[test]
#[ignore = "performance measurement; run with the benchmark driver"]
fn vba_bench() {
    let workload = std::env::var("VBA_BENCH_ONLY").expect("workload name");
    let backend = match std::env::var("VBA_BENCH_BACKEND").as_deref() {
        Ok("vm3") => Backend::Interpreter,
        Ok("jit") => Backend::Jit,
        _ => panic!("set VBA_BENCH_BACKEND to vm3 or jit"),
    };
    let (default_size, argument) = CASES
        .lines()
        .skip(1)
        .find_map(|line| {
            let mut columns = line.split('\t');
            (columns.next()? == workload)
                .then(|| (columns.next().unwrap(), columns.next().unwrap()))
        })
        .expect("known workload");
    let n: i32 = std::env::var("VBA_BENCH_SIZE")
        .as_deref()
        .unwrap_or(default_size)
        .parse()
        .expect("positive Long workload size");
    assert!(n > 0);
    assert_ne!(
        argument, "sheet",
        "a spreadsheet host must supply this case"
    );

    let projection = Arc::new(PortableComProjection::new());
    projection.register_object("Bench.WorksheetFunction", Arc::new(WorksheetFunctions));
    let host = HostBuilder::new()
        .profile(native_host_profile())
        .policy(HostPolicy::native())
        .portable_objects(projection)
        .build();
    let mut args = vec![Variant::from_i32(n)];
    match argument {
        "worksheet" => args.push(
            host.com()
                .create_object_variant(Variant::from_string("Bench.WorksheetFunction"))
                .expect("worksheet-function host"),
        ),
        "file" => {
            let scratch = std::env::current_dir().unwrap().join("lines.txt");
            assert!(!scratch.exists(), "start in a new scratch directory");
            args.push(Variant::from_string(scratch.to_str().expect("UTF-8 path")));
        }
        "long" => (),
        _ => panic!("unknown argument kind {argument}"),
    }

    let manifest = sym::SymbolProjectManifest {
        project_name: "VBAProject".into(),
        project_kind: sym::ProjectKind::Library,
        modules: [
            ("B_Work", sym::ModuleKind::Procedural, WORK),
            ("B_Acc", sym::ModuleKind::Class, ACC),
        ]
        .into_iter()
        .map(|(name, kind, source)| sym::ModuleUnit {
            module_name: name.into(),
            module_kind: kind,
            attributes: sym::ModuleAttributes::named(name),
            source: source.into(),
        })
        .collect(),
        references: vec![],
        reference_projects: vec![],
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let start = Instant::now();
    let module = Module::load_project(&manifest, &CatalogTypeLibResolver)
        .expect("bind benchmark project")
        .with_host(host);
    let session = module.prepare(backend).expect("prepare benchmark project");
    let procedure = session
        .bind("B_Work", &workload, ProcedureKind::Function)
        .expect("bind workload");
    let setup_ms = start.elapsed().as_secs_f64() * 1000.0;
    let run = || {
        session
            .call_bound(&procedure, &args, 1_000_000_000)
            .expect("run workload")
            .as_f64()
            .expect("Double checksum")
            .to_bits()
    };
    let expected = run();
    let mut samples = [0.0; REPS];
    for ms in &mut samples {
        let before = Instant::now();
        let result = run();
        *ms = before.elapsed().as_secs_f64() * 1000.0;
        assert_eq!(result, expected, "unstable checksum for {workload}");
    }
    let mut sorted = samples;
    sorted.sort_by(f64::total_cmp);
    let median = sorted[REPS / 2];
    print!("BENCH\t{workload}\t{n}\t{expected:016X}\t{median:.6}\t{setup_ms:.6}");
    for ms in samples {
        print!("\t{ms:.6}");
    }
    println!();
}
