//! VBA performance benchmark against Excel VBA (dc-tey4.7): the workloads in
//! `conformance/perf/vba-bench/B_Work.bas` run under VM3 and the JIT with the sizes Excel used
//! (`BenchExcel.bas`). Each run compiles and executes a fresh project; the same project with
//! size 0 gives the compile-and-setup time, so run time = full - empty. Three runs, median.
//! A measurement, not a check: `cargo test --release -p calcvba-differential --test vba_bench
//! -- --ignored --nocapture --test-threads=1`.

use std::time::Instant;

use calcvba_differential::{Executor, run_modules};
use calcvba_symbol::manifest::ModuleKind;

const WORK: &str = include_str!("../../../conformance/perf/vba-bench/B_Work.bas");
const ACC: &str = include_str!("../../../conformance/perf/vba-bench/B_Acc.cls");

/// Must match `RunAll` in BenchExcel.bas.
const SIZES: &[(&str, u32)] = &[
    ("LoopLong", 5_000_000),
    ("LoopDouble", 3_000_000),
    ("StrBuild", 50_000),
    ("StrFuncs", 200_000),
    ("ArraySort", 200_000),
    ("Coll", 100_000),
    ("Dict", 100_000),
    ("ClassCalls", 1_000_000),
    ("VariantMix", 1_000_000),
    ("Amort", 20_000),
    ("UdtArray", 1_000_000),
];
const REPS: usize = 3;

fn once(executor: Executor, name: &str, n: u32) -> (f64, String) {
    let main =
        format!("Public result As Double\nSub Main()\n    result = B_Work.{name}({n})\nEnd Sub\n");
    let modules = [
        ("Main", ModuleKind::Procedural, main.as_str()),
        ("B_Acc", ModuleKind::Class, ACC),
        ("B_Work", ModuleKind::Procedural, WORK),
    ];
    let t = Instant::now();
    let out = run_modules(executor, &modules, "VBAProject");
    let ms = t.elapsed().as_secs_f64() * 1000.0;
    let value = match (&out.unsupported, &out.result) {
        (Some(u), _) => format!("unsupported: {u}"),
        (None, Ok(v)) => format!("{:?}", v.first()),
        (None, Err(e)) => format!("error: {e}"),
    };
    (ms, value)
}

fn median(mut v: Vec<f64>) -> f64 {
    v.sort_by(f64::total_cmp);
    v[v.len() / 2]
}

#[test]
#[ignore = "measurement; run on request"]
fn vba_bench() {
    let filter = std::env::var("VBA_BENCH_ONLY").unwrap_or_default();
    println!("executor\tworkload\tcompile_ms\trun_ms\tvalue");
    for executor in [Executor::Vm3, Executor::Jit] {
        for &(name, n) in SIZES {
            if !filter.is_empty() && !filter.split(',').any(|f| f == name) {
                continue;
            }
            let empty: Vec<f64> = (0..REPS).map(|_| once(executor, name, 0).0).collect();
            let mut full = Vec::new();
            let mut value = String::new();
            for _ in 0..REPS {
                let (ms, v) = once(executor, name, n);
                full.push(ms);
                value = v;
            }
            let compile = median(empty);
            let run = (median(full) - compile).max(0.0);
            println!("{executor:?}\t{name}\t{compile:.1}\t{run:.1}\t{value}");
        }
    }
}
