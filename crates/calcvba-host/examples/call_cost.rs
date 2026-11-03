//! cargo run --release -p calcvba-host --example call_cost
use calcvba_host::{Backend, Module, ProcedureKind};
use calcvba_runtime::Variant;
use std::{hint::black_box, time::Instant};
fn measure(mut call: impl FnMut()) -> (f64, Vec<f64>) {
    for _ in 0..1000 {
        call();
    }
    let mut samples = Vec::new();
    for _ in 0..7 {
        let start = Instant::now();
        for _ in 0..100_000 {
            call();
        }
        samples.push(start.elapsed().as_nanos() as f64 / 100_000.0);
    }
    samples.sort_by(f64::total_cmp);
    (samples[3], samples)
}
fn main() -> Result<(), calcvba_host::Error> {
    let module = Module::load(
        "Math",
        "Public Function Add2(ByVal x As Double, ByVal y As Double) As Double\nAdd2 = x + y\nEnd Function\n",
    )?;
    for backend in [Backend::Interpreter, Backend::Jit] {
        let session = module.prepare(backend)?;
        let dynamic = session.bind("Math", "Add2", ProcedureKind::Function)?;
        let typed = session.bind_typed::<(f64, f64), f64>("Math", "Add2")?;
        let args = [Variant::from_f64(19.0), Variant::from_f64(23.0)];
        let (median, samples) = measure(|| {
            assert_eq!(
                session
                    .call_bound(black_box(&dynamic), black_box(&args), 1000)
                    .unwrap()
                    .as_f64(),
                Some(42.0)
            );
        });
        println!("{backend:?} bound-dynamic ns/call median={median:.2} samples={samples:?}");
        let (median, samples) = measure(|| {
            assert_eq!(
                session
                    .call_typed(black_box(&typed), black_box((19.0, 23.0)), 1000)
                    .unwrap(),
                42.0
            );
        });
        println!("{backend:?} typed ns/call median={median:.2} samples={samples:?}");
    }
    Ok(())
}
