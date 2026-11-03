//! Warm host-call cost. Run with --release; compilation is outside the timed loop.
use calcvba_spreadsheet_compat::{Backend, Module, Value};
use std::{hint::black_box, time::Instant};

fn main() -> Result<(), calcvba_spreadsheet_compat::Error> {
    let module = Module::load(
        "Math",
        "Public Function Add2(ByVal x As Double, ByVal y As Double) As Double\nAdd2 = x + y\nEnd Function\n",
    )?;
    for backend in [Backend::Interpreter, Backend::Jit] {
        let session = module.prepare(backend)?;
        let args = [Value::Number(19.0), Value::Number(23.0)];
        for _ in 0..1000 {
            black_box(session.call("Add2", &args, 1000)?);
        }
        let mut samples = Vec::new();
        for _ in 0..7 {
            let start = Instant::now();
            for _ in 0..100_000 {
                black_box(session.call(black_box("Add2"), black_box(&args), 1000)?);
            }
            samples.push(start.elapsed().as_nanos() as f64 / 100_000.0);
        }
        samples.sort_by(f64::total_cmp);
        println!(
            "{backend:?} dynamic ns/call median={:.2} samples={samples:?}",
            samples[3]
        );
    }
    Ok(())
}
