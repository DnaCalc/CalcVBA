//! cargo run --release -p calcvba-host --example host_call_cost
//! Reference native leaf excludes session validation/budget; no speed guarantee.
use calcvba_host::provider::Provider;
use calcvba_host_api::{self as api, RuntimeProvider, SessionExt};
use std::{hint::black_box, time::Instant};
#[path = "../tests/support/typed_host.rs"]
mod fixture;

#[inline(never)]
extern "C" fn native_read(offset: i32) -> f64 {
    40.0 + f64::from(offset)
}
fn measure(mut call: impl FnMut() -> f64) -> (f64, Vec<f64>) {
    for _ in 0..1000 {
        assert_eq!(black_box(call()), 42.0);
    }
    let mut samples = Vec::new();
    for _ in 0..7 {
        let start = Instant::now();
        for _ in 0..100_000 {
            assert_eq!(black_box(call()), 42.0);
        }
        samples.push(start.elapsed().as_nanos() as f64 / 100_000.0);
    }
    samples.sort_by(f64::total_cmp);
    (samples[3], samples)
}
fn main() -> Result<(), api::Error> {
    let (median, samples) = measure(|| native_read(black_box(1)) + 1.0);
    println!("plain-native-reference ns/call median={median:.2} samples={samples:?}");
    let object = api::HostObject::new(
        fixture::IID,
        vec![api::HostSlot::double_long(|_, offset| {
            Ok(native_read(offset))
        })],
    );
    for backend in fixture::backends() {
        let session = Provider::default().load(&fixture::project(), backend, api::HOST_API_ABI)?;
        let entry = session.bind_typed::<(api::ObjectArg<'_>, i32), f64>("Api", "ReadRange")?;
        let (median, samples) = measure(|| {
            session
                .call_typed(black_box(&entry), black_box((object.argument(), 1)), 1000)
                .unwrap()
        });
        println!("{backend:?} API-to-VBA-to-host ns/call median={median:.2} samples={samples:?}");
    }
    Ok(())
}
