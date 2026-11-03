//! Minimal WebAssembly export proving the interpreter-only embedding path.
//! The product API remains the Rust Module/Session interface.
use calcvba_spreadsheet_compat::{Backend, Error, Module, Value};

#[unsafe(no_mangle)]
pub extern "C" fn calcvba_add2_probe(a: f64, b: f64) -> f64 {
    let result = (|| {
        let module = Module::load("Math", "Function Add2(a, b): Add2 = a + b: End Function\n")?;
        module.prepare(Backend::Interpreter)?.call(
            "Add2",
            &[Value::Number(a), Value::Number(b)],
            1000,
        )
    })();
    match result {
        Ok(Value::Number(value)) => value,
        _ => f64::NAN,
    }
}

#[unsafe(no_mangle)]
pub extern "C" fn calcvba_budget_probe() -> i32 {
    let module = Module::load(
        "Budget",
        "Function Forever()\nOn Error Resume Next\nDo\nLoop\nEnd Function\n",
    )
    .expect("probe source");
    i32::from(matches!(
        module
            .prepare(Backend::Interpreter)
            .expect("interpreter")
            .call("Forever", &[], 10),
        Err(Error::BudgetExceeded)
    ))
}

#[unsafe(no_mangle)]
pub extern "C" fn calcvba_matrix_probe() -> f64 {
    let module = Module::load(
        "Matrix",
        "Function Pick(values)\nPick = values(2, 1)\nEnd Function\n",
    )
    .expect("probe source");
    let values = Value::Array {
        rows: 2,
        columns: 2,
        cells: vec![
            Value::Number(1.0),
            Value::Boolean(true),
            Value::Number(42.0),
            Value::Error(2042),
        ],
    };
    match module
        .prepare(Backend::Interpreter)
        .expect("interpreter")
        .call("Pick", &[values], 100)
    {
        Ok(Value::Number(number)) => number,
        _ => f64::NAN,
    }
}
