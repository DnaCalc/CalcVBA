use calcvba_spreadsheet_compat::{Backend, Module, Value};

fn main() -> Result<(), calcvba_spreadsheet_compat::Error> {
    let module = Module::load("Math", "Function Add2(a, b): Add2 = a + b: End Function\n")?;
    for function in module.functions() {
        println!("{}: {:?}", function.qualified_name, function.parameters);
    }
    let session = module.prepare(Backend::Interpreter)?;
    println!(
        "{:?}",
        session.call("Add2", &[Value::Number(19.0), Value::Number(23.0)], 1000)?
    );
    Ok(())
}
