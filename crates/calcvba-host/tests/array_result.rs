//! dc-snz.26: a fixed array assigned to a typed dynamic-array function result loads.
use calcvba_host::{Backend, Module};

const SOURCE: &str = "Public Function Grid() As Double()
    Dim a(1 To 2, 1 To 3) As Double
    a(2, 3) = 6
    Grid = a
End Function
Public Function Local() As Variant
    Dim d() As Double
    Dim a(1 To 2) As Double
    d = a
    Local = UBound(d)
End Function
";

#[test]
fn fixed_array_result_of_a_typed_dynamic_array_function_loads_and_prepares() {
    let module = Module::load("Api", SOURCE).expect("module loads");
    let mut backends = vec![Backend::Interpreter];
    if cfg!(feature = "jit") {
        backends.push(Backend::Jit);
    }
    for backend in backends {
        module.prepare(backend).expect("session prepares");
    }
}
