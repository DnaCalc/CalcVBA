use calcvba_differential::{Canon, Executor, run};

#[test]
fn clnglng_keeps_exact_endpoints_and_target_radix_width() {
    let source = r#"
Option Explicit
Public result As String
Sub Main()
    Dim v As Variant, s As String, d As Variant, x As Double
    s = "9223372036854775807"
    d = CDec(s)
    x = -(2 ^ 63)
    result = CStr(CLngLng(s)) & ";" & CStr(CLngLng(d)) & ";" & CStr(CLngLng(x))
    result = result & ";" & CStr(CLngLng("&HFFFFFFFF")) & ";" & CStr(CLngLng("&HFFFFFFFFFFFFFFFF"))
    v = CLngLng("-9223372036854775808")
    result = result & ";" & CStr(v) & ";" & CStr(CLngLng("3.5")) & ";" & CStr(CLngLng("-2.5"))
    On Error Resume Next
    v = 41
    v = CLngLng("9223372036854775808")
    result = result & ";" & v & "," & Err.Number
    Err.Clear
    v = CLngLng(2 ^ 63)
    result = result & ";" & v & "," & Err.Number
End Sub
"#;
    for executor in [Executor::Vm3, Executor::Jit] {
        let outcome = run(executor, source);
        assert!(
            outcome.unsupported.is_none(),
            "{executor:?}: {:?}",
            outcome.unsupported
        );
        let values = outcome
            .result
            .unwrap_or_else(|e| panic!("{executor:?}: {e}"));
        assert!(values.contains(&Canon::Str("9223372036854775807;9223372036854775807;-9223372036854775808;4294967295;-1;-9223372036854775808;4;-2;41,6;41,6".into())), "{executor:?}: {values:?}");
    }
}
