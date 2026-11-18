use calcvba_differential::{Canon, Executor, canon, run};
use calcvba_runtime::Variant;

#[test]
fn financial_cashflows_keep_excel_bits_and_errors_on_both_engines() {
    let source = r#"
Option Explicit
Public negativeGuess As Double, namedGuess As Double, modified As Double
Public errors As String
Sub Main()
    Dim d() As Double, v As Variant
    ReDim d(-2 To 2)
    d(-2) = -100: d(-1) = 30: d(0) = 40: d(1) = 50: d(2) = 60
    negativeGuess = IRR(d, -0.5)
    namedGuess = IRR(Guess:=0.2, ValueArray:=d)
    ReDim d(1 To 3)
    d(1) = -100: d(2) = 50: d(3) = 60
    modified = MIRR(ReinvestRate:=0.12, ValueArray:=d, FinanceRate:=0.1)
    v = 41
    On Error Resume Next
    v = MIRR(d, -1, 0.1)
    errors = v & "," & Err.Number
    Err.Clear
    d(1) = 100
    v = MIRR(d, 0.1, 0.1)
    errors = errors & ";" & v & "," & Err.Number
    Err.Clear
    Erase d
    v = IRR(d)
    errors = errors & ";" & v & "," & Err.Number
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
        for bits in [0x3FCFDB68EA9C0995, 0x3FCFDB68EA9BECBF, 0x3FB3B86EA20AE5D0] {
            assert!(
                values.contains(&canon(&Variant::from_f64(f64::from_bits(bits)))),
                "{executor:?}: {values:?}"
            );
        }
        assert!(
            values.contains(&Canon::Str("41,5;41,11;41,5".into())),
            "{executor:?}: {values:?}"
        );
    }
}
