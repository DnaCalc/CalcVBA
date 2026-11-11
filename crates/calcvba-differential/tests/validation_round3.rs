//! Backend disagreements found by sweeping the validation corpus and the conformance fixtures
//! through both executors (dc-zxpq family): the JIT must compile or fail the way VM3 does.

use calcvba_differential::{Canon, Executor, RunOutcome, canon, run};
use calcvba_runtime::Variant;

fn outcome(executor: Executor, source: &str) -> RunOutcome {
    run(executor, source)
}

fn both_agree(source: &str) -> RunOutcome {
    let vm3 = outcome(Executor::Vm3, source);
    let jit = outcome(Executor::Jit, source);
    assert!(
        vm3.unsupported.is_none(),
        "vm3 unsupported: {:?}",
        vm3.unsupported
    );
    assert!(
        jit.unsupported.is_none(),
        "jit unsupported: {:?}",
        jit.unsupported
    );
    assert_eq!(vm3.result, jit.result);
    assert_eq!(vm3.raised, jit.raised);
    assert_eq!(vm3.err, jit.err);
    vm3
}

fn first(outcome: &RunOutcome) -> Canon {
    outcome.result.as_ref().unwrap().first().cloned().unwrap()
}

#[test]
fn join_of_a_string_array_into_a_string_destination() {
    let o = both_agree(
        "Public r As Variant\nFunction J(ByVal n As Long) As String\n    Dim a() As String\n    Dim i As Long\n    ReDim a(n)\n    For i = 0 To n\n        a(i) = Chr$(65 + i)\n    Next i\n    J = VBA.Join$(a, \"-\")\nEnd Function\nSub Main()\n    r = J(3)\nEnd Sub\n",
    );
    assert_eq!(first(&o), canon(&Variant::from_string("A-B-C-D")));
}

#[test]
fn declare_call_compiles_on_the_jit_and_raises_like_the_interpreter() {
    // Off Windows (or under host policy) the call raises at run time on both backends; the JIT no
    // longer declines the whole module at load (the riff corpus project).
    let o = both_agree(
        "Declare Function HostPing Lib \"host\" Alias \"ping\" (ByVal x As Long) As Long\n\nSub Main()\n    Dim y\n    y = HostPing(3)\nEnd Sub\n",
    );
    assert!(o.raised);
}

#[test]
fn host_sensitive_interaction_builtins_compile_and_agree() {
    let o = both_agree(
        "Sub Main()\n    Dim a\n    Dim b\n    a = Shell(7)\n    b = Environ(4)\nEnd Sub\n",
    );
    assert!(o.raised || o.result.is_ok());
}

#[test]
fn fault_descriptions_agree_between_backends() {
    for source in [
        "Sub Main()\n    Dim x\n    Dim o As Object\n    Set o = 5\nEnd Sub\n",
        "Sub Main()\n    Dim a() As Long\n    ReDim a(5 To 1)\nEnd Sub\n",
        "Sub Main()\n    Dim x As Long\n    Dim v As Variant\n    For Each v In x\n    Next v\nEnd Sub\n",
    ] {
        let o = both_agree(source);
        assert!(o.raised, "{source}");
    }
}

#[test]
fn sin_and_chr_take_variant_temporaries() {
    // Sin(phase * PI2) and Chr(CInt("&H" & t)) (Riff, VBA-Web): the argument is a Variant temp.
    let o = both_agree(
        "Const PI2 As Double = 6.283185307179586
Public r As Variant
Function F(ByVal phase As Double) As Double
    Dim s As Double
    s = Sin(phase * PI2)
    F = s
End Function
Sub Main()
    r = F(0.25)
End Sub
",
    );
    assert_eq!(first(&o), canon(&Variant::from_f64(1.0)));
    for source in [
        "Public r As Variant
Function F(ByVal t As String) As String
    Dim s As String
    s = VBA.Chr(VBA.CInt(\"&H\" & t))
    F = s
End Function
Sub Main()
    r = F(\"41\")
End Sub
",
        "Public r As Variant
Function F(ByVal t As String) As String
    Dim s As String
    s = Chr(CInt(\"&H\" & t))
    F = s
End Function
Sub Main()
    r = F(\"41\")
End Sub
",
    ] {
        let o = both_agree(source);
        assert_eq!(first(&o), canon(&Variant::from_string("A")));
    }
    // A Variant that cannot be converted raises the same error on both.
    let o = both_agree(
        "Sub Main()
    Dim v As Variant
    Dim s As String
    v = \"abc\"
    s = Chr(v)
End Sub
",
    );
    assert!(o.raised);
}

#[test]
fn left_and_right_take_a_variant_length_temporary() {
    // VBA-Web JoinUrl: Right(s, Len(s) - 1) has a Variant-typed length temp.
    let o = both_agree(
        "Public r As Variant
Function F(ByVal s As String) As String
    If Left(s, 1) = \"/\" Then
        s = Right(s, Len(s) - 1)
    End If
    If Right(s, 1) = \"/\" Then
        s = Left(s, Len(s) - 1)
    End If
    F = s
End Function
Sub Main()
    r = F(\"/a/b/\")
End Sub
",
    );
    assert_eq!(first(&o), canon(&Variant::from_string("a/b")));
}
