//! Defects found by DNA Calc's VBA validation corpus (dc-2z5n), each with the Excel behaviour
//! its minimal repro measured: both executors must agree and give Excel's answer.

use calcvba_differential::{Canon, Executor, RunOutcome, canon, run};
use calcvba_runtime::Variant;

fn first(executor: Executor, source: &str) -> Canon {
    let outcome: RunOutcome = run(executor, source);
    assert!(
        outcome.unsupported.is_none(),
        "{executor:?} unsupported: {:?}",
        outcome.unsupported
    );
    let snap = outcome
        .result
        .unwrap_or_else(|e| panic!("{executor:?} run failed: {e}"));
    snap.first()
        .cloned()
        .unwrap_or_else(|| panic!("{executor:?}: empty snapshot"))
}

fn both(source: &str, expected: &Canon) {
    for executor in [Executor::Vm3, Executor::Jit] {
        assert_eq!(&first(executor, source), expected, "{executor:?}");
    }
}

#[test]
fn single_line_for_loop_with_colon_separators() {
    both(
        "Public r As Variant\nSub Main()\n    Dim i As Long, s As Long\n    For i = 1 To 3: s = s + i: Next i\n    r = s\nEnd Sub\n",
        &canon(&Variant::from_i32(6)),
    );
}

#[test]
fn optional_parameter_typed_with_enum_uses_its_default() {
    both(
        "Public r As Variant\nEnum E\n    A = 0\n    B = 1\nEnd Enum\nFunction F(Optional m As E = E.B) As Long\n    F = m\nEnd Function\nSub Main()\n    r = F()\nEnd Sub\n",
        &canon(&Variant::from_i32(1)),
    );
    both(
        "Public r As Variant\nEnum E\n    A = 0\n    B = 1\nEnd Enum\nFunction F(Optional m As E = 0) As Long\n    F = m\nEnd Function\nSub Main()\n    r = F()\nEnd Sub\n",
        &canon(&Variant::from_i32(0)),
    );
}

#[test]
fn single_line_loops_keep_statements_after_the_closing_keyword() {
    both(
        "Public r As Variant\nSub Main()\n    Dim i As Long, n As Long, s As Long\n    For i = 1 To 3: s = s + i: Next i: s = s * 10\n    Do While n < 4: n = n + 1: Loop: s = s + n\n    While n < 6: n = n + 1: Wend: s = s + n\n    r = s\nEnd Sub\n",
        &canon(&Variant::from_i32(70)),
    );
}

#[test]
fn member_attribute_statement_inside_a_procedure_body_is_ignored() {
    both(
        "Public r As Variant
Function F() As Long
Attribute F.VB_Description = \"demo\"
Attribute F.VB_UserMemId = 0
    F = 7
End Function
Sub Main()
    r = F()
End Sub
",
        &canon(&Variant::from_i32(7)),
    );
}

#[test]
fn enum_member_initialised_with_a_library_qualified_constant() {
    both(
        "Public r As Variant
Public Enum CompareMethod
    BinaryCompare = VBA.vbBinaryCompare
    TextCompare = VBA.vbTextCompare
End Enum
Sub Main()
    r = CompareMethod.TextCompare * 10 + CompareMethod.BinaryCompare
End Sub
",
        &canon(&Variant::from_i32(10)),
    );
}
