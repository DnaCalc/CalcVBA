//! `MidB(s, start[, length]) = x` replaces bytes of the string's UTF-16LE storage (MS-VBAL
//! 5.4.3.x). The MidB expectations are Excel VBA's answers (probe 2026-11-17: the six cases
//! below, UTF-16 units as hex). Also `InStrB` under `Option Compare Text` (byte position, text
//! compare; from the spec).

use calcvba_differential::{Canon, Executor, run};

/// Run on VM3 and the JIT, require identical results, and return the first global.
fn both(source: &str) -> Canon {
    let vm3 = run(Executor::Vm3, source);
    assert!(
        vm3.unsupported.is_none(),
        "vm3: {:?}\n{source}",
        vm3.unsupported
    );
    let vm3_values = vm3
        .result
        .clone()
        .unwrap_or_else(|e| panic!("vm3: {e}\n{source}"));
    let jit = run(Executor::Jit, source);
    if jit.unsupported.is_none() {
        assert_eq!(
            jit.result.as_ref().ok(),
            Some(&vm3_values),
            "jit vs vm3\n{source}"
        );
    } else {
        panic!("jit declined: {:?}\n{source}", jit.unsupported);
    }
    vm3_values[0].clone()
}

fn units(body: &str) -> Canon {
    // Hex of each UTF-16 unit of `s`, comma separated.
    both(&format!(
        "Public r As Variant\nSub Main()\nDim s As String, i As Long, t As String\n{body}\n\
         For i = 1 To Len(s)\nt = t & Hex$(AscW(Mid$(s, i, 1))) & \",\"\nNext\nr = t\nEnd Sub\n"
    ))
}

#[test]
fn midb_replaces_a_whole_character_at_an_even_byte() {
    // spec: s = "abcd", MidB(s, 3, 2) = "X" replaces bytes 3..4 (the character "b").
    assert_eq!(
        units("s = \"abcd\": MidB(s, 3, 2) = \"X\""),
        Canon::Str("61,58,63,64,".into())
    );
}

#[test]
fn midb_at_an_odd_start_replaces_the_high_byte() {
    // spec: byte 2 is the high byte of "a" (0x0061); "X" low byte 0x58 lands there -> 0x5861.
    assert_eq!(
        units("s = \"abcd\": MidB(s, 2, 1) = \"X\""),
        Canon::Str("5861,62,63,64,".into())
    );
}

#[test]
fn midb_without_length_and_dollar_form_replace_available_bytes() {
    assert_eq!(
        units("s = \"abcd\": MidB$(s, 3) = \"XY\""),
        Canon::Str("61,58,59,64,".into())
    );
}

#[test]
fn midb_never_extends_the_string() {
    assert_eq!(
        units("s = \"ab\": MidB(s, 3, 8) = \"WXYZ\""),
        Canon::Str("61,57,".into())
    );
}

#[test]
fn midb_without_length_replaces_as_many_bytes_as_the_replacement_has() {
    assert_eq!(
        units("s = \"abcd\": MidB(s, 1) = \"XYZ\""),
        Canon::Str("58,59,5A,64,".into())
    );
}

#[test]
fn midb_at_an_odd_start_splits_characters() {
    // Bytes 4..6 take the first three bytes of "XY" (58 00 59).
    assert_eq!(
        units("s = \"abcd\": MidB(s, 4, 3) = \"XY\""),
        Canon::Str("61,5862,5900,64,".into())
    );
}

#[test]
fn instrb_honours_option_compare_text() {
    let source = "Option Compare Text\nPublic r As Variant\nSub Main()\nr = CStr(InStrB(\"ABC\", \"b\"))\nEnd Sub\n";
    assert_eq!(both(source), Canon::Str("3".into()));
}
