use calcvba_syntax::parser::parse;

fn errors(src: &str) -> Vec<String> {
    parse(src)
        .errors()
        .iter()
        .map(|e| format!("{}@{}", e.message, e.offset))
        .collect()
}

#[test]
fn bare_named_arguments_with_to_and_copies() {
    // Recorded macro: PrintOut with named arguments, no parentheses (Excel compiles it).
    let src =
        "Sub t()\n    ActiveWindow.SelectedSheets.PrintOut From:=1, To:=2, Copies:=1\nEnd Sub\n";
    assert_eq!(errors(src), Vec::<String>::new());
}
