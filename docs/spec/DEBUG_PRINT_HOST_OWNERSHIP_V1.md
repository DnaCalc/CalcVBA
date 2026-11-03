# Debug output

The compiler recognises Debug.Print and retains its evaluated arguments and
formatting. VM3 and the JIT send the resulting text through DiagnosticsHal and
on_debug_print. Hosts choose its destination. A host does not need to emulate
a Debug COM object. Bare Print uses the console callback separately.

`crates/oxvba-host/tests/debug_and_console_print.rs` checks evaluation, tabs,
semicolons and routing on the current execution paths.
