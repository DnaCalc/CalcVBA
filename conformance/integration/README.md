# Project source fixtures

`projects/<case>/main` contains project modules; `references/<project>` contains
referenced modules. Names use `<ModuleName>.<kind>.bas` (`proc`, `class`,
`document`, `form`, or `extension`).

The lexer corpus and VM3/JIT source corpus read these files. Project-level
reference and event behaviour is exercised by the focused tests in
`crates/oxvba-host/tests` and `crates/oxvba-embed/tests`.
