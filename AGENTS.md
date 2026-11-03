# Working on CalcVBA

Read `CHARTER.md`, `README.md` and `docs/ARCHITECTURE.md` for scope and code layout.

- Match real Office VBA behaviour; turn corrected differences into tests.
- Keep VM3 and the JIT. Do not hide a JIT failure by falling back to VM3.
- Runtime execution and callbacks belong on the owning VBA thread. Compiler
  work may run in parallel.
- Keep changes direct. Use beads for real behaviour gaps, not process tasks.
  Do not add evidence zips, certification, formal proofs, governance checks,
  handoff documents, worksets or document-consistency checks.
- Before each commit, run `git pull --no-rebase` and `./scripts/check.ps1`
  (fmt, strict Clippy and workspace tests). Fix failures without weakening
  behaviour tests, review the diff, commit and push.
- End commit messages with `Co-Authored-By: Codex <noreply@openai.com>`.
- Never change Trust Center, macro security, Excel COM registrations or HKLM.
  Excel oracle runs must handle compile/runtime dialogs through UI Automation
  and clean up only their recorded, owned processes. Use visible VBE Debug ->
  Compile for compile diagnostics; `Application.Run` alone is not a compile check.
- Never use merged Excel cells; use Center Across Selection when needed.
- Do not give elapsed-time estimates for development work.
