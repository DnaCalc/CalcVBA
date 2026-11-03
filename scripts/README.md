# Scripts

`check.ps1` is the repository check: fmt, strict Clippy and workspace tests.
It excludes four registry-writing COM record fixtures, just as the previous
NoRegistry check did. It does not run live Excel tests.

`run-smoke.ps1` runs the small CLI fixture. `run-*-oracle.ps1` scripts run
specific Excel/VBA comparisons. The current hosting comparisons accept a
matching `-RuntimeTestExecutable` built with `cargo test --no-run`.
`run-xlsm-suite.ps1` runs the saved-workbook corpus.

The Excel helper scripts retain process ownership, UIA modal handling and
explicit VBE compilation. `test-excel-vba-oracle.ps1` checks that handling
without launching Excel. Never change macro security or registrations to make
an oracle run succeed; never dismiss or terminate another run's processes.
Generated observations belong under `temp/`, not in the source tree.
