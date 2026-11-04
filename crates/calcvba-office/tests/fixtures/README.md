These 33 workbooks are the saved Excel inputs used by the round 14–15 host
comparisons: seven original project cases and the 26 cases listed in
`calcvba-host/tests/support/xlsm/cases.json`. The JSON files retain the VBE code
pane snapshots and Excel return values/cells from those comparisons.

The reader tests compare persisted module roles, references, attributes and
source bodies with the snapshots. The owned Excel batch executes source read
from each saved workbook on both VM3 and the JIT and compares all captured
results. The snapshots do not supply the compiler's source.

`../formats` contains native Excel saves of the same UserForm in XLSM, XLSB,
XLS and XLAM, plus its VBE-exported form pair. The form contains a textbox and
button and a non-ASCII caption for resource/codepage checks.
