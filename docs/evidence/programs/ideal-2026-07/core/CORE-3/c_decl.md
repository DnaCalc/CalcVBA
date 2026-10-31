# Declare ordinal source validation

2026-10-31. Owner `bd-59co.2.4.5` remains **in-progress**. This change checks
ordinal alias literals before building an external call plan. Broader PtrSafe,
signature, AddressOf and declaration legality remain with that owner; Windows
import certification stays with WIN-14.

[MS-VBAL external declarations](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/75679e90-7e14-420d-af11-f83ffaf60418)
requires an ordinal alias to contain an integer literal valued from 0 through
32,767. The [number-token grammar](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/685ad840-accb-4bdb-8bfd-f3d88498547a)
permits decimal, hexadecimal and octal spellings, including octal without O;
the integer-literal production excludes signs and type suffixes. Those rules
decide the static target, independently of the rejected Excel attempts below.

The scanner now rejects invalid aliases with `SYM-E-INVALID-DECLARE-ORDINAL`.
The error points to the original quoted Alias token. The symbol keeps the
original spelling and a checked ordinal value; binding emits a decimal alias
for the existing native loader. Thus `#&H7FFF`, `#&O77777`, `#&77777` and
`#00032767` select the same ordinal as `#32767`. Values above 32,767, malformed
digits, signs, floats, suffixes and empty ordinal strings fail before loading.
Named aliases and library names decode their outer delimiters and doubled
quotes once; a literal leading quote cannot turn a named export into an ordinal.

## Fresh development observations

The unchanged owned supervisor, worker and guardian from `30ac2cb6` were used
through retained source-specific runner wrappers. Each wrapper freezes a source
and expectation into an existing evidence-contract shape before Excel starts.
The worker validates the sealed descriptor, actual injected source, ownership
and guardian evidence as usual. The default harness catalog was not changed.
No DLL entry is invoked by these Excel probes: a successful module returns a
constant after a clean forced VBE Compile.

Six targeted probes pass the full development contract:

| Alias | Forced compile | Runtime |
|---|---|---|
| `#0` | ok | constant returned |
| `#1` | ok | constant returned |
| `#32767` | ok | constant returned |
| `#-1` | Invalid ordinal | not run |
| `#1.0` | Invalid ordinal | not run |
| `#` | Invalid ordinal | not run |

Their negative observations include the exact selected Alias token, full line
and successful linked dismissal. Every accepted probe reports owned-process-zero.
Excel 16.0 build 20527, Windows x64, Safe Mode, observed AutomationSecurity 1.
No Office security or registry setting was changed. Pre-existing Excel PID 3144
and its original start time survived; no owned Excel remained after the probes.

Six other attempts remain **rejected/diagnostic**, not accepted oracle passes:
`#32768`, `#65535`, `#65536` and `#1x` display Invalid ordinal but expose an empty
selection and a declaration line, which the token-required contract rejects.
`#+1` captures and dismisses Invalid ordinal, but its preselected success
expectation was wrong. `#01` is rewritten to `#1` by the editor during injection;
the exact-source check rejects it before Compile. These limitations are retained,
and neither post-dismiss snapshots nor inferred tokens repair their authority.

`declare-ordinals-2026-10-31/raw.zip` retains the source-specific wrappers, plans,
source, raw observations, results and failures. Its manifest records every raw
byte length/SHA-256 and the archive/implementation hashes. Text remains inside
the archive under the repository LF-only policy. The overrange/no-selection
authority issue remains explicit in CURRENT_BLOCKERS.

## Checks and remaining scope

Focused source tests cover valid radix forms, range/grammar failures, exact UTF-8
Alias spans under LF/CRLF with preceding Unicode, inactive conditional branches,
active branch offsets and escaped quotes. Binder tests prove canonical native
descriptors and rejection before a call plan exists. Eight Windows VM3/JIT tests
pass, including the equivalent ordinal spellings' missing-entry Err 452,
description, LastDllError preservation and handle balance; the opt-in owned
workbook test remains ignored in the ordinary suite.

Independent review found and corrected the escaped-quote classification bug,
then found no blocking code issue. Full `scripts/meta-check.ps1 -NoArtifacts`
passes: 3,649 tests, 179 ignored, formatting, strict Clippy, governance, path
stability and workspace check. Runtime execution and ABI models did not change; existing formal
manifest/policy checks remain applicable, and no new Kani model is claimed.

Next: establish a separate exact evidence contract for compile failures where
VBE genuinely selects no token, without weakening the token-required cases or
inventing a highlighted token. Continue the broader declaration/AddressOf
requirements under `bd-59co.2.4.5`, then structural/native-string aliases under
the existing WIN-9 owners. No canonical capability row is promoted here.
