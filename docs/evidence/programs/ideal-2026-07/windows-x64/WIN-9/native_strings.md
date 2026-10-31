# Native String argument temporaries

2026-10-31. Bounded owner `bd-59co.3.10.10` is closed for the scope below; wider native-string parity remains
open under `bd-59co.3.10.6`. This change covers byte mutations in initialized
String buffers whose native LPSTR pointers are unchanged. Nullable/replaced
pointer ownership, locale breadth and mixed As Any shapes remain open. No
canonical row or pinned certification is promoted.

Unlike scalar/Variant cells, native String arguments use a fresh ANSI temporary
for each argument, including repeated caller variables. The alias classifier
retains a String-temporary class so mixed declarations still refuse before
entry, while matching String arguments fall through to ordinary fresh staging.
Each temporary is read back in argument order; the last argument therefore
writes the final caller value. Existing scalar and Variant pointer reuse remains.

## Fresh Excel characterizations

The controlled DLL compares raw LPSTR-cell/buffer addresses, writes X into the
first buffer's byte 0, checks readthrough, then writes Y into the second buffer's
byte 1. It retains no pointers and changes no allocation ownership. The ByRef
mask uses bits 1=shared cell, 2=shared buffer, 4=readthrough. The ByVal mask uses
bits 2 and 4. All inputs provide at least four bytes of initialized payload.

Expected values were deliberately unspecified before these six runs. They are
characterizations of actual VBA behavior, not matches to frozen predictions:

| Case | Exact returned observation |
|---|---|
| repeated ByRef String | `0\|aYcd\|4321` |
| parenthesized copy | `0\|aYcd\|4321` |
| forwarding | `0\|aYcd\|4321` |
| repeated ByVal String | `0\|aYcd\|4321` |
| distinct caller variables | `0\|Xbcd\|aYcd\|4321` |
| embedded NUL | `0\|4\|aY:d\|4321` |

The colon in the last report replaces NUL only for display; the length remains
4 and the regression checks full bytes. Mask zero establishes distinct ByRef
cells/buffers and no readthrough, or distinct buffers for ByVal. The source saves
Err.LastDllError before conversion helpers. Every run passes the unchanged owned
supervisor with command-578 Compile, exact source/project authority and
owned-process-zero cleanup. Excel 16.0 build 20527, Windows x64, AMD Ryzen 7
4800H, Safe Mode, observed AutomationSecurity 1. No Office security or registry
setting changed; pre-existing Excel PID 3144 and its original start time survive.

`string-alias-2026-10-31/raw.zip` retains the exact fixture source/DLL, wrappers,
input modules, selected descriptors, plans, guardian records, environment,
ownership/results and the pre-fix regression. The manifest records every member
length/SHA-256, archive and current implementation hashes, original paths, build
argv, compiler and CPU. PDB/import-library files are omitted. Archive text keeps
its original bytes under the repository LF-only policy.

## Checks and remaining work

The two positive regressions fail before the repair because repeated String
arguments are refused. Afterward VM3/JIT match all six characterizations,
including copies, forwarding, full-length NUL payloads, separate variables and
LastDllError. Mixed ByRef String/Any in either order still reports the explicit open
gap before the fixture executes (native call delta zero). The previous general
refusal control now exercises repeated Any. Runtime handle counts return to
baseline after each engine drops; this is not an OLE heap allocation counter.

The native parity suite passes 14 tests with the opt-in owned-workbook test
ignored; the 19 existing string controls pass. Some older controls retain JIT
declines outside this scoped repair and earn no new parity credit. Full
`scripts/meta-check.ps1 -NoArtifacts` passes: 3,655 tests, 179 ignored, format,
strict Clippy, governance, path stability and workspace check. Independent
source, evidence and hash review is clean.

No existing formal model covers this native staging boundary; the owning
`FTODO-NATIVE-ALIAS-20261031` remains open. The native Kani probe still reports
no such cargo command. Next work remains pointer replacement/nullable ownership,
locale breadth, structural carriers, foreign Boolean payload 1 and later early
COM/callback/reentry and pinned certification under their existing owners.
