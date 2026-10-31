# VBE selection diagnostic

2026-10-31 historical failed-run evidence. The later [owned supervisor repair](oracle-supervisor-2026-10-31/README.md) resolves BLK-EXCEL-COMPILE-20261031 on the development host and supersedes the remaining-work statements below. This earlier attempt remains rejected. No Windows capability or certification credit is granted.

The guardian called ExpandToEnclosingUnit with a nonexistent type,
`Windows.Automation.TextUnit`. Its catch path discarded the observed selection.
The correct enum is `System.Windows.Automation.Text.TextUnit`.

A separately held owned VBE diagnostic established that the Document element
supports TextPattern and returns the actual selected token. Its range clone
was valid; the enum lookup failed. DocumentRange returned the injected module.
A PID/start/executable-verified native-OM read independently saw the token and
full line while the same compile modal was still visible. Held-worker attempts
were diagnostic only and never count as an accepted oracle run.

With only the enum corrected, the ordinary repository supervisor produced an
immutable pre-dismiss record containing:

```text
Compile error: Sub or Function not defined
MissingOracleSymbol
    RunProbe = MissingOracleSymbol(1)
```

The complete production attempt is retained in
`vbe-selection-diagnostic-2026-10-31/run/`. It **failed** overall: the owned dialog
fragment returned no UIA button descendants, dismissal returned false, and the
supervisor rejected contradictory compile-status evidence. Cleanup was
owned-process-zero. No post-dismiss COM snapshot repaired that authority.

Experimental broader UIA lookup, native-parent checks and scan/timing changes
did not establish acceptance and were removed. The shipped implementation
change is only the enum correction. The next repair must establish exact owned
button resolution and a coherent asynchronous compile/evidence handoff, keeping
capture-before-dismiss and successful linked dismissal mandatory.

`scripts/test-excel-vba-oracle.ps1` now resolves the guardian's UIA type literals
against the installed Windows assemblies and rejects an invalid-namespace
mutation. The test performs no GUI operation. The complete oracle contract and
hostile-evidence suite passes. Existing Linux policy-only coverage does not
claim to exercise Windows UIA assemblies.

Office 16.0 build 20527, Safe Mode, observed AutomationSecurity=1. No Trust
Center, macro-security or registry setting was changed. Only recorded owned
Excel/helper processes were cleaned. Broader supervisor promotion and pinned
environment execution remain with the existing owners.

Implementation checkpoint: d2f2cc83cbc26f85896a9870be88e2a0f099a586
