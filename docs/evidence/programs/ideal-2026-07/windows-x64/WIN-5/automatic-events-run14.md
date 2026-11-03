# Automatic events, run 14

Owner: bd-59co.3.6.10. Status: in-progress.
Compiler routes now retain the coclass source IID, DISPID, exact parameter modes
and nominal argument IIDs. Wrong native handler declarations fail compilation;
nine tested descriptor/signature mutations fail verification. Hosts provide document module
identities, without registering event names, callbacks or handler procedures.

Both owning runtimes subscribe on document attachment and WithEvents assignment,
discard old subscriptions on rebinding/Nothing, reject stale generations, and
release subscriptions outside borrowed state. Private handlers re-enter the same
session, globals, Err and block budget. Class_Terminate retains readable WithEvents
sources until its body returns; reset detaches and rewires the retained documents.
A controlled host-service test also delivers events with no active macro or
native call context, then checks private-handler execution, reset and subscription
disposal on both backends. This exposed missing nominal metadata for document
types and unused interface-only parameters; the compiler now retains those IIDs
without depending on a member call in the source. Native idle message transport
and cross-apartment delivery remain separate from this controlled runtime test.

Owned, visibly VBE-compiled Excel and VM3/native JIT agree on Worksheet_Change,
WithEvents Change, nested writes, handler order, same-source rebinding,
disconnection, release/termination and Workbook_BeforeClose Boolean cancellation:
`1:1:WA1;DA1;|nested=3:3:WA1;DA1;WA2;DA2;|rebind=4:WA3;DA3;|release=4:1:True:DA5;|close=1`.
The automatic-events oracle repeats the engine scenario after session reset on
each backend. Saved/reopened events and mixed XLSM projects independently retain
the same behavior and compare all used cells (see WIN-4/xlsm-projects-run14.md).
Application WithEvents also agrees on WorkbookBeforeClose's native workbook
argument, document-before-application order, shared workbook identity and Boolean
copyback: `1:1:DA:True`. It exposed missing private object unboxing/interface
conversion at a typed VBA call boundary, now corrected without member lookup.

The first native attempt exposed an implicit-Me Variant carrier mismatch, now
corrected. Only successful retained final runs count as passing observations.
The NoRegistry repository check passes 3,786 tests / 210 suites, 186 ignored;
four registry-writing record fixtures are excluded without credit. Strict Clippy,
format, governance and path stability pass. [The source manifest](run14-automatic-workbooks.json)
and [archive](run14-automatic-workbooks.zip) retain the actual sources, native
binary hashes, VBE/guardian/cleanup evidence and full check log. A final audit
finds zero live owned processes across 132 recorded Excel/helper identities.

Remaining delivery scope stays explicit under this bead and bd-59co.3.6.5 (certification support: bd-59co.3.6.4):
general source-class incoming dispatch; structural/multiple ByRef native arguments;
complete handler error policy and cross-apartment delivery; Workbook_Open/startup
lifecycle; reload, resurrection and multiple source interfaces sharing DISPIDs;
borrowed-session entry and pinned certification. Installed Lean abstractions do
not prove native/TLS/parked-owner refinement; see the formal backlog. No broad
event row or certification promotion is made.
Source-free event images also need a complete descriptor-retention proof;
historical routes without native descriptors are outside this source-compiled proof.
