# Automatic document events, run 15

Owner: bd-59co.3.6.10. Status: in-progress pending the final repository check. Native source-compiled event and workbook comparisons pass; the final ownership audit finds 394 recorded identities and zero live or unverified owned processes.

The retained owned run `run15-m-automatic` visibly compiles the unchanged VBA
project, then compares VM3 and native JIT against Excel. Both runtimes repeat
the original nesting/order/rebind/release/cancellation scenario after reset,
reload, and detach/reset/reattach. Reload reverses source-module order and asserts
that the workbook document's class number changes. Each produces the same VBA observation.
Attachment uses module names across reload, so class ordering is not identity.
An idempotent attachment does not unsubscribe and reorder existing handlers.

Native idle delivery is also observed on both runtimes. A separate STA attaches
to the exact owned Excel HWND/PID and writes Z1 while the primary session has no
active invocation. The primary STA pumps its message queue. Its private document
handler runs exactly once and returns `1:DZ1;`, matching Excel's idle observation.
The test retains the actual native connection-point subscription throughout.

Controlled callbacks check duplicate attachment, detach, reset, reload, stale
callback rejection, stale procedure handles and disposal. The callback epoch
counter lives outside resettable event state. A weak callback retained after
session disposal is inert rather than entering a replacement instance.

Saved/reopened workbook probes compare startup Open delivery, multi-area Change
Target objects and area enumeration, BeforeSave's ByVal SaveAsUI plus ByRef
Cancel, and custom public document methods/properties/fields through Object and
Variant. Private members reject with error 438. The host supplies lifecycle
timing and imported IID/DISPID; the engine contains no Excel event-name rule.

Unhandled Change and BeforeSave error 11 abort their callers despite caller
On Error Resume Next. The owned UIA guardian captures and ends Excel's exact
runtime dialog. CalcVBA returns the same error and cell effects. Cancel is copied
before the callback fault aborts the outer call; canceled saves leave package
hashes unchanged. Idle errors remain available through take_event_error; hosts
own debugger and user-interface policy.

The installed Lean 4.29.1 DocumentEvents abstraction checks idempotent attach,
detached rejection and monotonic reset epochs. It does not prove native COM/TLS,
raw writeback cells, Rust storage or generated-code refinement. The owning formal
backlog records those obligations without passing credit.

Broader delivery remains open under bd-59co.3.6.5: multiple/structural native
ByRef authorities, cross-apartment transport, general source-class incoming COM
serving, source-free event descriptor retention, borrowed-session entry and
pinned certification. The corpus does not promote any canonical capability row.

Repository check run15-check10 stopped at the synthetic dispatch fixture: it rejected the legal combined method/property-get flags used by dynamic value calls. The fixture now accepts those flags only in its dynamic mode, and the test asserts all six dynamic calls use exactly those flags. Both focused runtime paths pass; the complete run15-check11 rerun is pending. The [manifest](../WIN-5/run15-automatic-workbooks.json) and [archive](../WIN-5/run15-automatic-workbooks.zip) retain all 25 final native comparisons and explicitly mark repository-check status. No commit or bead completion is justified until that check is green.
