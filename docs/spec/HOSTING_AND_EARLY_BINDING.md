# Hosting and early binding

Hosts supply application objects and type information. CalcVBA owns VBA syntax,
binding, runtime values and execution. Excel object names, worksheet errors,
calculation scheduling and spreadsheet policy belong to DNA Calc or another
host, not to an engine-built substitute for Excel.

## Entry and state

`calcvba-host-api` defines projects, modules, ordered library references,
procedure signatures, values, bound handles, host objects and runtime providers.
`calcvba-host::provider::Provider` compiles and loads them on VM3 or the JIT.
Loading initialises project state without calling Main. Calls retain globals,
Static variables, Err and object identity. Separate sessions have separate state.
Reset clears project state; reload prepares replacement code before publishing
it. Bound handles check session ownership and generation.

The typed scalar entry supports admitted Long/Double/host-object families using
physical slots. The owning value entry additionally supports tested strings,
Boolean, Variant, arrays, records, classes, internal calls, ByRef aliases and
error handling. Admission checks the reachable call graph before body effects.
Unsupported bodies/signatures fail explicitly; no JIT-to-VM fallback occurs.
Compiler-owned signatures, not runtime name lookups, determine passing rules.

DNA Calc uses begin_invocation, InvocationFrame, read_view, resume(quantum),
execution budgets and reload. Retained activations keep the original code and
project state alive; completed effects are not replayed when a slice resumes.
These facilities serve the calculation scheduler. Invocation metadata belongs
to the host and is not an additional VBA language feature.

## Native objects and calls

Imported type libraries provide IIDs, member kinds, DISPIDs, vtable slots,
parameter/result/wire types and defaults. Dual calls use the retained native
vtable plan and do not retry through IDispatch. Pure dispinterfaces use an
explicit compiler-prebound Invoke plan. Object/Variant calls and untyped
members resolve dynamically. Both backends use shared marshalling and error
translation; nominal interfaces and native object results retain COM identity.

Host tables can supply portable typed objects without native Windows COM.
HAL host and callback traits do not require Send or Sync. Arc host ownership
remains for the DNA Calc API; it does not permit cross-thread execution.
Native interface references remain on the owning apartment. The host controls
lifecycle and supplies the root objects; it must not pass another thread's raw
COM interface to a session. The runtime executes on its owning VBA thread.

## Events

Compiler metadata selects source interfaces and private event handlers.
Hosts attach document identities by module name and supply lifecycle timing.
Owning sessions wire WithEvents and document sources through retained native
connection points. Native sinks retain the same-thread Rc callback directly,
without a TLS callback-ID registry. Supported synchronous events enter the same VM3/JIT session;
tested Boolean ByRef arguments copy back before returning to the native caller.
Idle delivery needs an STA message pump. Reset/reload/detach reject stale
callbacks and release subscriptions. Unhandled handler errors abort the
triggering call; take_event_error exposes idle errors to the host.

General structural/custom event signatures, the older package-session event
path, custom-vtable sources, broader cross-process transport and JIT-backed COM
serving remain incomplete. Ordinary out-of-process COM events are in scope;
MTA/free-threaded VBA execution is not. Compare actual native behaviour with
controlled COM fixtures and opt-in Excel tests before extending support.

## Ownership and portability

Code caches share immutable code/metadata, not mutable VBA state. Owning
arguments, object results, arrays and records retain their runtime owners.
ByRef calls preserve actual aliases and use private cells for temporary values.
Native callbacks share the active session context and must unwind safely.

DNA Calc implements the HAL descriptor and policy interface; browser hosts use
host services to supply or refuse UI/filesystem/native operations. Browser
builds select the interpreter without default JIT features. The older
spreadsheet adapter is documented in `adapters/spreadsheet-compat/README.md`;
it is separate from the neutral engine API.
