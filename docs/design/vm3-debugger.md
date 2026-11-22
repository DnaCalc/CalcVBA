# Source-level VBA debugging in VM3

A host selects `BackendPolicy::Debugging` before loading the project and
starts a fresh invocation on VM3. The proposed debugger pauses at a source
statement and retains the invocation, including the original host read view.
[Engine selection](../ARCHITECTURE.md#embedding-and-project-state)

## Runtime hooks

### Source locations and breakpoints

The syntax tree exposes byte ranges, and conditional compilation preserves
offsets in the active source. OxIR has procedure names, local names and types,
and `StmtBoundary` markers for temporary cleanup and event delivery. Its
statement index comes from lowering a statement list. Numeric line labels
feed `Erl`. A source debugger needs a map across these layers.
[Syntax ranges](../../crates/calcvba-syntax/src/red.rs),
[conditional compilation](../../crates/calcvba-symbol/src/cond_comp.rs),
[statement lowering](../../crates/calcvba-oxir/src/elaborate/lower.rs)

Capture each executable statement's span before binding expands it into Core
IR. Carry that source identity through elaboration to a breakable OxIR entry
location (unit, procedure, block and instruction). Record the original module
and source revision alongside the span. Use the same map for call sites and
for source-visible locals, including `Static` storage and the function result.
Generated cleanup, error landing pads and scaffolding can have no breakable
location. Compilations with source maps should keep them with their code
generation, keyed to the exact source bytes. Source-less images can report
that source breakpoints are unavailable.

A breakpoint request names a project, module revision and line or statement range.
Return the resolved statement and its actual span, or an unbound breakpoint
with a reason. This gives the editor a placement for the marker and handles
continued lines, several statements on one line and inactive `#If` branches.
Resolve pending breakpoints again after a successfully compiled edit.

### Pause, continue and step

VM3's general dispatch loop already retains each frame's next block and
instruction, live aliases and saved error continuation. It runs to completion
and its host entry unwinds afterward. The host's sliced typed entry already
retains an owning `TypedStack`, a session lease, a read view and a budget.
That continuation counts operations and uses a separate physical PC. Scalar
value calls can also enter `TypedInterpreter` directly.
[VM3 dispatch](../../crates/calcvba-vm3/src/lib.rs),
[typed execution](../../crates/calcvba-vm3/src/typed.rs),
[host continuation](../../crates/calcvba-host/src/invocation.rs)

Start with a retained invocation for the general VM3 stack and a debug host
entry that routes through it, including calls that would normally use the
scalar path. Reuse the host's continuation lifecycle and read-view scope.
Expose a distinct stop result containing a reason, source position and pause
identifier. A later typed debug entry would use the same stop rules with a
physical-PC translation and typed-slot inspection.

At a source entry, complete the ordinary boundary cleanup and then offer a
stop before the next statement's effects. Persist the continuation after any
completed cleanup. A resume executes the stopped statement once and may stop
again on a later visit to the same location. Step Into stops at the next
reachable source statement, including a VBA callee. Step Over waits for the
next statement in the selected activation or a surviving caller. Step Out
waits for a statement in a surviving caller once the selected activation has
left. If the invocation completes first, report its ordinary outcome.
Identify an activation separately from its procedure or stack depth, so
recursion, error unwinding and sibling calls preserve the meaning of the step.
A breakpoint can stop inside a callee while an Over or Out is pending. Keep
the ordinary execution budget across pauses. Represent a debugger stop
separately from a VBA fault or budget error.

The UI posts a pause request, which VM3 observes cooperatively at a safe
boundary. Runtime execution, inspection and callbacks all run on the owning
VBA thread.
Native calls and synchronous callbacks finish on that thread, including their
return values and ByRef writebacks. Defer a requested pause there until control
reaches a resumable boundary. A retained stop inside a VBA callback will need
resumable call state for the entire chain. For explicit VBA breaks, preserve
`Stop` during binding and retain the already evaluated `Debug.Assert`
condition so a break decision can use it. Today `Stop` binds away and
`Debug.Assert` evaluates its arguments for effects. [Statement binding](../../crates/calcvba-bind/src/stmt.rs)

### Call stack and locals

Return source-qualified stack frames with opaque identities, the suspended
statement for the top frame and the call site for each ancestor. Use source
symbols to show arguments, locals, the return variable, `Me` and applicable
Static and module state. The existing `Vm3::slot` reads a result snapshot of
globals and entry locals. Add a pause-scoped read path for a particular frame.
Follow ByRef aliases to their live backing storage and use the selected
frame's project to resolve names. [VM3 frame and slot storage](../../crates/calcvba-vm3/src/lib.rs)

Inspect raw storage. Ordinary VBA value reads can instantiate `As New` objects
or run default getters. Show the stored type and value without activating
those paths. Copy scalar, string and error display data and expose aggregates
through bounded, lazy views. For foreign objects show identity/type metadata
available from storage. Keep object owners and addresses inside the session.
Bind all frame and value handles to the session, code revision and current
pause. Invalidate them on resume, replacement or cancellation.

### Immediate window

Connect `Debug.Print` to the host's existing diagnostics stream. For an input
such as `? name + 1`, use the VBA parser and binder with the selected frame's
lexical scope and bind the debugger input to its live slots. Start with scalar
expressions over stored values. Explicit commands that call procedures, set
variables or access properties need execution in a separate debugger
activation with the same owning thread, coercions, error handling, budget and
host capability checks as the suspended invocation. Preserve that activation's
distinction from the parked stack, in particular its error mode and `Resume`
targets. Measure visibility and the effect of a failed command on `Err` in
Excel before defining those rules. Make expression evaluation a user command
so expanding Locals never has the effects of running Immediate input.

### Edit-and-continue

`OwnedSession::reload` prepares a new session, advances its generation and
publishes it on success. The old continuation retains its old code and state.
Continuing an edited program in the same invocation needs a different
operation. [Session reload](../../crates/calcvba-host/src/cache.rs)

Compile edits against a source revision and validate a complete replacement
before publishing anything. For a first increment, investigate body edits to
inactive procedures with compatible signatures and unchanged global, Static,
class and record layouts. Future calls can then target replacement bodies in
the same session through stable procedure identities. Treat live `AddressOf`
callbacks, event routes and project-object methods as code owners when deciding
if a replacement is safe. Parked frames must keep their original code, storage
and copyback targets. A change to an active body requires a verified map for its
next statement, call return sites and live error or loop continuations, plus
compatible local and temporary storage. Require a restart when the edit fails
those checks. A rejected edit leaves the previous continuation resumable.
Check VBA's permitted edits and exact resume locations against the VBE before
expanding the supported edit shapes.

## DNA Calc and Formula Explorer

The DNA Calc views (`dc-etcxb`) and Formula Explorer can share a debugger
surface over their host's session commands. From a formula's VBA call or a
procedure in the project view, start a debug invocation with its original
arguments and read view. Send breakpoint, continue, pause, step, inspect,
Immediate and apply-edit commands to the owner thread. Return stop events,
source spans and value views to the editor, stack and locals panes. Keep the
spreadsheet formula trace alongside the VBA call entry, and publish a formula
result only after the invocation returns. Queue competing execution in the
leased VBA session until the invocation completes or is cancelled. Drop the
continuation on reset or close through the owning thread.

Validate the first breakpoint/step/inspect slice with authored probes for
recursion, ByRef, same-line statements, linked projects, error handlers and
native callback boundaries. Compile authored VBA probes through the VBE's
Debug > Compile command in an owned Excel session to compare stop locations,
visible values, Immediate effects and accepted edits.
Runtime work belongs to the VM3 epic `dc-17aqo`. Shared view work belongs to
`dc-etcxb`, with Formula Explorer hosting under `dc-0f1k`.
