# Owning source classes, records and ParamArray - run 11

2026-11-01. Delivery bd-59co.2.9.18 is closed for the bounded cases below on VM3 and
native JIT. CORE-JIT-LOWERING remains in-progress; no Windows certification credit.

Record slots own descriptor-backed VbaRecord values; array slots own canonical
SAFEARRAY descriptors. Only source-declared Variant slots and ParamArray elements
use Variants. Record fields, fixed arrays in records, value copies, return values,
repeated live ByRef aliases and host schema checks are exercised. Wrong schemas
and typed array element types reject before the body changes canonical globals.
A cross-project ByRef global read now uses the target frame's actual unit store.

ParamArray packs a zero-based Variant SAFEARRAY, including an empty 0:-1 packet.
Source calls preserve writable private arguments and repeated/live caller aliases;
copying the packet detaches it. Failed element coercion preserves the caller.
Fixed prefixes, twenty host tail arguments, IsArray, object/Nothing cells and
nested forwarding are covered with one-operation slices on both backends.

Provider::with_class_modules explicitly selects class modules for a single source
project. The public catalogue contains procedural entries only. Class objects
use the runtime's canonical IUnknown carrier and existing field store. New,
Class_Initialize, statically declared source member calls, primitive/object fields,
object results, Is and Class_Terminate execute the original typed functions.
Source class call targets resolve from declared class metadata before execution;
unproven dynamic receivers remain refused. Constructors/finalizers participate
in call preflight. Statement boundaries and normal frame return release owned
locals and drain termination work without holding a runtime state borrow.

Each session/reset generation has a same-thread object owner. Deferred queues
include that owner, preventing identical instance IDs in separate sessions from
routing to the wrong project. Host-held stale class values reject on class entry.
Expiry discards orphan parked boxes, including cascaded fields; scope guards
restore nested ownership. Cancellation at constructor/termination slices and reset
leave the session usable. Native host slots validate a factory vtable before
reading the larger portable header; source objects are never cast to that header.

Owned Excel VBA receipts force VBE compile command 578 with a ready PID-scoped
modal guardian. calcvba11-aggregates-final observes record=42:owned:7, copy=41,
aliases=44, param=9:9:0, empty=0:-1, returns=42:42:owned:7 and prefix=15.
calcvba11-paramarray-errors observes private=5, repeat=9:9, copy=9:27:9,
bad=13:9:9, nested=0:0:True and object=True. calcvba11-classes-final observes
same=True|start=40|next=42|distinct=False|journal=IIT42T40. All owned Excel
processes are cleaned up. PID 3144 and unrelated Excel processes are excluded.
The harness seals each additional class module and verifies its injected source.
It dismisses only an exact, PID-checked informational search coach with Escape;
security, trust, recovery files, registrations and HKLM are unchanged.

run11-aggregates-classes.json/.zip retain source and raw focused/full-check,
formal and oracle receipts. Fresh review checks alias/store authority, generic
IUnknown versus factory-header access, finalizer ownership and cancellation.
Four installed Lean abstractions are checked; concrete aggregate/owner/finalizer
refinement and the existing offline Miri blocker remain in EXTENDED_TODO.md.

Residual delivery path: bd-59co.2.9.9 and bd-59co.2.10.7 retain richer numeric and
fixed-string carriers, ReDim/Erase/ForEach, project interfaces/Implements, As New,
predeclared instances, public-field/default-member breadth, aggregate class fields,
optional/named/ParamArray member arguments, source-graph class import/activation and source-free publication. Native COM record
ABI, native callbacks and certification retain WIN-4/WIN-5 and CORE-7 owners. The four-formal,
128-slot, 64-frame limits remain. No compatibility member or native guessed slot
was added; BLK-WORKSHEETFUNCTION-VTABLE-20261101 awaits the owner.

Final check: 3758 passed in 201 suites, 181 ignored; four registry-writing fixtures excluded without credit. Strict workspace Clippy, formatting, governance and path stability pass. Fresh review found no unresolved issue in the bounded cases; the remaining behavior and formal questions stay with the named open owners.
