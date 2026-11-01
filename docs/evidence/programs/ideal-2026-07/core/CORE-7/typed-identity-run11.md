# Typed Is and native object identity â€” run 11

2026-11-01. Scoped delivery bd-59co.2.9.17 is closed; broad owners bd-59co.2.9.9 and
bd-59co.3.5.3 remain open. CORE-JIT-LOWERING and WCC-EARLY-COMPLEX retain their
existing incomplete statuses; this is development evidence, not certification.

The owning physical layout admits original CompareObjectIs on VM3/native JIT.
The shared single-operation kernel compares retained identity without resolving
object defaults. Object and object-bearing Variant values agree; two Nothing
references compare True, one Nothing compares False, and non-object Variants
raise 424 through the existing retained error/Resume machinery.

Native preparation queries and retains the exact canonical IUnknown before
publishing a table, including nominal interface and opaque Variant results.
Its owner spans every wrapper/callback capture. Comparison performs no COM call,
name/type-info lookup or Invoke. A failed identity QI releases the previously
acquired custom interface. Native pointers never become portable table headers.
Foreign runtime IUnknown carriers retain pointer identity; independently acquired
native COM interfaces require the native preparation boundary.

Authority: Microsoft's [Is operator](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/is-operator)
and [QueryInterface identity rules](https://learn.microsoft.com/en-us/windows/win32/com/rules-for-implementing-queryinterface).
The registration-free fixture has distinct custom-interface addresses with one
IUnknown, proves zero QI during Is and zero final references, and checks failed
identity preparation. Existing native fixtures cover interface-result comparisons,
Dictionary object/Variant round-trips, errors and zero Invoke/name/type-info reads.
Portable cases exercise copies, same/distinct/Nothing and bad inputs on both
backends, resuming one operation at a time and without calling default getters.

Owned Excel probe calcvba11-identity forced VBE compile command 578, then observed
same=True, different=False, boxed=True, nothing=True, one=False and error 424 for
Empty/Null/Long/String/CVErr. PID 48180 was owned and cleaned up; PID 3144 was
excluded. No settings, COM registrations or HKLM changed. The exact probe and
supervisor are retained in typed-identity-run11.bas/run-typed-identity-oracle.ps1.

Four installed Lean 4.29.1 ownership/continuation/reentry/handle abstractions are
checked separately. They do not prove native memory or code-generation refinement;
the existing Miri dependency/sysroot blocker remains without passing credit.
Source, command and raw oracle/check hashes are retained in run11-identity.json
and run11-identity.zip beside this document.

Residual: custom IUnknown-only serving/import breadth, nominal conversions/arrays,
proxy/certification and general typed carriers remain with the existing owners.
BLK-WORKSHEETFUNCTION-VTABLE-20261101 is unchanged; no fallback or slot was added.

Final repository check: 3749 passed in 199 suites, 181 ignored; four registry-writing fixtures excluded without credit. Strict workspace/all-target Clippy, formatting, governance and path stability pass. Fresh-eyes review checked default bypass, raw/runtime identity, failed acquisition and wrapper/result lifetimes; no material issue remains in this scoped delivery.
