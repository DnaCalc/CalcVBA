# Typed native values â€” run 10

2026-11-01. Scoped delivery bd-59co.3.5.8 is closed on the bounded native values/output/array subset.
WCC-EARLY-COMPLEX and WIN-1 carrier/certification producers remain open.

The extension retains descriptor plans for native Variant, output parameters,
nominal interface results and intrinsic SAFEARRAY signatures. Original scalar
entry ABIs remain unchanged. Member and result identities are resolved before
publishing the table; runtime calls consume retained slots and interface indices.
Only declared Variant/array values use native VARIANT cells. No universal Variant
procedure frame, Invoke fallback, live member-name or type-info lookup is added.
Native calls remain synchronous operations between continuation positions.

Authority: Microsoft's [x64 calling convention](https://learn.microsoft.com/en-us/cpp/build/x64-calling-convention)
requires indirect aggregate argument storage; these native temporary cells have
16-byte alignment. The [VARIANT contract](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-variant)
and [SAFEARRAY element ownership](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraygetelement)
govern the existing native conversion helpers reused here. The runtime's 16-byte
Variant core is not reinterpreted as the larger native Windows record-capable union;
the broader exact-layout producer bd-59co.3.2.5 is still open. No layout or
certification credit is claimed from this boundary extension.

Targeted embedding checks pass: six native/metadata cases, the default coercion
case, and eight neighboring value-slot cases; the owned Excel case is separate.
Real Dictionary Add/Exists/Item/Keys/RemoveAll makes 22 native calls on each backend and no
Invoke calls, preserving Null, CVErr, exact UTF-16, array values and duplicate-key
faults. The registration-free fixture proves shared versus distinct native ByRef
cells, writes before failure, private parenthesized copies, global writeback,
nominal interface results, array inputs/results and typed I4 SAFEARRAY replacement.
Zero native/IErrorInfo references and zero Invoke/name/type-info reads remain at
drop. A wrong array payload fails before native entry. Calls use one-operation
slices on both runtimes. Scalar native BSTR out-cells are also released on failure.
An exact QI'd IDispatch owner spans object-to-VARIANT conversion; a real Dictionary
object round-trip retains its independent result owner and balances the fixture
references at drop. Optional native Long defaults and mixed zero/one-argument
calls to the same retained member produce True/False without Invoke or name reads.

Fresh review caught an actual array-output binder gap: semantic Variant alone
discarded the ByRef SAFEARRAY wire direction. The binder now uses the wire shape
for positional and named arguments; the test asserts replacement values 144/233,
rather than merely checking the unchanged array's tag.

The owned Excel receipt `calcvba10-native-excel-interfaces` passes on PID 53816,
then records owned-process-zero. Application scalar getters and its native
WorksheetFunction interface-result getter use native slots on both runtimes with
no Invoke delta. Two earlier attempts (PIDs 27992 and 22248) refused unsupported
pure dispinterface members and also record owned-process-zero; they earn no
passing runtime credit. Installed type information and Microsoft's published
[Range](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.range?view=excel-pia)
and [WorksheetFunction](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.worksheetfunction?view=excel-pia)
contracts agree that Value/Sum have no custom vtable entries. That subset stays
blocked under zero Invoke; see BLK-WORKSHEETFUNCTION-VTABLE-20261101. No native
Range array or WorksheetFunction Sum execution is claimed.

Records, custom IUnknown
interfaces, nominal interface arrays, synchronous callback/reentry visibility
during a native call, and full deployment/certification remain with
bd-59co.3.5.3 and the WIN-1/WIN-14 producers.

General owning `Is` expressions and native identity across separately prepared
wrappers remain outside this delivery. CompareObjectIs is not admitted by the
owning layout, and a portable factory header must not be mistaken for native
IUnknown identity. The next delivery under bd-59co.2.9.9/bd-59co.3.5.3 must admit
the original Is instruction, retain native identity at preparation, and check
Nothing/Variant/interface-result cases against VBA, with no default getter or
Invoke/name lookup during comparison. Microsoft's
[QueryInterface identity rules](https://learn.microsoft.com/en-us/windows/win32/com/rules-for-implementing-queryinterface)
define the native identity authority. The current ownership tests give no Is credit.
Opaque object-bearing native Variants currently retain ownership/round-trip only.
Recovery into a nominal source interface still needs a retained QI conversion plan
under the same owners; a portable header IID comparison is not VBA's QI conversion.

Final repository check: 3747 passed in 198 suites, 181 ignored; four registry-writing fixtures excluded without credit. Formatting, strict workspace/all-target
Clippy, governance and path stability pass. The allocation-corrected source also
passes the owned native replay calcvba10-native-publish (PID 19144, then
owned-process-zero). Optional replay calcvba10-optional-final owns/cleans PID 47652.
Source/check/oracle hashes are retained in [run10-native.json](run10-native.json)
and [run10-native.zip](run10-native.zip). Earlier failed/intermediate attempts earn
no passing credit. WCC-EARLY-COMPLEX stays planned with its residual owners.
