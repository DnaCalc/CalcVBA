# Pointer helpers

`VarPtr`, `StrPtr` and `ObjPtr` return what VBA returns: the real address of the operand's
storage, the BSTR it holds, or the interface pointer it holds. Writes through those addresses
(`CopyMemory`, a native callee) reach the variable, and later reads see them, on VM3 and the
JIT alike. The Excel x64 facts are in `docs/design/runtime-variant-automation-layout-2026-11-13.md`
(section 2 and step 7), and `crates/calcvba-engine/tests/real_addresses.rs` checks them on
both backends.

| Operand | `VarPtr` | `StrPtr` |
|---|---|---|
| Variant variable, Variant element or field | the `VARIANT` cell | the BSTR at +8 when it holds a String, else a statement-temporary conversion |
| Typed scalar, String or object variable | the cell's payload at +8 (the value, the BSTR slot, the interface pointer slot) | the BSTR (String) |
| UDT variable | the record's memory | — |
| Array element `a(i, j)` | `pvData` + offset, first dimension fastest (error 9 out of range) | the element's BSTR |
| UDT field `t.f`, `a(i).f` | the field inside the record | the field's BSTR |
| Class field (inside the class) | the field's cell, stable for the instance's life | as a variable |
| Expression `VarPtr(n + 1)`, `StrPtr("lit")` | a statement temporary | the temporary's BSTR |

`ObjPtr` returns the object's interface pointer; for a foreign COM object held through the
binding table it is that object's native dispatch pointer (step 8 makes cells hold interface
pointers directly).

One runtime walker, `calcvba_runtime::address::locate`, finds the storage from the engine cell
of a local, global, temporary or class field, so both backends give the same address. The cells
it starts from never move during their variable's life: frame locals and temporaries are
pre-sized vectors, globals fixed vectors, class fields boxed cells kept until the instance is
freed. A `LongPtr` Declare argument is passed as the integer itself; there is no pointer
registry and no write-back after the call.

Known differences, each in a later step of the layout plan: `vbNullString` is still folded to
`""`, so `StrPtr(vbNullString)` is not 0 (step 9); a typed variable bound to a ByRef Variant
parameter is still a copy rather than a `VT_BYREF` cell (step 7b); `ReDim` does not yet keep the
descriptor address (step 7c). The AOT native lane addresses variables and temporaries only; an
element or field operand is refused there.
