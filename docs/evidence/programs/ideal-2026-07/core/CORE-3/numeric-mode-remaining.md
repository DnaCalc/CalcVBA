# NumericMode and object default values

2026-10-31. Broad NumericMode/coercion delivery remains **in-progress** under
`bd-59co.2.4.3`. Run-7 leaf `bd-59co.2.4.9` covers the bounded numeric arithmetic
default-value behavior observed in run 6; it does not complete this family.

The [idle Excel readings](../../windows-x64/WIN-4/worksheet-errors.md) retain a
Range argument as an object while `+1` and `*2` read its default value. Numeric 41
produces 42/82; confirmed text "3" and blank produce 4/1. Bad text and multiple-cell
arithmetic yield #VALUE! after Excel maps the unhandled VBA error.

VM3 and JIT now resolve object defaults through the existing project/host member
dispatch before the pure numeric arithmetic kernel. The original parameter/local
is preserved. No Excel/Range name or host-entry scalar conversion was added to
the engine. The host supplies the real default member and maps runtime faults to
worksheet errors; returned CVErr values remain distinct.

Checks: `cargo test -p oxvba-embed --test object_default_arithmetic` covers exact
numeric/text/blank values, bad text, multi-value defaults, Nothing, a host default
failure, error handlers, TypeName and object identity. The focused
`jit_project_objects::untyped_project_default_arithmetic_preserves_identity_and_caller_frame`
test proves the project-class route on both backends, including carrier balance.
The live owned-Excel integration test additionally exercises the real Range path
when its supervisor environment is supplied; ordinary ignored runs earn no credit.
See [run-7 evidence](../CORE-7/typed-host-callbacks.md) for execution receipts.

Remaining accepted scope belongs to `bd-59co.2.4.3`: other context-specific default
coercions (including general comparison/concatenation), exact recursive/deep default
behavior beyond the declared 64-resolution limit, the broader NumericMode families,
and Double Resume Next fault/store ordering. Excel observes error 6 with stored
+infinity (7FF0000000000000); that separate gap is not claimed by this delivery.
