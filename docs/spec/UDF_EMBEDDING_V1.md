# Spreadsheet function embedding

DNA Calc's `crates/calc-vba` consumes the neutral `calcvba-host-api` and
`calcvba-host` interfaces described in [hosting](HOSTING_AND_EARLY_BINDING.md).
It supplies its own application objects, type library, formula conversions,
calculation context and scheduler. Its invocation slices, read views, budgets,
reloads and handle generations are active consumers of the engine API.

The older value/Range adapter remains in `adapters/spreadsheet-compat`.
See that adapter's README for its value conversions and limits. It is not an
engine-owned Excel object model. Adding a spreadsheet member belongs in the
host's glue/object library. Macro/document hosting uses the neutral project
and event interfaces, not a synthetic worksheet function protocol.
