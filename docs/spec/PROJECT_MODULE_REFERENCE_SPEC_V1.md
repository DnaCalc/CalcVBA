# Projects, modules and references

The compiler loads a project and its referenced-project closure once, preserving
reference order. It preprocesses modules for the selected target, parses them
with the shared VBA parser, builds declarations/providers and binds typed calls.
`.basproj` is the project format. A bounded legacy .vbp importer reads module
lists; it does not implement the VB6 project system or forms.

Procedural, class and document modules have distinct VBA roles. Public/Private,
Option Private Module, predeclared instances, attributes, property kinds and
ordered references affect name lookup. Public procedure signatures retain
parameter/result types, ByVal/ByRef, Optional defaults and ParamArray. Classes
retain fields, methods, interfaces and lifecycle procedures. Documents retain
their host-supplied type and predeclared identity.

Providers supply referenced source/export surfaces, VBA library declarations,
host object libraries and imported COM metadata. The binder resolves a name
once and carries its selected declaration and call shape into Core IR/OxIR.
Runtime helpers must not replace that decision with dependency-specific names.
Genuinely dynamic Object/Variant calls retain VBA's dynamic dispatch behaviour.

VM3 and the JIT link the same closure. Each project has its own mutable global
store; shared references in a diamond use the same referenced project instance.
Compiled images retain the closure and imported descriptors. Full cross-project
field/type/visibility behaviour is still incomplete; the project, binder and
VM3/JIT tests define the supported cases.
