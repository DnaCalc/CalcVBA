# VBA type semantics

Language reference for in-scope VBA behaviour. Public VBA specifications and
Excel/VBA observations decide uncertain cases. See [architecture](../ARCHITECTURE.md)
for current implementation limits; this reference does not claim complete support.

## Declared Type Categories

### Scalar Declared Types

Supported VBA scalar declared types:

- `Boolean`
- `Byte`
- `Integer`
- `Long`
- `LongLong`
- `LongPtr`
- `Single`
- `Double`
- `Currency`
- `Date`
- `String`
- `Variant`

`LongLong` and `LongPtr` are VBA7/platform-sensitive. `LongPtr` is a source
semantic type whose runtime carrier depends on target pointer width.

Default initialization:

- numeric scalars: zero value of the declared type;
- `Boolean`: `False`;
- variable-length `String`: empty string;
- object/class/interface variables: `Nothing`;
- `Variant`: `Empty`;
- dynamic arrays: unallocated array descriptor state;
- fixed arrays and UDT fields: recursively initialized element/field defaults.

### Decimal

`Decimal` is not an ordinary declared variable type in VBA. It is a Variant
subtype produced by Variant-valued expressions such as `CDec(...)` and by COM
or runtime boundaries that yield `VT_DECIMAL`.

OxVba rules:

### String Types

String descriptors must distinguish:

- variable-length `String`, represented internally by `BStr`;
- fixed-length `String * N`, including fixed strings in UDT fields;
- `vbNullString` as a value expression/projection case, not a distinct declared
  type.

Semantic and verified-artifact facts must record fixed length where present, default
initialization, assignment truncation/padding behavior, and cleanup ownership.

### Arrays

Array descriptors must record:

- dynamic versus fixed/static array;
- rank;
- declared bounds per dimension where known;
- `Option Base` influence for omitted lower bounds;
- element declared type;
- element runtime carrier;
- resize/`ReDim`/`Preserve` legality;
- array-as-parameter shape and ByRef writeback requirements;
- fixed array field shape inside UDTs.

Arrays are declared type constructors. `SafeArray` is the runtime/interop
carrier used for array values and COM projection.

### Array Bounds And Shape

Array bounds are shape metadata, not the element declared type. Compiler-owned
semantic facts and the verified artifact
must preserve both:

- the element declared type and carrier;
- the rank, lower/upper bound, and provenance of each dimension.

Static/fixed arrays and fixed array fields in UDTs carry their shape in the
declaration/storage descriptor. Dynamic arrays carry their element type in the
declared type, while their current shape is runtime allocation state produced
by `ReDim`, array literal/runtime helpers, ParamArray binding, or COM SAFEARRAY
projection.

Bounds rules:

- explicit `lower To upper` clauses win over any module default;
- omitted lower bounds use the applicable `Option Base` default at compile
  time, and the descriptor must record that provenance;
- `ReDim` changes the runtime allocation shape of a resizable array, not its
  element declared type;
- `ReDim Preserve` cannot change rank or element type and can resize only the
  last dimension under VBA-compatible rules;
- `LBound` and `UBound` read runtime shape, and dimension arguments are
  one-based;
- `ParamArray` binding creates a Variant array with lower bound 0, including
  the empty lower-0/upper--1 shape when there are no extra arguments;
- `Array(...)` function lower-bound behavior must be represented by
  `ArrayFunction` provenance and closed against MS-VBAL plus Office oracle
  evidence before JIT specialization relies on it;
- COM SAFEARRAY projection must preserve the lower bound and element type
  metadata per dimension exactly at the boundary.

Illustrative semantic fact shape (non-normative representation):

```rust
pub enum ArrayStorageKind {
    Dynamic,
    Fixed { shape: ArrayShapeDescriptor },
    StaticLocal { shape: ArrayShapeDescriptor },
    UdtFieldFixed { shape: ArrayShapeDescriptor },
    ParamArray,
    ComSafeArrayProjection,
}

pub struct ArrayShapeDescriptor {
    pub rank: usize,
    pub dimensions: Vec<ArrayDimensionBounds>,
    pub provenance: ArrayShapeProvenance,
}

pub struct ArrayDimensionBounds {
    pub lower: i32,
    pub upper: i32,
}

pub enum ArrayShapeProvenance {
    ExplicitToClause,
    OptionBaseDefault,
    RuntimeReDim,
    ReDimPreserve,
    ArrayFunction,
    ParamArrayBinding,
    ComSafeArray,
}
```

### UDTs

UDTs are nominal declared types, not structural aliases. Descriptors must
record:

- `UdtTypeId`;
- project/module/name identity and visibility;
- ordered fields;
- each field's declared type id;
- fixed-string lengths;
- fixed-array field bounds and element type;
- nested UDT field references;
- initialization, copy, assignment, and cleanup rules;
- native ABI materialization policy, explicitly separate from internal
  descriptor-backed semantics.

Internal storage is a backend decision, but it must preserve nominal UDT
semantics and ownership. JIT/native lowering must not treat a UDT as a platform
struct unless a separately verified ABI descriptor proves the layout.

### Enums

Enums are nominal declared types with `Long`-compatible values. Descriptors must
record:

- `EnumTypeId`;
- project/module/name identity and visibility;
- ordered member names and constant values;
- underlying carrier as `Long`;
- assignment/coercion behavior where VBA treats enum values as numeric.

### Object, Class, Interface, And Imported COM Types

Object-like declared types are central, not edge cases.

Descriptor forms:

- `Object`: late-bound object reference.
- `Class { class_id }`: VBA class module object type.
- `Interface { interface_id }`: interface implemented by a class or imported
  from type information.
- `ComClass { typelib_id, coclass_id }`: imported COM coclass.
- `ComInterface { typelib_id, interface_id, dispatch_kind }`: imported COM
  interface, dispatch interface, or dual interface.
- `WithEventsObject { object_type, event_source_id }`: object variable that
  participates in event binding.
- `AsNewObject { object_type, activation_policy }`: object variable with
  automatic instantiation semantics.

Required object metadata:

- object identity carrier: `ObjectRef`;
- default interface and default member metadata;
- `IDispatch`/vtable strategy for imported COM types;
- `Implements` interface mapping from class member implementations to interface
  members;
- `WithEvents` variable identity and event-handler prefix binding;
- `RaiseEvent` and event dispatch signatures;
- `Set` assignment compatibility, `Nothing`, and object lifetime rules;
- class initialization/termination hooks where in scope;
- default instance behavior where in scope.

`Nothing` is a value state for object variables, not a declared type.

## Value States That Are Not Declared Types

These must be represented explicitly as runtime value states, expression states,
or call argument states, not ordinary declared types:

- `Empty`: default value of a `Variant`.
- `Null`: Variant/database null value state.
- `Error`/`CVErr`: Variant error value state.
- `Nothing`: object reference value state.
- missing optional argument: call argument state for `Optional` parameters.
- omitted/default parameter value: call binding state, with optional default
  value metadata.
- `vbNullString`: string value/projection state, not a declared type.

JIT and VM3 snapshots may observe these states through retained `Variant`,
`ObjectRef`, or call-frame metadata, but they must not appear as declared slot
types.

## Procedure Type Descriptors

Procedure semantic facts must capture the full source signature and resolved
calling shape. This is required for VM3/JIT execution, COM projection, native
`Declare`, and exported callables.

Procedure descriptors must record:

- procedure kind: `Sub`, `Function`, `Property Get`, `Property Let`,
  `Property Set`, event declaration, event handler, external Declare, exported
  callable;
- module/class/interface ownership;
- visibility and module kind;
- procedure name and property/event pairing identity;
- parameter list in source order;
- parameter names and case-insensitive name identity;
- parameter role: positional, optional, ParamArray, property value parameter,
  implicit current object, event signature parameter;
- source parameter mechanism: explicit `ByRef`, explicit `ByVal`, or omitted;
- resolved runtime mechanism, including the default `ByRef` rule and the
  special property-LHS value-parameter ByVal semantics;
- `Optional` and default-value expression/data value;
- missing optional argument state where distinct from default assignment;
- `ParamArray` element type, Variant-array shape, and binding rules;
- declared parameter type id;
- parameter array designator and resizable-array parameter type where present;
- return type id where applicable;
- property accessor pairing rules and property value type;
- event signature compatibility and handler binding;
- implicit method current-object (`Me`) descriptor;
- external Declare ABI surface where applicable;
- source span and resolved executable-procedure identity.

Illustrative semantic fact shape (non-normative representation):

```rust
pub struct ProcedureSignatureDescriptor {
    pub procedure_id: ProcedureId,
    pub name: NameId,
    pub owner: ProcedureOwnerId,
    pub kind: ProcedureKind,
    pub visibility: ProcedureVisibility,
    pub parameters: Vec<ParameterDescriptor>,
    pub return_type: Option<VbaTypeId>,
    pub property_group: Option<PropertyGroupId>,
    pub event_signature: Option<EventSignatureDescriptor>,
    pub implicit_current_object: Option<ImplicitCurrentObjectDescriptor>,
    pub external_abi: Option<ExternalAbiDescriptorId>,
    pub source_span: Option<SourceSpanId>,
    pub executable_procedure: Option<ResolvedProcedureId>,
}

pub struct ParameterDescriptor {
    pub index: usize,
    pub name: NameId,
    pub role: ParameterRole,
    pub source_mechanism: SourceParameterMechanism,
    pub resolved_mechanism: ResolvedParameterMechanism,
    pub declared_type: VbaTypeId,
    pub array_parameter: Option<ArrayParameterDescriptor>,
    pub optional: OptionalParameterDescriptor,
    pub param_array: Option<ParamArrayDescriptor>,
    pub slot: Option<usize>,
}

pub enum ParameterRole {
    Positional,
    Optional,
    ParamArray,
    PropertyValue,
    ImplicitCurrentObject,
    EventSignatureOnly,
}

pub enum SourceParameterMechanism {
    Omitted,
    ExplicitByRef,
    ExplicitByVal,
    ImplementationInjected,
}

pub enum ResolvedParameterMechanism {
    ByRef,
    ByVal,
    PropertyValueByVal,
    EventSignatureOnly,
}

pub enum OptionalParameterDescriptor {
    Required,
    Optional {
        default_value: OptionalDefaultValue,
        missing_state: OptionalMissingStatePolicy,
    },
}

pub enum OptionalDefaultValue {
    Explicit(ConstantValueId),
    DeclaredTypeDefault,
    VariantMissingError448,
    AuthorityPending,
}

pub enum OptionalMissingStatePolicy {
    AssignDefaultLocal,
    PreserveMissingArgumentState,
}

pub struct ParamArrayDescriptor {
    pub element_type: VbaTypeId,
    pub array_lower_bound: i32,
    pub empty_upper_bound: i32,
}

pub struct ImplicitCurrentObjectDescriptor {
    pub declared_type: VbaTypeId,
    pub mechanism: ResolvedParameterMechanism,
    pub accessible_name: NameId,
    pub assignable: bool,
}
```

`AuthorityPending` is an editor/unverified-analysis state for a behavior still
awaiting public-specification or Excel/VBA adjudication. It cannot enter Core
IR, verified OxIR/OxImage, or executable backend admission.

`ByRef` compatibility is exact declared type compatibility for non-object,
non-Variant parameters unless a documented VBA rule permits a different
binding. `ByVal` may coerce according to assignment and call-context rules.
Call-site aliasing, temporary-local creation, optional/default assignment,
ParamArray construction, and boundary writeback are defined in
[`VBA_EXPRESSION_CALL_SEMANTICS_V1.md`](VBA_EXPRESSION_CALL_SEMANTICS_V1.md).

## Interop Type Projection

Interop descriptors project from semantic types:

- COM projection: `VARIANT`, `BSTR`, `SAFEARRAY`, `IDispatch`, `IUnknown`,
  vtable interface pointer, `EXCEPINFO`, HRESULT/ArgErr behavior.
- Native `Declare` projection: platform calling convention, pointer width,
  string/array/variant marshaling, ByRef writeback, return value projection.
- Exported callable projection: inbound ABI shape, slot population, ByRef
  writeback, cleanup, and error return policy.

COM and native projections consume verified semantic descriptors. Generated
code and runtime helpers must not discover these shapes through ambient symbols
or side channels; concrete descriptor representation belongs to the artifact
and Windows subsystem contracts.
