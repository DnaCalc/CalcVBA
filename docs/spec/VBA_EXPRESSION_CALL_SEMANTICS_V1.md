# VBA expression and call semantics

Language reference for in-scope VBA behaviour. Public VBA specifications and
Excel/VBA observations decide uncertain cases. See [architecture](../ARCHITECTURE.md)
for current implementation limits; this reference does not claim complete support.

## Expression Classification Model

Every relevant expression must retain the following semantic facts; this shape
is illustrative:

```rust
pub struct ExpressionSemanticsDescriptor {
    pub expression_id: ExpressionId,
    pub source_span: Option<SourceSpanId>,
    pub classification: ExpressionClassification,
    pub declared_type: VbaTypeId,
    pub value_state_space: ValueStateSpace,
    pub carrier_hint: RuntimeCarrierKind,
    pub default_member_policy: DefaultMemberPolicy,
}

pub enum ExpressionClassification {
    Value,
    Variable,
    Property,
    Function,
    Subroutine,
    UnboundMember,
    ObjectMember,
    ArrayElement,
    UdtField,
    WithExpression,
    AddressOf,
}
```

Classification matters for assignment and calls. In particular, a ByRef
argument that is a variable has different aliasing behavior from a ByRef
argument whose expression is a value, function, property, or unbound member.
Object expressions can also trigger default-member binding before ordinary
Let-coercion.

## Coercion Model

OxVba needs explicit descriptors for both coercion families:

```rust
pub struct CoercionDescriptor {
    pub source_declared_type: VbaTypeId,
    pub source_value_state_space: ValueStateSpace,
    pub target_declared_type: VbaTypeId,
    pub kind: CoercionKind,
    pub static_status: CoercionStaticStatus,
    pub runtime_failure: RuntimeFailurePolicy,
    pub source_anchor: Option<SpecAnchorId>,
}

pub enum CoercionKind {
    Let,
    Set,
    ExplicitConversionIntrinsic,
    BoundaryProjection,
}

pub enum CoercionStaticStatus {
    Valid,
    Invalid,
    RequiresRuntimeCheck,
    AuthorityPending,
}
```

Let-coercion covers ordinary value assignment and most non-object parameter
binding. It must include numeric, Boolean, Date, String, fixed-length String,
Byte array, non-Byte array, UDT, Error/CVErr, Null, Empty, Variant, class,
Object, and Nothing cases.

Set-coercion covers object-reference assignment. It must not be collapsed into
Let-coercion. `Set` assignment, `Property Set`, object parameters, `Nothing`,
class/interface conformance, and COM object projection all depend on this
separation.

The semantic facts must preserve whether coercion was checked statically,
requires a runtime helper, or still awaits public-specification or Excel/VBA
adjudication. `AuthorityPending` cannot enter verified artifacts or executable
backend admission. A typed JIT fast path may inline only authority-resolved rows
whose coercion behavior is proved or guarded.

## Operator Semantics

```rust
pub struct OperatorSemanticsDescriptor {
    pub operator: VbaOperator,
    pub left_declared_type: Option<VbaTypeId>,
    pub right_declared_type: Option<VbaTypeId>,
    pub result_declared_type: VbaTypeId,
    pub result_value_state_space: ValueStateSpace,
    pub coercion_policy: OperatorCoercionPolicy,
    pub error_policy: RuntimeFailurePolicy,
    pub compare_mode: Option<ModuleCompareMode>,
    pub helper: Option<HelperAbiId>,
}
```

Required operator families:

The `+` operator is a named risk surface because it can mean addition or string
concatenation depending on operand value types. The `&` operator is the forced
string-concatenation path. Relational string comparisons must preserve the
module `Option Compare` mode. `Is` is object identity, not value equality.

An operator fast path is admissible only when its semantic row is authoritative
and verified. VM3/JIT fixtures test that shared meaning against public
specification or Excel/VBA evidence; current interpreter behavior cannot
self-authorize the expected result.

## Assignment And Property Semantics

Let assignment:

- may omit the `Let` keyword in source;
- uses Let-coercion from the expression value to the target declared type;
- handles assignment into variables, properties, array elements, UDT fields,
  and fixed-length strings;
- must preserve Null, Empty, Error/CVErr, and default-member behavior.

Set assignment:

- requires the `Set` keyword;
- uses Set-coercion to assign object references;
- handles `Nothing`, class/interface compatibility, `Object`, imported COM
  object/interface descriptors, and object lifetime.

Property descriptors must preserve:

- `Property Get`, `Property Let`, and `Property Set` kind;
- shared property name pairing rules;
- equivalent property parameter lists;
- value parameter declared type;
- the rule that the `Property Let` value type matches the corresponding
  `Property Get` return type;
- the rule that `Property Set` value type is Object, Variant, or a named class;
- the rule that a property-LHS value parameter has runtime ByVal semantics even
  when its source mechanism is omitted or written as ByRef.

Default-member binding is a distinct semantic fact. A Let-style context may need
to bind through an object's default value member, while a Set-style context
assigns the object reference itself.

## Call Binding Semantics

Every call site should lower to a call descriptor:

```rust
pub struct CallSiteDescriptor {
    pub call_site_id: CallSiteId,
    pub target: CallableTargetDescriptor,
    pub specified_arguments: Vec<ArgumentDescriptor>,
    pub resolved_bindings: Vec<ArgumentBindingDescriptor>,
    pub param_array_binding: Option<ParamArrayBindingDescriptor>,
    pub invocation_kind: InvocationKind,
    pub error_policy: RuntimeFailurePolicy,
}

pub struct ArgumentBindingDescriptor {
    pub parameter_index: usize,
    pub argument_index: Option<usize>,
    pub parameter_slot: Option<usize>,
    pub mechanism: ParameterMechanism,
    pub argument_classification: Option<ExpressionClassification>,
    pub declared_type_check: CallTypeCompatibility,
    pub coercion: Option<CoercionDescriptor>,
    pub storage: ArgumentStorageKind,
}

pub enum ArgumentStorageKind {
    DirectByRefAlias { source_slot: usize },
    LocalByValCopy,
    LocalByRefExpressionTemp,
    OptionalDefaultLocal,
    OptionalMissingState,
    ParamArrayElement,
    BoundaryProjectedTemp,
}
```

Static call compatibility must preserve the MS-VBAL mapping sequence:

- positional arguments map left-to-right;
- named arguments map by case-insensitive name value;
- duplicate mappings are incompatible;
- too many positional arguments are incompatible unless the last parameter is a
  `ParamArray`;
- any non-optional parameter without a mapped argument is incompatible;
- omitted positional values mapped to non-optional parameters are incompatible;
- ByVal non-object arguments require valid Let-coercion to the parameter type;
- ByVal object arguments must be Object, class, or Variant-compatible;
- ByRef non-object, non-Variant arguments require exact declared type match;
- ByRef Variant parameters require the MS-VBAL Variant-specific binding rule
  and must not be silently narrowed to a primitive exact-match rule;
- ByRef object arguments require object/class declared compatibility.

Runtime call binding must preserve:

- runtime error 450 for wrong argument count or invalid property assignment,
  except when a `ParamArray` consumes extra positional arguments;
- runtime error 448 for unknown named arguments, duplicate mappings, and
  omitted positional values mapped to non-optional parameters;
- runtime error 449 for unmapped non-optional parameters;
- ByVal parameters as procedure-local variables initialized from the mapped
  argument after the required coercion;
- ByRef variable arguments as aliases to the caller variable;
- ByRef value/function/property/unbound arguments as procedure-local temporary
  variables where the invocation rules allow that binding;
- unmapped optional parameters as procedure-local variables initialized from
  their default value, or as the explicit missing/default state required for
  Variant optional parameters;
- `ParamArray` as a new Variant array with lower bound 0, containing extra
  positional arguments, or an empty lower-0/upper--1 array when no extras are
  supplied.

COM and native Declare calls reuse this descriptor shape, but add boundary
projection and writeback descriptors. A COM/natively projected ByRef is not a
reason to weaken core VBA ByRef semantics; it is an external ABI projection
from the semantic binding.

## Method And Event Invocation

Method calls include an implicit current-object parameter:

- the implicit current object is ByVal;
- it has procedure extent;
- its declared type is the class containing the method;
- it is exposed through `Me`;
- it is not assignable by user code.

Events and event handlers must use the same signature compatibility model as
procedure invocation. Event declaration parameters describe the arguments
required for `RaiseEvent`; they are not ordinary variable bindings by
themselves.

Default function/subroutine invocation on objects must be represented as a
resolved call target, not as an untracked runtime fallback. Late-bound COM
default members must carry the same named/default argument facts used by the
COM descriptor.
