//! Host contracts only. Supply an implementation from the application's runtime
//! composition; this crate links neither compiler, interpreter nor JIT.
mod project_graph;
pub use project_graph::*;
mod native_callback;
pub use native_callback::{
    NativeCallbackSync, with_native_call, with_native_callback_context, with_native_callback_sync,
};
mod invocation;
pub mod object;
pub use invocation::*;
mod scalar;
pub mod value;
pub use object::{HostObject, HostSlot, InterfaceId, ObjectArg};
pub use oxvba_runtime::Variant;
pub use scalar::*;
use std::{any::Any, marker::PhantomData, rc::Rc};
pub use value::*;

pub const HOST_API_ABI: u32 = 2;
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Backend {
    Interpreter,
    Jit,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProcedureKind {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Visibility {
    Public,
    Private,
    Friend,
}
/// Compiler signature projection. Nominal identities and known ranks are kept.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Type {
    Variant,
    Boolean,
    Byte,
    Integer,
    Long,
    LongLong,
    Single,
    Double,
    Currency,
    Date,
    Decimal,
    String,
    FixedString(u32),
    Object {
        name: Option<String>,
        iid: Option<InterfaceId>,
    },
    Array {
        element: Box<Type>,
        rank: Option<usize>,
    },
    Record(String),
    Procedure,
}
#[derive(Debug, Clone, PartialEq)]
pub struct Parameter {
    pub name: String,
    pub ty: Type,
    pub by_ref: bool,
    pub optional: bool,
    pub param_array: bool,
    pub default: Option<Variant>,
    pub array_element: Option<ArrayElement>,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ArrayElement {
    Scalar(Type),
    Record(Vec<ArrayElement>),
    FixedArray {
        element: Box<ArrayElement>,
        bounds: Vec<(i32, usize)>,
    },
}
#[derive(Debug, Clone, PartialEq)]
pub struct Procedure {
    pub project: String,
    pub module: String,
    pub name: String,
    pub kind: ProcedureKind,
    pub visibility: Visibility,
    pub option_private_module: bool,
    pub parameters: Vec<Parameter>,
    pub return_type: Option<Type>,
}

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("host API ABI mismatch: requested {requested}, provider {provided}")]
    AbiMismatch { requested: u32, provided: u32 },
    #[error("compile: {0}")]
    Compile(String),
    #[error("unknown or ambiguous public procedure: {0}")]
    UnknownProcedure(String),
    #[error("unsupported signature: {0}")]
    UnsupportedSignature(String),
    #[error("stale procedure handle or handle from another session")]
    StaleHandle,
    #[error("invocation has returned, faulted, or been cancelled")]
    InvocationFinished,
    #[error("procedure argument count/type mismatch")]
    ArgumentMismatch,
    #[error("execution budget exhausted")]
    BudgetExceeded,
    #[error("session reentry requires the active typed context and a supported typed entry")]
    Reentry,
    #[error("backend is unavailable")]
    BackendUnavailable,
    #[error("backend does not support this module: {0}")]
    UnsupportedBackend(String),
    #[error("backend: {0}")]
    Backend(String),
    #[error("VBA error {number}: {description}")]
    Runtime {
        number: i32,
        description: String,
        vba_source: String,
    },
}

/// Opaque provider payload, with no exposed code pointer. Providers must validate
/// owner and generation on every call, including handles created by other providers.
#[derive(Clone)]
pub struct BoundProcedure {
    payload: Rc<dyn Any>,
    generation: u64,
}
impl BoundProcedure {
    pub fn generation(&self) -> u64 {
        self.generation
    }
    /// Provider construction; wrapping a payload grants no executable authority.
    pub fn from_provider<T: Any>(payload: T, generation: u64) -> Self {
        Self {
            payload: Rc::new(payload),
            generation,
        }
    }
    pub fn provider_payload<T: Any>(&self) -> Option<&T> {
        self.payload.downcast_ref()
    }
}
pub struct TypedProcedure<A, R> {
    bound: BoundProcedure,
    signature: PhantomData<fn() -> (A, R)>,
}
impl<A, R> TypedProcedure<A, R> {
    pub fn generation(&self) -> u64 {
        self.bound.generation()
    }
}

#[derive(Clone)]
pub struct SourceModule {
    pub name: String,
    pub source: String,
}
#[derive(Clone)]
pub struct Project {
    pub name: String,
    pub modules: Vec<SourceModule>,
    pub libraries: Vec<ObjectLibrary>,
    pub references: Vec<TypeLibraryReference>,
}
impl Project {
    pub fn module(name: impl Into<String>, source: impl Into<String>) -> Self {
        Self {
            name: "HostProject".into(),
            modules: vec![SourceModule {
                name: name.into(),
                source: source.into(),
            }],
            libraries: vec![],
            references: vec![],
        }
    }
}
/// Ordered external references are resolved by the application-supplied runtime
/// composition. Type-library parsing and native COM policy stay with the provider.
#[derive(Clone)]
pub struct TypeLibraryReference {
    pub name: String,
    pub libid: Option<String>,
    pub major: Option<u16>,
    pub minor: Option<u16>,
    pub lcid: Option<u32>,
    pub import_library: Option<String>,
}
impl TypeLibraryReference {
    pub fn named(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            libid: None,
            major: None,
            minor: None,
            lcid: None,
            import_library: None,
        }
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MemberKind {
    Method,
    PropertyGet,
}
#[derive(Clone)]
pub struct ObjectMember {
    pub name: String,
    pub slot: usize,
    pub kind: MemberKind,
    pub parameters: Vec<ScalarType>,
    pub result: ScalarType,
    pub is_default: bool,
}
#[derive(Clone)]
pub struct ObjectInterface {
    pub name: String,
    pub iid: InterfaceId,
    pub members: Vec<ObjectMember>,
}
#[derive(Clone)]
pub struct ObjectLibrary {
    pub name: String,
    pub interfaces: Vec<ObjectInterface>,
}

#[derive(Debug, Clone, Copy)]
pub struct ProviderInfo {
    pub host_api_abi: u32,
    pub typed_entry_abi: u32,
    pub jit_available: bool,
    /// This version is a linked Rust API; each owner frees its own allocations.
    pub allocator: &'static str,
}
pub trait RuntimeProvider {
    /// Compile an explicit, ordered closure of procedural source projects.
    fn load_graph(
        &self,
        _graph: &ProjectGraph,
        _libraries: &[TypedObjectLibrary],
        _backend: Backend,
        _requested_abi: u32,
    ) -> Result<Box<dyn Session>, Error> {
        Err(Error::UnsupportedSignature(
            "provider has no source-project graph API".into(),
        ))
    }
    /// Additive value-slot capability; the original scalar API remains ABI 2.
    fn value_abi(&self) -> Option<u32> {
        None
    }
    fn load_with_object_libraries(
        &self,
        project: &Project,
        libraries: &[TypedObjectLibrary],
        backend: Backend,
        requested_abi: u32,
    ) -> Result<Box<dyn Session>, Error> {
        if libraries.is_empty() {
            self.load(project, backend, requested_abi)
        } else {
            Err(Error::UnsupportedSignature(
                "provider has no extended typed object-library API".into(),
            ))
        }
    }
    fn info(&self) -> ProviderInfo;
    fn load(
        &self,
        project: &Project,
        backend: Backend,
        requested_abi: u32,
    ) -> Result<Box<dyn Session>, Error>;
}
/// Same-thread session owner. Rust object borrows and opaque call handles never
/// authorize concurrent access. Reload prepares replacement before invalidation.
pub trait Session {
    fn reload_graph(&mut self, _graph: &ProjectGraph) -> Result<(), Error> {
        Err(Error::UnsupportedSignature(
            "provider has no source-project graph API".into(),
        ))
    }
    fn bind_values(
        &self,
        _module: &str,
        _name: &str,
        _arguments: &[ValueType],
        _result: ValueType,
    ) -> Result<BoundProcedure, Error> {
        Err(Error::UnsupportedSignature(
            "provider has no extended value-slot API".into(),
        ))
    }
    fn type_bound_values(
        &self,
        _handle: &BoundProcedure,
        _arguments: &[ValueType],
        _result: ValueType,
    ) -> Result<BoundProcedure, Error> {
        Err(Error::UnsupportedSignature(
            "provider has no extended value-slot API".into(),
        ))
    }
    fn call_values(
        &self,
        _handle: &BoundProcedure,
        _args: &[ValueSlot<'_>],
        _context: &mut TypedContext,
    ) -> Result<Value, Error> {
        Err(Error::UnsupportedSignature(
            "provider has no extended value-slot API".into(),
        ))
    }
    fn begin_values(
        &self,
        _handle: &BoundProcedure,
        _args: &[ValueSlot<'_>],
        _frame: InvocationFrame,
        _max_blocks: u64,
    ) -> Result<Box<dyn ValueInvocation>, Error> {
        Err(Error::UnsupportedSignature(
            "provider has no extended value-slot API".into(),
        ))
    }
    fn procedures(&self) -> &[Procedure];
    fn bind(&self, module: &str, name: &str, kind: ProcedureKind) -> Result<BoundProcedure, Error>;
    fn call_bound(
        &self,
        handle: &BoundProcedure,
        args: &[Variant],
        max_steps: u64,
    ) -> Result<Variant, Error>;
    fn bind_slots(
        &self,
        module: &str,
        name: &str,
        arguments: &[ScalarType],
        result: ScalarType,
    ) -> Result<BoundProcedure, Error>;
    fn type_bound_slots(
        &self,
        handle: &BoundProcedure,
        arguments: &[ScalarType],
        result: ScalarType,
    ) -> Result<BoundProcedure, Error>;
    fn call_slots(
        &self,
        handle: &BoundProcedure,
        args: &[Slot<'_>],
        context: &mut TypedContext,
    ) -> Result<ScalarBits, Error>;
    /// Copy/retain arguments into an owning frame bound to this code/state
    /// generation. A later reload cannot retarget this activation.
    fn begin_invocation(
        &self,
        handle: &BoundProcedure,
        args: &[Slot<'_>],
        frame: InvocationFrame,
        max_blocks: u64,
    ) -> Result<Box<dyn Invocation>, Error>;
    fn reset(&self) -> Result<(), Error>;
    fn reload(&mut self, project: &Project) -> Result<(), Error>;
}
pub trait SessionExt: Session {
    fn bind_typed<A: Arguments, R: Scalar>(
        &self,
        module: &str,
        name: &str,
    ) -> Result<TypedProcedure<A, R>, Error> {
        self.bind_slots(module, name, &A::types(), R::TYPE)
            .map(|bound| TypedProcedure {
                bound,
                signature: PhantomData,
            })
    }
    fn bind_typed_handle<A: Arguments, R: Scalar>(
        &self,
        handle: &BoundProcedure,
    ) -> Result<TypedProcedure<A, R>, Error> {
        self.type_bound_slots(handle, &A::types(), R::TYPE)
            .map(|bound| TypedProcedure {
                bound,
                signature: PhantomData,
            })
    }
    fn call_typed<A: Arguments, R: Scalar>(
        &self,
        handle: &TypedProcedure<A, R>,
        args: A,
        blocks: u64,
    ) -> Result<R, Error> {
        self.call_typed_in_context(
            handle,
            args,
            &mut TypedContext {
                remaining_blocks: blocks,
            },
        )
    }
    /// Reborrow the context supplied to a host callback for bounded synchronous
    /// reentry. Nested blocks consume the outer call's remaining budget.
    fn call_typed_in_context<A: Arguments, R: Scalar>(
        &self,
        handle: &TypedProcedure<A, R>,
        args: A,
        context: &mut TypedContext,
    ) -> Result<R, Error> {
        let count = A::COUNT;
        let slots = args.slots();
        self.call_slots(&handle.bound, &slots[..count], context)
            .map(R::from_bits)
    }
}
impl<T: Session + ?Sized> SessionExt for T {}
