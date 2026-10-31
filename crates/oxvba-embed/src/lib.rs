//! Host-neutral procedure catalogue and persistent, generation-checked calls.
//! Spreadsheet adapters live outside this engine, in adapters/spreadsheet-compat.
use oxvba_bundle::{ArrayElementType, coreir::CoreConst};
use oxvba_hal::{HostServices, adapters::builder::HostBuilder};
use oxvba_oxir::{FuncId, OxProgram, OxTy};
use oxvba_runtime::Variant;
use oxvba_symbol::{TypeLibResolver, manifest as sym};
use std::{
    cell::{Cell, Ref, RefCell},
    marker::PhantomData,
    rc::Rc,
    sync::Arc,
};
mod cache;
use cache::PreparedCode;
pub use cache::{CacheStats, CodeCache, OwnedSession};
pub use oxvba_bundle::ProcedureKind;
pub use oxvba_rt_abi::host_entry::{Arguments, Scalar, ScalarType, TYPED_ENTRY_ABI};
pub use oxvba_symbol::model::Visibility;
pub mod provider;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Backend {
    Interpreter,
    Jit,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Parameter {
    pub name: String,
    pub ty: OxTy,
    pub array_element: Option<ArrayElementType>,
    pub by_ref: bool,
    pub optional: bool,
    pub param_array: bool,
    pub default: Option<CoreConst>,
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
    pub return_type: Option<OxTy>,
    id: FuncId,
}

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("compile: {0}")]
    Compile(String),
    #[error("unknown or ambiguous public procedure: {0}")]
    UnknownProcedure(String),
    #[error("unsupported signature: {0}")]
    UnsupportedSignature(String),
    #[error("stale procedure handle or handle from another session")]
    StaleHandle,
    #[error("procedure argument count/type mismatch")]
    ArgumentMismatch,
    #[error("execution budget exhausted")]
    BudgetExceeded,
    #[error("session is already executing")]
    Reentry,
    #[error("JIT backend is not enabled")]
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

#[derive(Clone)]
pub struct Module {
    program: OxProgram,
    procedures: Vec<Procedure>,
    host: Arc<dyn HostServices>,
}
impl Module {
    pub fn load(module_name: &str, source: &str) -> Result<Self, Error> {
        let manifest = sym::SymbolProjectManifest {
            project_name: "HostProject".into(),
            project_kind: sym::ProjectKind::Library,
            modules: vec![sym::ModuleUnit {
                module_name: module_name.into(),
                module_kind: sym::ModuleKind::Procedural,
                attributes: sym::ModuleAttributes::named(module_name),
                source: source.into(),
            }],
            references: vec![],
            reference_projects: vec![],
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
        };
        Self::load_project(&manifest, &oxvba_symbol::CatalogTypeLibResolver)
    }
    /// The host supplies ordered references, type libraries and module identities.
    /// No host library name, object model, ProgID or activation is invented here.
    pub fn load_project(
        manifest: &sym::SymbolProjectManifest,
        resolver: &dyn TypeLibResolver,
    ) -> Result<Self, Error> {
        let (core, catalogue) = oxvba_bind::bind_program_with_procedures(manifest, resolver)
            .map_err(|e| Error::Compile(e.to_string()))?;
        let mut program =
            oxvba_oxir::elaborate::elaborate(&core).map_err(|e| Error::Compile(e.to_string()))?;
        // A source-module host has explicit bound procedures, not an implicit
        // program run entry or diagnostic entry-local snapshot. Initialization
        // remains the separate compiler-owned global initializer.
        program.entry = None;
        oxvba_oxir::verify_program(&program)
            .map_err(|e| Error::Compile(format!("invalid program: {e:?}")))?;
        if !program.classes.is_empty() || !program.record_layouts.is_empty() {
            return Err(Error::UnsupportedSignature(
                "project classes/UDTs still require the full project-session owner".into(),
            ));
        }
        if program.funcs.iter().any(|f| {
            f.blocks
                .iter()
                .any(|b| matches!(b.terminator, oxvba_oxir::OxTerminator::Halt))
        }) {
            return Err(Error::UnsupportedSignature(
                "End requires whole-project reset semantics".into(),
            ));
        }
        let procedures = catalogue
            .into_iter()
            .map(|p| {
                let func = &program.funcs[p.index];
                Procedure {
                    project: manifest.project_name.clone(),
                    module: p.module,
                    name: p.name,
                    kind: func.kind,
                    visibility: p.visibility,
                    option_private_module: p.option_private_module,
                    parameters: func
                        .locals
                        .iter()
                        .take(func.param_count)
                        .zip(&core.procs[p.index].params)
                        .map(|(local, param)| Parameter {
                            name: local.name.clone(),
                            ty: local.ty.clone(),
                            array_element: local.array_element.clone(),
                            by_ref: param.by_ref,
                            optional: param.optional,
                            param_array: param.variadic,
                            default: param.default.clone(),
                        })
                        .collect(),
                    return_type: func.return_local.map(|id| func.locals[id.0].ty.clone()),
                    id: FuncId(p.index),
                }
            })
            .collect();
        Ok(Self {
            program,
            procedures,
            host: HostBuilder::new().build(),
        })
    }
    pub fn with_host(mut self, host: Arc<dyn HostServices>) -> Self {
        self.host = host;
        self
    }
    pub fn procedures(&self) -> impl ExactSizeIterator<Item = &Procedure> {
        self.procedures.iter()
    }
    pub fn com_interfaces(&self) -> &[oxvba_oxir::ComInterface] {
        &self.program.com_interfaces
    }
    pub fn prepare(&self, backend: Backend) -> Result<Session<'_>, Error> {
        Session::new(self, backend)
    }
    pub fn into_session(self, backend: Backend) -> Result<OwnedSession, Error> {
        OwnedSession::new(self, backend)
    }
}

/// Bind-once dynamic adapter. Arguments/results retain exact VBA carriers.
/// Host ByRef arguments are private copies; native ByRef host ABI remains open.
#[derive(Clone)]
pub struct BoundProcedure {
    owner: Rc<()>,
    generation: u64,
    id: FuncId,
    metadata: usize,
}
impl BoundProcedure {
    pub fn generation(&self) -> u64 {
        self.generation
    }
}

/// Contains no public address and cannot execute without a matching live session.
pub struct TypedProcedure<A, R> {
    bound: BoundProcedure,
    #[cfg(feature = "jit")]
    native: Option<oxvba_jit::TypedEntry<A, R>>,
    signature: PhantomData<fn() -> (A, R)>,
}

enum Executable<'a> {
    Interpreter(RefCell<oxvba_vm3::Vm3<'a>>),
    #[cfg(feature = "jit")]
    Jit(
        Rc<oxvba_jit::CompiledImage<'a>>,
        RefCell<oxvba_jit::JitSession<'a>>,
    ),
}

pub struct Session<'a> {
    module: &'a Module,
    host: &'a dyn HostServices,
    executable: Executable<'a>,
    typed: Vec<Result<oxvba_vm3::typed::TypedInterpreter<'a>, String>>,
    owner: Rc<()>,
    generation: u64,
    executing: RefCell<()>,
    typed_depth: Cell<u32>,
    typed_context: Cell<*mut calcvba_host_api::TypedContext>,
}
struct TypedActivation<'a> {
    _executing: Ref<'a, ()>,
    depth: &'a Cell<u32>,
    context: &'a Cell<*mut calcvba_host_api::TypedContext>,
}
thread_local! {
    // Bound the whole synchronous engine stack, including callbacks that cross
    // multiple independent sessions on the same host thread.
    static TYPED_CALL_DEPTH: Cell<u32> = const { Cell::new(0) };
}
impl Drop for TypedActivation<'_> {
    fn drop(&mut self) {
        TYPED_CALL_DEPTH.with(|depth| depth.set(depth.get() - 1));
        self.depth.set(self.depth.get() - 1);
        if self.depth.get() == 0 {
            self.context.set(std::ptr::null_mut());
        }
    }
}
const INITIALIZATION_BUDGET: u64 = 1_000_000;
impl<'a> Session<'a> {
    fn new(module: &'a Module, backend: Backend) -> Result<Self, Error> {
        PreparedCode::compile(module, backend, None)?.start(&*module.host)
    }
    fn from_code(
        module: &'a Module,
        host: &'a dyn HostServices,
        code: PreparedCode<'a>,
    ) -> Result<Self, Error> {
        let executable = match code {
            PreparedCode::Interpreter(_) => Executable::Interpreter(RefCell::new(
                oxvba_vm3::Vm3::link_with_budget(
                    &[&module.program],
                    host,
                    Some(INITIALIZATION_BUDGET),
                )
                .map_err(vm_error)?,
            )),
            #[cfg(feature = "jit")]
            PreparedCode::Jit(_, image) => {
                let mut state = image.start_session(host).map_err(jit_error)?;
                let result = image
                    .initialize_session(&mut state, INITIALIZATION_BUDGET)
                    .map_err(jit_error)?;
                if result.raised {
                    return Err(Error::Runtime {
                        number: result.err.number,
                        description: result.err.description,
                        vba_source: result.err.source,
                    });
                }
                Executable::Jit(image, RefCell::new(state))
            }
        };
        let typed = module
            .program
            .funcs
            .iter()
            .map(|func| {
                oxvba_vm3::typed::TypedInterpreter::bind_with_interfaces(
                    func,
                    &module.program.com_interfaces,
                )
            })
            .collect();
        Ok(Self {
            module,
            host,
            executable,
            typed,
            owner: Rc::new(()),
            generation: 1,
            executing: RefCell::new(()),
            typed_depth: Cell::new(0),
            typed_context: Cell::new(std::ptr::null_mut()),
        })
    }
    /// Reinitialize globals/Static/Err using retained code. Bound handles stay valid.
    pub fn reset(&self) -> Result<(), Error> {
        let _executing = self
            .executing
            .try_borrow_mut()
            .map_err(|_| Error::Reentry)?;
        match &self.executable {
            Executable::Interpreter(state) => {
                let next = oxvba_vm3::Vm3::link_with_budget(
                    &[&self.module.program],
                    self.host,
                    Some(INITIALIZATION_BUDGET),
                )
                .map_err(vm_error)?;
                *state.try_borrow_mut().map_err(|_| Error::Reentry)? = next;
            }
            #[cfg(feature = "jit")]
            Executable::Jit(image, state) => {
                let mut next = image.start_session(self.host).map_err(jit_error)?;
                let result = image
                    .initialize_session(&mut next, INITIALIZATION_BUDGET)
                    .map_err(jit_error)?;
                if result.raised {
                    return Err(Error::Runtime {
                        number: result.err.number,
                        description: result.err.description,
                        vba_source: result.err.source,
                    });
                }
                *state.try_borrow_mut().map_err(|_| Error::Reentry)? = next;
            }
        }
        Ok(())
    }
    pub fn backend(&self) -> Backend {
        match &self.executable {
            Executable::Interpreter(_) => Backend::Interpreter,
            #[cfg(feature = "jit")]
            Executable::Jit(..) => Backend::Jit,
        }
    }
    pub fn procedures(&self) -> impl ExactSizeIterator<Item = &Procedure> {
        self.module.procedures()
    }
    pub fn com_interfaces(&self) -> &[oxvba_oxir::ComInterface] {
        self.module.com_interfaces()
    }
    pub fn bind(
        &self,
        module: &str,
        name: &str,
        kind: ProcedureKind,
    ) -> Result<BoundProcedure, Error> {
        let module_key = oxvba_symbol::model::fold_identifier(module);
        let name_key = oxvba_symbol::model::fold_identifier(name);
        let mut candidates = self.module.procedures.iter().enumerate().filter(|(_, p)| {
            p.visibility == Visibility::Public
                && oxvba_symbol::model::fold_identifier(&p.module) == module_key
                && oxvba_symbol::model::fold_identifier(&p.name) == name_key
                && p.kind == kind
        });
        let (index, procedure) = candidates
            .next()
            .ok_or_else(|| Error::UnknownProcedure(format!("{module}.{name}")))?;
        if candidates.next().is_some() {
            return Err(Error::UnknownProcedure(format!("{module}.{name}")));
        }
        Ok(BoundProcedure {
            owner: self.owner.clone(),
            generation: self.generation,
            id: procedure.id,
            metadata: index,
        })
    }
    fn validate(&self, handle: &BoundProcedure) -> Result<(), Error> {
        if handle.generation != self.generation || !Rc::ptr_eq(&handle.owner, &self.owner) {
            return Err(Error::StaleHandle);
        }
        Ok(())
    }
    pub fn bind_typed<A: Arguments, R: Scalar>(
        &self,
        module: &str,
        name: &str,
    ) -> Result<TypedProcedure<A, R>, Error> {
        let bound = self.bind(module, name, ProcedureKind::Function)?;
        self.bind_typed_handle(&bound)
    }
    /// Add a checked native signature to an already resolved Function/Property Get.
    pub fn bind_typed_handle<A: Arguments, R: Scalar>(
        &self,
        bound: &BoundProcedure,
    ) -> Result<TypedProcedure<A, R>, Error> {
        self.validate(bound)?;
        let typed = self.typed[bound.id.0]
            .as_ref()
            .map_err(|e| Error::UnsupportedSignature(e.clone()))?;
        if typed.layout().parameters != A::types() || typed.layout().result != R::TYPE {
            return Err(Error::ArgumentMismatch);
        }
        #[cfg(feature = "jit")]
        let native = match &self.executable {
            Executable::Interpreter(_) => None,
            Executable::Jit(image, _) => {
                Some(image.bind_typed::<A, R>(bound.id).map_err(jit_error)?)
            }
        };
        Ok(TypedProcedure {
            bound: bound.clone(),
            #[cfg(feature = "jit")]
            native,
            signature: PhantomData,
        })
    }
    /// No name lookup, Variant argument vector, Variant frame or allocation.
    pub fn call_typed<A: Arguments, R: Scalar>(
        &self,
        handle: &TypedProcedure<A, R>,
        args: A,
        max_blocks: u64,
    ) -> Result<R, Error> {
        self.call_typed_in_context(
            handle,
            args,
            &mut calcvba_host_api::TypedContext {
                remaining_blocks: max_blocks,
            },
        )
    }
    pub fn call_typed_in_context<A: Arguments, R: Scalar>(
        &self,
        handle: &TypedProcedure<A, R>,
        args: A,
        context: &mut calcvba_host_api::TypedContext,
    ) -> Result<R, Error> {
        self.validate(&handle.bound)?;
        let _activation = self.enter_typed(context)?;
        let result = if !args.is_valid() {
            Err(6)
        } else {
            match &self.executable {
                Executable::Interpreter(_) => {
                    let typed = self.typed[handle.bound.id.0]
                        .as_ref()
                        .expect("bound typed interpreter");
                    let slots = args.slots();
                    typed
                        .invoke_slots(&slots[..A::COUNT], context)
                        .map(R::from_bits)
                }
                #[cfg(feature = "jit")]
                Executable::Jit(image, _) => handle
                    .native
                    .as_ref()
                    .expect("bound native entry")
                    .invoke_in_context(image, args, context),
            }
        };
        result.map_err(|number| self.typed_error(number))
    }
    fn enter_typed(
        &self,
        context: &mut calcvba_host_api::TypedContext,
    ) -> Result<TypedActivation<'_>, Error> {
        let executing = self.executing.try_borrow().map_err(|_| Error::Reentry)?;
        let pointer = context as *mut _;
        if self.typed_depth.get() != 0 && self.typed_context.get() != pointer {
            return Err(Error::Reentry);
        }
        if !TYPED_CALL_DEPTH.with(|depth| {
            if depth.get() >= 64 {
                false
            } else {
                depth.set(depth.get() + 1);
                true
            }
        }) {
            return Err(self.typed_error(28));
        }
        self.typed_context.set(pointer);
        self.typed_depth.set(self.typed_depth.get() + 1);
        Ok(TypedActivation {
            _executing: executing,
            depth: &self.typed_depth,
            context: &self.typed_context,
        })
    }
    fn typed_error(&self, number: i32) -> Error {
        if number == -1 {
            return Error::BudgetExceeded;
        }
        match &self.executable {
            Executable::Interpreter(state) => state.borrow_mut().record_typed_entry_error(number),
            #[cfg(feature = "jit")]
            Executable::Jit(_, state) => state.borrow_mut().record_typed_entry_error(number),
        }
        Error::Runtime {
            number,
            description: oxvba_runtime::default_error_message(number).into(),
            vba_source: self.module.program.unit_name.clone(),
        }
    }
    pub fn bind_slots(
        &self,
        bound: &BoundProcedure,
        arguments: &[ScalarType],
        result: ScalarType,
    ) -> Result<(), Error> {
        self.validate(bound)?;
        let layout = self.typed[bound.id.0]
            .as_ref()
            .map_err(|e| Error::UnsupportedSignature(e.clone()))?
            .layout();
        if layout.parameters != arguments || layout.result != result {
            return Err(Error::ArgumentMismatch);
        }
        Ok(())
    }
    pub fn call_slots(
        &self,
        bound: &BoundProcedure,
        args: &[calcvba_host_api::Slot<'_>],
        context: &mut calcvba_host_api::TypedContext,
    ) -> Result<calcvba_host_api::ScalarBits, Error> {
        self.validate(bound)?;
        let _activation = self.enter_typed(context)?;
        let typed = self.typed[bound.id.0]
            .as_ref()
            .map_err(|e| Error::UnsupportedSignature(e.clone()))?;
        if args.len() != typed.layout().parameters.len()
            || args
                .iter()
                .zip(&typed.layout().parameters)
                .any(|(a, t)| a.ty() != *t)
        {
            return Err(Error::ArgumentMismatch);
        }
        let result = match &self.executable {
            Executable::Interpreter(_) => typed.invoke_slots(args, context),
            #[cfg(feature = "jit")]
            Executable::Jit(image, _) => image.invoke_typed_slots(bound.id, args, context),
        };
        result.map_err(|number| self.typed_error(number))
    }
    /// Explicit dynamic adapter: exact-arity arguments already prepared by host.
    pub fn call_bound(
        &self,
        handle: &BoundProcedure,
        args: &[Variant],
        max_steps: u64,
    ) -> Result<Variant, Error> {
        self.validate(handle)?;
        let _executing = self
            .executing
            .try_borrow_mut()
            .map_err(|_| Error::Reentry)?;
        let procedure = &self.module.procedures[handle.metadata];
        if args.len() != procedure.parameters.len()
            || args
                .iter()
                .zip(&procedure.parameters)
                .any(|(value, p)| !prepared_argument_matches(value, &p.ty))
        {
            return Err(Error::ArgumentMismatch);
        }
        match &self.executable {
            Executable::Interpreter(state) => {
                let mut state = state.try_borrow_mut().map_err(|_| Error::Reentry)?;
                state.set_execution_budget(max_steps);
                state
                    .invoke_function_values(handle.id, args.to_vec())
                    .map_err(vm_error)
            }
            #[cfg(feature = "jit")]
            Executable::Jit(image, state) => {
                let mut state = state.try_borrow_mut().map_err(|_| Error::Reentry)?;
                let mut result = image
                    .invoke_in_session(&mut state, handle.id, args, max_steps)
                    .map_err(jit_error)?;
                if result.raised {
                    return Err(Error::Runtime {
                        number: result.err.number,
                        description: result.err.description,
                        vba_source: result.err.source,
                    });
                }
                Ok(result.values.pop().unwrap_or_else(Variant::empty))
            }
        }
    }
}

fn prepared_argument_matches(value: &Variant, ty: &OxTy) -> bool {
    use oxvba_runtime::VarType as V;
    let tag = value.vtype();
    match ty {
        OxTy::Variant => true,
        OxTy::Bool => tag == V::Boolean,
        OxTy::Byte => tag == V::Byte,
        OxTy::Integer => tag == V::Integer,
        OxTy::Long => tag == V::Long,
        OxTy::LongLong => tag == V::LongLong,
        OxTy::Single => tag == V::Single,
        OxTy::Double => tag == V::Double,
        OxTy::Currency => tag == V::Currency,
        OxTy::Date => tag == V::Date,
        OxTy::Decimal => tag == V::Decimal,
        OxTy::Str | OxTy::FixedStr(_) => tag == V::String,
        OxTy::Object(_) => tag == V::Object,
        OxTy::Array(_, _) => tag == V::ArrayVariant,
        OxTy::Record(_) => tag == V::Record,
        OxTy::ProcRef => tag == V::ProcRef,
    }
}

fn vm_error(error: oxvba_vm3::Vm3Error) -> Error {
    match error {
        oxvba_vm3::Vm3Error::BudgetExceeded => Error::BudgetExceeded,
        oxvba_vm3::Vm3Error::Fault(f) => Error::Runtime {
            number: f.code,
            description: f.message,
            vba_source: f.source.unwrap_or_default(),
        },
        other => Error::Backend(other.to_string()),
    }
}
#[cfg(feature = "jit")]
fn jit_error(error: oxvba_jit::JitError) -> Error {
    match error {
        oxvba_jit::JitError::BudgetExceeded => Error::BudgetExceeded,
        oxvba_jit::JitError::Unsupported(reason) => Error::UnsupportedBackend(reason),
        other => Error::Backend(other.to_string()),
    }
}
