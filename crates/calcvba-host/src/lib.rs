//! Host-neutral procedure catalogue and persistent, generation-checked calls.
//! Spreadsheet adapters live outside this engine, in adapters/spreadsheet-compat.
use calcvba_bundle::{ArrayElementType, coreir::CoreConst};
use calcvba_hal::{HostServices, adapters::builder::HostBuilder};
use calcvba_oxir::{FuncId, OxProgram, OxTy};
use calcvba_rt_abi::host_entry::LinkedTarget;
use calcvba_runtime::Variant;
use calcvba_symbol::{TypeLibResolver, manifest as sym};
use std::{
    cell::{Cell, Ref, RefCell},
    marker::PhantomData,
    rc::Rc,
    sync::Arc,
};
mod automatic_session;
mod cache;
mod events;
mod invocation;
mod typed_stack;
use cache::PreparedCode;
pub use cache::{CacheStats, CodeCache, OwnedSession};
pub use calcvba_bundle::ProcedureKind;
pub use calcvba_host_api::ScalarType;
pub use calcvba_rt_abi::host_entry::{Arguments, Scalar, TYPED_ENTRY_ABI, TypedContext};
pub use calcvba_symbol::model::Visibility;
mod project_graph;
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

#[derive(Debug, Clone, thiserror::Error)]
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
    reference_programs: Vec<OxProgram>,
    procedures: Vec<Procedure>,
    host: Arc<dyn HostServices>,
    documents: Vec<DocumentModule>,
}
#[derive(Clone, Debug)]
struct DocumentModule {
    name: String,
    host_type: String,
    class: usize,
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
        Self::load_project(&manifest, &calcvba_symbol::CatalogTypeLibResolver)
    }
    /// The host supplies ordered references, type libraries and module identities.
    /// No host library name, object model, ProgID or activation is invented here.
    pub fn load_project(
        manifest: &sym::SymbolProjectManifest,
        resolver: &dyn TypeLibResolver,
    ) -> Result<Self, Error> {
        let (core, catalogue) = calcvba_bind::bind_program_with_procedures(manifest, resolver)
            .map_err(|e| Error::Compile(e.to_string()))?;
        let mut program =
            calcvba_oxir::elaborate::elaborate(&core).map_err(|e| Error::Compile(e.to_string()))?;
        // A source-module host has explicit bound procedures, not an implicit
        // program run entry or diagnostic entry-local snapshot. Initialization
        // remains the separate compiler-owned global initializer.
        program.entry = None;
        calcvba_oxir::verify_program(&program)
            .map_err(|e| Error::Compile(format!("invalid program: {e:?}")))?;
        if program.funcs.iter().any(|f| {
            f.blocks
                .iter()
                .any(|b| matches!(b.terminator, calcvba_oxir::OxTerminator::Halt))
        }) {
            return Err(Error::UnsupportedSignature(
                "End requires whole-project reset semantics".into(),
            ));
        }
        let procedures = catalogue
            .into_iter()
            .filter(|p| {
                manifest.modules.iter().any(|m| {
                    m.module_name.eq_ignore_ascii_case(&p.module)
                        && m.module_kind == sym::ModuleKind::Procedural
                })
            })
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
        let documents = manifest
            .modules
            .iter()
            .filter(|m| m.module_kind == sym::ModuleKind::Document)
            .filter_map(|m| {
                Some(DocumentModule {
                    name: m.module_name.clone(),
                    host_type: m.attributes.host_document_type.clone()?,
                    class: program
                        .classes
                        .iter()
                        .position(|c| c.name.eq_ignore_ascii_case(&m.module_name))?,
                })
            })
            .collect();
        Ok(Self {
            program,
            reference_programs: vec![],
            procedures,
            host: HostBuilder::new().build(),
            documents,
        })
    }
    /// Compile leaf-first units, retaining independent canonical stores.
    pub fn load_project_graph(
        manifests: &[sym::SymbolProjectManifest],
        resolver: &dyn TypeLibResolver,
    ) -> Result<Self, Error> {
        let mut modules = manifests
            .iter()
            .map(|m| Self::load_project(m, resolver))
            .collect::<Result<Vec<_>, _>>()?;
        let mut root = modules
            .pop()
            .ok_or_else(|| Error::Compile("empty project closure".into()))?;
        root.reference_programs = modules.into_iter().map(|m| m.program).collect();
        Ok(root)
    }
    fn programs(&self) -> Vec<&OxProgram> {
        self.reference_programs
            .iter()
            .chain(std::iter::once(&self.program))
            .collect()
    }
    pub fn with_host(mut self, host: Arc<dyn HostServices>) -> Self {
        self.host = host;
        self
    }
    pub fn procedures(&self) -> impl ExactSizeIterator<Item = &Procedure> {
        self.procedures.iter()
    }
    pub fn com_interfaces(&self) -> &[calcvba_oxir::ComInterface] {
        &self.program.com_interfaces
    }
    pub fn event_routes(&self) -> &[calcvba_bundle::EventRoute] {
        &self.program.event_routes
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
    native: Option<calcvba_jit::TypedEntry<A, R>>,
    signature: PhantomData<fn() -> (A, R)>,
}

enum Executable<'a> {
    Interpreter(RefCell<calcvba_vm3::Vm3<'a>>),
    #[cfg(feature = "jit")]
    Jit(
        Rc<calcvba_jit::CompiledImage<'a>>,
        RefCell<calcvba_jit::JitSession<'a>>,
    ),
}

pub struct Session<'a> {
    module: &'a Module,
    programs: Vec<&'a OxProgram>,
    host: &'a dyn HostServices,
    executable: Executable<'a>,
    typed: Vec<Result<calcvba_vm3::typed::TypedInterpreter<'a>, String>>,
    owning_typed: Vec<Vec<Result<calcvba_vm3::typed::TypedInterpreter<'a>, String>>>,
    owner: Rc<()>,
    generation: u64,
    executing: RefCell<()>,
    typed_depth: Cell<u32>,
    typed_context: Cell<*mut calcvba_host_api::TypedContext>,
    suspended: Cell<u32>,
    object_owner: RefCell<calcvba_runtime::object_owner::RuntimeObjectOwner>,
    automatic: Rc<events::AutomaticEvents>,
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
        let object_owner = calcvba_runtime::object_owner::RuntimeObjectOwner::fresh()
            .ok_or_else(|| Error::Backend("object owner identity exhausted".into()))?;
        let _object_scope = object_owner.enter();
        let executable = match code {
            PreparedCode::Interpreter(_) => Executable::Interpreter(RefCell::new(
                calcvba_vm3::Vm3::link_with_budget(
                    &module.programs(),
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
                if module.programs().iter().any(|p| !p.classes.is_empty()) {
                    Err("source class sessions require owning activations".into())
                } else {
                    calcvba_vm3::typed::TypedInterpreter::bind_with_interfaces(
                        func,
                        &module.program.com_interfaces,
                    )
                }
            })
            .collect();
        let programs = module.programs();
        let owning_typed = programs
            .iter()
            .map(|program| {
                program
                    .funcs
                    .iter()
                    .map(|func| {
                        calcvba_vm3::typed::TypedInterpreter::bind_with_programs(
                            func, program, &programs,
                        )
                    })
                    .collect()
            })
            .collect();
        let event_state = match &executable {
            Executable::Interpreter(state) => state.borrow().typed_event_state(),
            #[cfg(feature = "jit")]
            Executable::Jit(_, state) => state.borrow().typed_event_state(),
        };
        Ok(Self {
            automatic: events::AutomaticEvents::new(event_state),
            object_owner: RefCell::new(object_owner),
            module,
            programs,
            host,
            executable,
            typed,
            owning_typed,
            owner: Rc::new(()),
            generation: 1,
            executing: RefCell::new(()),
            typed_depth: Cell::new(0),
            typed_context: Cell::new(std::ptr::null_mut()),
            suspended: Cell::new(0),
        })
    }
    /// Reinitialize globals/Static/Err using retained code. Bound handles stay valid.
    pub fn reset(&self) -> Result<(), Error> {
        if self.suspended.get() != 0 {
            return Err(Error::Reentry);
        }
        let _executing = self
            .executing
            .try_borrow_mut()
            .map_err(|_| Error::Reentry)?;
        let next_owner = calcvba_runtime::object_owner::RuntimeObjectOwner::fresh()
            .ok_or_else(|| Error::Backend("object owner identity exhausted".into()))?;
        let _object_scope = next_owner.enter();
        self.automatic.clear();
        match &self.executable {
            Executable::Interpreter(state) => {
                let next = calcvba_vm3::Vm3::link_with_budget(
                    &self.module.programs(),
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
        *self.object_owner.borrow_mut() = next_owner;
        let state = match &self.executable {
            Executable::Interpreter(state) => state.borrow().typed_event_state(),
            #[cfg(feature = "jit")]
            Executable::Jit(_, state) => state.borrow().typed_event_state(),
        };
        *self.automatic.state.borrow_mut() = state;
        let documents = self
            .automatic
            .documents
            .borrow()
            .iter()
            .map(|((_, class), source)| (*class, source.clone()))
            .collect::<Vec<_>>();
        self.automatic.documents.borrow_mut().clear();
        drop(_executing);
        for (class, source) in documents {
            let document = self
                .module
                .documents
                .iter()
                .find(|d| d.class == class)
                .ok_or(Error::ArgumentMismatch)?;
            self.attach_document(&document.name, &source)?;
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
    pub fn com_interfaces(&self) -> &[calcvba_oxir::ComInterface] {
        self.module.com_interfaces()
    }
    pub fn bind(
        &self,
        module: &str,
        name: &str,
        kind: ProcedureKind,
    ) -> Result<BoundProcedure, Error> {
        let module_key = calcvba_symbol::model::fold_identifier(module);
        let name_key = calcvba_symbol::model::fold_identifier(name);
        let mut candidates = self.module.procedures.iter().enumerate().filter(|(_, p)| {
            p.visibility == Visibility::Public
                && calcvba_symbol::model::fold_identifier(&p.module) == module_key
                && calcvba_symbol::model::fold_identifier(&p.name) == name_key
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
        self.enter_activation(context, false)
    }
    fn enter_activation(
        &self,
        context: &mut calcvba_host_api::TypedContext,
        resuming: bool,
    ) -> Result<TypedActivation<'_>, Error> {
        if self.suspended.get() != 0 && self.typed_depth.get() == 0 && !resuming {
            return Err(Error::Reentry);
        }
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
            description: calcvba_runtime::default_error_message(number).into(),
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
            .or(self.owning_typed.last().expect("root entries")[bound.id.0].as_ref())
            .map_err(|e| Error::UnsupportedSignature(e.clone()))?
            .layout();
        if !layout.has_result
            || !matches!(
                layout.result.legacy(),
                Some(ScalarType::Long | ScalarType::Double | ScalarType::HostObject)
            )
        {
            return Err(Error::UnsupportedSignature(
                "typed host Sub entry requires an explicit void carrier".into(),
            ));
        }
        if arguments.len() > layout.parameters.len()
            || layout.parameters[..arguments.len()] != *arguments
            || layout.optional_parameters[arguments.len()..]
                .iter()
                .any(|p| !*p)
            || layout.result != result
        {
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
            .or(self.owning_typed.last().expect("root entries")[bound.id.0].as_ref())
            .map_err(|e| Error::UnsupportedSignature(e.clone()))?;
        if args.len() > typed.layout().parameters.len()
            || typed.layout().optional_parameters[args.len()..]
                .iter()
                .any(|p| !p)
            || args
                .iter()
                .zip(&typed.layout().parameters)
                .any(|(a, t)| a.ty() != *t)
        {
            return Err(Error::ArgumentMismatch);
        }
        if !matches!(
            typed.layout().result.legacy(),
            Some(ScalarType::Long | ScalarType::Double)
        ) {
            return Err(Error::UnsupportedSignature(
                "owning object results require begin_invocation".into(),
            ));
        }
        if self.typed[bound.id.0].is_err() {
            let mut frame = self.new_activation(bound, args)?;
            let status = self.execute_frame(&mut frame, context, u64::MAX)?;
            if status != 0 {
                return Err(Error::Backend("unexpected finite-slice suspension".into()));
            }
            return Ok(frame.raw_result());
        }
        let result = match &self.executable {
            Executable::Interpreter(_) => self.typed[bound.id.0]
                .as_ref()
                .expect("pure typed entry")
                .invoke_slots(args, context),
            #[cfg(feature = "jit")]
            Executable::Jit(image, _) => image.invoke_typed_slots(bound.id, args, context),
        };
        result.map_err(|number| self.typed_error(number))
    }
    fn bind_values(
        &self,
        bound: &BoundProcedure,
        arguments: &[calcvba_host_api::ValueType],
        result: calcvba_host_api::ValueType,
    ) -> Result<(), Error> {
        self.validate(bound)?;
        let layout = self.owning_typed.last().expect("root entries")[bound.id.0]
            .as_ref()
            .map_err(|e| Error::UnsupportedSignature(e.clone()))?
            .layout();
        for (index, argument) in arguments.iter().enumerate() {
            if layout.param_array.is_some_and(|tail| index >= tail) {
                continue;
            }
            if let Some(iid) = argument.interface()
                && layout.parameter_iids.get(index) != Some(&Some(iid))
            {
                return Err(Error::ArgumentMismatch);
            }
        }
        if let Some(iid) = result.interface() {
            let nominal = self.module.procedures[bound.metadata]
                .return_type
                .as_ref()
                .and_then(|ty| {
                    let OxTy::Object(calcvba_oxir::ObjClass::ComIface(id)) = ty else {
                        return None;
                    };
                    let calcvba_oxir::ComInterface::Com(meta) =
                        self.module.program.com_interfaces.get(id.0)?
                    else {
                        return None;
                    };
                    meta.iid.map(provider::interface_id)
                });
            if nominal != Some(iid) {
                return Err(Error::ArgumentMismatch);
            }
        }
        let arguments: Vec<_> = arguments.iter().map(|ty| ty.carrier()).collect();
        if let Some(tail) = layout.param_array {
            if arguments.len() < tail
                || arguments[..tail] != layout.parameters[..tail]
                || layout.result != result.carrier()
            {
                return Err(Error::ArgumentMismatch);
            }
            return Ok(());
        }
        if arguments.len() > layout.parameters.len()
            || layout.parameters[..arguments.len()] != arguments
            || layout.optional_parameters[arguments.len()..]
                .iter()
                .any(|p| !p)
            || layout.result != result.carrier()
        {
            return Err(Error::ArgumentMismatch);
        }
        Ok(())
    }
    fn call_values(
        &self,
        bound: &BoundProcedure,
        args: &[calcvba_host_api::ValueSlot<'_>],
        context: &mut TypedContext,
    ) -> Result<calcvba_host_api::Value, Error> {
        self.validate(bound)?;
        let _activation = self.enter_typed(context)?;
        let mut frame = self.new_value_activation(bound, args)?;
        if self.execute_frame(&mut frame, context, u64::MAX)? != 0 {
            return Err(Error::Backend("unexpected synchronous suspension".into()));
        }
        Ok(frame.result())
    }
    fn call_values_with_writebacks(
        &self,
        bound: &BoundProcedure,
        args: &mut [calcvba_host_api::Value],
        context: &mut TypedContext,
    ) -> Result<calcvba_host_api::Value, Error> {
        self.validate(bound)?;
        let _activation = self.enter_typed(context)?;
        let inputs = args
            .iter()
            .map(|v| v.borrow_slot().ok_or(Error::ArgumentMismatch))
            .collect::<Result<Vec<_>, _>>()?;
        let mut frame = self.new_value_activation(bound, &inputs)?;
        frame.enable_writebacks(args.len())?;
        drop(inputs);
        let result = self.execute_frame(&mut frame, context, u64::MAX);
        // VBA mutations made before an unhandled error still reach the caller.
        for (index, value) in frame.take_writebacks().map_err(|n| self.typed_error(n))? {
            args[index] = value;
        }
        if result? != 0 {
            return Err(Error::Backend("unexpected synchronous suspension".into()));
        }
        Ok(frame.result())
    }
    fn new_activation(
        &self,
        bound: &BoundProcedure,
        args: &[calcvba_host_api::Slot<'_>],
    ) -> Result<typed_stack::TypedStack, Error> {
        if args.len() > calcvba_host_api::MAX_TYPED_ARGS {
            return Err(Error::ArgumentMismatch);
        }
        let mut values = [calcvba_host_api::ValueSlot::Long(0); calcvba_host_api::MAX_TYPED_ARGS];
        for (i, arg) in args.iter().enumerate() {
            values[i] = (*arg).into();
        }
        self.new_value_activation(bound, &values[..args.len()])
    }
    fn new_value_activation(
        &self,
        bound: &BoundProcedure,
        args: &[calcvba_host_api::ValueSlot<'_>],
    ) -> Result<typed_stack::TypedStack, Error> {
        let _object_scope = self.object_owner.borrow().enter();
        // Reject the complete reachable typed call graph before any body effect.
        let root = self.module.reference_programs.len();
        self.validate_reachable(root, bound.id)?;
        let typed = self.owning_typed.last().expect("root entries")[bound.id.0]
            .as_ref()
            .map_err(|e| Error::UnsupportedSignature(e.clone()))?;
        let tail_owner = if let Some(tail) = typed.layout().param_array {
            if args.len() < tail {
                return Err(Error::ArgumentMismatch);
            }
            let values = args[tail..]
                .iter()
                .map(|arg| {
                    // SAFETY: the borrowed ValueSlot retains its exact input carrier.
                    unsafe { calcvba_host_api::value::retain_bits(arg.ty(), arg.bits()) }
                        .map(calcvba_rt_abi::host_entry::value::to_variant)
                        .map_err(|_| Error::ArgumentMismatch)
                })
                .collect::<Result<Vec<_>, _>>()?;
            Some(calcvba_host_api::Value::Array(Some(
                calcvba_runtime::safe_array::SafeArray::from_variants(values),
            )))
        } else {
            None
        };
        let packed;
        let args = if let Some(owner) = &tail_owner {
            let tail = typed.layout().param_array.expect("prepared tail");
            packed = args[..tail]
                .iter()
                .copied()
                .chain(std::iter::once(owner.borrow_slot().expect("array owner")))
                .collect::<Vec<_>>();
            &packed[..]
        } else {
            args
        };
        if args.len() > typed.layout().parameters.len()
            || typed.layout().optional_parameters[args.len()..]
                .iter()
                .any(|p| !p)
            || args
                .iter()
                .zip(&typed.layout().parameters)
                .any(|(arg, ty)| arg.ty() != *ty)
        {
            return Err(Error::ArgumentMismatch);
        }
        let defaults = self.module.procedures[bound.metadata]
            .parameters
            .iter()
            .enumerate()
            .skip(args.len())
            .map(|(i, p)| {
                let value = p
                    .default
                    .as_ref()
                    .map(provider::default_value)
                    .unwrap_or_else(|| {
                        if typed.layout().parameters[i] == calcvba_host_api::ValueType::Variant {
                            Variant::from_error_code(0x8002_0004u32 as i32)
                        } else if typed.layout().parameters[i] == ScalarType::HostObject {
                            Variant::nothing()
                        } else {
                            Variant::empty()
                        }
                    });
                calcvba_rt_abi::host_entry::value::from_variant(typed.layout().parameters[i], value)
                    .map_err(|_| Error::ArgumentMismatch)
            })
            .collect::<Result<Vec<_>, _>>()?;
        let mut prepared = args.to_vec();
        prepared.extend(
            defaults
                .iter()
                .map(|v| v.borrow_slot().expect("admitted optional carrier")),
        );
        self.new_internal_stack(root, bound.id, &prepared)
    }
    fn resume_frame(
        &self,
        bound: &BoundProcedure,
        frame: &mut typed_stack::TypedStack,
        context: &mut TypedContext,
        operations: u64,
        resuming: bool,
    ) -> Result<i32, Error> {
        self.validate(bound)?;
        let _activation = self.enter_activation(context, resuming)?;
        self.execute_frame(frame, context, operations)
    }
    fn execute_frame(
        &self,
        frame: &mut typed_stack::TypedStack,
        context: &mut TypedContext,
        operations: u64,
    ) -> Result<i32, Error> {
        let _object_scope = self.object_owner.borrow().enter();
        let result = frame.resume(
            &self.programs,
            &self.owning_typed,
            context,
            operations,
            self.host,
            |unit, id, frame, context, operations| match &self.executable {
                Executable::Interpreter(_) => self.owning_typed[unit][id.0]
                    .as_ref()
                    .expect("bound owning entry")
                    .resume(frame, context, operations),
                #[cfg(feature = "jit")]
                Executable::Jit(image, _) => {
                    // SAFETY: frame was built from this exact checked function and
                    // the same leased canonical store; active guards span entry.
                    unsafe { image.resume_activation_in(unit, id, frame, context, operations) }
                }
            },
        );
        result.map_err(|number| frame.error(number, &self.module.program.unit_name))
    }
    /// Explicit dynamic adapter: exact-arity arguments already prepared by host.
    pub fn call_bound(
        &self,
        handle: &BoundProcedure,
        args: &[Variant],
        max_steps: u64,
    ) -> Result<Variant, Error> {
        let _object_scope = self.object_owner.borrow().enter();
        self.validate(handle)?;
        if self.suspended.get() != 0 {
            return Err(Error::Reentry);
        }
        let _executing = self
            .executing
            .try_borrow_mut()
            .map_err(|_| Error::Reentry)?;
        let procedure = &self.module.procedures[handle.metadata];
        if args.len() != procedure.parameters.len()
            || args.iter().zip(&procedure.parameters).any(|(value, p)| {
                !prepared_argument_matches(value, &p.ty)
                    || value.as_object_ref().is_some_and(|object| {
                        object.is_project_instance()
                            && object.object_owner_id()
                                != Some(
                                    calcvba_runtime::object_owner::RuntimeObjectOwner::current()
                                        .id(),
                                )
                    })
            })
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
    use calcvba_runtime::VarType as V;
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

fn vm_error(error: calcvba_vm3::Vm3Error) -> Error {
    match error {
        calcvba_vm3::Vm3Error::BudgetExceeded => Error::BudgetExceeded,
        calcvba_vm3::Vm3Error::Fault(f) => Error::Runtime {
            number: f.code,
            description: f.message,
            vba_source: f.source.unwrap_or_default(),
        },
        other => Error::Backend(other.to_string()),
    }
}
#[cfg(feature = "jit")]
fn jit_error(error: calcvba_jit::JitError) -> Error {
    match error {
        calcvba_jit::JitError::BudgetExceeded => Error::BudgetExceeded,
        calcvba_jit::JitError::Unsupported(reason) => Error::UnsupportedBackend(reason),
        other => Error::Backend(other.to_string()),
    }
}
