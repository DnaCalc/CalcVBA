//! Same-thread immutable code ownership, shared by neutral sessions and adapters.
use super::*;
#[cfg(feature = "jit")]
use oxvba_jit::JitConfiguration;
use std::collections::VecDeque;
#[cfg(not(feature = "jit"))]
pub(crate) struct JitConfiguration;

#[derive(Clone)]
pub(crate) enum PreparedCode<'a> {
    Interpreter(&'a Module),
    #[cfg(feature = "jit")]
    Jit(&'a Module, Rc<oxvba_jit::CompiledImage<'a>>),
}
impl<'a> PreparedCode<'a> {
    pub(crate) fn compile(
        module: &'a Module,
        backend: Backend,
        configuration: Option<&JitConfiguration>,
    ) -> Result<Self, Error> {
        match backend {
            Backend::Interpreter => Ok(Self::Interpreter(module)),
            #[cfg(feature = "jit")]
            Backend::Jit => {
                let image = match configuration {
                    Some(configuration) => oxvba_jit::JitEngine
                        .compile_image_with_configuration(&module.programs(), configuration),
                    None => oxvba_jit::JitEngine.compile_image(&module.programs()),
                }
                .map_err(jit_error)?;
                Ok(Self::Jit(module, Rc::new(image)))
            }
            #[cfg(not(feature = "jit"))]
            Backend::Jit => {
                let _ = configuration;
                Err(Error::BackendUnavailable)
            }
        }
    }
    pub(crate) fn start(self, host: &'a dyn HostServices) -> Result<Session<'a>, Error> {
        let module = match &self {
            Self::Interpreter(m) => *m,
            #[cfg(feature = "jit")]
            Self::Jit(m, _) => *m,
        };
        Session::from_code(module, host, self)
    }
}
self_cell::self_cell! {
    struct PreparedModule { owner: Module, #[covariant] dependent: PreparedCode, }
}
pub(crate) struct SessionOwner {
    prepared: Rc<PreparedModule>,
    host: Arc<dyn HostServices>,
}
impl SessionOwner {
    fn start(&self) -> Result<Session<'_>, Error> {
        self.prepared.borrow_dependent().clone().start(&*self.host)
    }
}
self_cell::self_cell! {
    pub(crate) struct SessionCell { owner: SessionOwner, #[not_covariant] dependent: Session, }
}

pub struct OwnedSession {
    pub(crate) cell: Rc<SessionCell>,
}
impl OwnedSession {
    pub(crate) fn new(mut module: Module, backend: Backend) -> Result<Self, Error> {
        let host = module.host.clone();
        module.host = HostBuilder::new().build();
        let prepared = Rc::new(PreparedModule::try_new(module, |m| {
            PreparedCode::compile(m, backend, None)
        })?);
        Self::from_prepared(prepared, host)
    }
    fn from_prepared(
        prepared: Rc<PreparedModule>,
        host: Arc<dyn HostServices>,
    ) -> Result<Self, Error> {
        SessionCell::try_new(SessionOwner { prepared, host }, SessionOwner::start).map(|cell| {
            Self {
                cell: Rc::new(cell),
            }
        })
    }
    pub fn backend(&self) -> Backend {
        self.cell.with_dependent(|_, s| s.backend())
    }
    /// Ownership diagnostic only, never an executable address.
    pub fn code_identity(&self) -> usize {
        Rc::as_ptr(&self.cell.borrow_owner().prepared) as usize
    }
    pub fn procedures(&self) -> impl ExactSizeIterator<Item = &Procedure> {
        self.cell
            .borrow_owner()
            .prepared
            .borrow_owner()
            .procedures()
    }
    pub fn com_interfaces(&self) -> &[oxvba_oxir::ComInterface] {
        self.cell
            .borrow_owner()
            .prepared
            .borrow_owner()
            .com_interfaces()
    }
    pub fn bind(
        &self,
        module: &str,
        name: &str,
        kind: ProcedureKind,
    ) -> Result<BoundProcedure, Error> {
        self.cell.with_dependent(|_, s| s.bind(module, name, kind))
    }
    pub fn bind_typed<A: Arguments, R: Scalar>(
        &self,
        module: &str,
        name: &str,
    ) -> Result<TypedProcedure<A, R>, Error> {
        self.cell.with_dependent(|_, s| s.bind_typed(module, name))
    }
    pub fn call_bound(
        &self,
        handle: &BoundProcedure,
        args: &[Variant],
        max_steps: u64,
    ) -> Result<Variant, Error> {
        self.cell
            .with_dependent(|_, s| s.call_bound(handle, args, max_steps))
    }
    pub fn bind_typed_handle<A: Arguments, R: Scalar>(
        &self,
        bound: &BoundProcedure,
    ) -> Result<TypedProcedure<A, R>, Error> {
        self.cell.with_dependent(|_, s| s.bind_typed_handle(bound))
    }
    pub fn call_typed<A: Arguments, R: Scalar>(
        &self,
        handle: &TypedProcedure<A, R>,
        args: A,
        max_blocks: u64,
    ) -> Result<R, Error> {
        self.cell
            .with_dependent(|_, s| s.call_typed(handle, args, max_blocks))
    }
    pub fn reset(&self) -> Result<(), Error> {
        self.cell.with_dependent(|_, s| s.reset())
    }
    pub fn bind_slots(
        &self,
        bound: &BoundProcedure,
        arguments: &[ScalarType],
        result: ScalarType,
    ) -> Result<(), Error> {
        self.cell
            .with_dependent(|_, s| s.bind_slots(bound, arguments, result))
    }
    pub fn call_slots(
        &self,
        bound: &BoundProcedure,
        args: &[calcvba_host_api::Slot<'_>],
        context: &mut calcvba_host_api::TypedContext,
    ) -> Result<calcvba_host_api::ScalarBits, Error> {
        self.cell
            .with_dependent(|_, s| s.call_slots(bound, args, context))
    }
    pub fn call_typed_in_context<A: Arguments, R: Scalar>(
        &self,
        handle: &TypedProcedure<A, R>,
        args: A,
        context: &mut calcvba_host_api::TypedContext,
    ) -> Result<R, Error> {
        self.cell
            .with_dependent(|_, s| s.call_typed_in_context(handle, args, context))
    }
    /// Build replacement first. On success invalidate generations before old code
    /// can be released. Failed preparation preserves current code and handles.
    pub fn reload(&mut self, module: Module) -> Result<(), Error> {
        let (owner, generation) = self.cell.with_dependent(|_, s| {
            (
                s.owner.clone(),
                s.generation
                    .checked_add(1)
                    .ok_or_else(|| Error::Backend("generation exhausted".into())),
            )
        });
        let generation = generation?;
        let mut next = Self::new(module, self.backend())?;
        Rc::get_mut(&mut next.cell)
            .expect("unpublished replacement")
            .with_dependent_mut(|_, s| {
                s.owner = owner;
                s.generation = generation;
            });
        std::mem::swap(self, &mut next);
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct CacheStats {
    pub hits: u64,
    pub misses: u64,
    pub compilations: u64,
    pub evictions: u64,
    pub bypasses: u64,
    pub entries: usize,
    /// Key/description bytes only; generated code retained by live sessions is excluded.
    pub description_bytes: usize,
}
#[derive(PartialEq, Eq)]
struct Key {
    image: Vec<u8>,
    signature: String,
    target: String,
    host: String,
    backend: Backend,
}
impl Key {
    fn bytes(&self) -> usize {
        self.image
            .len()
            .saturating_add(self.signature.len())
            .saturating_add(self.target.len())
            .saturating_add(self.host.len())
    }
}
struct Entry {
    key: Key,
    prepared: Rc<PreparedModule>,
}
pub struct CodeCache {
    max_entries: usize,
    max_description_bytes: usize,
    entries: VecDeque<Entry>,
    stats: CacheStats,
}
impl CodeCache {
    pub fn new(max_entries: usize, max_description_bytes: usize) -> Self {
        Self {
            max_entries,
            max_description_bytes,
            entries: VecDeque::new(),
            stats: CacheStats::default(),
        }
    }
    pub fn stats(&self) -> CacheStats {
        self.stats
    }
    pub fn clear(&mut self) {
        self.stats.evictions += self.entries.len() as u64;
        self.entries.clear();
        self.stats.entries = 0;
        self.stats.description_bytes = 0;
    }
    pub fn prepare(&mut self, module: Module, backend: Backend) -> Result<OwnedSession, Error> {
        #[cfg(feature = "jit")]
        let configuration = match backend {
            Backend::Interpreter => None,
            Backend::Jit => Some(JitConfiguration::native().map_err(jit_error)?),
        };
        #[cfg(not(feature = "jit"))]
        let configuration = match backend {
            Backend::Interpreter => None,
            Backend::Jit => return Err(Error::BackendUnavailable),
        };
        self.prepare_inner(module, backend, configuration.as_ref())
    }
    #[cfg(feature = "jit")]
    pub fn prepare_with_configuration(
        &mut self,
        module: Module,
        configuration: &JitConfiguration,
    ) -> Result<OwnedSession, Error> {
        self.prepare_inner(module, Backend::Jit, Some(configuration))
    }
    fn prepare_inner(
        &mut self,
        mut module: Module,
        backend: Backend,
        configuration: Option<&JitConfiguration>,
    ) -> Result<OwnedSession, Error> {
        #[cfg(feature = "jit")]
        let target = configuration
            .map(JitConfiguration::identity)
            .unwrap_or_else(|| {
                format!(
                    "interpreter:{}:{}:typed={TYPED_ENTRY_ABI}",
                    std::env::consts::ARCH,
                    std::env::consts::OS
                )
            });
        #[cfg(not(feature = "jit"))]
        let target = {
            let _ = configuration;
            format!(
                "interpreter:{}:{}:typed={TYPED_ENTRY_ABI}",
                std::env::consts::ARCH,
                std::env::consts::OS
            )
        };
        let key = Key {
            image: serde_json::to_vec(&module.programs())
                .map_err(|e| Error::Compile(e.to_string()))?,
            signature: format!("{:?}|{:?}", module.procedures, module.documents),
            target,
            host: format!("{:?}|{:?}", module.host.descriptor(), module.host.policy()),
            backend,
        };
        let host = module.host.clone();
        module.host = HostBuilder::new().build();
        if let Some(index) = self.entries.iter().position(|e| e.key == key) {
            let entry = self.entries.remove(index).expect("located entry");
            let prepared = entry.prepared.clone();
            self.entries.push_back(entry);
            self.stats.hits += 1;
            return OwnedSession::from_prepared(prepared, host);
        }
        self.stats.misses += 1;
        let prepared = Rc::new(PreparedModule::try_new(module, |m| {
            PreparedCode::compile(m, backend, configuration)
        })?);
        self.stats.compilations += u64::from(backend == Backend::Jit);
        let bytes = key.bytes();
        if self.max_entries == 0 || bytes > self.max_description_bytes {
            self.stats.bypasses += 1;
        } else {
            while self.entries.len() >= self.max_entries
                || self.stats.description_bytes > self.max_description_bytes - bytes
            {
                let old = self.entries.pop_front().expect("capacity exceeded");
                self.stats.description_bytes -= old.key.bytes();
                self.stats.evictions += 1;
            }
            self.stats.description_bytes += bytes;
            self.entries.push_back(Entry {
                key,
                prepared: prepared.clone(),
            });
            self.stats.entries = self.entries.len();
        }
        OwnedSession::from_prepared(prepared, host)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn exhausted_generation_never_wraps_or_invalidates_the_live_session() {
        let module = Module::load(
            "Math",
            "Public Function Answer() As Long\nAnswer = 42\nEnd Function\n",
        )
        .unwrap();
        let mut session = module.clone().into_session(Backend::Interpreter).unwrap();
        Rc::get_mut(&mut session.cell)
            .expect("unshared session")
            .with_dependent_mut(|_, s| s.generation = u64::MAX);
        let handle = session.bind_typed::<(), i32>("Math", "Answer").unwrap();
        assert!(
            matches!(session.reload(module), Err(Error::Backend(message)) if message == "generation exhausted")
        );
        assert_eq!(session.call_typed(&handle, (), 100).unwrap(), 42);
    }
}
