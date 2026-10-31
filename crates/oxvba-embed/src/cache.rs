//! Bounded, same-thread immutable code reuse. Project state belongs to each invocation owner.
use super::*;
use std::collections::VecDeque;

self_cell::self_cell! {
    struct PreparedModule {
        owner: Module,
        #[covariant]
        dependent: PreparedCode,
    }
}

struct SessionOwner {
    prepared: Rc<PreparedModule>,
    host: Arc<dyn HostServices>,
    spreadsheet: Option<Arc<spreadsheet::SpreadsheetHost>>,
}

impl SessionOwner {
    fn new(prepared: Rc<PreparedModule>, host: Arc<dyn HostServices>, spreadsheet: bool) -> Self {
        let spreadsheet = spreadsheet.then(|| Arc::new(spreadsheet::SpreadsheetHost::new()));
        let host: Arc<dyn HostServices> = match &spreadsheet {
            Some(host) => host.clone(),
            None => host,
        };
        Self {
            prepared,
            host,
            spreadsheet,
        }
    }
    fn start(&self) -> Result<Session<'_>, Error> {
        self.prepared
            .borrow_dependent()
            .clone()
            .into_session(&*self.host, self.spreadsheet.as_deref())
    }
}

self_cell::self_cell! {
    struct SessionCell {
        owner: SessionOwner,
        #[not_covariant]
        dependent: Session,
    }
}

/// An owned instance of the same borrowed Session. Mutable project and host-object storage
/// is independent even when compiled code/metadata is shared with another instance.
pub struct OwnedSession(SessionCell);

impl OwnedSession {
    pub fn backend(&self) -> Backend {
        self.0.with_dependent(|_, session| session.backend())
    }
    /// Diagnostic ownership identity, valid within this process. It is never a callable handle.
    pub fn code_identity(&self) -> usize {
        Rc::as_ptr(&self.0.borrow_owner().prepared) as usize
    }
    pub fn reset(&self) -> Result<(), Error> {
        self.0.with_dependent(|_, session| session.reset())
    }
    pub fn call(&self, name: &str, args: &[Value], max_steps: u64) -> Result<Value, Error> {
        self.0
            .with_dependent(|_, session| session.call(name, args, max_steps))
    }
    pub fn call_with_context(
        &self,
        name: &str,
        args: &[Value],
        max_steps: u64,
        context: &mut dyn SpreadsheetContext,
    ) -> Result<Value, Error> {
        self.0
            .with_dependent(|_, session| session.call_with_context(name, args, max_steps, context))
    }
    fn from_prepared(
        prepared: Rc<PreparedModule>,
        host: Arc<dyn HostServices>,
        spreadsheet: bool,
    ) -> Result<Self, Error> {
        SessionCell::try_new(
            SessionOwner::new(prepared, host, spreadsheet),
            SessionOwner::start,
        )
        .map(Self)
    }
}

impl Module {
    pub fn into_session(mut self, backend: Backend) -> Result<OwnedSession, Error> {
        let host = self.host.clone();
        let spreadsheet = self.spreadsheet.is_some();
        self.host = HostBuilder::new().build();
        self.spreadsheet = None;
        let prepared =
            PreparedModule::try_new(self, |module| PreparedCode::compile(module, backend, None))?;
        OwnedSession::from_prepared(Rc::new(prepared), host, spreadsheet)
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
    /// Retained key/image-description bytes. This does not measure Cranelift arena allocation
    /// or code retained by live caller-owned sessions after cache eviction.
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

/// A caller-owned LRU for verified compiled modules. It shares code/semantic metadata, never
/// globals, Static, Err or a host-object table. Rust borrowing serializes lookups on one thread;
/// there is no cross-thread or on-disk code-cache claim.
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
                    "interpreter:{}:{}:carrier={}:{}:v2",
                    std::env::consts::ARCH,
                    std::env::consts::OS,
                    std::mem::size_of::<Variant>(),
                    std::mem::align_of::<Variant>()
                )
            });
        #[cfg(not(feature = "jit"))]
        let target = format!(
            "interpreter:{}:{}:carrier={}:{}:v2",
            std::env::consts::ARCH,
            std::env::consts::OS,
            std::mem::size_of::<Variant>(),
            std::mem::align_of::<Variant>()
        );
        let key = Key {
            image: serde_json::to_vec(&module.program)
                .map_err(|e| Error::Compile(e.to_string()))?,
            // Folded defaults are compiler metadata not yet serialized by OxIR; they must
            // participate in identity even when two function bodies are byte-identical.
            signature: format!("{:?}", module.functions),
            target,
            host: format!(
                "{:?}|{:?}|spreadsheet={}",
                module.host.descriptor(),
                module.host.policy(),
                module.spreadsheet.is_some()
            ),
            backend,
        };
        let host = module.host.clone();
        let spreadsheet = module.spreadsheet.is_some();
        module.host = HostBuilder::new().build();
        module.spreadsheet = None;
        if let Some(index) = self.entries.iter().position(|entry| entry.key == key) {
            let entry = self.entries.remove(index).expect("located cache entry");
            let prepared = entry.prepared.clone();
            self.entries.push_back(entry);
            self.stats.hits += 1;
            return OwnedSession::from_prepared(prepared, host, spreadsheet);
        }
        self.stats.misses += 1;
        let prepared = Rc::new(PreparedModule::try_new(module, |module| {
            PreparedCode::compile(module, backend, configuration)
        })?);
        self.stats.compilations += u64::from(backend == Backend::Jit);
        let bytes = key.bytes();
        if self.max_entries == 0 || bytes > self.max_description_bytes {
            self.stats.bypasses += 1;
        } else {
            while self.entries.len() >= self.max_entries
                || self.stats.description_bytes > self.max_description_bytes - bytes
            {
                let old = self
                    .entries
                    .pop_front()
                    .expect("capacity exceeded by retained entries");
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
        OwnedSession::from_prepared(prepared, host, spreadsheet)
    }
}
