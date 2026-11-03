//! Thin adapter over neutral code/session ownership.
use super::*;
pub struct OwnedSession {
    inner: calcvba_host::OwnedSession,
    functions: Vec<(Function, calcvba_host::Procedure)>,
    spreadsheet: Option<Arc<spreadsheet::SpreadsheetHost>>,
    bound: Vec<calcvba_host::BoundProcedure>,
}
impl OwnedSession {
    fn wrap(
        inner: calcvba_host::OwnedSession,
        functions: Vec<(Function, calcvba_host::Procedure)>,
        spreadsheet: Option<Arc<spreadsheet::SpreadsheetHost>>,
    ) -> Result<Self, Error> {
        let bound = bind_functions(&functions, |p| inner.bind(&p.module, &p.name, p.kind))?;
        Ok(Self {
            inner,
            functions,
            spreadsheet,
            bound,
        })
    }
    pub(crate) fn from_module(module: Module, backend: Backend) -> Result<Self, Error> {
        let (module, functions, spreadsheet) = split_module(module);
        Self::wrap(
            module.into_session(backend).map_err(engine_error)?,
            functions,
            spreadsheet,
        )
    }
    pub fn backend(&self) -> Backend {
        self.inner.backend()
    }
    pub fn code_identity(&self) -> usize {
        self.inner.code_identity()
    }
    pub fn reset(&self) -> Result<(), Error> {
        self.inner.reset().map_err(engine_error)
    }
    pub fn call(&self, name: &str, args: &[Value], max_steps: u64) -> Result<Value, Error> {
        let call = || self.call_inner(name, args, max_steps);
        if self.spreadsheet.is_some() {
            spreadsheet::without_context(call)
        } else {
            call()
        }
    }
    pub fn call_with_context(
        &self,
        name: &str,
        args: &[Value],
        max_steps: u64,
        context: &mut dyn SpreadsheetContext,
    ) -> Result<Value, Error> {
        spreadsheet::with_context(context, || self.call_inner(name, args, max_steps))
    }
    fn call_inner(&self, name: &str, args: &[Value], max_steps: u64) -> Result<Value, Error> {
        call_function(
            &self.functions,
            self.spreadsheet.as_deref(),
            name,
            args,
            |index, values| {
                self.inner
                    .call_bound(&self.bound[index], values, max_steps)
                    .map_err(engine_error)
            },
        )
    }
}
type SplitModule = (
    calcvba_host::Module,
    Vec<(Function, calcvba_host::Procedure)>,
    Option<Arc<spreadsheet::SpreadsheetHost>>,
);
fn split_module(module: Module) -> SplitModule {
    #[expect(
        clippy::arc_with_non_send_sync,
        reason = "The existing adapter passes Arc-owned hosts to the engine; calls stay on the session thread."
    )]
    let spreadsheet = module
        .spreadsheet
        .as_ref()
        .map(|_| Arc::new(spreadsheet::SpreadsheetHost::new()));
    let engine = if let Some(host) = &spreadsheet {
        module.engine.with_host(host.clone())
    } else {
        module.engine
    };
    (engine, module.functions, spreadsheet)
}
pub struct CodeCache(calcvba_host::CodeCache);
impl CodeCache {
    pub fn new(max_entries: usize, max_description_bytes: usize) -> Self {
        Self(calcvba_host::CodeCache::new(
            max_entries,
            max_description_bytes,
        ))
    }
    pub fn stats(&self) -> CacheStats {
        self.0.stats()
    }
    pub fn clear(&mut self) {
        self.0.clear();
    }
    pub fn prepare(&mut self, module: Module, backend: Backend) -> Result<OwnedSession, Error> {
        let (module, functions, spreadsheet) = split_module(module);
        OwnedSession::wrap(
            self.0.prepare(module, backend).map_err(engine_error)?,
            functions,
            spreadsheet,
        )
    }
    #[cfg(feature = "jit")]
    pub fn prepare_with_configuration(
        &mut self,
        module: Module,
        configuration: &calcvba_jit::JitConfiguration,
    ) -> Result<OwnedSession, Error> {
        let (module, functions, spreadsheet) = split_module(module);
        OwnedSession::wrap(
            self.0
                .prepare_with_configuration(module, configuration)
                .map_err(engine_error)?,
            functions,
            spreadsheet,
        )
    }
}
