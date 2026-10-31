//! Portable COM projection: universal host object registration API.
//!
//! Available on every platform (AP-5). On Windows, host-registered objects
//! coexist with real COM objects (portable registry is checked first).
//! On Linux/macOS, it's the only COM-like mechanism. On WASM, the JS host
//! registers objects into it.

use oxvba_runtime::Variant;
use std::collections::HashMap;
use std::sync::{Arc, Mutex};

/// Factory for creating portable dispatch objects.
pub trait PortableObjectFactory: Send + Sync {
    fn create(&self) -> Box<dyn PortableDispatch>;
}

/// A portable dispatch result that should be retained as another COM-like
/// object by the host projection.
pub struct PortableObjectResult {
    pub prog_id: String,
    pub object: Box<dyn PortableDispatch>,
}

impl PortableObjectResult {
    pub fn new(prog_id: impl Into<String>, object: Box<dyn PortableDispatch>) -> Self {
        Self {
            prog_id: prog_id.into(),
            object,
        }
    }
}

/// A typed portable call returns values/writebacks or transfers a child object
/// for the host to retain. No raw/native COM pointer is fabricated.
pub enum PortableTypedCallResult {
    Value(oxvba_runtime::RuntimeCallResult),
    Object(PortableObjectResult),
}

/// Host-registered object with dynamic dispatch and an explicit typed-call option.
///
/// This portable host API carries retained [`Variant`] values.
pub trait PortableDispatch: Send + Sync {
    /// Explicit compiler-bound host subset. Implementations select their own
    /// fixed member identity/signature; no default name-based retry is allowed.
    /// This is a portable host adapter, not native COM transport evidence.
    fn invoke_early_bound(
        &self,
        _member: &crate::TypeLibMemberMetadata,
        _args: &[crate::DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        Err("portable object does not implement typed calls".into())
    }

    fn invoke(&self, member: &str, args: &[Variant]) -> Result<Variant, String>;
    fn get(&self, member: &str) -> Result<Variant, String>;
    fn put(&self, member: &str, value: Variant) -> Result<(), String>;
    fn member_names(&self) -> Vec<String>;

    fn invoke_object(
        &self,
        _member: &str,
        _args: &[Variant],
    ) -> Option<Result<PortableObjectResult, String>> {
        None
    }

    fn get_object(&self, _member: &str) -> Option<Result<PortableObjectResult, String>> {
        None
    }
}

/// Registry of host-provided COM-like objects, keyed by ProgID.
#[derive(Default)]
pub struct PortableComProjection {
    registry: Mutex<HashMap<String, Arc<dyn PortableObjectFactory>>>,
}

impl PortableComProjection {
    pub fn new() -> Self {
        Self {
            registry: Mutex::new(HashMap::new()),
        }
    }

    /// Register an object factory for a ProgID. Any VBA `CreateObject()` call
    /// matching this ProgID will use this factory.
    pub fn register_object(&self, prog_id: &str, factory: Arc<dyn PortableObjectFactory>) {
        let mut reg = self.registry.lock().expect("portable registry lock");
        reg.insert(prog_id.to_string(), factory);
    }

    pub fn unregister_object(&self, prog_id: &str) {
        let mut reg = self.registry.lock().expect("portable registry lock");
        reg.remove(prog_id);
    }

    pub fn has_object(&self, prog_id: &str) -> bool {
        let reg = self.registry.lock().expect("portable registry lock");
        reg.contains_key(prog_id)
    }

    pub fn create_object(&self, prog_id: &str) -> Option<Box<dyn PortableDispatch>> {
        let reg = self.registry.lock().expect("portable registry lock");
        reg.get(prog_id).map(|factory| factory.create())
    }
}

impl std::fmt::Debug for PortableComProjection {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let count = self.registry.lock().map(|r| r.len()).unwrap_or(0);
        f.debug_struct("PortableComProjection")
            .field("registered_count", &count)
            .finish()
    }
}

#[derive(Debug, Default)]
pub struct PortableComBridge;
