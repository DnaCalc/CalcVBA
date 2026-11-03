//! Runtime composition for clients that depend only on calcvba-host-api.
use crate::{Module, OwnedSession};
use calcvba_host_api as api;
use oxvba_com::{
    ComInterfaceIid, TypeLibMemberMetadata, TypeLibParamType as P, TypeLibWireType as W,
};
use oxvba_symbol::{TypeLibResolver, manifest as sym};
use std::{rc::Rc, sync::Arc};

/// Linked runtime facade. Native applications select VM3 or JIT; interpreter-only
/// builds do not link JIT. The host supplies its services and owns native objects.
pub struct Provider {
    host: Arc<dyn oxvba_hal::HostServices>,
    resolver: Rc<dyn TypeLibResolver>,
    object_libraries: Vec<api::TypedObjectLibrary>,
    class_modules: Vec<String>,
}
impl Default for Provider {
    fn default() -> Self {
        Self {
            host: oxvba_hal::adapters::builder::HostBuilder::new().build(),
            resolver: Rc::new(oxvba_symbol::CatalogTypeLibResolver),
            object_libraries: vec![],
            class_modules: vec![],
        }
    }
}
impl Provider {
    pub fn with_host(host: Arc<dyn oxvba_hal::HostServices>) -> Self {
        Self {
            host,
            ..Self::default()
        }
    }
    /// Explicit source module kinds, supplied by the host rather than inferred from text.
    pub fn with_class_modules(mut self, names: Vec<String>) -> Self {
        self.class_modules = names;
        self
    }
    pub fn with_type_libraries(mut self, resolver: Rc<dyn TypeLibResolver>) -> Self {
        self.resolver = resolver;
        self
    }
    /// Prepare compiler-owned metadata without publishing any callable address.
    pub fn compile_graph(&self, graph: &api::ProjectGraph) -> Result<Module, api::Error> {
        if !self.class_modules.is_empty() {
            return Err(api::Error::UnsupportedSignature(
                "class module selection currently requires a single project".into(),
            ));
        }
        let mut manifests = crate::project_graph::manifests(graph)?;
        if graph
            .projects
            .iter()
            .any(|node| !node.project.libraries.is_empty())
        {
            return Err(api::Error::UnsupportedSignature(
                "source graph requires the explicit typed object-library catalogue".into(),
            ));
        }
        validate_typed_libraries(&self.object_libraries)?;
        for manifest in &mut manifests {
            for library in &self.object_libraries {
                if manifest.references.iter().any(|r| match r {
                    sym::ProjectReference::TypeLibrary { name, .. } => {
                        name.eq_ignore_ascii_case(&library.name)
                    }
                    sym::ProjectReference::Project {
                        referenced_project_name,
                    } => referenced_project_name.eq_ignore_ascii_case(&library.name),
                    _ => false,
                }) {
                    return Err(api::Error::Compile(
                        "duplicate graph library reference".into(),
                    ));
                }
                manifest
                    .references
                    .push(sym::ProjectReference::TypeLibrary {
                        name: library.name.clone(),
                        guid: None,
                        version_major: Some(1),
                        version_minor: Some(0),
                        lcid: None,
                        import_lib: None,
                    });
            }
        }
        let resolver = Libraries(&self.object_libraries, &*self.resolver);
        let mut module = Module::load_project_graph(&manifests, &resolver).map_err(error)?;
        // Reuse the normal compiler retention pass for every unit, including
        // untyped Object/Variant defaults absent from source declarations.
        for program in module
            .reference_programs
            .iter_mut()
            .chain(std::iter::once(&mut module.program))
        {
            retain_interfaces(program, &self.object_libraries, &resolver)?;
        }
        Ok(module.with_host(self.host.clone()))
    }
    pub fn compile(&self, project: &api::Project) -> Result<Module, api::Error> {
        let mut classes = std::collections::HashSet::new();
        for name in &self.class_modules {
            if !classes.insert(name.to_ascii_lowercase())
                || !project
                    .modules
                    .iter()
                    .any(|m| m.name.eq_ignore_ascii_case(name))
            {
                return Err(api::Error::Compile(
                    "unknown or duplicate source class module".into(),
                ));
            }
        }
        validate_libraries(&project.libraries)?;
        let libraries: Vec<_> = project
            .libraries
            .iter()
            .map(legacy_library)
            .chain(self.object_libraries.iter().cloned())
            .collect();
        validate_typed_libraries(&libraries)?;
        let mut references = std::collections::HashSet::new();
        for name in project
            .references
            .iter()
            .map(|r| &r.name)
            .chain(libraries.iter().map(|l| &l.name))
        {
            if name.trim().is_empty() || !references.insert(name.to_ascii_lowercase()) {
                return Err(api::Error::UnsupportedSignature(
                    "empty or duplicate library reference".into(),
                ));
            }
        }
        let manifest = sym::SymbolProjectManifest {
            project_name: project.name.clone(),
            project_kind: sym::ProjectKind::Library,
            modules: project
                .modules
                .iter()
                .map(|m| sym::ModuleUnit {
                    module_name: m.name.clone(),
                    module_kind: if classes.contains(&m.name.to_ascii_lowercase()) {
                        sym::ModuleKind::Class
                    } else {
                        sym::ModuleKind::Procedural
                    },
                    attributes: sym::ModuleAttributes::named(&m.name),
                    source: m.source.clone(),
                })
                .collect(),
            references: project
                .references
                .iter()
                .map(|r| sym::ProjectReference::TypeLibrary {
                    name: r.name.clone(),
                    guid: r.libid.clone(),
                    version_major: r.major,
                    version_minor: r.minor,
                    lcid: r.lcid,
                    import_lib: r.import_library.clone(),
                })
                .chain(
                    libraries
                        .iter()
                        .map(|l| sym::ProjectReference::TypeLibrary {
                            name: l.name.clone(),
                            guid: None,
                            version_major: Some(1),
                            version_minor: Some(0),
                            lcid: None,
                            import_lib: None,
                        }),
                )
                .collect(),
            reference_projects: vec![],
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
        };
        let resolver = Libraries(&libraries, &*self.resolver);
        let mut module = Module::load_project(&manifest, &resolver).map_err(error)?;
        // An untyped Object/Variant can arrive from any explicitly supplied
        // host interface, even when no source declaration mentions that type.
        // Retain its authoritative default plan before publishing code/handles.
        retain_interfaces(&mut module.program, &libraries, &resolver)?;
        Ok(module.with_host(self.host.clone()))
    }
}
fn retain_interfaces(
    program: &mut oxvba_oxir::OxProgram,
    libraries: &[api::TypedObjectLibrary],
    resolver: &dyn TypeLibResolver,
) -> Result<(), api::Error> {
    for library in libraries {
        for interface in &library.interfaces {
            let iid = com_iid(interface.iid);
            if program.com_interfaces.iter().any(|i| {
                matches!(i,
                    oxvba_oxir::ComInterface::Com(meta) if meta.iid == Some(iid))
            }) {
                continue;
            }
            let request = oxvba_com::TypeLibResolveRequest {
                reference_name: library.name.clone(),
                requested_coclass: Some(interface.name.clone()),
                importlib_hint: None,
                libid_hint: None,
                major_version_hint: Some(1),
                minor_version_hint: Some(0),
                lcid_hint: None,
            };
            let blob = resolver
                .resolve(&request)
                .ok_or_else(|| api::Error::Compile("host interface metadata unavailable".into()))?;
            program.com_interfaces.push(oxvba_oxir::ComInterface::Com(
                oxvba_com::TypeLibInterfaceMetadata {
                    name: format!("{}.{}", library.name, interface.name),
                    iid: Some(iid),
                    members: blob.members,
                },
            ));
        }
    }
    Ok(())
}
impl api::RuntimeProvider for Provider {
    fn load_graph(
        &self,
        graph: &api::ProjectGraph,
        libraries: &[api::TypedObjectLibrary],
        backend: api::Backend,
        requested_abi: u32,
    ) -> Result<Box<dyn api::Session>, api::Error> {
        if requested_abi != api::HOST_API_ABI {
            return Err(api::Error::AbiMismatch {
                requested: requested_abi,
                provided: api::HOST_API_ABI,
            });
        }
        let provider = Self {
            host: self.host.clone(),
            resolver: self.resolver.clone(),
            class_modules: self.class_modules.clone(),
            object_libraries: self
                .object_libraries
                .iter()
                .chain(libraries)
                .cloned()
                .collect(),
        };
        let module = provider.compile_graph(graph)?;
        let session = module
            .into_session(match backend {
                api::Backend::Interpreter => crate::Backend::Interpreter,
                api::Backend::Jit => crate::Backend::Jit,
            })
            .map_err(error)?;
        let procedures = catalogue(&session);
        Ok(Box::new(Session {
            session,
            procedures,
            provider,
            graph: Some(graph.clone()),
        }))
    }
    fn value_abi(&self) -> Option<u32> {
        Some(api::TYPED_VALUE_ABI)
    }
    fn load_with_object_libraries(
        &self,
        project: &api::Project,
        libraries: &[api::TypedObjectLibrary],
        backend: api::Backend,
        requested_abi: u32,
    ) -> Result<Box<dyn api::Session>, api::Error> {
        let provider = Self {
            host: self.host.clone(),
            resolver: self.resolver.clone(),
            class_modules: self.class_modules.clone(),
            object_libraries: self
                .object_libraries
                .iter()
                .chain(libraries)
                .cloned()
                .collect(),
        };
        provider.load(project, backend, requested_abi)
    }
    fn info(&self) -> api::ProviderInfo {
        api::ProviderInfo {
            host_api_abi: api::HOST_API_ABI,
            typed_entry_abi: api::TYPED_ENTRY_ABI,
            jit_available: cfg!(feature = "jit"),
            allocator: "linked Rust; each owner drops its own allocations",
        }
    }
    fn load(
        &self,
        project: &api::Project,
        backend: api::Backend,
        requested_abi: u32,
    ) -> Result<Box<dyn api::Session>, api::Error> {
        if requested_abi != api::HOST_API_ABI {
            return Err(api::Error::AbiMismatch {
                requested: requested_abi,
                provided: api::HOST_API_ABI,
            });
        }
        let module = self.compile(project)?;
        let session = module
            .into_session(match backend {
                api::Backend::Interpreter => crate::Backend::Interpreter,
                api::Backend::Jit => crate::Backend::Jit,
            })
            .map_err(error)?;
        let procedures = catalogue(&session);
        Ok(Box::new(Session {
            session,
            procedures,
            graph: None,
            provider: Self {
                host: self.host.clone(),
                resolver: self.resolver.clone(),
                class_modules: self.class_modules.clone(),
                object_libraries: self.object_libraries.clone(),
            },
        }))
    }
}
struct Session {
    session: OwnedSession,
    graph: Option<api::ProjectGraph>,
    procedures: Vec<api::Procedure>,
    provider: Provider,
}
struct TypedBound(crate::BoundProcedure);
fn typed_bound(handle: &api::BoundProcedure) -> Result<&crate::BoundProcedure, api::Error> {
    let bound = handle
        .provider_payload::<TypedBound>()
        .ok_or(api::Error::StaleHandle)?;
    if handle.generation() != bound.0.generation() {
        return Err(api::Error::StaleHandle);
    }
    Ok(&bound.0)
}
fn dynamic_bound(handle: &api::BoundProcedure) -> Result<&crate::BoundProcedure, api::Error> {
    let bound = handle
        .provider_payload::<crate::BoundProcedure>()
        .ok_or(api::Error::StaleHandle)?;
    if handle.generation() != bound.generation() {
        return Err(api::Error::StaleHandle);
    }
    Ok(bound)
}
impl api::Session for Session {
    fn reload_graph(&mut self, graph: &api::ProjectGraph) -> Result<(), api::Error> {
        let module = self.provider.compile_graph(graph)?;
        self.session.reload(module).map_err(error)?;
        self.procedures = catalogue(&self.session);
        self.graph = Some(graph.clone());
        Ok(())
    }
    fn bind_values(
        &self,
        module: &str,
        name: &str,
        arguments: &[api::ValueType],
        result: api::ValueType,
    ) -> Result<api::BoundProcedure, api::Error> {
        let kind = if result == api::ValueType::Void {
            crate::ProcedureKind::Sub
        } else {
            crate::ProcedureKind::Function
        };
        let handle = self.session.bind(module, name, kind).map_err(error)?;
        self.session
            .bind_values(&handle, arguments, result)
            .map_err(error)?;
        let generation = handle.generation();
        Ok(api::BoundProcedure::from_provider(
            TypedBound(handle),
            generation,
        ))
    }
    fn type_bound_values(
        &self,
        handle: &api::BoundProcedure,
        arguments: &[api::ValueType],
        result: api::ValueType,
    ) -> Result<api::BoundProcedure, api::Error> {
        let handle = dynamic_bound(handle)?;
        self.session
            .bind_values(handle, arguments, result)
            .map_err(error)?;
        Ok(api::BoundProcedure::from_provider(
            TypedBound(handle.clone()),
            handle.generation(),
        ))
    }
    fn call_values(
        &self,
        handle: &api::BoundProcedure,
        args: &[api::ValueSlot<'_>],
        context: &mut api::TypedContext,
    ) -> Result<api::Value, api::Error> {
        self.session
            .call_values(typed_bound(handle)?, args, context)
            .map_err(error)
    }
    fn call_values_with_writebacks(
        &self,
        handle: &api::BoundProcedure,
        args: &mut [api::Value],
        context: &mut api::TypedContext,
    ) -> Result<api::Value, api::Error> {
        self.session
            .call_values_with_writebacks(typed_bound(handle)?, args, context)
            .map_err(error)
    }
    fn begin_values(
        &self,
        handle: &api::BoundProcedure,
        args: &[api::ValueSlot<'_>],
        frame: api::InvocationFrame,
        max_blocks: u64,
    ) -> Result<Box<dyn api::ValueInvocation>, api::Error> {
        self.session
            .begin_values(typed_bound(handle)?, args, frame, max_blocks)
            .map_err(error)
    }
    fn begin_invocation(
        &self,
        handle: &api::BoundProcedure,
        args: &[api::Slot<'_>],
        frame: api::InvocationFrame,
        max_blocks: u64,
    ) -> Result<Box<dyn api::Invocation>, api::Error> {
        let bound = typed_bound(handle)?;
        self.session
            .begin_invocation(bound, args, frame, max_blocks)
            .map_err(error)
    }
    fn procedures(&self) -> &[api::Procedure] {
        &self.procedures
    }
    fn bind(
        &self,
        module: &str,
        name: &str,
        kind: api::ProcedureKind,
    ) -> Result<api::BoundProcedure, api::Error> {
        let handle = self
            .session
            .bind(module, name, kind_to_engine(kind))
            .map_err(error)?;
        let generation = handle.generation();
        Ok(api::BoundProcedure::from_provider(handle, generation))
    }
    fn call_bound(
        &self,
        handle: &api::BoundProcedure,
        args: &[api::Variant],
        max_steps: u64,
    ) -> Result<api::Variant, api::Error> {
        let handle = dynamic_bound(handle)?;
        self.session
            .call_bound(handle, args, max_steps)
            .map_err(error)
    }
    fn bind_slots(
        &self,
        module: &str,
        name: &str,
        arguments: &[api::ScalarType],
        result: api::ScalarType,
    ) -> Result<api::BoundProcedure, api::Error> {
        let handle = self
            .session
            .bind(module, name, crate::ProcedureKind::Function)
            .map_err(error)?;
        self.session
            .bind_slots(&handle, arguments, result)
            .map_err(error)?;
        let generation = handle.generation();
        Ok(api::BoundProcedure::from_provider(
            TypedBound(handle),
            generation,
        ))
    }
    fn call_slots(
        &self,
        handle: &api::BoundProcedure,
        args: &[api::Slot<'_>],
        context: &mut api::TypedContext,
    ) -> Result<api::ScalarBits, api::Error> {
        let bound = handle
            .provider_payload::<TypedBound>()
            .ok_or(api::Error::StaleHandle)?;
        if handle.generation() != bound.0.generation() {
            return Err(api::Error::StaleHandle);
        }
        self.session
            .call_slots(&bound.0, args, context)
            .map_err(error)
    }
    fn type_bound_slots(
        &self,
        handle: &api::BoundProcedure,
        arguments: &[api::ScalarType],
        result: api::ScalarType,
    ) -> Result<api::BoundProcedure, api::Error> {
        let handle = dynamic_bound(handle)?;
        self.session
            .bind_slots(handle, arguments, result)
            .map_err(error)?;
        Ok(api::BoundProcedure::from_provider(
            TypedBound(handle.clone()),
            handle.generation(),
        ))
    }
    fn reset(&self) -> Result<(), api::Error> {
        self.session.reset().map_err(error)
    }
    fn reload(&mut self, project: &api::Project) -> Result<(), api::Error> {
        if let Some(mut graph) = self.graph.clone() {
            let root = graph
                .projects
                .iter_mut()
                .find(|n| n.project.name.eq_ignore_ascii_case(&graph.root))
                .ok_or_else(|| api::Error::Compile("missing graph root".into()))?;
            if !root.project.name.eq_ignore_ascii_case(&project.name) {
                return Err(api::Error::Compile(
                    "reload changes source graph root identity".into(),
                ));
            }
            root.project = project.clone();
            return api::Session::reload_graph(self, &graph);
        }
        let module = self.provider.compile(project)?;
        self.session.reload(module).map_err(error)?;
        self.procedures = catalogue(&self.session);
        Ok(())
    }
}
pub(crate) fn error(value: crate::Error) -> api::Error {
    match value {
        crate::Error::Compile(v) => api::Error::Compile(v),
        crate::Error::UnknownProcedure(v) => api::Error::UnknownProcedure(v),
        crate::Error::UnsupportedSignature(v) => api::Error::UnsupportedSignature(v),
        crate::Error::StaleHandle => api::Error::StaleHandle,
        crate::Error::ArgumentMismatch => api::Error::ArgumentMismatch,
        crate::Error::BudgetExceeded => api::Error::BudgetExceeded,
        crate::Error::Reentry => api::Error::Reentry,
        crate::Error::BackendUnavailable => api::Error::BackendUnavailable,
        crate::Error::UnsupportedBackend(v) => api::Error::UnsupportedBackend(v),
        crate::Error::Backend(v) => api::Error::Backend(v),
        crate::Error::Runtime {
            number,
            description,
            vba_source,
        } => api::Error::Runtime {
            number,
            description,
            vba_source,
        },
    }
}
fn kind_to_engine(v: api::ProcedureKind) -> crate::ProcedureKind {
    match v {
        api::ProcedureKind::Sub => crate::ProcedureKind::Sub,
        api::ProcedureKind::Function => crate::ProcedureKind::Function,
        api::ProcedureKind::PropertyGet => crate::ProcedureKind::PropertyGet,
        api::ProcedureKind::PropertyLet => crate::ProcedureKind::PropertyLet,
        api::ProcedureKind::PropertySet => crate::ProcedureKind::PropertySet,
    }
}
fn catalogue(session: &OwnedSession) -> Vec<api::Procedure> {
    session
        .procedures()
        .map(|p| api::Procedure {
            project: p.project.clone(),
            module: p.module.clone(),
            name: p.name.clone(),
            kind: match p.kind {
                crate::ProcedureKind::Sub => api::ProcedureKind::Sub,
                crate::ProcedureKind::Function => api::ProcedureKind::Function,
                crate::ProcedureKind::PropertyGet => api::ProcedureKind::PropertyGet,
                crate::ProcedureKind::PropertyLet => api::ProcedureKind::PropertyLet,
                crate::ProcedureKind::PropertySet => api::ProcedureKind::PropertySet,
            },
            visibility: match p.visibility {
                crate::Visibility::Public => api::Visibility::Public,
                crate::Visibility::Private => api::Visibility::Private,
                crate::Visibility::Friend => api::Visibility::Friend,
            },
            option_private_module: p.option_private_module,
            parameters: p
                .parameters
                .iter()
                .map(|p| api::Parameter {
                    name: p.name.clone(),
                    ty: ty(&p.ty, session),
                    by_ref: p.by_ref,
                    optional: p.optional,
                    param_array: p.param_array,
                    default: p.default.as_ref().map(default_value),
                    array_element: p.array_element.as_ref().map(array_element),
                })
                .collect(),
            return_type: p.return_type.as_ref().map(|t| ty(t, session)),
        })
        .collect()
}
fn ty(value: &oxvba_oxir::OxTy, session: &OwnedSession) -> api::Type {
    use {api::Type as T, oxvba_oxir::OxTy as O};
    match value {
        O::Variant => T::Variant,
        O::Bool => T::Boolean,
        O::Byte => T::Byte,
        O::Integer => T::Integer,
        O::Long => T::Long,
        O::LongLong => T::LongLong,
        O::Single => T::Single,
        O::Double => T::Double,
        O::Currency => T::Currency,
        O::Date => T::Date,
        O::Decimal => T::Decimal,
        O::Str => T::String,
        O::FixedStr(n) => T::FixedString(*n),
        O::ProcRef => T::Procedure,
        O::Object(oxvba_oxir::ObjClass::ComIface(i) | oxvba_oxir::ObjClass::Iface(i)) => {
            let interface = session.com_interfaces().get(i.0);
            T::Object {
                name: interface.map(|i| i.name().to_owned()),
                iid: interface.and_then(|i| match i {
                    oxvba_oxir::ComInterface::Com(meta) => meta.iid.map(interface_id),
                    _ => None,
                }),
            }
        }
        O::Object(_) => T::Object {
            name: None,
            iid: None,
        },
        O::Record(id) => T::Record(format!("record#{}", id.0)),
        O::Array(element, shape) => T::Array {
            element: Box::new(ty(element, session)),
            rank: match shape {
                oxvba_oxir::ArrayShape::Dynamic => None,
                oxvba_oxir::ArrayShape::Fixed { rank } => Some(*rank as usize),
            },
        },
    }
}
fn array_element(value: &oxvba_bundle::ArrayElementType) -> api::ArrayElement {
    use {
        api::{ArrayElement as A, Type as T},
        oxvba_bundle::ArrayElementType as E,
    };
    let scalar = match value {
        E::Variant => T::Variant,
        E::Integer => T::Integer,
        E::Long => T::Long,
        E::LongLong => T::LongLong,
        E::Byte => T::Byte,
        E::Single => T::Single,
        E::Double => T::Double,
        E::Currency => T::Currency,
        E::Date => T::Date,
        E::String => T::String,
        E::FixedString(n) => T::FixedString(*n as u32),
        E::Boolean => T::Boolean,
        E::Record(fields) => return A::Record(fields.iter().map(array_element).collect()),
        E::FixedArray { element, bounds } => {
            return A::FixedArray {
                element: Box::new(array_element(element)),
                bounds: bounds.iter().map(|b| (b.lower, b.len)).collect(),
            };
        }
    };
    A::Scalar(scalar)
}
pub(super) fn default_value(value: &oxvba_bundle::coreir::CoreConst) -> api::Variant {
    use {api::Variant as V, oxvba_bundle::coreir::CoreConst as C};
    match value {
        C::Empty => V::empty(),
        C::Null => V::null(),
        C::Nothing => V::nothing(),
        C::Bool(b) => V::from_bool(*b),
        C::I16(n) => V::from_i16(*n),
        C::I32(n) => V::from_i32(*n),
        C::I64(n) => V::from_i64(*n),
        C::F64(n) => V::from_f64(f64::from_bits(*n)),
        C::F32(n) => V::from_f32(f32::from_bits(*n)),
        C::Currency(n) => V::from_currency_scaled_i64(*n),
        C::Date(n) => V::from_date_f64(f64::from_bits(*n)),
        C::Str(s) => V::from_string(s.clone()),
    }
}
pub(crate) fn interface_id(iid: ComInterfaceIid) -> api::InterfaceId {
    let mut bytes = [0; 16];
    bytes[..4].copy_from_slice(&iid.data1.to_le_bytes());
    bytes[4..6].copy_from_slice(&iid.data2.to_le_bytes());
    bytes[6..8].copy_from_slice(&iid.data3.to_le_bytes());
    bytes[8..].copy_from_slice(&iid.data4);
    bytes
}
struct Libraries<'a>(&'a [api::TypedObjectLibrary], &'a dyn TypeLibResolver);
impl TypeLibResolver for Libraries<'_> {
    fn interface_iid(
        &self,
        request: &oxvba_com::TypeLibResolveRequest,
        blob: &oxvba_com::TypeLibMetadataBlob,
    ) -> Option<ComInterfaceIid> {
        let Some(library) = self
            .0
            .iter()
            .find(|l| l.name.eq_ignore_ascii_case(&request.reference_name))
        else {
            return self.1.interface_iid(request, blob);
        };
        let interface = match &request.requested_coclass {
            Some(name) => library
                .interfaces
                .iter()
                .find(|i| i.name.eq_ignore_ascii_case(name))?,
            None => library.interfaces.first()?,
        };
        Some(ComInterfaceIid {
            data1: u32::from_le_bytes(interface.iid[..4].try_into().ok()?),
            data2: u16::from_le_bytes(interface.iid[4..6].try_into().ok()?),
            data3: u16::from_le_bytes(interface.iid[6..8].try_into().ok()?),
            data4: interface.iid[8..].try_into().ok()?,
        })
    }
    fn resolve(
        &self,
        request: &oxvba_com::TypeLibResolveRequest,
    ) -> Option<oxvba_com::TypeLibMetadataBlob> {
        let Some(library) = self
            .0
            .iter()
            .find(|l| l.name.eq_ignore_ascii_case(&request.reference_name))
        else {
            return self.1.resolve(request);
        };
        let interface = match &request.requested_coclass {
            Some(name) => library
                .interfaces
                .iter()
                .find(|i| i.name.eq_ignore_ascii_case(name))?,
            None => library.interfaces.first()?,
        };
        let bound = interface
            .members
            .iter()
            .map(|m| m.slot + 1)
            .max()
            .unwrap_or(0) as u16;
        let iid = ComInterfaceIid {
            data1: u32::from_le_bytes(interface.iid[..4].try_into().ok()?),
            data2: u16::from_le_bytes(interface.iid[4..6].try_into().ok()?),
            data3: u16::from_le_bytes(interface.iid[6..8].try_into().ok()?),
            data4: interface.iid[8..].try_into().ok()?,
        };
        let members: Vec<_> = interface
            .members
            .iter()
            .map(|m| TypeLibMemberMetadata {
                name: m.name.clone(),
                token: m.slot as i32 + 1,
                vtable_slot: Some(m.slot as u16),
                requires_argument: !m.parameters.is_empty(),
                invoke_kind: match m.kind {
                    api::TypedMemberKind::Method => oxvba_com::TypeLibMemberInvokeKind::Method,
                    api::TypedMemberKind::PropertyGet => {
                        oxvba_com::TypeLibMemberInvokeKind::PropertyGet
                    }
                    api::TypedMemberKind::PropertyLet => {
                        oxvba_com::TypeLibMemberInvokeKind::PropertyPut
                    }
                    api::TypedMemberKind::PropertySet => {
                        oxvba_com::TypeLibMemberInvokeKind::PropertyPutRef
                    }
                },
                parameter_names: (0..m.parameters.len()).map(|i| format!("arg{i}")).collect(),
                parameter_optional: vec![false; m.parameters.len()],
                parameter_optional_defaults: vec![],
                is_default_member: m.is_default,
                parameter_types: m.parameters.iter().map(|p| param_type(*p)).collect(),
                parameter_wire_types: m.parameters.iter().map(|p| self.wire_type(*p)).collect(),
                parameter_iids: m
                    .parameters
                    .iter()
                    .map(|p| p.interface().map(com_iid))
                    .collect(),
                return_type: (m.result != api::ValueType::Void).then(|| param_type(m.result)),
                return_wire_type: (m.result != api::ValueType::Void)
                    .then(|| self.wire_type(m.result)),
                callconv_is_stdcall: false,
                is_dual: false,
                interface_iid: Some(iid),
                source_typekind: None,
                vtable_slot_bound: Some(bound),
            })
            .collect();
        Some(oxvba_com::TypeLibMetadataBlob {
            identity: oxvba_com::TypeLibResolvedIdentity {
                reference_name: request.reference_name.clone(),
                requested_coclass: request.requested_coclass.clone(),
                importlib: "calcvba-host-slots-v1".into(),
                libid: None,
                major_version: 1,
                minor_version: 0,
                lcid: None,
                cache_key: format!("host-slots:{}:{}", library.name, interface.name),
            },
            activation_prog_id: None,
            member_name_to_token: members.iter().map(|m| (m.name.clone(), m.token)).collect(),
            members,
            events: vec![],
            coclass_names: library.interfaces.iter().map(|i| i.name.clone()).collect(),
        })
    }
}
fn com_iid(iid: api::InterfaceId) -> ComInterfaceIid {
    ComInterfaceIid {
        data1: u32::from_le_bytes(iid[..4].try_into().expect("IID")),
        data2: u16::from_le_bytes(iid[4..6].try_into().expect("IID")),
        data3: u16::from_le_bytes(iid[6..8].try_into().expect("IID")),
        data4: iid[8..].try_into().expect("IID"),
    }
}
impl Libraries<'_> {
    fn wire_type(&self, ty: api::ValueType) -> W {
        if let api::ValueType::Interface(iid) = ty {
            for library in self.0 {
                if let Some(interface) = library.interfaces.iter().find(|i| i.iid == iid) {
                    return W::InterfacePointer {
                        name: format!("{}.{}", library.name, interface.name),
                    };
                }
            }
        }
        W::Automation(param_type(ty))
    }
}
fn param_type(t: api::ValueType) -> P {
    match t.carrier() {
        api::ValueType::Long => P::Long,
        api::ValueType::Double => P::Double,
        api::ValueType::Boolean => P::Boolean,
        api::ValueType::String => P::String,
        api::ValueType::Variant => P::Variant,
        api::ValueType::HostObject => P::Object,
        _ => P::Variant,
    }
}
fn legacy_library(library: &api::ObjectLibrary) -> api::TypedObjectLibrary {
    api::TypedObjectLibrary {
        name: library.name.clone(),
        interfaces: library
            .interfaces
            .iter()
            .map(|interface| api::TypedObjectInterface {
                name: interface.name.clone(),
                iid: interface.iid,
                members: interface
                    .members
                    .iter()
                    .map(|m| api::TypedObjectMember {
                        name: m.name.clone(),
                        slot: m.slot,
                        kind: match m.kind {
                            api::MemberKind::Method => api::TypedMemberKind::Method,
                            api::MemberKind::PropertyGet => api::TypedMemberKind::PropertyGet,
                        },
                        parameters: m.parameters.iter().copied().map(Into::into).collect(),
                        result: m.result.into(),
                        is_default: m.is_default,
                    })
                    .collect(),
            })
            .collect(),
    }
}
fn validate_typed_libraries(libraries: &[api::TypedObjectLibrary]) -> Result<(), api::Error> {
    let invalid = |why: &str| api::Error::UnsupportedSignature(why.into());
    let mut names = std::collections::HashSet::new();
    let mut iids = std::collections::HashSet::new();
    for library in libraries {
        if library.name.trim().is_empty() || !names.insert(library.name.to_ascii_lowercase()) {
            return Err(invalid("duplicate/empty typed host library"));
        }
        let mut interfaces = std::collections::HashSet::new();
        for interface in &library.interfaces {
            if interface.name.trim().is_empty()
                || !interfaces.insert(interface.name.to_ascii_lowercase())
                || interface.iid == [0; 16]
                || !iids.insert(interface.iid)
            {
                return Err(invalid("duplicate/empty typed host interface or IID"));
            }
            let mut slots = std::collections::HashSet::new();
            let mut members = std::collections::HashSet::new();
            let mut default = None;
            for member in &interface.members {
                let name = member.name.to_ascii_lowercase();
                if member.slot >= u16::MAX as usize
                    || !slots.insert(member.slot)
                    || name.trim().is_empty()
                    || !members.insert((name.clone(), member.kind as u8))
                    || api::value_slot_signature(&member.parameters, member.result).is_none()
                {
                    return Err(invalid("invalid typed host slot/signature"));
                }
                if member.is_default {
                    if default.as_ref().is_some_and(|old| *old != name) {
                        return Err(invalid("ambiguous typed default member"));
                    }
                    default = Some(name);
                }
            }
        }
    }
    for ty in libraries
        .iter()
        .flat_map(|l| &l.interfaces)
        .flat_map(|i| &i.members)
        .flat_map(|m| m.parameters.iter().chain(std::iter::once(&m.result)))
    {
        if let Some(iid) = ty.interface()
            && !iids.contains(&iid)
        {
            return Err(invalid(
                "typed member refers to an undeclared interface IID",
            ));
        }
    }
    Ok(())
}
fn validate_libraries(libraries: &[api::ObjectLibrary]) -> Result<(), api::Error> {
    let invalid = |why: &str| api::Error::UnsupportedSignature(why.into());
    let mut names = std::collections::HashSet::new();
    let mut iids = std::collections::HashSet::new();
    for library in libraries {
        if library.name.is_empty() || !names.insert(library.name.to_ascii_lowercase()) {
            return Err(invalid("duplicate/empty host library"));
        }
        let mut interfaces = std::collections::HashSet::new();
        for interface in &library.interfaces {
            if interface.name.is_empty()
                || !interfaces.insert(interface.name.to_ascii_lowercase())
                || interface.iid == [0; 16]
                || !iids.insert(interface.iid)
            {
                return Err(invalid("duplicate/empty host interface or IID"));
            }
            let mut slots = std::collections::HashSet::new();
            let mut members = std::collections::HashSet::new();
            let mut defaults = 0;
            for member in &interface.members {
                defaults += usize::from(member.is_default);
                if member.slot >= u16::MAX as usize
                    || !slots.insert(member.slot)
                    || member.name.trim().is_empty()
                    || !members.insert((member.name.to_ascii_lowercase(), member.kind as u8))
                    || defaults > 1
                    || !matches!(member.parameters.as_slice(), [] | [api::ScalarType::Long])
                    || member.result != api::ScalarType::Double
                {
                    return Err(invalid(
                        "host slots require a unique bounded slot, Double result and zero or one Long argument",
                    ));
                }
            }
        }
    }
    Ok(())
}
