#![cfg(windows)]
use oxvba_com::{
    DUAL_FIXTURE_INTERFACE_IID, OptionalParamDefault, SourceTypeKind,
    TypeLibMemberInvokeKind as Kind, TypeLibMemberMetadata, TypeLibMetadataBlob,
    TypeLibParamType as Ty, TypeLibResolveRequest, TypeLibResolvedIdentity, TypeLibWireType,
};
use oxvba_embed::{Backend, Module, ProcedureKind};
use oxvba_symbol::{TypeLibResolver, manifest as sym};
use std::sync::atomic::Ordering;

struct HostTypes;
impl TypeLibResolver for HostTypes {
    fn resolve(&self, request: &TypeLibResolveRequest) -> Option<TypeLibMetadataBlob> {
        if request.reference_name != "TypedHost" {
            return None;
        }
        let member = TypeLibMemberMetadata {
            name: "Count".into(),
            token: 6007,
            vtable_slot: Some(7),
            requires_argument: false,
            invoke_kind: Kind::PropertyGet,
            parameter_names: vec![],
            parameter_optional: vec![],
            parameter_optional_defaults: Vec::<OptionalParamDefault>::new(),
            is_default_member: false,
            parameter_wire_types: vec![],
            parameter_iids: vec![],
            parameter_types: vec![],
            return_type: Some(Ty::Long),
            return_wire_type: Some(TypeLibWireType::Automation(Ty::Long)),
            callconv_is_stdcall: true,
            is_dual: true,
            interface_iid: Some(DUAL_FIXTURE_INTERFACE_IID),
            source_typekind: Some(SourceTypeKind::Interface),
            vtable_slot_bound: Some(40),
        };
        Some(TypeLibMetadataBlob {
            identity: TypeLibResolvedIdentity {
                reference_name: request.reference_name.clone(),
                requested_coclass: request.requested_coclass.clone(),
                importlib: "test-owned-host".into(),
                libid: None,
                major_version: 1,
                minor_version: 0,
                lcid: None,
                cache_key: "neutral-host-test-v1".into(),
            },
            activation_prog_id: None,
            member_name_to_token: vec![(member.name.clone(), member.token)],
            members: vec![member],
            events: vec![],
            coclass_names: vec!["Calculator".into()],
        })
    }
}
#[test]
fn host_supplied_object_uses_declared_iid_and_slot_on_each_backend() {
    let mut backends = vec![Backend::Interpreter];
    if cfg!(feature = "jit") {
        backends.push(Backend::Jit);
    }
    for backend in backends {
        let before = oxvba_runtime::current_thread_live_handle_counts();
        let (pointer, references) = oxvba_com::create_oxvba_dual_vtable_object_observed();
        {
            let host = oxvba_hal::adapters::builder::HostBuilder::new()
                .profile(oxvba_hal::model::HalProfileId::Windows)
                .policy(oxvba_hal::model::HostPolicy::interactive_dev())
                .build();
            // SAFETY: fixture owns one retained native reference; HAL accepts
            // ownership on this thread and releases it when the host is dropped.
            let object = unsafe {
                host.com()
                    .bind_native_dispatch_object_variant("Neutral.TestRoot", pointer)
            }
            .unwrap();
            let manifest = sym::SymbolProjectManifest {
                project_name: "Neutral".into(), project_kind: sym::ProjectKind::Library,
                modules: vec![sym::ModuleUnit { module_name: "Api".into(), module_kind: sym::ModuleKind::Procedural, attributes: sym::ModuleAttributes::named("Api"), source: "Public Function ReadCount(ByVal root As TypedHost.Calculator) As Long\nReadCount = root.Count\nEnd Function\n".into() }],
                references: vec![sym::ProjectReference::TypeLibrary { name: "TypedHost".into(), guid: None, version_major: Some(1), version_minor: Some(0), lcid: None, import_lib: None }],
                reference_projects: vec![], conditional_constants: Default::default(), conditional_compilation_target: Default::default(),
            };
            let module = Module::load_project(&manifest, &HostTypes)
                .unwrap()
                .with_host(host.clone());
            assert!(!module.com_interfaces().is_empty());
            let session = module.prepare(backend).unwrap();
            let bound = session
                .bind("Api", "ReadCount", ProcedureKind::Function)
                .unwrap();
            let before = host.com().com_dispatch_transport_counts();
            assert_eq!(
                session
                    .call_bound(&bound, &[object], 1000)
                    .unwrap()
                    .as_i32(),
                Some(7)
            );
            let after = host.com().com_dispatch_transport_counts();
            assert_eq!(after.0 - before.0, 1);
            assert_eq!(after.1 - before.1, 0);
        }
        assert_eq!(references.references.load(Ordering::SeqCst), 0);
        assert_eq!(references.invoke_calls.load(Ordering::SeqCst), 0);
        assert_eq!(references.name_lookups.load(Ordering::SeqCst), 0);
        assert_eq!(references.type_info_reads.load(Ordering::SeqCst), 0);
        assert!(
            before
                .balance_to(oxvba_runtime::current_thread_live_handle_counts())
                .is_zero()
        );
    }
}
