//! Compiler -> OxIR -> each runtime -> HAL -> native vtable, without registration.
use super::{Observation, canon};
use oxvba_com::{
    DUAL_FIXTURE_INTERFACE_IID, OptionalParamDefault, SourceTypeKind,
    TypeLibMemberInvokeKind as Kind, TypeLibMemberMetadata as Member, TypeLibMetadataBlob,
    TypeLibParamType as Ty, TypeLibResolveRequest, TypeLibResolvedIdentity, TypeLibWireType,
};
use oxvba_hal::{
    adapters::builder::HostBuilder,
    model::{ComInvocationStrategy, HalProfileId, HostPolicy},
};
use oxvba_host::FinalErr;
use oxvba_jit::JitEngine;
use oxvba_runtime::{Variant, current_thread_live_handle_counts};
use oxvba_symbol::{TypeLibResolver, manifest as sym};
use oxvba_vm3::Vm3;

fn member(name: &str, slot: u16, kind: Kind, params: Vec<Ty>, ret: Option<Ty>) -> Member {
    Member {
        name: name.into(),
        // Deliberately unrelated to the fixture's runtime dispatch IDs. Its
        // GetTypeInfo/GetIDsOfNames/Invoke cannot supply these members at all.
        token: 6000 + i32::from(slot),
        vtable_slot: Some(slot),
        requires_argument: !params.is_empty(),
        invoke_kind: kind,
        parameter_names: (0..params.len()).map(|n| format!("arg{n}")).collect(),
        parameter_optional: vec![false; params.len()],
        parameter_optional_defaults: vec![OptionalParamDefault::Required; params.len()],
        is_default_member: false,
        parameter_wire_types: params
            .iter()
            .copied()
            .map(TypeLibWireType::Automation)
            .collect(),
        parameter_iids: vec![None; params.len()],
        parameter_types: params,
        return_type: ret,
        return_wire_type: ret.map(TypeLibWireType::Automation),
        callconv_is_stdcall: true,
        is_dual: true,
        interface_iid: Some(DUAL_FIXTURE_INTERFACE_IID),
        source_typekind: Some(SourceTypeKind::Interface),
        vtable_slot_bound: Some(40),
    }
}

#[derive(Clone, Copy)]
enum DescriptorFault {
    None,
    Slot,
    Interface,
    DispatchOnly,
}
struct FixtureTypes {
    fault: DescriptorFault,
}
impl TypeLibResolver for FixtureTypes {
    fn resolve(&self, request: &TypeLibResolveRequest) -> Option<TypeLibMetadataBlob> {
        if !request.reference_name.eq_ignore_ascii_case("EarlyHost") {
            return None;
        }
        let mut members = vec![
            member("Count", 7, Kind::PropertyGet, vec![], Some(Ty::Long)),
            member("Exists", 8, Kind::Method, vec![Ty::Long], Some(Ty::Boolean)),
            member("Caption", 21, Kind::PropertyGet, vec![], Some(Ty::String)),
            member(
                "DoubleValue",
                20,
                Kind::PropertyGet,
                vec![],
                Some(Ty::Double),
            ),
            member("Mutate", 28, Kind::Method, vec![Ty::ByRefLong], None),
            member("RaiseError", 11, Kind::Method, vec![], Some(Ty::Long)),
            member("Value", 9, Kind::PropertyPut, vec![Ty::ByRefVariant], None),
            member(
                "AliasLongs",
                39,
                Kind::Method,
                vec![Ty::ByRefLong, Ty::ByRefLong],
                Some(Ty::Long),
            ),
        ];
        match self.fault {
            DescriptorFault::DispatchOnly => {
                let mut typed = member("TypedCount", 7, Kind::PropertyGet, vec![], Some(Ty::Long));
                typed.token = oxvba_com::TEST_DISPID_COUNT;
                typed.vtable_slot = None;
                typed.vtable_slot_bound = None;
                typed.is_dual = false;
                typed.source_typekind = Some(SourceTypeKind::Dispatch);
                typed.interface_iid = Some(oxvba_com::ComInterfaceIid::from_guid(
                    &oxvba_com::windows_client::IID_IDISPATCH,
                ));
                members = vec![typed];
            }
            DescriptorFault::None => {}
            DescriptorFault::Slot => members[0].vtable_slot = Some(40),
            DescriptorFault::Interface => {
                let mut iid = DUAL_FIXTURE_INTERFACE_IID;
                iid.data1 ^= 0x100;
                members[0].interface_iid = Some(iid);
            }
        }
        Some(TypeLibMetadataBlob {
            constants: Vec::new(),
            identity: TypeLibResolvedIdentity {
                reference_name: request.reference_name.clone(),
                requested_coclass: request.requested_coclass.clone(),
                importlib: "test-owned-early-host".into(),
                libid: None,
                major_version: 1,
                minor_version: 0,
                lcid: None,
                cache_key: "test-owned-early-host-v1".into(),
            },
            activation_prog_id: Some("OxVba.Run5Fixture".into()),
            member_name_to_token: members.iter().map(|m| (m.name.clone(), m.token)).collect(),
            members,
            events: vec![],
            coclass_names: vec!["Calculator".into()],
        })
    }
}

fn run(source: &str, jit: bool, fault: DescriptorFault) -> Observation {
    let before = current_thread_live_handle_counts();
    let native_references;
    let observation = {
        let mut policy = HostPolicy::native();
        // This preference controls genuinely dynamic calls only. Typed calls
        // must remain vtable calls even when the host prefers late dispatch.
        policy.com_invocation_strategy = ComInvocationStrategy::DispatchOnly;
        let host = HostBuilder::new()
            .profile(HalProfileId::Windows)
            .policy(policy)
            .build();
        let (fixture, references) = if matches!(fault, DescriptorFault::DispatchOnly) {
            (oxvba_com::create_oxvba_test_dispatch().cast(), None)
        } else {
            let (pointer, count) = oxvba_com::create_oxvba_dual_vtable_object_observed();
            (pointer, Some(count))
        };
        native_references = references;
        // SAFETY: fresh fixture carries one retained IDispatch reference; HAL
        // takes ownership and releases it with the host's bindings map.
        let bound = unsafe {
            host.com()
                .bind_native_dispatch_object_variant("OxVba.Run5Fixture", fixture)
        }
        .expect("bind test-owned dual interface");
        let manifest = sym::SymbolProjectManifest {
            project_name: "Main".into(),
            project_kind: sym::ProjectKind::Source,
            modules: vec![sym::ModuleUnit {
                module_name: "Main".into(),
                module_kind: sym::ModuleKind::Procedural,
                attributes: sym::ModuleAttributes::named("Main"),
                source: source.into(),
            }],
            references: vec![sym::ProjectReference::TypeLibrary {
                name: "EarlyHost".into(),
                guid: None,
                version_major: Some(1),
                version_minor: Some(0),
                lcid: None,
                import_lib: None,
            }],
            reference_projects: vec![],
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
        };
        let core =
            oxvba_bind::bind_program(&manifest, &FixtureTypes { fault }).expect("bind typed host");
        let program = oxvba_oxir::elaborate::elaborate(&core).expect("elaborate typed host");
        oxvba_oxir::verify::verify_program(&program).expect("verify typed host");
        assert!(
            program
                .funcs
                .iter()
                .flat_map(|f| &f.blocks)
                .flat_map(|b| &b.instrs)
                .any(|i| matches!(i, oxvba_oxir::OxInst::ComCallEarly { .. }))
        );
        let mut observation = if jit {
            let code = JitEngine
                .compile_image(&[&program])
                .expect("compile early COM on JIT");
            let outcome = code.run(host.as_ref()).expect("execute JIT early COM");
            Observation {
                values: outcome
                    .values
                    .iter()
                    .take(program.globals.len())
                    .map(canon)
                    .collect(),
                err: FinalErr {
                    number: outcome.err.number,
                    source: outcome.err.source,
                    description: outcome.err.description,
                    last_dll_error: outcome.err.last_dll_error,
                },
                raised: outcome.raised,
                transport: (0, 0),
            }
        } else {
            let mut vm = Vm3::link(&[&program], host.as_ref()).expect("link early COM");
            let outcome = vm.run_entry();
            Observation {
                values: (0..program.globals.len())
                    .map(|i| canon(&vm.slot(i).unwrap()))
                    .collect(),
                err: FinalErr {
                    number: vm.err_number(),
                    source: vm.err_source().into(),
                    description: vm.err_description().into(),
                    last_dll_error: vm.last_dll_error(),
                },
                raised: outcome.is_err(),
                transport: (0, 0),
            }
        };
        observation.transport = host.com().com_dispatch_transport_counts();
        drop(bound);
        observation
    };
    assert!(
        before
            .balance_to(current_thread_live_handle_counts())
            .is_zero(),
        "early COM owned carrier leak"
    );
    if let Some(references) = native_references {
        assert_eq!(
            references
                .references
                .load(std::sync::atomic::Ordering::SeqCst),
            0,
            "every native COM/QI reference must be released on success and refusal/error"
        );
        assert_eq!(
            references
                .error_info_references
                .load(std::sync::atomic::Ordering::SeqCst),
            0,
            "the native fixture's error objects must release creator and thread references"
        );
        assert_eq!(
            references
                .invoke_calls
                .load(std::sync::atomic::Ordering::SeqCst),
            0,
            "no Invoke attempt, including failed attempts, may hide behind successful-call counters"
        );
        assert_eq!(
            references
                .name_lookups
                .load(std::sync::atomic::Ordering::SeqCst),
            0
        );
        assert_eq!(
            references
                .type_info_reads
                .load(std::sync::atomic::Ordering::SeqCst),
            0
        );
    }
    observation
}

fn compare(source: &str, invalid_slot: bool) -> Observation {
    let fault = if invalid_slot {
        DescriptorFault::Slot
    } else {
        DescriptorFault::None
    };
    let vm = run(source, false, fault);
    assert_eq!(run(source, true, fault), vm, "typed COM backend observable");
    vm
}

#[test]
fn compiler_descriptor_drives_native_vtable_on_both_backends() {
    let result = compare(
        r#"
Public count As Long
Public caption As String
Public number As Double
Public exists As Boolean
Public changed As Long
Sub Main()
    Dim obj As EarlyHost.Calculator
    Set obj = CreateObject("OxVba.Run5Fixture")
    count = obj.Count
    caption = obj.Caption
    number = obj.DoubleValue
    exists = obj.Exists(arg0:=42.5)
    changed = 19
    obj.Mutate changed
    obj.Mutate (changed)
    obj.Value = 42
    Set obj = Nothing
End Sub
"#,
        false,
    );
    assert!(!result.raised, "{result:?}");
    assert_eq!(result.err.number, 0);
    assert_eq!(
        &result.values[..5],
        &[
            canon(&Variant::from_i32(7)),
            canon(&Variant::from_string("vtable-text")),
            canon(&Variant::from_f64(-9876.25)),
            canon(&Variant::from_bool(true)),
            canon(&Variant::from_i32(1019)),
        ]
    );
    assert_eq!(result.transport, (7, 0), "zero IDispatch fallback");
}

#[test]
fn early_vtable_hresult_preserves_rich_error_without_retry() {
    let result = compare(
        r#"
Public captured As Long
Sub Main()
    Dim obj As EarlyHost.Calculator
    Set obj = CreateObject("OxVba.Run5Fixture")
    On Error Resume Next
    obj.RaiseError
    captured = Err.Number
    Set obj = Nothing
End Sub
"#,
        false,
    );
    assert!(!result.raised);
    assert_ne!(result.err.number, 0);
    assert_eq!(result.err.source, "OxVba.DualFixture");
    assert_eq!(
        result.err.description,
        "controlled vtable error via SetErrorInfo"
    );
    assert_eq!(result.transport, (0, 0)); // counters count successful calls
}

#[test]
fn unsafe_descriptor_is_refused_without_late_fallback() {
    let result = compare(
        r#"
Public count As Long
Sub Main()
    Dim obj As EarlyHost.Calculator
    Set obj = CreateObject("OxVba.Run5Fixture")
    On Error Resume Next
    count = obj.Count
    Set obj = Nothing
End Sub
"#,
        true,
    );
    assert!(!result.raised);
    assert!(
        result.err.description.contains("SlotOutOfBounds"),
        "{result:?}"
    );
    assert_eq!(result.transport, (0, 0));
}

#[test]
fn native_byref_aliases_and_equal_distinct_cells_stay_distinct() {
    let result = compare(
        r#"
Public same As Long
Public distinct As Long
Public lhs As Long
Public rhs As Long
Private Sub ThroughAlias(ByRef value As Long)
    Dim obj As EarlyHost.Calculator
    Set obj = CreateObject("OxVba.Run5Fixture")
    same = obj.AliasLongs(value, lhs)
    Set obj = Nothing
End Sub
Sub Main()
    Dim obj As EarlyHost.Calculator
    Set obj = CreateObject("OxVba.Run5Fixture")
    lhs = 7
    ThroughAlias lhs
    lhs = 7
    rhs = 7
    distinct = obj.AliasLongs(lhs, rhs)
    Set obj = Nothing
End Sub
"#,
        false,
    );
    assert!(!result.raised, "{result:?}");
    assert_eq!(
        &result.values[..4],
        &[4242, 4108, 41, 8].map(|n| canon(&Variant::from_i32(n)))
    );
    assert_eq!(result.transport, (2, 0));
}

#[test]
fn registered_typelib_dictionary_executes_early_on_both_backends() {
    let source = r#"
Public result As Long
Public count As Long
Public exists As Boolean
Public missing As Variant
Public after As Long
Public duplicate As Long
Public errorSource As String
Public errorDescription As String
Sub Main()
    Dim obj As Scripting.Dictionary
    Set obj = CreateObject("Scripting.Dictionary")
    obj.Add "a", 19
    obj.Item("a") = 42
    result = obj.Item("a")
    count = obj.Count
    exists = obj.Exists("a")
    missing = obj.Item("absent")
    after = obj.Count
    On Error Resume Next
    obj.Add "a", 7
    duplicate = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    Set obj = Nothing
End Sub
"#;
    let reference = sym::ProjectReference::TypeLibrary {
        name: "Scripting".into(),
        guid: Some("{420B2830-E718-11CF-893D-00A0C9054228}".into()),
        version_major: Some(1),
        version_minor: Some(0),
        lcid: None,
        import_lib: None,
    };
    let vm = super::run_referenced(source, false, vec![reference.clone()]);
    let jit = super::run_referenced(source, true, vec![reference]);
    assert_eq!(jit, vm);
    assert!(!vm.raised, "{vm:?}");
    assert_eq!(
        &vm.values[..6],
        &[
            canon(&Variant::from_i32(42)),
            canon(&Variant::from_i32(1)),
            canon(&Variant::from_bool(true)),
            canon(&Variant::empty()),
            canon(&Variant::from_i32(2)),
            canon(&Variant::from_i32(457)),
        ]
    );
    assert_eq!(
        vm.transport,
        (7, 0),
        "real imported dual interface cannot silently use Invoke"
    );
    // Excel uses its project name (VBAProject); this fixture's project is Main.
    assert_eq!(vm.values[6], canon(&Variant::from_string("Main")));
    assert_eq!(
        vm.values[7],
        canon(&Variant::from_string(
            "This key is already associated with an element of this collection"
        ))
    );
}

#[test]
fn missing_declared_interface_raises_430_without_retry() {
    let source = r#"
Public count As Long
Sub Main()
    Dim obj As EarlyHost.Calculator
    Set obj = CreateObject("OxVba.Run5Fixture")
    On Error Resume Next
    count = obj.Count
    Set obj = Nothing
End Sub
"#;
    let vm = run(source, false, DescriptorFault::Interface);
    assert_eq!(run(source, true, DescriptorFault::Interface), vm);
    assert!(!vm.raised);
    assert_eq!(vm.err.number, 430);
    assert_eq!(
        vm.err.description,
        oxvba_runtime::default_error_message(430)
    );
    assert_eq!(vm.transport, (0, 0));
}

#[test]
fn pure_dispinterface_uses_declared_dispid_without_name_lookup() {
    let source = r#"
Public count As Long
Sub Main()
    Dim obj As EarlyHost.Calculator
    Set obj = CreateObject("OxVba.Run5Fixture")
    count = obj.TypedCount
    Set obj = Nothing
End Sub
"#;
    // The server knows Count, not TypedCount; the compiler descriptor's token
    // supplies the binding. A name-based recovery path cannot pass this case.
    let vm = run(source, false, DescriptorFault::DispatchOnly);
    assert_eq!(run(source, true, DescriptorFault::DispatchOnly), vm);
    assert!(!vm.raised, "{vm:?}");
    assert_eq!(vm.values[0], canon(&Variant::from_i32(7)));
    assert_eq!(
        vm.transport,
        (0, 1),
        "explicit dispatch-only transport, never a dual fallback"
    );
}
