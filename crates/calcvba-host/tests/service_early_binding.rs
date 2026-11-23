//! Host-neutral early binding over service-backed portable host objects and
//! cross-library type-library references (both VM3 and JIT).
use calcvba_com::{
    ComInterfaceIid, OptionalParamDefault as D, SourceTypeKind, TypeLibAppObject,
    TypeLibEventDispatchPath, TypeLibEventMetadata, TypeLibMemberInvokeKind as K,
    TypeLibMemberMetadata, TypeLibMetadataBlob, TypeLibParamType as P, TypeLibResolveRequest,
    TypeLibResolvedIdentity, TypeLibWireType as W,
};
use calcvba_hal::{
    HostServices,
    adapters::builder::HostBuilder,
    error::HalResult,
    model::{CapabilityId, HalDescriptor, HalProfileId, HostPolicy},
    traits::*,
};
use calcvba_host::{Backend, ProcedureKind, provider::Provider};
use calcvba_host_api::{
    self as api,
    object::{
        DynamicMemberKind, EarlyMemberCall, EventCallback, EventDescriptor, EventSubscription,
        HostEnumerator, HostObject, HostObjectServices, InterfaceId,
    },
};
use calcvba_runtime::{ObjectRef, Variant};
use calcvba_symbol::TypeLibResolver;
use std::{
    cell::{Cell, RefCell},
    collections::BTreeMap,
    rc::Rc,
    sync::Arc,
};

const IID_ROOT: InterfaceId = [0x10; 16];
const IID_SHEET: InterfaceId = [0x11; 16];
const IID_SHEET_EVENTS: InterfaceId = [0x12; 16];
const IID_ITEM: InterfaceId = [0x13; 16];
const IID_PANEL: InterfaceId = [0x21; 16];
const IID_CONTENT: InterfaceId = [0x22; 16];

const ROOT_CLSID: &str = "{10101010-1010-1010-1010-101010101010}";
const SHEET_EVENTS_TEXT: &str = "{12121212-1212-1212-1212-121212121212}";

fn com_iid(bytes: InterfaceId) -> ComInterfaceIid {
    ComInterfaceIid {
        data1: u32::from_le_bytes(bytes[..4].try_into().unwrap()),
        data2: u16::from_le_bytes(bytes[4..6].try_into().unwrap()),
        data3: u16::from_le_bytes(bytes[6..8].try_into().unwrap()),
        data4: bytes[8..].try_into().unwrap(),
    }
}

fn disp_member(
    iid: InterfaceId,
    name: &str,
    token: i32,
    kind: K,
    params: &[(&str, P, bool, D)],
    ret: Option<(P, W)>,
    is_default: bool,
) -> TypeLibMemberMetadata {
    TypeLibMemberMetadata {
        name: name.into(),
        token,
        vtable_slot: None,
        requires_argument: params.iter().any(|(_, _, opt, _)| !opt),
        invoke_kind: kind,
        parameter_names: params.iter().map(|(n, _, _, _)| (*n).into()).collect(),
        parameter_optional: params.iter().map(|(_, _, opt, _)| *opt).collect(),
        parameter_optional_defaults: params.iter().map(|(_, _, _, d)| d.clone()).collect(),
        is_default_member: is_default,
        parameter_wire_types: params
            .iter()
            .map(|(_, t, _, _)| W::Automation(*t))
            .collect(),
        parameter_iids: vec![None; params.len()],
        parameter_types: params.iter().map(|(_, t, _, _)| *t).collect(),
        return_type: ret.as_ref().map(|(t, _)| *t),
        return_wire_type: ret.map(|(_, w)| w),
        callconv_is_stdcall: true,
        is_dual: false,
        interface_iid: Some(com_iid(iid)),
        source_typekind: Some(SourceTypeKind::Dispatch),
        vtable_slot_bound: Some(7),
    }
}

#[allow(clippy::too_many_arguments)]
fn dual_member(
    iid: InterfaceId,
    slot: u16,
    bound: u16,
    name: &str,
    token: i32,
    kind: K,
    params: &[(&str, P, bool, D)],
    ret: Option<(P, W)>,
) -> TypeLibMemberMetadata {
    let mut m = disp_member(iid, name, token, kind, params, ret, false);
    m.vtable_slot = Some(slot);
    m.is_dual = true;
    m.source_typekind = Some(SourceTypeKind::Interface);
    m.vtable_slot_bound = Some(bound);
    m
}

struct TwoLibResolver;
impl TypeLibResolver for TwoLibResolver {
    fn interface_iid(
        &self,
        request: &TypeLibResolveRequest,
        _: &TypeLibMetadataBlob,
    ) -> Option<ComInterfaceIid> {
        let name = request.requested_coclass.as_deref()?;
        if request.reference_name.eq_ignore_ascii_case("AlphaLib") {
            match name.to_ascii_lowercase().as_str() {
                "root" => Some(com_iid(IID_ROOT)),
                "sheetdoc" => Some(com_iid(IID_SHEET)),
                "item" => Some(com_iid(IID_ITEM)),
                _ => None,
            }
        } else if request.reference_name.eq_ignore_ascii_case("BetaLib") {
            match name.to_ascii_lowercase().as_str() {
                "panel" => Some(com_iid(IID_PANEL)),
                "content" => Some(com_iid(IID_CONTENT)),
                _ => None,
            }
        } else {
            None
        }
    }

    fn resolve(&self, request: &TypeLibResolveRequest) -> Option<TypeLibMetadataBlob> {
        let id = |lib: &str, coclass: Option<&str>| TypeLibResolvedIdentity {
            reference_name: lib.into(),
            requested_coclass: coclass.map(str::to_owned),
            importlib: format!("test:{lib}"),
            libid: None,
            major_version: 1,
            minor_version: 0,
            lcid: Some(0),
            cache_key: format!("test:{lib}:{}", coclass.unwrap_or("*")),
        };
        if request.reference_name.eq_ignore_ascii_case("AlphaLib") {
            let app_objects = vec![TypeLibAppObject {
                coclass: "Root".into(),
                clsid: ROOT_CLSID.into(),
                creatable: false,
            }];
            let coclasses = vec!["Root".into(), "SheetDoc".into(), "Item".into()];
            let Some(coclass) = request.requested_coclass.as_deref() else {
                return Some(TypeLibMetadataBlob {
                    constants: vec![],
                    identity: id("AlphaLib", None),
                    activation_prog_id: None,
                    member_name_to_token: vec![],
                    members: vec![],
                    events: vec![],
                    coclass_names: coclasses,
                    coclass_prog_ids: vec![],
                    app_objects,
                    extensible: false,
                });
            };
            let (members, events) = match coclass.to_ascii_lowercase().as_str() {
                "root" => (
                    vec![disp_member(
                        IID_ROOT,
                        "Lookup",
                        10,
                        K::PropertyGet,
                        &[("Key", P::String, false, D::Required)],
                        Some((
                            P::Object,
                            W::InterfacePointer {
                                name: "AlphaLib.Item".into(),
                            },
                        )),
                        false,
                    )],
                    vec![],
                ),
                "sheetdoc" => (
                    vec![
                        dual_member(
                            IID_SHEET,
                            197,
                            300,
                            "Name",
                            110,
                            K::PropertyGet,
                            &[],
                            Some((P::String, W::Automation(P::String))),
                        ),
                        dual_member(
                            IID_SHEET,
                            210,
                            300,
                            "Lookup",
                            197,
                            K::PropertyGet,
                            &[("Key", P::String, false, D::Required)],
                            Some((
                                P::Object,
                                W::InterfacePointer {
                                    name: "AlphaLib.Item".into(),
                                },
                            )),
                        ),
                    ],
                    vec![TypeLibEventMetadata {
                        name: "Changed".into(),
                        token: 1545,
                        callback_arity: 1,
                        parameter_types: vec![P::Object],
                        parameter_iids: vec![Some(com_iid(IID_ITEM))],
                        dispatch_path: TypeLibEventDispatchPath::Dispatch,
                        connection_point_iid: Some(SHEET_EVENTS_TEXT.into()),
                        dispatch_member_id: Some(1545),
                        coclass: Some("SheetDoc".into()),
                    }],
                ),
                "item" => (
                    vec![
                        disp_member(
                            IID_ITEM,
                            "Value",
                            0,
                            K::PropertyGet,
                            &[],
                            Some((P::Variant, W::Automation(P::Variant))),
                            true,
                        ),
                        disp_member(
                            IID_ITEM,
                            "Value",
                            0,
                            K::PropertyPut,
                            &[
                                ("Mode", P::Variant, true, D::OptionalVariant),
                                ("RHS", P::Variant, false, D::Required),
                            ],
                            None,
                            true,
                        ),
                        disp_member(
                            IID_ITEM,
                            "Panel",
                            50,
                            K::PropertyGet,
                            &[],
                            Some((
                                P::Object,
                                W::InterfacePointer {
                                    name: "BetaLib.Panel".into(),
                                },
                            )),
                            false,
                        ),
                        disp_member(
                            IID_ITEM,
                            "Combine5",
                            51,
                            K::Method,
                            &[
                                ("A", P::Long, false, D::Required),
                                ("B", P::Long, false, D::Required),
                                ("C", P::Long, false, D::Required),
                                ("D", P::Long, false, D::Required),
                                ("E", P::Long, false, D::Required),
                            ],
                            Some((P::Long, W::Automation(P::Long))),
                            false,
                        ),
                        disp_member(
                            IID_ITEM,
                            "Format",
                            52,
                            K::PropertyGet,
                            &[
                                ("First", P::Variant, true, D::OptionalVariant),
                                ("Second", P::Variant, true, D::OptionalVariant),
                                ("Third", P::Variant, true, D::OptionalVariant),
                                ("Fourth", P::Variant, true, D::OptionalVariant),
                                ("Fifth", P::Variant, true, D::OptionalVariant),
                            ],
                            Some((P::String, W::Automation(P::String))),
                            false,
                        ),
                        disp_member(
                            IID_ITEM,
                            "Bump",
                            53,
                            K::Method,
                            &[("N", P::ByRefLong, false, D::Required)],
                            Some((P::Long, W::Automation(P::Long))),
                            false,
                        ),
                        disp_member(
                            IID_ITEM,
                            "NamedMix",
                            54,
                            K::Method,
                            &[
                                ("Id", P::Long, false, D::Required),
                                ("Label", P::String, true, D::OptionalVariant),
                                ("Flag", P::Boolean, true, D::OptionalVariant),
                                ("OutN", P::ByRefLong, true, D::OptionalVariant),
                            ],
                            Some((P::String, W::Automation(P::String))),
                            false,
                        ),
                        disp_member(
                            IID_ITEM,
                            "DynamicView",
                            55,
                            K::PropertyGet,
                            &[],
                            Some((P::Object, W::Automation(P::Object))),
                            false,
                        ),
                    ],
                    vec![],
                ),
                _ => return None,
            };
            return Some(TypeLibMetadataBlob {
                constants: vec![],
                identity: id("AlphaLib", Some(coclass)),
                activation_prog_id: None,
                member_name_to_token: members.iter().map(|m| (m.name.clone(), m.token)).collect(),
                members,
                events,
                coclass_names: coclasses,
                coclass_prog_ids: vec![],
                app_objects,
                extensible: false,
            });
        }
        if request.reference_name.eq_ignore_ascii_case("BetaLib") {
            let coclasses = vec!["Panel".into(), "Content".into()];
            let Some(coclass) = request.requested_coclass.as_deref() else {
                return Some(TypeLibMetadataBlob {
                    constants: vec![],
                    identity: id("BetaLib", None),
                    activation_prog_id: None,
                    member_name_to_token: vec![],
                    members: vec![],
                    events: vec![],
                    coclass_names: coclasses,
                    coclass_prog_ids: vec![],
                    app_objects: vec![],
                    extensible: false,
                });
            };
            let members = match coclass.to_ascii_lowercase().as_str() {
                "panel" => vec![dual_member(
                    IID_PANEL,
                    36,
                    42,
                    "Content",
                    114,
                    K::PropertyGet,
                    &[],
                    Some((
                        P::Object,
                        W::InterfacePointer {
                            name: "BetaLib.Content".into(),
                        },
                    )),
                )],
                "content" => vec![
                    dual_member(
                        IID_CONTENT,
                        18,
                        51,
                        "Text",
                        0,
                        K::PropertyGet,
                        &[],
                        Some((P::String, W::Automation(P::String))),
                    ),
                    dual_member(
                        IID_CONTENT,
                        19,
                        51,
                        "Text",
                        0,
                        K::PropertyPut,
                        &[("RHS", P::String, false, D::Required)],
                        None,
                    ),
                ],
                _ => return None,
            };
            return Some(TypeLibMetadataBlob {
                constants: vec![],
                identity: id("BetaLib", Some(coclass)),
                activation_prog_id: None,
                member_name_to_token: members.iter().map(|m| (m.name.clone(), m.token)).collect(),
                members,
                events: vec![],
                coclass_names: coclasses,
                coclass_prog_ids: vec![],
                app_objects: vec![],
                extensible: false,
            });
        }
        None
    }
}

#[derive(Default)]
struct World {
    cells: RefCell<BTreeMap<String, Variant>>,
    panel_text: RefCell<String>,
    callbacks: RefCell<Vec<EventCallback>>,
    early_calls: Cell<usize>,
}

struct Node {
    world: Rc<World>,
    kind: NodeKind,
}
#[derive(Clone)]
enum NodeKind {
    Root,
    Sheet(String),
    Item(String),
    Panel,
    Content,
}

fn make_object(world: Rc<World>, kind: NodeKind) -> HostObject {
    let iid = match &kind {
        NodeKind::Root => IID_ROOT,
        NodeKind::Sheet(_) => IID_SHEET,
        NodeKind::Item(_) => IID_ITEM,
        NodeKind::Panel => IID_PANEL,
        NodeKind::Content => IID_CONTENT,
    };
    let node = Rc::new(Node { world, kind });
    let identity = Rc::as_ptr(&node).cast();
    // SAFETY: `node` is retained by both payload and services for the object's lifetime.
    unsafe { HostObject::new_with_services(iid, vec![], node.clone(), identity, node) }
}

impl HostObjectServices for Node {
    fn subscribe_event(
        &self,
        descriptor: &EventDescriptor,
        callback: EventCallback,
    ) -> Result<EventSubscription, i32> {
        if descriptor.interface == IID_SHEET_EVENTS && descriptor.dispatch_id == 1545 {
            self.world.callbacks.borrow_mut().push(callback);
            return Ok(EventSubscription::new(()));
        }
        Err(438)
    }

    fn enumerator(&self, _: &mut api::TypedContext) -> Result<HostEnumerator, i32> {
        let mut items = vec![
            api::Value::Long(10),
            api::Value::Long(20),
            api::Value::Long(30),
        ]
        .into_iter();
        Ok(HostEnumerator::new(move |_| Ok(items.next())))
    }

    fn invoke_dynamic(
        &self,
        call: &api::object::DynamicMemberCall,
        _args: &mut [Variant],
        _: &mut api::TypedContext,
    ) -> Result<Variant, api::TypedFault> {
        if let NodeKind::Item(key) = &self.kind
            && call.name.is_none()
            && matches!(call.kind, DynamicMemberKind::Get)
        {
            return Ok(self
                .world
                .cells
                .borrow()
                .get(key)
                .cloned()
                .unwrap_or_else(Variant::empty));
        }
        Err(api::TypedFault {
            number: 438,
            description: format!("unexpected dynamic {:?}", call.name),
            source: None,
            help_file: None,
            help_context: None,
        })
    }

    fn invoke_early(
        &self,
        call: &EarlyMemberCall<'_>,
        args: &mut [Variant],
        _: &mut api::TypedContext,
    ) -> Result<Variant, api::TypedFault> {
        self.world.early_calls.set(self.world.early_calls.get() + 1);
        match (&self.kind, call.member_id, call.kind) {
            (NodeKind::Root, 10, DynamicMemberKind::Get)
            | (NodeKind::Sheet(_), 197, DynamicMemberKind::Get) => {
                let key = args[0].as_bstr().unwrap().as_str().to_owned();
                let item = make_object(self.world.clone(), NodeKind::Item(key));
                Ok(Variant::from_object_ref(item.object_ref()))
            }
            (NodeKind::Sheet(name), 110, DynamicMemberKind::Get) => {
                assert_eq!(call.vtable_slot, Some(197));
                Ok(Variant::from_string(name.clone()))
            }
            (NodeKind::Item(key), 0, DynamicMemberKind::Get) => Ok(self
                .world
                .cells
                .borrow()
                .get(key)
                .cloned()
                .unwrap_or_else(Variant::empty)),
            (NodeKind::Item(key), 0, DynamicMemberKind::Let) => {
                self.world
                    .cells
                    .borrow_mut()
                    .insert(key.clone(), args.last().unwrap().clone());
                Ok(Variant::empty())
            }
            (NodeKind::Item(_), 50, DynamicMemberKind::Get) => {
                let panel = make_object(self.world.clone(), NodeKind::Panel);
                Ok(Variant::from_object_ref(panel.object_ref()))
            }
            (NodeKind::Item(_), 51, DynamicMemberKind::Method) => {
                let sum: i32 = args.iter().map(|v| v.as_i32().unwrap()).sum();
                Ok(Variant::from_i32(sum))
            }
            (NodeKind::Item(key), 52, DynamicMemberKind::Get) => {
                let fourth = args.get(3).and_then(Variant::as_bool).unwrap_or(false);
                Ok(Variant::from_string(if fourth {
                    format!("[EXT]{key}")
                } else {
                    key.clone()
                }))
            }
            (NodeKind::Item(_), 53, DynamicMemberKind::Method) => {
                let n = args[0].as_i32().unwrap();
                args[0] = Variant::from_i32(n + 5);
                Ok(Variant::from_i32(n * 2))
            }
            (NodeKind::Item(_), 54, DynamicMemberKind::Method) => {
                // Early-bound named arguments are lowered into declared positional slots.
                let id = args[0].as_i32().unwrap();
                let flag = args[2].as_bool().unwrap();
                let cur = args[3].as_i32().unwrap();
                args[3] = Variant::from_i32(cur + 10);
                Ok(Variant::from_string(format!("{id}:{flag}")))
            }
            (NodeKind::Item(key), 55, DynamicMemberKind::Get) => {
                let clone = make_object(self.world.clone(), NodeKind::Item(key.clone()));
                Ok(Variant::from_object_ref(clone.object_ref()))
            }
            (NodeKind::Panel, 114, DynamicMemberKind::Get) => {
                assert_eq!(call.vtable_slot, Some(36));
                let content = make_object(self.world.clone(), NodeKind::Content);
                Ok(Variant::from_object_ref(content.object_ref()))
            }
            (NodeKind::Content, 0, DynamicMemberKind::Get) => {
                assert_eq!(call.vtable_slot, Some(18));
                Ok(Variant::from_string(self.world.panel_text.borrow().clone()))
            }
            (NodeKind::Content, 0, DynamicMemberKind::Let) => {
                assert_eq!(call.vtable_slot, Some(19));
                *self.world.panel_text.borrow_mut() =
                    args[0].as_bstr().unwrap().as_str().to_owned();
                Ok(Variant::empty())
            }
            _ => Err(api::TypedFault {
                number: 438,
                description: format!("unexpected member {} ({:?})", call.member_id, call.kind),
                source: None,
                help_file: None,
                help_context: None,
            }),
        }
    }
}

struct TestHost {
    base: Arc<dyn HostServices>,
    root: std::sync::Mutex<Option<Variant>>,
}
// SAFETY: test runs strictly on one thread.
unsafe impl Send for TestHost {}
// SAFETY: test runs strictly on one thread.
unsafe impl Sync for TestHost {}

impl HostServices for TestHost {
    fn profile(&self) -> HalProfileId {
        self.base.profile()
    }
    fn descriptor(&self) -> HalDescriptor {
        self.base.descriptor()
    }
    fn policy(&self) -> &HostPolicy {
        self.base.policy()
    }
    fn console(&self) -> &dyn ConsoleHal {
        self.base.console()
    }
    fn ui(&self) -> &dyn UiInteractionHal {
        self.base.ui()
    }
    fn events(&self) -> &dyn EventPumpHal {
        self.base.events()
    }
    fn fs(&self) -> &dyn FileSystemHal {
        self.base.fs()
    }
    fn process(&self) -> &dyn ProcessEnvHal {
        self.base.process()
    }
    fn com(&self) -> &dyn ComHal {
        self
    }
    fn time_locale(&self) -> &dyn TimeLocaleHal {
        self.base.time_locale()
    }
    fn dynlink(&self) -> &dyn DynamicLinkHal {
        self.base.dynlink()
    }
    fn diag(&self) -> &dyn DiagnosticsHal {
        self.base.diag()
    }
    fn project_catalog(&self) -> Option<&dyn ProjectCatalogHal> {
        self.base.project_catalog()
    }
    fn project_references(&self) -> Option<&dyn ProjectReferenceHal> {
        self.base.project_references()
    }
    fn project_mutation(&self) -> Option<&dyn ProjectMutationHal> {
        self.base.project_mutation()
    }
}

impl ComHal for TestHost {
    fn describe_object(&self, _: ObjectRef) -> HalResult<Option<calcvba_com::ComObjectDescriptor>> {
        Ok(None)
    }
    fn app_object_variant(&self, class: Variant) -> HalResult<Variant> {
        if class
            .as_bstr()
            .is_some_and(|s| s.as_str().eq_ignore_ascii_case(ROOT_CLSID))
            && let Some(root) = self.root.lock().unwrap().clone()
        {
            return Ok(root);
        }
        Err(calcvba_hal::error::HalError::adapter_fault(
            HalProfileId::Null,
            CapabilityId::ComActivationDispatch,
            "app_object_variant",
            "unknown app object",
        )
        .with_host_error_code(429))
    }
    fn subscribe_event(
        &self,
        _: ObjectRef,
        _: calcvba_com::ComMemberToken,
    ) -> HalResult<calcvba_com::ComSubscriptionToken> {
        unreachable!()
    }
    fn poll_event_callback(&self) -> HalResult<Option<calcvba_com::ComCallbackPayload>> {
        Ok(None)
    }
    fn event_callback_subscription(
        &self,
        _: calcvba_com::ComCallbackToken,
    ) -> HalResult<calcvba_com::ComSubscriptionToken> {
        unreachable!()
    }
    fn event_callback_arity(&self, _: calcvba_com::ComCallbackToken) -> HalResult<usize> {
        unreachable!()
    }
    fn resolve_typelib_reference(
        &self,
        _: &TypeLibResolveRequest,
    ) -> HalResult<TypeLibResolvedIdentity> {
        unreachable!()
    }
    fn load_typelib_metadata(&self, _: &TypeLibResolvedIdentity) -> HalResult<TypeLibMetadataBlob> {
        unreachable!()
    }
    fn invalidate_typelib_cache(
        &self,
        _: calcvba_com::TypeLibCacheScope,
        _: Option<&str>,
    ) -> HalResult<Variant> {
        Ok(Variant::empty())
    }
}

#[test]
fn service_backed_early_binding_handles_high_slots_five_args_named_byref_cross_lib_and_events() {
    let mut project = api::Project::module(
        "Api",
        r#"
Public Summary As String
Public Function ReadSummary() As String
    ReadSummary = Summary
End Function
"#,
    );
    project.modules.push(api::SourceModule {
        name: "Doc1".into(),
        source: r#"
Private Sub SheetDoc_Changed(ByVal Target As AlphaLib.Item)
    Dim n As Long
    Dim bumped As Long
    Dim total As Long
    Dim totalDyn As Long
    Dim eachVal As Variant
    Dim direct As Variant
    Dim dynObj As Variant
    Dim dynVal As Variant
    Dim mixed As String
    Dim outVal As Long
    n = 4
    outVal = 3
    bumped = Target.Bump(n)
    mixed = Target.NamedMix(Flag:=True, OutN:=outVal, Id:=7)
    Target.Panel.Content.Text = Me.Name & ":" & Target.Format(Fourth:=True)
    Me.Lookup("C1").Value = Target.Combine5(1, 2, 3, 4, 5)
    Lookup("D1").Value = Target.Value
    direct = Target
    Set dynObj = Target.DynamicView
    dynVal = dynObj
    For Each eachVal In Target
        total = total + CLng(eachVal)
    Next eachVal
    For Each eachVal In dynObj
        totalDyn = totalDyn + CLng(eachVal)
    Next eachVal
    Summary = Target.Panel.Content.Text & "|" & CStr(Me.Lookup("C1").Value) & "|" & CStr(Lookup("D1").Value) & "|" & CStr(direct) & "|" & CStr(bumped) & ":" & CStr(n) & "|" & CStr(total) & "|" & mixed & ":" & CStr(outVal) & "|" & CStr(totalDyn) & "|" & CStr(dynVal)
End Sub
"#
        .into(),
    });
    project.references = vec![
        api::TypeLibraryReference::named("AlphaLib"),
        api::TypeLibraryReference::named("BetaLib"),
    ];

    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let world = Rc::new(World::default());
        world
            .cells
            .borrow_mut()
            .insert("A1".into(), Variant::from_i32(42));
        let root_obj = make_object(world.clone(), NodeKind::Root);
        let host = Arc::new(TestHost {
            base: HostBuilder::new().build(),
            root: std::sync::Mutex::new(Some(Variant::from_object_ref(root_obj.object_ref()))),
        });
        let module = Provider::with_host(host)
            .with_type_libraries(Rc::new(TwoLibResolver))
            .with_document_modules(vec![("Doc1".into(), "AlphaLib.SheetDoc".into())])
            .compile(&project)
            .unwrap();
        let session = module.into_session(backend).unwrap();
        let sheet = make_object(world.clone(), NodeKind::Sheet("Main".into()));
        session.attach_document("Doc1", &sheet).unwrap();
        assert_eq!(world.callbacks.borrow().len(), 1);

        let target = make_object(world.clone(), NodeKind::Item("A1".into()));
        let mut budget = api::TypedContext {
            remaining_blocks: 100_000,
        };
        let cb = world.callbacks.borrow()[0].clone();
        api::with_native_call(&mut budget, || cb(&mut [api::Value::Object(Some(target))])).unwrap();

        let read = session
            .bind("Api", "ReadSummary", ProcedureKind::Function)
            .unwrap();
        let summary = session.call_values(&read, &[], &mut budget).unwrap();
        let api::Value::String(text) = summary else {
            panic!("expected String, got {summary:?}");
        };
        assert_eq!(
            text.as_str(),
            "Main:[EXT]A1|15|42|42|8:9|60|7:true:13|60|42",
            "{backend:?}"
        );
        assert!(world.early_calls.get() >= 10, "{backend:?}");
    }
}
