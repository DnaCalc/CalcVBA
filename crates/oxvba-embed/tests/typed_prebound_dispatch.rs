#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_host_api as api;
use oxvba_com::{
    OptionalParamDefault as D, TypeLibMemberInvokeKind as K, TypeLibParamType as P,
    TypeLibWireType as W,
};
use oxvba_embed::{Backend, ProcedureKind, provider::Provider};
use oxvba_hal::{
    adapters::builder::HostBuilder,
    model::{HalProfileId, HostPolicy},
};
use std::rc::Rc;
#[path = "support/prebound_dispatch.rs"]
mod fixture;

struct Types;
impl oxvba_symbol::TypeLibResolver for Types {
    fn resolve(
        &self,
        _: &oxvba_com::TypeLibResolveRequest,
    ) -> Option<oxvba_com::TypeLibMetadataBlob> {
        let make =
            |name: &str, token, kind, params: Vec<P>, result| oxvba_com::TypeLibMemberMetadata {
                name: name.into(),
                token,
                vtable_slot: None,
                requires_argument: !params.is_empty(),
                invoke_kind: kind,
                parameter_names: (0..params.len()).map(|i| format!("arg{i}")).collect(),
                parameter_optional: vec![false; params.len()],
                parameter_optional_defaults: vec![D::Required; params.len()],
                is_default_member: token == 0,
                parameter_wire_types: params.iter().copied().map(W::Automation).collect(),
                parameter_iids: vec![None; params.len()],
                parameter_types: params,
                return_type: result,
                return_wire_type: result.map(W::Automation),
                callconv_is_stdcall: true,
                is_dual: false,
                interface_iid: Some(fixture::IID),
                source_typekind: Some(oxvba_com::SourceTypeKind::Dispatch),
                vtable_slot_bound: Some(7),
            };
        let mut members = vec![
            make("Value", 0, K::PropertyGet, vec![], Some(P::Variant)),
            make("Value", 0, K::PropertyPut, vec![P::Variant], None),
            make(
                "Mix",
                12,
                K::Method,
                vec![P::Long, P::Double],
                Some(P::Double),
            ),
            make(
                "Alias",
                13,
                K::Method,
                vec![P::ByRefLong, P::ByRefLong],
                Some(P::Long),
            ),
            make("Fail", 14, K::Method, vec![], Some(P::Double)),
            make("Same", 15, K::PropertyGet, vec![], Some(P::Object)),
            make("Echo", 16, K::Method, vec![P::String], Some(P::String)),
        ];
        members[5].return_wire_type = Some(W::InterfacePointer {
            name: "Dispatch.Native".into(),
        });
        let mut sum = make("Sum", 17, K::Method, vec![P::Variant; 30], Some(P::Double));
        sum.parameter_optional[1..].fill(true);
        sum.parameter_optional_defaults[1..].fill(D::OptionalVariant);
        members.push(sum);
        members.push(make("Opaque", 15, K::PropertyGet, vec![], Some(P::Object)));
        Some(oxvba_com::TypeLibMetadataBlob {
            identity: oxvba_com::TypeLibResolvedIdentity {
                reference_name: "Dispatch".into(),
                requested_coclass: Some("Native".into()),
                importlib: "owned-dispatch-fixture".into(),
                libid: None,
                major_version: 1,
                minor_version: 0,
                lcid: Some(0),
                cache_key: "owned-dispatch-fixture".into(),
            },
            activation_prog_id: Some("Owned.Dispatch".into()),
            member_name_to_token: members.iter().map(|m| (m.name.clone(), m.token)).collect(),
            members,
            events: vec![],
            coclass_names: vec!["Native".into()],
        })
    }
}
fn call(
    session: &oxvba_embed::OwnedSession,
    name: &str,
    args: &[api::ValueSlot<'_>],
    result: api::ValueType,
) -> api::Value {
    let bound = session
        .bind(
            "Api",
            name,
            if result == api::ValueType::Void {
                ProcedureKind::Sub
            } else {
                ProcedureKind::Function
            },
        )
        .unwrap();
    session
        .bind_values(
            &bound,
            &args.iter().map(|a| a.ty()).collect::<Vec<_>>(),
            result,
        )
        .unwrap();
    let mut call = session
        .begin_values(
            &bound,
            args,
            api::InvocationFrame::new(1, None, Rc::new(())),
            1000,
        )
        .unwrap();
    loop {
        if let api::ValueProgress::Returned(value) = call.resume(1).unwrap() {
            return value;
        }
    }
}
#[test]
fn pure_dispatch_uses_retained_id_accessor_and_types_without_lookup_or_retry() {
    for backend in [Backend::Interpreter, Backend::Jit]
        .into_iter()
        .filter(|b| *b != Backend::Jit || cfg!(feature = "jit"))
    {
        let host = HostBuilder::new()
            .profile(HalProfileId::Windows)
            .policy(HostPolicy::interactive_dev())
            .build();
        let mut project = api::Project::module(
            "Api",
            r#"
Public Function ReadValue(ByVal obj As Dispatch.Native) As Variant
    ReadValue = obj.Value
End Function
Public Sub WriteValue(ByVal obj As Dispatch.Native, ByVal value As Variant)
    obj.Value = value
End Sub
Public Function Add(ByVal obj As Dispatch.Native) As Double
    Add = obj.Mix(arg1:=2.5, arg0:=3)
End Function
Public Function AliasValue(ByVal obj As Dispatch.Native) As Long
    Dim value As Long
    value = 2
    AliasValue = obj.Alias(value, value) * 100 + value
End Function
Public Function DefaultValue(ByVal obj As Object) As Double
    DefaultValue = obj + 1#
End Function
Public Function SameObject(ByVal obj As Dispatch.Native) As Boolean
    Dim other As Dispatch.Native
    Set other = obj.Same
    SameObject = obj Is other
End Function
Public Function CastOpaque(ByVal obj As Dispatch.Native) As Double
    Dim raw As Object, typed As Dispatch.Native
    Set raw = obj.Opaque
    Set typed = raw
    CastOpaque = typed.Value
End Function
Public Function Echo(ByVal obj As Dispatch.Native, ByVal text As String) As String
    Echo = obj.Echo(text)
End Function
Public Function OptionalArities(ByVal obj As Dispatch.Native, ByVal x As Variant, ByVal y As Variant) As Double
    OptionalArities = obj.Sum(x) + obj.Sum(x, y)
End Function
Public Function ErrorInfo(ByVal obj As Dispatch.Native) As String
    On Error GoTo Failed
    ErrorInfo = obj.Fail
    Exit Function
Failed:
    ErrorInfo = Err.Number & "|" & Err.Description & "|" & Err.Source & "|" & Err.HelpFile & "|" & Err.HelpContext
End Function
"#,
        );
        project.references.push(api::TypeLibraryReference {
            name: "Dispatch".into(),
            libid: None,
            major: None,
            minor: None,
            lcid: None,
            import_library: Some("owned-dispatch-fixture".into()),
        });
        let module = Provider::with_host(host.clone())
            .with_type_libraries(Rc::new(Types))
            .compile(&project)
            .unwrap();
        let interfaces: Vec<_> = module
            .com_interfaces()
            .iter()
            .filter_map(|i| match i {
                oxvba_oxir::ComInterface::Com(m) => Some(m.clone()),
                _ => None,
            })
            .collect();
        let members = &interfaces
            .iter()
            .find(|i| i.iid == Some(fixture::IID))
            .unwrap()
            .members;
        let (pointer, observed) = fixture::create();
        {
            // SAFETY: fixture transfers one retained same-apartment IDispatch.
            let raw = unsafe {
                host.com()
                    .bind_native_dispatch_object_variant("Owned.Dispatch", pointer.cast())
            }
            .unwrap();
            // SAFETY: these negative mutations either fail descriptor shape
            // admission before native entry or omit the mandatory index table.
            unsafe {
                assert!(
                    host.com()
                        .prepare_typed_object(raw.as_object_ref().unwrap(), members)
                        .is_err()
                );
                for mutation in 0..4 {
                    let mut invalid = members[0].clone();
                    match mutation {
                        0 => invalid.is_dual = true,
                        1 => invalid.vtable_slot = Some(7),
                        2 => {
                            invalid.interface_iid = Some(oxvba_com::ComInterfaceIid {
                                data1: 0,
                                data2: 0,
                                data3: 0,
                                data4: [0; 8],
                            })
                        }
                        _ => invalid.return_wire_type = Some(W::Automation(P::Long)),
                    }
                    assert!(
                        host.com()
                            .prepare_typed_object_with_interfaces(
                                raw.as_object_ref().unwrap(),
                                &[invalid],
                                &interfaces
                            )
                            .is_err()
                    );
                }
            }
            // SAFETY: the exact seven-slot fixture ABI matches these unmodified
            // compiler-retained descriptors and complete interface metadata.
            let object = unsafe {
                host.com().prepare_typed_object_with_interfaces(
                    raw.as_object_ref().unwrap(),
                    members,
                    &interfaces,
                )
            }
            .unwrap();
            host.com()
                .release_object_variant(raw.as_object_ref().unwrap())
                .unwrap();
            let session = module.into_session(backend).unwrap();
            let arg = api::ValueSlot::Object(object.argument());
            let queries = observed.queries.get();
            let x = api::Variant::from_f64(2.0);
            let y = api::Variant::from_f64(3.0);
            assert!(matches!(
                call(
                    &session,
                    "OptionalArities",
                    &[
                        arg,
                        api::ValueSlot::Variant(&x),
                        api::ValueSlot::Variant(&y)
                    ],
                    api::ValueType::Double
                ),
                api::Value::Double(7.0)
            ));
            assert!(matches!(
                call(&session, "Add", &[arg], api::ValueType::Double),
                api::Value::Double(5.5)
            ));
            assert!(matches!(
                call(&session, "AliasValue", &[arg], api::ValueType::Long),
                api::Value::Long(1313)
            ));
            let api::Value::Variant(value) =
                call(&session, "ReadValue", &[arg], api::ValueType::Variant)
            else {
                panic!("Variant expected")
            };
            assert_eq!(value.as_f64(), Some(41.0));
            assert!(matches!(
                call(&session, "DefaultValue", &[arg], api::ValueType::Double),
                api::Value::Double(42.0)
            ));
            assert_eq!(
                observed.queries.get(),
                queries,
                "scalar calls must not QI or look up names/types"
            );
            assert!(matches!(
                call(&session, "CastOpaque", &[arg], api::ValueType::Double),
                api::Value::Double(41.0)
            ));
            let value = api::Variant::from_string("owned dispatch");
            call(
                &session,
                "WriteValue",
                &[arg, api::ValueSlot::Variant(&value)],
                api::ValueType::Void,
            );
            let api::Value::Variant(read) =
                call(&session, "ReadValue", &[arg], api::ValueType::Variant)
            else {
                panic!("Variant expected")
            };
            assert_eq!(read.as_bstr().unwrap().as_str(), "owned dispatch");
            let api::Value::String(echo) = call(
                &session,
                "Echo",
                &[
                    arg,
                    api::ValueSlot::String(
                        &api::BStr::from_utf16_units(&[0xd800, 0, 0xdc00]).unwrap(),
                    ),
                ],
                api::ValueType::String,
            ) else {
                panic!("BSTR expected")
            };
            assert_eq!(echo.to_utf16_units(), vec![0xd800, 0, 0xdc00]);
            assert!(matches!(
                call(&session, "SameObject", &[arg], api::ValueType::Boolean),
                api::Value::Boolean(true)
            ));
            let api::Value::String(error) =
                call(&session, "ErrorInfo", &[arg], api::ValueType::String)
            else {
                panic!("String expected")
            };
            assert_eq!(
                error.as_str(),
                "13|dispatch fixture mismatch|Owned.Dispatch|dispatch-help.chm|71"
            );
            assert_eq!(observed.lookups.get(), 0);
            assert_eq!(observed.calls.borrow().len(), 13);
            assert_eq!(host.com().com_dispatch_transport_counts(), (0, 13));
        }
        assert_eq!(observed.references.get(), 0);
    }
}
