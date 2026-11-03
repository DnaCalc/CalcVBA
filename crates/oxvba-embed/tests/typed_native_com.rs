#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_host_api as api;
use oxvba_embed::{Backend, ProcedureKind, provider::Provider};
use oxvba_hal::{
    HostServices,
    adapters::builder::HostBuilder,
    model::{HalProfileId, HostPolicy},
};
use std::{rc::Rc, sync::Arc};
#[path = "support/native_identity.rs"]
mod native_identity;

#[test]
fn distinct_native_interfaces_share_retained_identity_and_failed_qi_does_not_leak() {
    use oxvba_symbol::TypeLibResolver;
    for backend in backends() {
        let host = host();
        let (pointer, observed) = native_identity::create();
        {
            // SAFETY: create transfers one owned same-apartment IDispatch ref.
            let raw = unsafe {
                host.com()
                    .bind_native_dispatch_object_variant("Owned.Identity", pointer.cast())
            }
            .unwrap();
            let metadata = FixtureTypes
                .resolve(&oxvba_com::TypeLibResolveRequest {
                    reference_name: "NativeFixture".into(),
                    requested_coclass: Some("Calculator".into()),
                    importlib_hint: None,
                    libid_hint: None,
                    major_version_hint: None,
                    minor_version_hint: None,
                    lcid_hint: None,
                })
                .unwrap();
            let mut member = metadata.members[0].clone();
            member.vtable_slot_bound = Some(8);
            let first = prepare(
                host.as_ref(),
                raw.as_object_ref().unwrap(),
                &[member.clone()],
            )
            .unwrap();
            member.interface_iid = Some(native_identity::OTHER_IID);
            let second = prepare(
                host.as_ref(),
                raw.as_object_ref().unwrap(),
                &[member.clone()],
            )
            .unwrap();
            assert_ne!(first.argument().as_ptr(), second.argument().as_ptr());
            assert!(first.same_identity(&second));
            observed.refuse_identity.set(true);
            let before = observed.references.get();
            assert!(prepare(host.as_ref(), raw.as_object_ref().unwrap(), &[member]).is_err());
            assert_eq!(observed.references.get(), before);
            observed.refuse_identity.set(false);
            host.com()
                .release_object_variant(raw.as_object_ref().unwrap())
                .unwrap();
            let module = Provider::default().compile(&api::Project::module("Api", "Public Function Same(ByVal a As Object, ByVal b As Object) As Boolean\nSame = a Is b\nEnd Function\n")).unwrap();
            let session = module.into_session(backend).unwrap();
            let queries = observed.queries.get();
            assert!(matches!(
                call_values_sliced(
                    &session,
                    "Same",
                    &[
                        api::ValueSlot::Object(first.argument()),
                        api::ValueSlot::Object(second.argument())
                    ],
                    api::ValueType::Boolean
                ),
                api::Value::Boolean(true)
            ));
            assert_eq!(observed.queries.get(), queries, "Is must not query COM");
            assert_eq!(observed.calls.get(), 0);
            assert_eq!(observed.dynamic.get(), 0);
        }
        assert_eq!(observed.references.get(), 0);
    }
}
fn backends() -> Vec<Backend> {
    let mut values = vec![Backend::Interpreter];
    if cfg!(feature = "jit") {
        values.push(Backend::Jit);
    }
    values
}
fn host() -> Arc<dyn HostServices> {
    HostBuilder::new()
        .profile(HalProfileId::Windows)
        .policy(HostPolicy::interactive_dev())
        .build()
}
fn prepare(
    host: &dyn HostServices,
    object: oxvba_runtime::ObjectRef,
    members: &[oxvba_com::TypeLibMemberMetadata],
) -> oxvba_hal::HalResult<api::HostObject> {
    // SAFETY: accepted descriptors come from the installed typelib or the
    // controlled fixture's exact vtable. Negative mutations either fail shape
    // validation or QI before any member entry can be published/called.
    unsafe { host.com().prepare_typed_object(object, members) }
}

fn project(source: &str) -> api::Project {
    let mut project = api::Project::module("Api", source);
    project.references.push(api::TypeLibraryReference {
        name: "Scripting".into(),
        libid: Some("{420B2830-E718-11CF-893D-00A0C9054228}".into()),
        major: Some(1),
        minor: Some(0),
        lcid: Some(0),
        import_library: None,
    });
    project
}
#[test]
fn installed_excel_typelib_retains_native_function_and_range_plans() {
    let mut project = api::Project::module(
        "Api",
        "Public Function Functions(ByVal app As Excel.Application) As Excel.WorksheetFunction\nSet Functions=app.WorksheetFunction\nEnd Function\nPublic Function SumArray(ByVal functions As Excel.WorksheetFunction, ByVal value As Variant) As Double\nSumArray=functions.Sum(value)\nEnd Function\nPublic Function ReadRange(ByVal value As Excel.Range) As Variant\nReadRange=value.Value\nEnd Function\n",
    );
    project.references.push(api::TypeLibraryReference {
        name: "Excel".into(),
        libid: Some("{00020813-0000-0000-C000-000000000046}".into()),
        major: Some(1),
        minor: Some(9),
        lcid: Some(0),
        import_library: None,
    });
    let module = Provider::with_host(host()).compile(&project).unwrap();
    let mut sums = 0;
    for interface in module.com_interfaces() {
        println!("retained {}", interface.name());
        for member in interface
            .com_members()
            .into_iter()
            .flatten()
            .filter(|m| m.name == "Sum" || m.name == "WorksheetFunction")
        {
            println!("selected: {member:?}");
            if member.name == "Sum" {
                sums += 1;
                assert!(!member.is_dual);
                assert_eq!(
                    member.source_typekind,
                    Some(oxvba_com::SourceTypeKind::Dispatch)
                );
                assert_eq!(member.vtable_slot, None);
                assert_eq!(member.vtable_slot_bound, Some(7));
                assert_eq!(
                    oxvba_com::typed_slots::plan::visible_signature(member, 1),
                    Some((vec![api::ValueType::Variant], api::ValueType::Double))
                );
            }
        }
    }
    assert_eq!(sums, 1);
}
fn call_values_sliced(
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
    let types: Vec<_> = args.iter().map(|arg| arg.ty()).collect();
    session.bind_values(&bound, &types, result).unwrap();
    let mut call = session
        .begin_values(
            &bound,
            args,
            api::InvocationFrame::new(10, None, Rc::new(())),
            10000,
        )
        .unwrap();
    loop {
        match call
            .resume(1)
            .unwrap_or_else(|error| panic!("{name}: {error:?}"))
        {
            api::ValueProgress::Suspended => {}
            api::ValueProgress::Returned(value) => return value,
        }
    }
}
#[test]
fn real_dictionary_variant_members_and_array_values_use_only_native_slots() {
    for backend in backends() {
        let host = host();
        let module = Provider::with_host(host.clone())
            .compile(&project(
                r#"
Public Sub AddValue(ByVal d As Scripting.Dictionary, ByVal key As Variant, ByVal value As Variant)
    d.Add key, value
End Sub
Public Function HasValue(ByVal d As Scripting.Dictionary, ByVal key As Variant) As Boolean
    HasValue = d.Exists(key)
End Function
Public Function ReadValue(ByVal d As Scripting.Dictionary, ByVal key As Variant) As Variant
    ReadValue = d.Item(key)
End Function
Public Sub PutValue(ByVal d As Scripting.Dictionary, ByVal key As Variant, ByVal value As Variant)
    d.Item(key) = value
End Sub
Public Function ReadKeys(ByVal d As Scripting.Dictionary) As Variant
    ReadKeys = d.Keys
End Function
Public Sub ClearValues(ByVal d As Scripting.Dictionary)
    d.RemoveAll
End Sub
Public Function Duplicate(ByVal d As Scripting.Dictionary, ByVal key As Variant) As String
    On Error GoTo Failed
    d.Add key, key
    Exit Function
Failed:
    Duplicate = Err.Description
End Function
Public Function Same(ByVal left As Variant, ByVal right As Variant) As Boolean
    Same = left Is right
End Function
"#,
            ))
            .unwrap();
        let members: Vec<_> = module
            .com_interfaces()
            .iter()
            .flat_map(|i| i.com_members().into_iter().flatten())
            .filter(|m| {
                ["Add", "Exists", "Item", "Keys", "RemoveAll"]
                    .iter()
                    .any(|name| m.name.eq_ignore_ascii_case(name))
            })
            .cloned()
            .collect();
        for member in &members {
            println!("Dictionary imported member: {member:?}");
        }
        let raw = host
            .com()
            .create_object_variant(api::Variant::from_string("Scripting.Dictionary"))
            .unwrap();
        let object = prepare(host.as_ref(), raw.as_object_ref().unwrap(), &members).unwrap();
        host.com()
            .release_object_variant(raw.as_object_ref().unwrap())
            .unwrap();
        let session = module.into_session(backend).unwrap();
        let baseline = host.com().com_dispatch_transport_counts();
        let array =
            api::Variant::from_safearray(oxvba_runtime::safe_array::SafeArray::from_variants(
                vec![api::Variant::from_i32(1), api::Variant::from_string("two")],
            ));
        for (index, input) in [
            api::Variant::from_f64(42.0),
            api::Variant::null(),
            api::Variant::from_error_code(2042),
            api::Variant::from_utf16_units(&[0xd800, 0, 0xdc00]),
            array,
        ]
        .into_iter()
        .enumerate()
        {
            let key = api::Variant::from_string(format!("key{index}"));
            let args = [
                api::ValueSlot::Object(object.argument()),
                api::ValueSlot::Variant(&key),
            ];
            assert!(matches!(
                call_values_sliced(
                    &session,
                    "AddValue",
                    &[args[0], args[1], api::ValueSlot::Variant(&input)],
                    api::ValueType::Void
                ),
                api::Value::Empty
            ));
            assert!(matches!(
                call_values_sliced(&session, "HasValue", &args, api::ValueType::Boolean),
                api::Value::Boolean(true)
            ));
            let api::Value::Variant(actual) =
                call_values_sliced(&session, "ReadValue", &args, api::ValueType::Variant)
            else {
                panic!("Variant expected")
            };
            assert_eq!(actual, input);
        }
        let key = api::Variant::from_string("key0");
        let changed = api::Variant::from_i32(73);
        call_values_sliced(
            &session,
            "PutValue",
            &[
                api::ValueSlot::Object(object.argument()),
                api::ValueSlot::Variant(&key),
                api::ValueSlot::Variant(&changed),
            ],
            api::ValueType::Void,
        );
        assert!(
            matches!(call_values_sliced(&session, "ReadValue", &[api::ValueSlot::Object(object.argument()),api::ValueSlot::Variant(&key)], api::ValueType::Variant), api::Value::Variant(v) if v == changed)
        );
        assert!(
            matches!(call_values_sliced(&session, "Duplicate", &[api::ValueSlot::Object(object.argument()),api::ValueSlot::Variant(&key)], api::ValueType::String), api::Value::String(v) if !v.is_empty())
        );
        let api::Value::Variant(keys) = call_values_sliced(
            &session,
            "ReadKeys",
            &[api::ValueSlot::Object(object.argument())],
            api::ValueType::Variant,
        ) else {
            panic!("array expected")
        };
        assert_eq!(keys.safearray_bounds_len().unwrap().1, 5);
        let counts = host.com().com_dispatch_transport_counts();
        assert_eq!(counts, (baseline.0 + 19, baseline.1));
        // Round-trip a separately owned native object through a declared Variant.
        // The native cell must carry IDispatch obtained through QI, and the
        // opaque result must retain its own independent native owner.
        let (pointer, observed) = oxvba_com::create_oxvba_dual_vtable_object_observed();
        {
            use oxvba_symbol::TypeLibResolver;
            // SAFETY: the fresh fixture transfers one same-apartment dispatch ref.
            let raw = unsafe {
                host.com()
                    .bind_native_dispatch_object_variant("Typed.Payload", pointer)
            }
            .unwrap();
            let metadata = FixtureTypes
                .resolve(&oxvba_com::TypeLibResolveRequest {
                    reference_name: "NativeFixture".into(),
                    requested_coclass: Some("Calculator".into()),
                    importlib_hint: None,
                    libid_hint: None,
                    major_version_hint: None,
                    minor_version_hint: None,
                    lcid_hint: None,
                })
                .unwrap();
            let fixture = prepare(
                host.as_ref(),
                raw.as_object_ref().unwrap(),
                &metadata.members,
            )
            .unwrap();
            host.com()
                .release_object_variant(raw.as_object_ref().unwrap())
                .unwrap();
            let payload = api::Variant::from_object_ref(fixture.object_ref());
            let key = api::Variant::from_string("native object");
            let args = [
                api::ValueSlot::Object(object.argument()),
                api::ValueSlot::Variant(&key),
            ];
            call_values_sliced(
                &session,
                "AddValue",
                &[args[0], args[1], api::ValueSlot::Variant(&payload)],
                api::ValueType::Void,
            );
            let result = call_values_sliced(&session, "ReadValue", &args, api::ValueType::Variant);
            assert!(
                matches!(&result, api::Value::Variant(value) if value.as_object_ref().is_some())
            );
            let api::Value::Variant(returned) = &result else {
                unreachable!()
            };
            let before = host.com().com_dispatch_transport_counts();
            assert!(matches!(
                call_values_sliced(
                    &session,
                    "Same",
                    &[
                        api::ValueSlot::Variant(&payload),
                        api::ValueSlot::Variant(returned)
                    ],
                    api::ValueType::Boolean
                ),
                api::Value::Boolean(true)
            ));
            assert_eq!(host.com().com_dispatch_transport_counts(), before);
            call_values_sliced(&session, "ClearValues", &[args[0]], api::ValueType::Void);
            assert!(
                observed
                    .references
                    .load(std::sync::atomic::Ordering::SeqCst)
                    > 0
            );
            drop(result);
        }
        assert_eq!(
            observed
                .references
                .load(std::sync::atomic::Ordering::SeqCst),
            0
        );
        assert_eq!(
            host.com().com_dispatch_transport_counts(),
            (baseline.0 + 22, baseline.1)
        );
    }
}
#[test]
fn real_dictionary_imported_members_use_typed_slots_on_both_runtimes() {
    for backend in backends() {
        let host = host();
        let project = project(
            r#"
Public Function Probe(ByVal d As Scripting.Dictionary) As Long
    d.CompareMode = 1
    Probe = d.Count + d.CompareMode
End Function
Public Sub Clear(ByVal d As Scripting.Dictionary)
    d.RemoveAll
End Sub
Public Function FailedPut(ByVal d As Scripting.Dictionary, ByVal bad As Long) As String
    On Error GoTo Describe
    d.CompareMode = bad
    Exit Function
Describe:
    FailedPut = Err.Description
End Function
"#,
        );
        let module = Provider::with_host(host.clone()).compile(&project).unwrap();
        let members: Vec<_> = module
            .com_interfaces()
            .iter()
            .flat_map(|i| i.com_members().into_iter().flatten())
            .filter(|m| oxvba_com::typed_slots::scalar_signature(m).is_some())
            .cloned()
            .collect();
        assert!(members.iter().any(|m| m.name.eq_ignore_ascii_case("Count")));
        for m in &members {
            assert!(m.is_dual && m.callconv_is_stdcall);
            assert_eq!(
                m.source_typekind,
                Some(oxvba_com::SourceTypeKind::Interface)
            );
        }
        let raw = host
            .com()
            .create_object_variant(api::Variant::from_string("Scripting.Dictionary"))
            .unwrap();
        let reference = raw.as_object_ref().unwrap();
        let count = members
            .iter()
            .find(|m| m.name.eq_ignore_ascii_case("Count"))
            .unwrap()
            .clone();
        let mut bad = count.clone();
        bad.vtable_slot_bound = Some(bad.vtable_slot.unwrap());
        assert!(prepare(host.as_ref(), reference.clone(), &[bad]).is_err());
        let mut bad = count.clone();
        bad.callconv_is_stdcall = false;
        assert!(prepare(host.as_ref(), reference.clone(), &[bad]).is_err());
        let mut bad = count.clone();
        bad.source_typekind = Some(oxvba_com::SourceTypeKind::Dispatch);
        assert!(prepare(host.as_ref(), reference.clone(), &[bad]).is_err());
        let mut wrong_iid = count.clone();
        wrong_iid.interface_iid = Some(oxvba_com::ComInterfaceIid {
            data1: 0x12345678,
            data2: 0x1234,
            data3: 0x1234,
            data4: [1; 8],
        });
        assert!(prepare(host.as_ref(), reference.clone(), &[wrong_iid]).is_err());
        let other_host = host.clone();
        let other_members = members.clone();
        let raw_identity = reference.raw();
        let error = std::thread::spawn(move || {
            prepare(
                other_host.as_ref(),
                oxvba_runtime::ObjectRef::from_compat_identity(raw_identity),
                &other_members,
            )
            .err()
            .map(|e| e.to_string())
        })
        .join()
        .unwrap()
        .unwrap();
        assert!(error.contains("different thread/apartment"));
        let object = prepare(host.as_ref(), reference.clone(), &members).unwrap();
        // The typed receiver owns its QI reference independently of the HAL binding.
        host.com()
            .release_object_variant(reference.clone())
            .unwrap();
        let session = module.into_session(backend).unwrap();
        let bound = session
            .bind("Api", "Probe", ProcedureKind::Function)
            .unwrap();
        session
            .bind_values(&bound, &[api::ValueType::HostObject], api::ValueType::Long)
            .unwrap();
        let frame = api::InvocationFrame::new(1, None, Rc::new(()));
        let mut call = session
            .begin_values(
                &bound,
                &[api::ValueSlot::Object(object.argument())],
                frame,
                1000,
            )
            .unwrap();
        loop {
            match call.resume(1).unwrap() {
                api::ValueProgress::Suspended => {}
                api::ValueProgress::Returned(api::Value::Long(value)) => {
                    assert_eq!(value, 1);
                    break;
                }
                _ => panic!("Long expected"),
            }
        }
        assert_eq!(host.com().com_dispatch_transport_counts(), (3, 0));
        let bound = session.bind("Api", "Clear", ProcedureKind::Sub).unwrap();
        session
            .bind_values(&bound, &[api::ValueType::HostObject], api::ValueType::Void)
            .unwrap();
        assert!(matches!(
            session
                .call_values(
                    &bound,
                    &[api::ValueSlot::Object(object.argument())],
                    &mut api::TypedContext {
                        remaining_blocks: 1000
                    }
                )
                .unwrap(),
            api::Value::Empty
        ));
        assert_eq!(host.com().com_dispatch_transport_counts(), (4, 0));
        let bound = session
            .bind("Api", "FailedPut", ProcedureKind::Function)
            .unwrap();
        session
            .bind_values(
                &bound,
                &[api::ValueType::HostObject, api::ValueType::Long],
                api::ValueType::String,
            )
            .unwrap();
        let api::Value::String(description) = session
            .call_values(
                &bound,
                &[
                    api::ValueSlot::Object(object.argument()),
                    api::ValueSlot::Long(-999),
                ],
                &mut api::TypedContext {
                    remaining_blocks: 1000,
                },
            )
            .unwrap()
        else {
            panic!("String expected")
        };
        assert_eq!(description.as_str(), "Invalid procedure call or argument");
        assert_eq!(host.com().com_dispatch_transport_counts(), (4, 0));
        drop(call);
        drop(object);
    }
}

#[test]
#[ignore = "requires PID/HWND from the owned Excel UIA oracle supervisor"]
fn owned_excel_imported_version_and_build_use_typed_slots() {
    use windows_sys::Win32::{
        System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx, CoUninitialize},
        UI::{
            Accessibility::AccessibleObjectFromWindow,
            WindowsAndMessaging::{EnumChildWindows, GetClassNameW, GetWindowThreadProcessId},
        },
    };
    struct ComScope;
    impl Drop for ComScope {
        fn drop(&mut self) {
            // SAFETY: balances this test's successful STA initialization.
            unsafe {
                CoUninitialize();
            }
        }
    }
    // SAFETY: initialize the current test thread; no pointer argument is retained.
    let initialized = unsafe { CoInitializeEx(std::ptr::null(), COINIT_APARTMENTTHREADED as u32) };
    assert!(
        initialized >= 0,
        "COM initialization failed: {initialized:08x}"
    );
    let _com = ComScope;
    let expected_pid: u32 = std::env::var("OXVBA_OWNED_EXCEL_PID")
        .unwrap()
        .parse()
        .unwrap();
    let hwnd_number: i32 = std::env::var("OXVBA_OWNED_EXCEL_HWND")
        .unwrap()
        .parse()
        .unwrap();
    assert_ne!(expected_pid, 0);
    assert_ne!(expected_pid, 3144);
    assert_ne!(hwnd_number, 0);
    let hwnd = hwnd_number as u32 as usize as windows_sys::Win32::Foundation::HWND;
    let mut actual = 0;
    // SAFETY: read only; actual is one aligned writable process-ID cell.
    assert_ne!(unsafe { GetWindowThreadProcessId(hwnd, &mut actual) }, 0);
    assert_eq!(actual, expected_pid);
    struct Search {
        process: u32,
        window: windows_sys::Win32::Foundation::HWND,
    }
    unsafe extern "system" fn find(
        hwnd: windows_sys::Win32::Foundation::HWND,
        parameter: isize,
    ) -> i32 {
        // SAFETY: synchronous EnumChildWindows retains this exact live Search.
        let search = unsafe { &mut *(parameter as *mut Search) };
        let mut process = 0;
        let mut class = [0u16; 64];
        // SAFETY: read-only window queries, initialized aligned output buffers.
        let len = unsafe {
            GetWindowThreadProcessId(hwnd, &mut process);
            GetClassNameW(hwnd, class.as_mut_ptr(), class.len() as i32)
        };
        if process == search.process
            && len == 6
            && class[..6] == [69, 88, 67, 69, 76, 55]
            && search.window.is_null()
        {
            search.window = hwnd;
        }
        1
    }
    let mut search = Search {
        process: expected_pid,
        window: std::ptr::null_mut(),
    };
    // SAFETY: enumerate only children of the checked owner window; Search
    // remains a unique live stack object until the callback walk returns.
    unsafe {
        EnumChildWindows(hwnd, Some(find), (&mut search as *mut Search) as isize);
    }
    assert!(!search.window.is_null(), "owned EXCEL7 child must be live");
    let iid = windows_sys::core::GUID {
        data1: 0x00020400,
        data2: 0,
        data3: 0,
        data4: [0xc0, 0, 0, 0, 0, 0, 0, 0x46],
    };
    let mut pointer = std::ptr::null_mut();
    // SAFETY: query the enumerated owned EXCEL7 public NativeOM; IID/output
    // cells remain live. A successful result transfers one IDispatch ref.
    let hr = unsafe { AccessibleObjectFromWindow(search.window, 0xfffffff0, &iid, &mut pointer) };
    assert!(
        hr >= 0 && !pointer.is_null(),
        "owned NativeOM failed: {hr:08x}"
    );
    let host = host();
    // SAFETY: successful NativeOM returned one same-apartment IDispatch reference.
    let window = unsafe {
        host.com()
            .bind_native_dispatch_object_variant("Owned.ExcelWindow", pointer)
    }
    .unwrap();
    let setup =
        oxvba_embed::Module::load("Setup", include_str!("support/prebound_excel_setup.bas"))
            .unwrap()
            .with_host(host.clone())
            .into_session(Backend::Interpreter)
            .unwrap();
    let bound = setup.bind("Setup", "App", ProcedureKind::Function).unwrap();
    let app = setup
        .call_bound(&bound, std::slice::from_ref(&window), 1000)
        .unwrap();
    let reference = app.as_object_ref().unwrap();
    let workbook_name = setup
        .bind("Setup", "WorkbookName", ProcedureKind::Function)
        .unwrap();
    let workbook_name = setup
        .call_bound(&workbook_name, std::slice::from_ref(&app), 1000)
        .unwrap();
    assert_eq!(
        workbook_name.as_bstr().unwrap().as_str(),
        std::env::var("OXVBA_OWNED_EXCEL_WORKBOOK").unwrap(),
        "refuse writes outside the supervisor's exact owned workbook"
    );
    for backend in backends() {
        let mut project = api::Project::module("Api", include_str!("support/prebound_excel.bas"));
        project.references.push(api::TypeLibraryReference {
            name: "Excel".into(),
            libid: Some("{00020813-0000-0000-C000-000000000046}".into()),
            major: Some(1),
            minor: Some(9),
            lcid: Some(0),
            import_library: None,
        });
        let module = Provider::with_host(host.clone()).compile(&project).unwrap();
        let parameter_iid = |procedure: &str| {
            let procedure = module.procedures().find(|p| p.name == procedure).unwrap();
            let oxvba_oxir::OxTy::Object(oxvba_oxir::ObjClass::ComIface(id)) =
                procedure.parameters[0].ty
            else {
                panic!("nominal parameter expected")
            };
            let oxvba_oxir::ComInterface::Com(meta) = &module.com_interfaces()[id.0] else {
                panic!("native interface expected")
            };
            meta.iid.unwrap()
        };
        let app_iid = parameter_iid("Version");
        let range_iid = parameter_iid("ReadRange");
        let interfaces: Vec<_> = module
            .com_interfaces()
            .iter()
            .filter_map(|i| match i {
                oxvba_oxir::ComInterface::Com(m) => Some(m.clone()),
                _ => None,
            })
            .collect();
        let members: Vec<_> = module
            .com_interfaces()
            .iter()
            .flat_map(|i| i.com_members().into_iter().flatten())
            .filter(|m| {
                m.interface_iid == Some(app_iid)
                    && m.invoke_kind == oxvba_com::TypeLibMemberInvokeKind::PropertyGet
                    && ["Version", "Build", "Height", "Visible", "WorksheetFunction"]
                        .iter()
                        .any(|name| m.name.eq_ignore_ascii_case(name))
            })
            .cloned()
            .collect();
        assert_eq!(members.len(), 5);
        for member in &members {
            println!("imported Excel member: {member:?}");
            assert!(oxvba_com::typed_slots::signature(member).is_some());
            assert!(
                member.name == "WorksheetFunction"
                    || member
                        .parameter_optional_defaults
                        .contains(&oxvba_com::OptionalParamDefault::Lcid)
            );
        }
        // SAFETY: unmodified compiler-retained Excel interfaces in this apartment.
        let object = unsafe {
            host.com().prepare_typed_object_with_interfaces(
                reference.clone(),
                &members,
                &interfaces,
            )
        }
        .unwrap();
        let range_members: Vec<_> = interfaces
            .iter()
            .flat_map(|i| &i.members)
            .filter(|m| {
                m.interface_iid == Some(range_iid)
                    && (m.name.eq_ignore_ascii_case("Value")
                        || m.name.eq_ignore_ascii_case("Value2")
                        || m.is_default_member)
                    && oxvba_com::typed_slots::plan::validate(m)
            })
            .cloned()
            .collect();
        for member in &range_members {
            println!("imported Range member: {member:?}");
        }
        let session = module.into_session(backend).unwrap();
        let before = host.com().com_dispatch_transport_counts();
        for (name, ty) in [
            ("Version", api::ValueType::String),
            ("Build", api::ValueType::Long),
            ("Height", api::ValueType::Double),
            ("Visible", api::ValueType::Boolean),
        ] {
            let bound = session.bind("Api", name, ProcedureKind::Function).unwrap();
            session
                .bind_values(&bound, &[api::ValueType::HostObject], ty)
                .unwrap();
            let mut call = session
                .begin_values(
                    &bound,
                    &[api::ValueSlot::Object(object.argument())],
                    api::InvocationFrame::new(1, None, Rc::new(())),
                    1000,
                )
                .unwrap();
            loop {
                match call.resume(1).unwrap() {
                    api::ValueProgress::Suspended => {}
                    api::ValueProgress::Returned(value) => {
                        println!("{backend:?} {name}={value:?}");
                        match value {
                            api::Value::String(v) => assert!(!v.is_empty()),
                            api::Value::Long(v) => assert!(v > 0),
                            api::Value::Double(v) => assert!(v.is_finite() && v > 0.0),
                            api::Value::Boolean(_) => assert_eq!(name, "Visible"),
                            _ => panic!("wrong native carrier"),
                        };
                        break;
                    }
                }
            }
        }
        assert_eq!(
            host.com().com_dispatch_transport_counts(),
            (before.0 + 4, before.1)
        );
        let baseline = host.com().com_dispatch_transport_counts();
        let api::Value::Object(Some(functions)) = call_values_sliced(
            &session,
            "GetFunctions",
            &[api::ValueSlot::Object(object.argument())],
            api::ValueType::HostObject,
        ) else {
            panic!("WorksheetFunction expected")
        };
        let array =
            api::Variant::from_safearray(oxvba_runtime::safe_array::SafeArray::from_variants(
                vec![api::Variant::from_f64(2.0), api::Variant::from_f64(3.0)],
            ));
        for input in [api::Variant::from_f64(7.0), array] {
            let api::Value::Double(sum) = call_values_sliced(
                &session,
                "SumArray",
                &[
                    api::ValueSlot::Object(functions.argument()),
                    api::ValueSlot::Variant(&input),
                ],
                api::ValueType::Double,
            ) else {
                panic!("Double sum expected")
            };
            assert_eq!(
                sum,
                if input.as_safearray().is_some() {
                    5.0
                } else {
                    7.0
                }
            );
        }
        let two = api::Variant::from_f64(2.0);
        let three = api::Variant::from_f64(3.0);
        assert!(matches!(
            call_values_sliced(
                &session,
                "SumDifferentArities",
                &[
                    api::ValueSlot::Object(functions.argument()),
                    api::ValueSlot::Variant(&two),
                    api::ValueSlot::Variant(&three)
                ],
                api::ValueType::Double
            ),
            api::Value::Double(7.0)
        ));
        assert!(matches!(
            call_values_sliced(
                &session,
                "AverageArray",
                &[
                    api::ValueSlot::Object(functions.argument()),
                    api::ValueSlot::Variant(&api::Variant::from_safearray(
                        oxvba_runtime::safe_array::SafeArray::from_variants(vec![
                            two.clone(),
                            three.clone()
                        ])
                    ))
                ],
                api::ValueType::Double
            ),
            api::Value::Double(2.5)
        ));
        assert!(matches!(
            call_values_sliced(
                &session,
                "SumTyped",
                &[
                    api::ValueSlot::Object(functions.argument()),
                    api::ValueSlot::Long(2),
                    api::ValueSlot::Double(3.0)
                ],
                api::ValueType::Double
            ),
            api::Value::Double(5.0)
        ));
        let strings = api::Variant::from_safearray(
            oxvba_runtime::safe_array::SafeArray::from_variants(vec![
                api::Variant::from_string("a"),
                api::Variant::from_string("b"),
            ]),
        );
        assert!(matches!(
            call_values_sliced(
                &session,
                "MissingMatch",
                &[
                    api::ValueSlot::Object(functions.argument()),
                    api::ValueSlot::Variant(&strings)
                ],
                api::ValueType::Long
            ),
            api::Value::Long(1004)
        ));
        for setup_name in ["Cell", "Block"] {
            let bound = setup
                .bind("Setup", setup_name, ProcedureKind::Function)
                .unwrap();
            let range = setup
                .call_bound(&bound, std::slice::from_ref(&app), 1000)
                .unwrap();
            // SAFETY: exact compiler-retained Range metadata and live owned Excel
            // receiver on this apartment; library also owns nominal result plans.
            let range_object = unsafe {
                host.com().prepare_typed_object_with_interfaces(
                    range.as_object_ref().unwrap(),
                    &range_members,
                    &interfaces,
                )
            }
            .unwrap();
            let arg = api::ValueSlot::Object(range_object.argument());
            let reads = ["ReadRange", "ReadValue2", "ReadDefault"];
            let writes = ["WriteRange", "WriteValue2", "WriteDefault"];
            for (read, write) in reads.into_iter().zip(writes) {
                let scalar = api::Variant::from_f64(42.0);
                let counts = host.com().com_dispatch_transport_counts();
                call_values_sliced(
                    &session,
                    write,
                    &[arg, api::ValueSlot::Variant(&scalar)],
                    api::ValueType::Void,
                );
                let api::Value::Variant(value) =
                    call_values_sliced(&session, read, &[arg], api::ValueType::Variant)
                else {
                    panic!("Range Variant expected")
                };
                if setup_name == "Cell" {
                    assert_eq!(
                        value.as_f64(),
                        Some(42.0),
                        "{backend:?} {write}/{read}: {value:?}"
                    );
                } else {
                    let array = value.as_safearray().expect("block must return 2D array");
                    assert_eq!(array.dimensions(), 2);
                    assert_eq!(array.variant_element(0).unwrap().as_f64(), Some(42.0));
                    assert_eq!(array.variant_element(3).unwrap().as_f64(), Some(42.0));
                }
                assert_eq!(
                    host.com().com_dispatch_transport_counts(),
                    (counts.0, counts.1 + 2)
                );
            }
            if setup_name == "Cell" {
                assert!(matches!(
                    call_values_sliced(&session, "RuntimeDefault", &[arg], api::ValueType::Double),
                    api::Value::Double(43.0)
                ));
            }
            assert!(
                matches!(call_values_sliced(&session, "SumRange", &[api::ValueSlot::Object(object.argument()), arg], api::ValueType::Double), api::Value::Double(v) if v == if setup_name == "Cell" {42.0} else {168.0})
            );
            let bounds = vec![oxvba_runtime::safe_array::SafeArrayBound { count: 2, lower: 1 }; 2];
            let packet = api::Variant::from_safearray(
                oxvba_runtime::safe_array::SafeArray::from_variants_nd(
                    bounds.clone(),
                    vec![
                        api::Variant::from_f64(1.25),
                        api::Variant::from_string("alpha"),
                        api::Variant::from_bool(true),
                        api::Variant::from_error_code(2042),
                    ],
                ),
            );
            for (read, write) in reads.into_iter().zip(writes) {
                let counts = host.com().com_dispatch_transport_counts();
                call_values_sliced(
                    &session,
                    write,
                    &[arg, api::ValueSlot::Variant(&packet)],
                    api::ValueType::Void,
                );
                let api::Value::Variant(value) =
                    call_values_sliced(&session, read, &[arg], api::ValueType::Variant)
                else {
                    panic!("Range Variant expected")
                };
                if setup_name == "Cell" {
                    assert_eq!(value.as_f64(), Some(1.25));
                } else {
                    let array = value.as_safearray().expect("block must return 2D array");
                    assert_eq!(array.bounds().unwrap(), bounds);
                    assert_eq!(array.variant_element(0).unwrap().as_f64(), Some(1.25));
                    assert_eq!(array.variant_element(2).unwrap().as_bool(), Some(true));
                    assert_eq!(
                        array
                            .variant_element(1)
                            .unwrap()
                            .as_bstr()
                            .unwrap()
                            .as_str(),
                        "alpha"
                    );
                    assert_eq!(
                        array.variant_element(3).unwrap().as_error_code(),
                        Some(2042)
                    );
                }
                assert_eq!(
                    host.com().com_dispatch_transport_counts(),
                    (counts.0, counts.1 + 2)
                );
                if setup_name == "Block" {
                    // Read separate actual cells: a writer/reader array roundtrip
                    // alone could hide a transposition in both directions. Core
                    // flat storage is row-major; VBA's packet is A1=1.25,
                    // B1=alpha, A2=True, B2=CVErr(2042).
                    for (helper, boolean) in [("CellA2", true), ("CellB1", false)] {
                        let at = setup
                            .bind("Setup", helper, ProcedureKind::Function)
                            .unwrap();
                        let cell = setup
                            .call_bound(&at, std::slice::from_ref(&app), 1000)
                            .unwrap();
                        // SAFETY: exact compiler Range table and owned workbook cell.
                        let cell_object = unsafe {
                            host.com().prepare_typed_object_with_interfaces(
                                cell.as_object_ref().unwrap(),
                                &range_members,
                                &interfaces,
                            )
                        }
                        .unwrap();
                        let api::Value::Variant(value) = call_values_sliced(
                            &session,
                            read,
                            &[api::ValueSlot::Object(cell_object.argument())],
                            api::ValueType::Variant,
                        ) else {
                            panic!("cell Variant expected")
                        };
                        if boolean {
                            assert_eq!(value.as_bool(), Some(true));
                        } else {
                            assert_eq!(value.as_bstr().unwrap().as_str(), "alpha");
                        }
                        host.com()
                            .release_object_variant(cell.as_object_ref().unwrap())
                            .unwrap();
                    }
                }
            }
            if setup_name == "Cell" {
                for (read, write) in reads.into_iter().zip(writes) {
                    for scalar in [
                        api::Variant::from_bool(true),
                        api::Variant::from_string("text"),
                        api::Variant::from_error_code(2042),
                        api::Variant::empty(),
                    ] {
                        call_values_sliced(
                            &session,
                            write,
                            &[arg, api::ValueSlot::Variant(&scalar)],
                            api::ValueType::Void,
                        );
                        let api::Value::Variant(value) =
                            call_values_sliced(&session, read, &[arg], api::ValueType::Variant)
                        else {
                            panic!("scalar Variant expected")
                        };
                        assert_eq!(value.vtype(), scalar.vtype());
                        match value.vtype() {
                            oxvba_runtime::VarType::Boolean => {
                                assert_eq!(value.as_bool(), scalar.as_bool())
                            }
                            oxvba_runtime::VarType::String => {
                                assert_eq!(value.as_bstr(), scalar.as_bstr())
                            }
                            oxvba_runtime::VarType::Error => {
                                assert_eq!(value.as_error_code(), scalar.as_error_code())
                            }
                            oxvba_runtime::VarType::Empty => {}
                            _ => panic!("unexpected scalar"),
                        }
                    }
                }
            }
            host.com()
                .release_object_variant(range.as_object_ref().unwrap())
                .unwrap();
        }
        for (setup_name, input, expected) in [
            ("DateCell", 45000.0, oxvba_runtime::VarType::Date),
            ("CurrencyCell", 1.234567, oxvba_runtime::VarType::Currency),
        ] {
            let bound = setup
                .bind("Setup", setup_name, ProcedureKind::Function)
                .unwrap();
            let range = setup
                .call_bound(&bound, std::slice::from_ref(&app), 1000)
                .unwrap();
            // SAFETY: unmodified compiler Range plans and the owned VBA-prepared
            // cell, in this same apartment. Formatting is oracle setup, not a
            // claim about the separate dynamic NumberFormat adapter.
            let range_object = unsafe {
                host.com().prepare_typed_object_with_interfaces(
                    range.as_object_ref().unwrap(),
                    &range_members,
                    &interfaces,
                )
            }
            .unwrap();
            let arg = api::ValueSlot::Object(range_object.argument());
            let number = api::Variant::from_f64(input);
            call_values_sliced(
                &session,
                "WriteValue2",
                &[arg, api::ValueSlot::Variant(&number)],
                api::ValueType::Void,
            );
            for read in ["ReadRange", "ReadDefault"] {
                let api::Value::Variant(value) =
                    call_values_sliced(&session, read, &[arg], api::ValueType::Variant)
                else {
                    panic!("formatted Variant expected")
                };
                println!("{backend:?} {setup_name}/{read}: {value:?}");
                assert_eq!(value.vtype(), expected);
                if expected == oxvba_runtime::VarType::Currency {
                    assert_eq!(value.as_currency_scaled_i64(), Some(12346));
                } else {
                    assert_eq!(value.as_date_f64(), Some(45000.0));
                }
            }
            let api::Value::Variant(value) =
                call_values_sliced(&session, "ReadValue2", &[arg], api::ValueType::Variant)
            else {
                panic!("Value2 Variant expected")
            };
            assert_eq!(value.as_f64(), Some(input));
            host.com()
                .release_object_variant(range.as_object_ref().unwrap())
                .unwrap();
        }
        assert!(!functions.header().iid.iter().all(|b| *b == 0));
        let counts = host.com().com_dispatch_transport_counts();
        assert!(counts.0 > baseline.0 && counts.1 > baseline.1);
        println!(
            "{backend:?}: retained WorksheetFunction/Range plans pass through one prebound Invoke per dispinterface call"
        );
    }
    host.com().release_object_variant(reference).unwrap();
    host.com()
        .release_object_variant(window.as_object_ref().unwrap())
        .unwrap();
}
struct FixtureTypes;
struct ComplexFixtureTypes;
impl oxvba_symbol::TypeLibResolver for ComplexFixtureTypes {
    fn resolve(
        &self,
        request: &oxvba_com::TypeLibResolveRequest,
    ) -> Option<oxvba_com::TypeLibMetadataBlob> {
        use oxvba_com::{
            OptionalParamDefault as D, TypeLibMemberInvokeKind as K, TypeLibParamType as P,
            TypeLibWireType as W,
        };
        let mut blob = FixtureTypes.resolve(request)?;
        let prototype = blob.members[0].clone();
        let member = |name: &str,
                      slot,
                      parameters: Vec<P>,
                      wires: Vec<W>,
                      result: Option<P>,
                      result_wire: Option<W>,
                      kind| {
            let mut member = prototype.clone();
            member.name = name.into();
            member.token = 7000 + i32::from(slot);
            member.vtable_slot = Some(slot);
            member.vtable_slot_bound = Some(42);
            member.invoke_kind = kind;
            member.parameter_names = (0..parameters.len()).map(|i| format!("arg{i}")).collect();
            member.parameter_optional = vec![false; parameters.len()];
            member.parameter_optional_defaults = vec![D::Required; parameters.len()];
            member.parameter_iids = vec![None; parameters.len()];
            member.requires_argument = !parameters.is_empty();
            member.parameter_types = parameters;
            member.parameter_wire_types = wires;
            member.return_type = result;
            member.return_wire_type = result_wire;
            member
        };
        blob.members.extend([
            {
                let mut optional = member(
                    "OptionalExists",
                    8,
                    vec![P::Long],
                    vec![W::Automation(P::Long)],
                    Some(P::Boolean),
                    Some(W::Automation(P::Boolean)),
                    K::Method,
                );
                optional.parameter_optional[0] = true;
                optional.parameter_optional_defaults[0] =
                    D::HasDefault(oxvba_com::ComDefaultValue::Int(42));
                optional.requires_argument = false;
                optional
            },
            member(
                "Payload",
                22,
                vec![],
                vec![],
                Some(P::Variant),
                Some(W::Automation(P::Variant)),
                K::PropertyGet,
            ),
            member(
                "Values",
                25,
                vec![],
                vec![],
                Some(P::Variant),
                Some(W::SafeArrayVariant),
                K::PropertyGet,
            ),
            member(
                "CheckValues",
                24,
                vec![P::Variant],
                vec![W::SafeArrayVariant],
                Some(P::Boolean),
                Some(W::Automation(P::Boolean)),
                K::Method,
            ),
            member(
                "Aliases",
                39,
                vec![P::ByRefLong; 2],
                vec![W::Automation(P::ByRefLong); 2],
                Some(P::Long),
                Some(W::Automation(P::Long)),
                K::Method,
            ),
            member(
                "FailAliases",
                40,
                vec![P::ByRefVariant; 2],
                vec![W::Automation(P::ByRefVariant); 2],
                Some(P::Long),
                Some(W::Automation(P::Long)),
                K::Method,
            ),
            member(
                "SelfValue",
                41,
                vec![],
                vec![],
                Some(P::Object),
                Some(W::InterfacePointer {
                    name: "NativeFixture.Calculator".into(),
                }),
                K::PropertyGet,
            ),
            member(
                "Integers",
                35,
                vec![],
                vec![],
                Some(P::Variant),
                Some(W::SafeArray {
                    element_vt: 3,
                    record_info: None,
                }),
                K::PropertyGet,
            ),
            member(
                "ChangeIntegers",
                37,
                vec![P::Variant],
                vec![W::ByRefSafeArray {
                    element_vt: 3,
                    record_info: None,
                }],
                None,
                None,
                K::Method,
            ),
        ]);
        blob.member_name_to_token = blob
            .members
            .iter()
            .map(|m| (m.name.clone(), m.token))
            .collect();
        Some(blob)
    }
}
#[test]
fn planned_native_outputs_preserve_aliases_errors_interfaces_arrays_and_cleanup() {
    use std::sync::atomic::Ordering;
    for backend in backends() {
        let (pointer, observed) = oxvba_com::create_oxvba_dual_vtable_object_observed();
        {
            let host = host();
            // SAFETY: fresh fixture transfers one same-apartment IDispatch ref.
            let raw = unsafe {
                host.com()
                    .bind_native_dispatch_object_variant("Owned.ComplexFixture", pointer)
            }
            .unwrap();
            let mut project = api::Project::module(
                "Api",
                r#"
Dim saved As Variant
Public Function OptionalArities(ByVal root As NativeFixture.Calculator) As String
    OptionalArities = root.OptionalExists() & ":" & root.OptionalExists(arg0:=17)
End Function
Public Function Payload(ByVal root As NativeFixture.Calculator) As Variant
    Payload = root.Payload
End Function
Public Function Aliases(ByVal root As NativeFixture.Calculator) As Long
    Dim value As Long
    Aliases = root.Aliases(value, value)
End Function
Public Function Distinct(ByVal root As NativeFixture.Calculator) As Long
    Dim first As Long, second As Long
    first = 1
    second = 1
    Distinct = root.Aliases(first, second)
End Function
Public Function FaultWrite(ByVal root As NativeFixture.Calculator) As Variant
    On Error GoTo Failed
    Dim value As Variant
    value = "owned"
    root.FailAliases value, value
Failed:
    FaultWrite = value
End Function
Public Function GlobalFault(ByVal root As NativeFixture.Calculator) As Variant
    On Error GoTo Failed
    saved = "owned"
    root.FailAliases saved, saved
Failed:
    GlobalFault = saved
End Function
Public Function PrivateFault(ByVal root As NativeFixture.Calculator) As Variant
    On Error GoTo Failed
    Dim value As Variant
    value = "owned"
    root.FailAliases (value), (value)
Failed:
    PrivateFault = value
End Function
Public Function FaultDescription(ByVal root As NativeFixture.Calculator) As String
    On Error GoTo Failed
    Dim first As Variant, second As Variant
    root.FailAliases first, second
Failed:
    FaultDescription = Err.Source & "|" & Err.Description
End Function
Public Function Child(ByVal root As NativeFixture.Calculator) As Long
    Dim value As NativeFixture.Calculator
    Set value = root.SelfValue
    Child = value.Count
End Function
Public Function ReturnChild(ByVal root As NativeFixture.Calculator) As NativeFixture.Calculator
    Set ReturnChild = root.SelfValue
End Function
Public Function Same(ByVal left As Object, ByVal right As Object) As Boolean
    Same = left Is right
End Function
Public Function SameResult(ByVal root As NativeFixture.Calculator) As Boolean
    SameResult = root Is root.SelfValue
End Function
Public Function ArrayValues(ByVal root As NativeFixture.Calculator) As Variant
    ArrayValues = root.Values
End Function
Public Function CheckArray(ByVal root As NativeFixture.Calculator, ByVal value As Variant) As Boolean
    CheckArray = root.CheckValues(value)
End Function
Public Function ChangeArray(ByVal root As NativeFixture.Calculator) As Variant
    Dim value As Variant
    value = root.Integers
    root.ChangeIntegers value
    ChangeArray = value
End Function
"#,
            );
            project
                .references
                .push(api::TypeLibraryReference::named("NativeFixture"));
            let module = Provider::with_host(host.clone())
                .with_type_libraries(Rc::new(ComplexFixtureTypes))
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
            let members: Vec<_> = interfaces
                .iter()
                .flat_map(|i| &i.members)
                .cloned()
                .collect();
            // SAFETY: all descriptors describe the exact fixture vtable above.
            let object = unsafe {
                host.com().prepare_typed_object_with_interfaces(
                    raw.as_object_ref().unwrap(),
                    &members,
                    &interfaces,
                )
            }
            .unwrap();
            host.com()
                .release_object_variant(raw.as_object_ref().unwrap())
                .unwrap();
            let session = module.into_session(backend).unwrap();
            let args = [api::ValueSlot::Object(object.argument())];
            assert!(
                matches!(call_values_sliced(&session,"OptionalArities",&args,api::ValueType::String), api::Value::String(value) if value.as_str()=="True:False")
            );
            assert!(
                matches!(call_values_sliced(&session,"Payload",&args,api::ValueType::Variant),api::Value::Variant(v) if v.as_i32()==Some(4242))
            );
            assert!(matches!(
                call_values_sliced(&session, "Aliases", &args, api::ValueType::Long),
                api::Value::Long(4242)
            ));
            assert!(matches!(
                call_values_sliced(&session, "Distinct", &args, api::ValueType::Long),
                api::Value::Long(4102)
            ));
            for name in ["FaultWrite", "GlobalFault"] {
                assert!(
                    matches!(call_values_sliced(&session,name,&args,api::ValueType::Variant),api::Value::Variant(v) if v.as_i32()==Some(42))
                );
            }
            assert!(
                matches!(call_values_sliced(&session,"PrivateFault",&args,api::ValueType::Variant), api::Value::Variant(v) if v.as_bstr().is_some_and(|v| v.as_str()=="owned"))
            );
            assert!(
                matches!(call_values_sliced(&session,"FaultDescription",&args,api::ValueType::String),api::Value::String(v) if v.as_str()==format!("{}|{}",oxvba_com::windows_test_dispatch::DUAL_RAISE_ERROR_SOURCE,oxvba_com::windows_test_dispatch::DUAL_RAISE_ERROR_DESCRIPTION))
            );
            assert_eq!(observed.error_info_references.load(Ordering::SeqCst), 0);
            assert!(matches!(
                call_values_sliced(&session, "Child", &args, api::ValueType::Long),
                api::Value::Long(7)
            ));
            let child =
                call_values_sliced(&session, "ReturnChild", &args, api::ValueType::HostObject);
            assert!(
                matches!(&child,api::Value::Object(Some(value)) if value.header().iid==object.header().iid)
            );
            assert_eq!(observed.references.load(Ordering::SeqCst), 4);
            let api::Value::Object(Some(child_object)) = &child else {
                panic!("object expected")
            };
            let before = host.com().com_dispatch_transport_counts();
            assert!(matches!(
                call_values_sliced(
                    &session,
                    "Same",
                    &[args[0], api::ValueSlot::Object(child_object.argument())],
                    api::ValueType::Boolean
                ),
                api::Value::Boolean(true)
            ));
            assert_eq!(host.com().com_dispatch_transport_counts(), before);
            assert!(matches!(
                call_values_sliced(&session, "SameResult", &args, api::ValueType::Boolean),
                api::Value::Boolean(true)
            ));
            drop(child);
            let api::Value::Variant(array) =
                call_values_sliced(&session, "ArrayValues", &args, api::ValueType::Variant)
            else {
                panic!("array expected")
            };
            assert_eq!(
                array,
                api::Variant::from_safearray(oxvba_runtime::safe_array::SafeArray::from_variants(
                    vec![
                        api::Variant::from_i32(13),
                        api::Variant::from_i32(21),
                        api::Variant::from_i32(34)
                    ]
                ))
            );
            let input = api::Variant::from_safearray(
                oxvba_runtime::safe_array::SafeArray::from_variants(vec![
                    api::Variant::from_i32(3),
                    api::Variant::from_i32(5),
                    api::Variant::from_i32(8),
                ]),
            );
            assert!(matches!(
                call_values_sliced(
                    &session,
                    "CheckArray",
                    &[args[0], api::ValueSlot::Variant(&input)],
                    api::ValueType::Boolean
                ),
                api::Value::Boolean(true)
            ));
            let api::Value::Variant(changed) =
                call_values_sliced(&session, "ChangeArray", &args, api::ValueType::Variant)
            else {
                panic!("array expected")
            };
            println!("{backend:?} changed array={changed:?}");
            assert_eq!(changed.array_element_vartype(), Some(3));
            assert_eq!(
                changed.as_safearray().unwrap().variant_elements().unwrap(),
                vec![api::Variant::from_i32(144), api::Variant::from_i32(233)]
            );
            assert_eq!(observed.references.load(Ordering::SeqCst), 2);
            let bound = session
                .bind("Api", "CheckArray", ProcedureKind::Function)
                .unwrap();
            let bad = api::Variant::from_string("not an array");
            let before = host.com().com_dispatch_transport_counts();
            assert!(matches!(
                session.call_values(
                    &bound,
                    &[args[0], api::ValueSlot::Variant(&bad)],
                    &mut api::TypedContext {
                        remaining_blocks: 1000
                    }
                ),
                Err(oxvba_embed::Error::Runtime { number: 13, .. })
            ));
            assert_eq!(host.com().com_dispatch_transport_counts(), before);
        }
        assert_eq!(observed.references.load(Ordering::SeqCst), 0);
        assert_eq!(observed.error_info_references.load(Ordering::SeqCst), 0);
        assert_eq!(observed.invoke_calls.load(Ordering::SeqCst), 0);
        assert_eq!(observed.name_lookups.load(Ordering::SeqCst), 0);
        assert_eq!(observed.type_info_reads.load(Ordering::SeqCst), 0);
    }
}
impl oxvba_symbol::TypeLibResolver for FixtureTypes {
    fn resolve(
        &self,
        request: &oxvba_com::TypeLibResolveRequest,
    ) -> Option<oxvba_com::TypeLibMetadataBlob> {
        if request.reference_name != "NativeFixture" {
            return None;
        }
        use oxvba_com::{TypeLibMemberInvokeKind as K, TypeLibParamType as T};
        let members: Vec<_> = [("Count", 7, K::PropertyGet), ("Fault", 11, K::Method)]
            .into_iter()
            .map(|(name, slot, kind)| oxvba_com::TypeLibMemberMetadata {
                name: name.into(),
                token: 7000 + slot as i32,
                vtable_slot: Some(slot),
                requires_argument: false,
                invoke_kind: kind,
                parameter_names: vec![],
                parameter_optional: vec![],
                parameter_optional_defaults: vec![],
                is_default_member: false,
                parameter_types: vec![],
                parameter_wire_types: vec![],
                parameter_iids: vec![],
                return_type: Some(T::Long),
                return_wire_type: Some(oxvba_com::TypeLibWireType::Automation(T::Long)),
                callconv_is_stdcall: true,
                is_dual: true,
                interface_iid: Some(oxvba_com::DUAL_FIXTURE_INTERFACE_IID),
                source_typekind: Some(oxvba_com::SourceTypeKind::Interface),
                vtable_slot_bound: Some(40),
            })
            .collect();
        Some(oxvba_com::TypeLibMetadataBlob {
            identity: oxvba_com::TypeLibResolvedIdentity {
                reference_name: request.reference_name.clone(),
                requested_coclass: request.requested_coclass.clone(),
                importlib: "test-owned-native-fixture".into(),
                libid: None,
                major_version: 1,
                minor_version: 0,
                lcid: None,
                cache_key: "test-owned-native-fixture".into(),
            },
            activation_prog_id: None,
            member_name_to_token: members.iter().map(|m| (m.name.clone(), m.token)).collect(),
            members,
            events: vec![],
            coclass_names: vec!["Calculator".into()],
        })
    }
}
#[test]
fn native_qi_and_error_info_owners_balance_without_dispatch_or_name_reads() {
    use std::sync::atomic::Ordering;
    for backend in backends() {
        let (pointer, observed) = oxvba_com::create_oxvba_dual_vtable_object_observed();
        {
            let host = host();
            // SAFETY: the fresh fixture transfers one same-apartment IDispatch reference.
            let raw = unsafe {
                host.com()
                    .bind_native_dispatch_object_variant("Owned.TypedFixture", pointer)
            }
            .unwrap();
            let mut project = api::Project::module(
                "Api",
                r#"
Public Function Count(ByVal root As NativeFixture.Calculator) As Long
    Count = root.Count
End Function
Public Function Describe(ByVal root As NativeFixture.Calculator) As String
    Dim value As Long
    On Error GoTo CatchFault
    value = root.Fault()
    Exit Function
CatchFault:
    Describe = Err.Source & "|" & Err.Description
End Function
"#,
            );
            project
                .references
                .push(api::TypeLibraryReference::named("NativeFixture"));
            let module = Provider::with_host(host.clone())
                .with_type_libraries(Rc::new(FixtureTypes))
                .compile(&project)
                .unwrap();
            let members: Vec<_> = module
                .com_interfaces()
                .iter()
                .flat_map(|i| i.com_members().into_iter().flatten())
                .cloned()
                .collect();
            let object = prepare(host.as_ref(), raw.as_object_ref().unwrap(), &members).unwrap();
            host.com()
                .release_object_variant(raw.as_object_ref().unwrap())
                .unwrap();
            assert_eq!(observed.references.load(Ordering::SeqCst), 2);
            let session = module.into_session(backend).unwrap();
            let bound = session
                .bind("Api", "Describe", ProcedureKind::Function)
                .unwrap();
            session
                .bind_values(
                    &bound,
                    &[api::ValueType::HostObject],
                    api::ValueType::String,
                )
                .unwrap();
            let api::Value::String(text) = session
                .call_values(
                    &bound,
                    &[api::ValueSlot::Object(object.argument())],
                    &mut api::TypedContext {
                        remaining_blocks: 1000,
                    },
                )
                .unwrap()
            else {
                panic!("String expected")
            };
            assert_eq!(
                text.as_str(),
                format!(
                    "{}|{}",
                    oxvba_com::windows_test_dispatch::DUAL_RAISE_ERROR_SOURCE,
                    oxvba_com::windows_test_dispatch::DUAL_RAISE_ERROR_DESCRIPTION
                )
            );
            assert_eq!(observed.error_info_references.load(Ordering::SeqCst), 0);
            let bound = session
                .bind("Api", "Count", ProcedureKind::Function)
                .unwrap();
            session
                .bind_values(&bound, &[api::ValueType::HostObject], api::ValueType::Long)
                .unwrap();
            let mut cancelled = session
                .begin_values(
                    &bound,
                    &[api::ValueSlot::Object(object.argument())],
                    api::InvocationFrame::new(2, None, Rc::new(())),
                    1000,
                )
                .unwrap();
            assert!(matches!(
                cancelled.resume(0).unwrap(),
                api::ValueProgress::Suspended
            ));
            drop(cancelled);
            assert_eq!(observed.references.load(Ordering::SeqCst), 2);
            let mut call = session
                .begin_values(
                    &bound,
                    &[api::ValueSlot::Object(object.argument())],
                    api::InvocationFrame::new(1, None, Rc::new(())),
                    1000,
                )
                .unwrap();
            assert!(matches!(
                call.resume(0).unwrap(),
                api::ValueProgress::Suspended
            ));
            drop(object);
            drop(session);
            loop {
                match call.resume(1).unwrap() {
                    api::ValueProgress::Suspended => {}
                    api::ValueProgress::Returned(api::Value::Long(value)) => {
                        assert_eq!(value, 7);
                        break;
                    }
                    _ => panic!("Long expected"),
                }
            }
            drop(call);
            assert_eq!(host.com().com_dispatch_transport_counts(), (1, 0));
        }
        assert_eq!(observed.references.load(Ordering::SeqCst), 0);
        assert_eq!(observed.error_info_references.load(Ordering::SeqCst), 0);
        assert_eq!(observed.invoke_calls.load(Ordering::SeqCst), 0);
        assert_eq!(observed.name_lookups.load(Ordering::SeqCst), 0);
        assert_eq!(observed.type_info_reads.load(Ordering::SeqCst), 0);
    }
}
#[test]
fn real_text_stream_long_to_bstr_calls_preserve_position_across_slices() {
    use std::io::Write;
    struct OwnedFile(std::path::PathBuf);
    impl Drop for OwnedFile {
        fn drop(&mut self) {
            std::fs::remove_file(&self.0).expect("remove owned text-stream input");
        }
    }
    for backend in backends() {
        let stamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let input = OwnedFile(std::env::temp_dir().join(format!(
            "calcvba-r9-stream-{}-{stamp}.txt",
            std::process::id()
        )));
        let mut file = std::fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&input.0)
            .unwrap();
        file.write_all(b"abcdef").unwrap();
        drop(file);
        let mut policy = HostPolicy::interactive_dev();
        policy.com_invocation_strategy = oxvba_hal::model::ComInvocationStrategy::DispatchOnly;
        let host = HostBuilder::new()
            .profile(HalProfileId::Windows)
            .policy(policy)
            .build();
        let setup = oxvba_embed::Module::load(
            "Setup",
            r#"
Public Function OpenStream(ByVal path As String) As Object
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set OpenStream = fs.OpenTextFile(path, 1)
End Function
"#,
        )
        .unwrap()
        .with_host(host.clone())
        .into_session(Backend::Interpreter)
        .unwrap();
        let bound = setup
            .bind("Setup", "OpenStream", ProcedureKind::Function)
            .unwrap();
        let raw = setup
            .call_bound(
                &bound,
                &[api::Variant::from_string(input.0.to_str().unwrap())],
                1000,
            )
            .unwrap();
        let project = project(
            r#"
Public Function ReadPart(ByVal stream As Scripting.TextStream, ByVal count As Long) As String
    ReadPart = stream.Read(count)
End Function
Public Sub CloseStream(ByVal stream As Scripting.TextStream)
    stream.Close
End Sub
"#,
        );
        let module = Provider::with_host(host.clone()).compile(&project).unwrap();
        let members: Vec<_> = module
            .com_interfaces()
            .iter()
            .flat_map(|i| i.com_members().into_iter().flatten())
            .filter(|m| {
                ["Read", "Close"]
                    .iter()
                    .any(|n| m.name.eq_ignore_ascii_case(n))
            })
            .cloned()
            .collect();
        assert_eq!(members.len(), 2);
        for member in &members {
            assert!(oxvba_com::typed_slots::signature(member).is_some());
        }
        let object = prepare(host.as_ref(), raw.as_object_ref().unwrap(), &members).unwrap();
        host.com()
            .release_object_variant(raw.as_object_ref().unwrap())
            .unwrap();
        let before = host.com().com_dispatch_transport_counts();
        let session = module.into_session(backend).unwrap();
        let bound = session
            .bind("Api", "ReadPart", ProcedureKind::Function)
            .unwrap();
        session
            .bind_values(
                &bound,
                &[api::ValueType::HostObject, api::ValueType::Long],
                api::ValueType::String,
            )
            .unwrap();
        for expected in ["abc", "def"] {
            let mut call = session
                .begin_values(
                    &bound,
                    &[
                        api::ValueSlot::Object(object.argument()),
                        api::ValueSlot::Long(3),
                    ],
                    api::InvocationFrame::new(1, None, Rc::new(())),
                    1000,
                )
                .unwrap();
            loop {
                match call.resume(1).unwrap() {
                    api::ValueProgress::Suspended => {}
                    api::ValueProgress::Returned(api::Value::String(text)) => {
                        assert_eq!(text.as_str(), expected);
                        break;
                    }
                    _ => panic!("BSTR expected"),
                }
            }
        }
        let bound = session
            .bind("Api", "CloseStream", ProcedureKind::Sub)
            .unwrap();
        session
            .bind_values(&bound, &[api::ValueType::HostObject], api::ValueType::Void)
            .unwrap();
        assert!(matches!(
            session
                .call_values(
                    &bound,
                    &[api::ValueSlot::Object(object.argument())],
                    &mut api::TypedContext {
                        remaining_blocks: 1000
                    }
                )
                .unwrap(),
            api::Value::Empty
        ));
        assert_eq!(
            host.com().com_dispatch_transport_counts(),
            (before.0 + 3, before.1)
        );
        drop(object);
        drop(session);
        drop(setup);
        drop(host);
        drop(input);
    }
}

#[test]
fn native_callback_into_typed_vba_shares_context_globals_view_and_error_routing() {
    use api::{RuntimeProvider, SessionExt};
    use std::{cell::Cell, ffi::c_void, rc::Weak};
    struct CallbackTypes(bool);
    impl oxvba_symbol::TypeLibResolver for CallbackTypes {
        fn resolve(
            &self,
            request: &oxvba_com::TypeLibResolveRequest,
        ) -> Option<oxvba_com::TypeLibMetadataBlob> {
            let mut blob = FixtureTypes.resolve(request)?;
            blob.members.truncate(1);
            blob.members[0].vtable_slot_bound = Some(9);
            let mut member = blob.members[0].clone();
            member.name = "CountByRef".into();
            member.token = 7008;
            member.vtable_slot = Some(8);
            member.invoke_kind = oxvba_com::TypeLibMemberInvokeKind::Method;
            member.requires_argument = true;
            member.parameter_names = vec!["amount".into()];
            member.parameter_optional = vec![false];
            member.parameter_optional_defaults = vec![oxvba_com::OptionalParamDefault::Required];
            member.parameter_iids = vec![None];
            member.parameter_types = vec![oxvba_com::TypeLibParamType::ByRefLong];
            member.parameter_wire_types = vec![oxvba_com::TypeLibWireType::Automation(
                oxvba_com::TypeLibParamType::ByRefLong,
            )];
            blob.members.push(member);
            if self.0 {
                for member in &mut blob.members {
                    member.vtable_slot = None;
                    member.vtable_slot_bound = Some(7);
                    member.is_dual = false;
                    member.source_typekind = Some(oxvba_com::SourceTypeKind::Dispatch);
                }
            }
            blob.member_name_to_token = blob
                .members
                .iter()
                .map(|m| (m.name.clone(), m.token))
                .collect();
            Some(blob)
        }
    }
    struct CallbackState {
        session: Weak<Box<dyn api::Session>>,
        tick: api::TypedProcedure<(i32,), i32>,
        fail: api::TypedProcedure<(), i32>,
        fail_next: Cell<bool>,
        callbacks: Cell<u32>,
    }
    // SAFETY CONTRACT: the registration below pins this exact state and output
    // for the synchronous native call, and disarms before either owner drops.
    unsafe extern "system" fn callback(context: *mut c_void, out: *mut i32) -> i32 {
        if context.is_null() || out.is_null() {
            return 0x80004003u32 as i32;
        }
        // SAFETY: the live registration retains this stable CallbackState.
        let state = unsafe { &*context.cast::<CallbackState>() };
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            api::with_native_callback_context(|context| {
                let session = state.session.upgrade().ok_or(api::Error::StaleHandle)?;
                let view = api::active_invocation().unwrap();
                assert_eq!(view.id, 71);
                assert_eq!(view.read_view.downcast_ref::<i32>(), Some(&83));
                assert!(matches!(session.reset(), Err(api::Error::Reentry)));
                assert!(matches!(
                    session.call_typed(&state.tick, (1,), 1000),
                    Err(api::Error::Reentry)
                ));
                state.callbacks.set(state.callbacks.get() + 1);
                let before = context.remaining_blocks;
                let result = if state.fail_next.replace(false) {
                    session.call_typed_in_context(&state.fail, (), context)
                } else {
                    session.call_typed_in_context(&state.tick, (1,), context)
                };
                assert!(
                    context.remaining_blocks < before,
                    "callback must consume the outer budget"
                );
                result
            })
            .and_then(|result| result)
        }));
        match result {
            Ok(Ok(value)) => {
                // SAFETY: the native Count caller supplies one writable LONG.
                unsafe {
                    out.write(value);
                }
                0
            }
            Ok(Err(api::Error::Runtime { number: 13, .. })) => 0x80020005u32 as i32,
            _ => 0x80004005u32 as i32,
        }
    }
    struct Registration(Rc<native_identity::Observation>);
    impl Drop for Registration {
        fn drop(&mut self) {
            self.0.callback.set(None);
            self.0.callback_context.set(std::ptr::null_mut());
        }
    }
    let source = r#"
Public total As Long
Private nesting As Boolean
Private nested As NativeFixture.Calculator
Public Function OnTick(ByVal amount As Long) As Long
    total = total + amount
    If nesting Then
        nesting = False
        OnTick = nested.CountByRef(total)
    Else
        OnTick = total
    End If
End Function
Public Function RaiseTick() As Long
    total = total + 1
    Err.Raise 13
End Function
Public Function Trigger(ByVal value As NativeFixture.Calculator) As Long
    total = 40
    Trigger = value.Count + total
End Function
Public Function Handled(ByVal value As NativeFixture.Calculator) As Long
    On Error Resume Next
    total = 40
    Handled = value.Count
    Handled = Err.Number
End Function
Public Function TriggerByRef(ByVal value As NativeFixture.Calculator) As Long
    Dim amount As Long
    total = 40
    amount = 2
    TriggerByRef = value.CountByRef(amount)
    TriggerByRef = TriggerByRef + amount + total
End Function
Public Function HandledByRef(ByVal value As NativeFixture.Calculator) As Long
    Dim amount As Long
    On Error Resume Next
    total = 40
    amount = 2
    HandledByRef = value.CountByRef(amount)
    HandledByRef = Err.Number + amount + total
End Function
Public Function TriggerGlobalByRef(ByVal value As NativeFixture.Calculator) As Long
    total = 2
    TriggerGlobalByRef = value.CountByRef(total)
    TriggerGlobalByRef = TriggerGlobalByRef * 100 + total
End Function
Public Function HandledGlobalByRef(ByVal value As NativeFixture.Calculator) As Long
    On Error Resume Next
    total = 2
    HandledGlobalByRef = value.CountByRef(total)
    HandledGlobalByRef = Err.Number * 100 + total
End Function
Public Function TriggerPrivateGlobalByRef(ByVal value As NativeFixture.Calculator) As Long
    total = 2
    TriggerPrivateGlobalByRef = value.CountByRef((total))
    TriggerPrivateGlobalByRef = TriggerPrivateGlobalByRef * 100 + total
End Function
Public Function TriggerNestedByRef(ByVal value As NativeFixture.Calculator) As Long
    Set nested = value
    nesting = True
    total = 2
    TriggerNestedByRef = value.CountByRef(total)
    TriggerNestedByRef = TriggerNestedByRef * 100 + total
    Set nested = Nothing
End Function
Public Function ReadTotal() As Long
    ReadTotal = total
End Function
"#;
    for (backend, dispatch) in backends().into_iter().flat_map(|b| [(b, false), (b, true)]) {
        let host = host();
        let (pointer, observed) = native_identity::create();
        observed.prebound.set(dispatch);
        {
            // SAFETY: create transfers the owned same-thread fixture dispatch reference.
            let raw = unsafe {
                host.com()
                    .bind_native_dispatch_object_variant("Owned.Callback", pointer.cast())
            }
            .unwrap();
            let mut project = api::Project::module("Api", source);
            project
                .references
                .push(api::TypeLibraryReference::named("NativeFixture"));
            let provider = Provider::with_host(host.clone())
                .with_type_libraries(Rc::new(CallbackTypes(dispatch)));
            let module = provider.compile(&project).unwrap();
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
                .find(|i| !i.members.is_empty())
                .unwrap()
                .members;
            // SAFETY: the fixture ABI matches these unchanged compiler plans;
            // dispatch additionally retains the complete compiler table ordering.
            let object = unsafe {
                host.com().prepare_typed_object_with_interfaces(
                    raw.as_object_ref().unwrap(),
                    members,
                    &interfaces,
                )
            }
            .unwrap();
            let session: Rc<Box<dyn api::Session>> = Rc::new(
                provider
                    .load(
                        &project,
                        match backend {
                            Backend::Interpreter => api::Backend::Interpreter,
                            Backend::Jit => api::Backend::Jit,
                        },
                        api::HOST_API_ABI,
                    )
                    .unwrap(),
            );
            let state = Box::new(CallbackState {
                session: Rc::downgrade(&session),
                tick: session.bind_typed("Api", "OnTick").unwrap(),
                fail: session.bind_typed("Api", "RaiseTick").unwrap(),
                fail_next: Cell::new(false),
                callbacks: Cell::new(0),
            });
            observed
                .callback_context
                .set((&*state as *const CallbackState).cast_mut().cast());
            observed.callback.set(Some(callback));
            let registration = Registration(observed.clone());
            let queries = observed.queries.get();
            for (name, expected, fail, expected_global) in [
                ("Trigger", 82, false, 41),
                ("Handled", 13, true, 41),
                ("TriggerByRef", 94, false, 41),
                ("HandledByRef", 66, true, 41),
                ("TriggerGlobalByRef", 1313, false, 13),
                ("HandledGlobalByRef", 1313, true, 13),
                ("TriggerPrivateGlobalByRef", 303, false, 3),
                ("TriggerNestedByRef", 2424, false, 24),
                ("Trigger", 82, false, 41),
            ] {
                state.fail_next.set(fail);
                let bound = session
                    .bind_values(
                        "Api",
                        name,
                        &[api::ValueType::HostObject],
                        api::ValueType::Long,
                    )
                    .unwrap();
                let mut call = session
                    .begin_values(
                        &bound,
                        &[api::ValueSlot::Object(object.argument())],
                        api::InvocationFrame::new(71, None, Rc::new(83)),
                        1000,
                    )
                    .unwrap();
                loop {
                    match call.resume(1).unwrap() {
                        api::ValueProgress::Suspended => {}
                        api::ValueProgress::Returned(value) => {
                            assert!(
                                matches!(value,api::Value::Long(v) if v==expected),
                                "{backend:?} {name}: {value:?}, expected {expected}"
                            );
                            break;
                        }
                    }
                }
                assert_eq!(
                    session
                        .call_typed::<(), i32>(
                            &session.bind_typed("Api", "ReadTotal").unwrap(),
                            (),
                            1000
                        )
                        .unwrap(),
                    expected_global
                );
                assert!(matches!(
                    api::with_native_callback_context(|_| ()),
                    Err(api::Error::Reentry)
                ));
            }
            assert_eq!(state.callbacks.get(), 10);
            assert_eq!(observed.calls.get(), 10);
            assert_eq!(observed.invokes.get(), if dispatch { 10 } else { 0 });
            assert_eq!(observed.queries.get(), queries);
            assert_eq!(observed.dynamic.get(), 0);
            let mut out = 0;
            // SAFETY: exercise the still-pinned registration outside a native
            // scope: it must refuse without entering or changing VBA state.
            let outside = unsafe {
                callback(
                    (&*state as *const CallbackState).cast_mut().cast(),
                    &mut out,
                )
            };
            assert!(outside < 0);
            assert_eq!(state.callbacks.get(), 10);
            drop(registration);
            drop(state);
            drop(object);
            host.com()
                .release_object_variant(raw.as_object_ref().unwrap())
                .unwrap();
        }
        assert_eq!(observed.references.get(), 0);
        assert_eq!(observed.dynamic.get(), 0);
    }
}

#[test]
fn connection_point_dispatch_sink_runs_typed_handlers_and_rejects_stale_targets() {
    use api::{RuntimeProvider, SessionExt};
    use oxvba_com::{ComValue, DispatchEventSinkConfig, windows_connection_point::WindowsEventArg};
    use std::cell::{Cell, RefCell};
    struct Target {
        session: std::rc::Weak<RefCell<Box<dyn api::Session>>>,
        handler: api::TypedProcedure<(i32,), i32>,
        deliveries: Cell<u32>,
        errors: Cell<u32>,
        last: Cell<i32>,
    }
    thread_local! {static TARGET:RefCell<Option<Rc<Target>>>=const {RefCell::new(None)};}
    struct TargetScope;
    impl Drop for TargetScope {
        fn drop(&mut self) {
            TARGET.with(|slot| {
                slot.borrow_mut().take();
            });
        }
    }
    struct Source(*mut oxvba_com::RawIDispatch, Rc<Cell<u32>>);
    impl Source {
        fn as_ptr(&self) -> *mut oxvba_com::RawIDispatch {
            self.0
        }
    }
    impl Drop for Source {
        fn drop(&mut self) {
            // SAFETY: the fixture transferred this one native owner on this thread.
            unsafe {
                let remaining = ((*(*self.0).vtbl).unknown.release)(self.0.cast());
                self.1.set(remaining);
            }
        }
    }
    struct Subscription(Option<oxvba_com::WindowsConnectionPointTransport>);
    impl Drop for Subscription {
        fn drop(&mut self) {
            if let Some(transport) = self.0.take() {
                // SAFETY: this guard owns the one advise transport until final release.
                let _ = unsafe { oxvba_com::unadvise_connection_point(transport) };
            }
        }
    }
    // SAFETY CONTRACT: source is the retained fixture IDispatch. The helper is
    // test-driver control: it asks the fixture to emit its declared native event.
    unsafe fn fire(source: *mut oxvba_com::RawIDispatch, value: i32) {
        use windows_sys::Win32::System::{
            Com::{DISPATCH_METHOD, DISPPARAMS},
            Variant::{VARIANT, VT_I4},
        };
        // SAFETY: initialized scalar VARIANT and exact live DISPPARAMS extents;
        // the owned fixture supports this declared DISPID and consumes no owner.
        unsafe {
            let mut argument: VARIANT = std::mem::zeroed();
            argument.Anonymous.Anonymous.vt = VT_I4;
            argument.Anonymous.Anonymous.Anonymous.lVal = value;
            let mut args = DISPPARAMS {
                rgvarg: &mut argument,
                rgdispidNamedArgs: std::ptr::null_mut(),
                cArgs: 1,
                cNamedArgs: 0,
            };
            let status = ((*(*source).vtbl).invoke)(
                source.cast(),
                oxvba_com::TEST_DISPID_FIRE_CHANGED,
                &oxvba_com::IID_NULL,
                0,
                DISPATCH_METHOD,
                &mut args,
                std::ptr::null_mut(),
                std::ptr::null_mut(),
                std::ptr::null_mut(),
            );
            assert!(status >= 0, "fixture event driver: {status:08x}");
        }
    }
    for backend in backends() {
        let project = api::Project::module(
            "Api",
            "Public total As Long\nPublic Function Receive(ByVal amount As Long) As Long\ntotal=total+amount\nReceive=total\nEnd Function\nPublic Function ReadTotal() As Long\nReadTotal=total\nEnd Function\n",
        );
        let session = Rc::new(RefCell::new(
            Provider::default()
                .load(
                    &project,
                    match backend {
                        Backend::Interpreter => api::Backend::Interpreter,
                        Backend::Jit => api::Backend::Jit,
                    },
                    api::HOST_API_ABI,
                )
                .unwrap(),
        ));
        let target = Rc::new(Target {
            session: Rc::downgrade(&session),
            handler: session.borrow().bind_typed("Api", "Receive").unwrap(),
            deliveries: Cell::new(0),
            errors: Cell::new(0),
            last: Cell::new(-1),
        });
        TARGET.with(|slot| {
            slot.replace(Some(target.clone()));
        });
        let scope = TargetScope;
        let released = Rc::new(Cell::new(u32::MAX));
        let source = Source(oxvba_com::create_oxvba_test_dispatch(), released.clone());
        let config = DispatchEventSinkConfig {
            event_dispatch_member: oxvba_com::TEST_EVENT_CHANGED,
            expected_arity: 1,
            connection_point_iid: None,
            on_event: Arc::new(|args| {
                let Some(target) = TARGET.with(|slot| slot.borrow().clone()) else {
                    return false;
                };
                let [WindowsEventArg::Value(ComValue::I32(value))] = args else {
                    return false;
                };
                target.deliveries.set(target.deliveries.get() + 1);
                // The Send+Sync sink callback carries no non-Send session. The
                // owning thread snapshots its host target before running VBA.
                let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                    api::with_native_callback_context(|context| {
                        let session = target.session.upgrade().ok_or(api::Error::StaleHandle)?;
                        session
                            .borrow()
                            .call_typed_in_context(&target.handler, (*value,), context)
                    })
                    .and_then(|result| result)
                }));
                match result {
                    Ok(Ok(value)) => {
                        target.last.set(value);
                        true
                    }
                    _ => {
                        target.errors.set(target.errors.get() + 1);
                        false
                    }
                }
            }),
        };
        // SAFETY: exact live owned source and its fixture-declared connection point IID.
        let transport = unsafe {
            oxvba_com::try_advise_dispatch_event_sink(
                source.as_ptr().cast(),
                oxvba_com::IID_OXVBA_TEST_DISPATCH_EVENTS_STR,
                config,
            )
        }
        .unwrap()
        .unwrap();
        let mut subscription = Subscription(Some(transport));
        // SAFETY: the source remains owned; an unavailable context must refuse the handler.
        unsafe {
            fire(source.as_ptr().cast(), 3);
        }
        assert_eq!(target.errors.get(), 1);
        let mut context = api::TypedContext {
            remaining_blocks: 200,
        };
        for value in [7, 9] {
            api::with_native_call(&mut context, || {
                // SAFETY: source remains owned across this synchronous event emission.
                unsafe {
                    fire(source.as_ptr().cast(), value);
                }
            });
        }
        assert_eq!(target.last.get(), 16);
        assert_eq!(target.deliveries.get(), 3);
        assert!(context.remaining_blocks < 200);
        session.borrow_mut().reload(&project).unwrap();
        let before = context.remaining_blocks;
        api::with_native_call(&mut context, || {
            // SAFETY: source remains owned; stale generation must reject before VBA entry.
            unsafe {
                fire(source.as_ptr().cast(), 3);
            }
        });
        assert_eq!(context.remaining_blocks, before);
        assert_eq!(target.errors.get(), 2);
        assert_eq!(target.last.get(), 16);
        let fresh = Rc::new(Target {
            session: Rc::downgrade(&session),
            handler: session.borrow().bind_typed("Api", "Receive").unwrap(),
            deliveries: Cell::new(0),
            errors: Cell::new(0),
            last: Cell::new(-1),
        });
        TARGET.with(|slot| {
            slot.replace(Some(fresh.clone()));
        });
        api::with_native_call(&mut context, || {
            // SAFETY: source remains owned; the replacement target has the live generation.
            unsafe {
                fire(source.as_ptr().cast(), 5);
            }
        });
        assert_eq!(fresh.last.get(), 5);
        assert_eq!(fresh.errors.get(), 0);
        let transport = subscription.0.take().unwrap();
        // SAFETY: take transfers this guard's sole advise/release authority.
        unsafe { oxvba_com::unadvise_connection_point(transport) }.unwrap();
        api::with_native_call(&mut context, || {
            // SAFETY: live source, with the sink now unadvised; no handler may execute.
            unsafe {
                fire(source.as_ptr().cast(), 7);
            }
        });
        assert_eq!(fresh.deliveries.get(), 1);
        let read = session
            .borrow()
            .bind_typed::<(), i32>("Api", "ReadTotal")
            .unwrap();
        assert_eq!(session.borrow().call_typed(&read, (), 1000).unwrap(), 5);
        assert!(matches!(
            api::with_native_callback_context(|_| ()),
            Err(api::Error::Reentry)
        ));
        drop(scope);
        assert_eq!(Rc::strong_count(&fresh), 1);
        drop(subscription);
        drop(source);
        assert_eq!(
            released.get(),
            0,
            "connection point/source references balance"
        );
    }
}

#[test]
fn native_double_boolean_aliases_remain_visible_during_prebound_and_native_callbacks() {
    use api::{RuntimeProvider, SessionExt};
    use std::{cell::Cell, ffi::c_void, rc::Weak};
    struct Types(bool);
    impl oxvba_symbol::TypeLibResolver for Types {
        fn resolve(
            &self,
            request: &oxvba_com::TypeLibResolveRequest,
        ) -> Option<oxvba_com::TypeLibMetadataBlob> {
            use oxvba_com::{TypeLibParamType as P, TypeLibWireType as W};
            let mut blob = FixtureTypes.resolve(request)?;
            let prototype = blob.members[0].clone();
            blob.members = [
                ("CountDouble", 9, P::ByRefDouble, P::Double),
                ("CountBoolean", 10, P::ByRefBoolean, P::Boolean),
            ]
            .into_iter()
            .map(|(name, slot, parameter, result)| {
                let mut member = prototype.clone();
                member.name = name.into();
                member.token = 7000 + slot;
                member.vtable_slot = Some(slot as u16);
                member.vtable_slot_bound = Some(11);
                member.invoke_kind = oxvba_com::TypeLibMemberInvokeKind::Method;
                member.requires_argument = true;
                member.parameter_names = vec!["amount".into(), "sameAmount".into()];
                member.parameter_optional = vec![false; 2];
                member.parameter_optional_defaults =
                    vec![oxvba_com::OptionalParamDefault::Required; 2];
                member.parameter_iids = vec![None; 2];
                member.parameter_types = vec![parameter; 2];
                member.parameter_wire_types = vec![W::Automation(parameter); 2];
                member.return_type = Some(result);
                member.return_wire_type = Some(W::Automation(result));
                if self.0 {
                    member.vtable_slot = None;
                    member.vtable_slot_bound = Some(7);
                    member.is_dual = false;
                    member.source_typekind = Some(oxvba_com::SourceTypeKind::Dispatch);
                }
                member
            })
            .collect();
            blob.member_name_to_token = blob
                .members
                .iter()
                .map(|m| (m.name.clone(), m.token))
                .collect();
            Some(blob)
        }
    }
    struct State {
        session: Weak<Box<dyn api::Session>>,
        tick: api::TypedProcedure<(), f64>,
        boolean_tick: api::BoundProcedure,
        calls: Cell<usize>,
    }
    // SAFETY CONTRACT: a registration retains this state and the native DOUBLE
    // result throughout the same-thread synchronous call, then disarms it.
    unsafe extern "system" fn callback(data: *mut c_void, output: *mut f64) -> i32 {
        // SAFETY: the scoped registration pins the exact State until disarmed.
        let state = unsafe { &*data.cast::<State>() };
        let value = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            api::with_native_callback_context(|context| {
                let session = state.session.upgrade().ok_or(api::Error::StaleHandle)?;
                let before = context.remaining_blocks;
                let result = session.call_typed_in_context(&state.tick, (), context);
                assert!(context.remaining_blocks < before);
                state.calls.set(state.calls.get() + 1);
                result
            })
            .and_then(|r| r)
        }));
        match value {
            Ok(Ok(value)) => {
                // SAFETY: the admitted native signature supplies one writable DOUBLE.
                unsafe { output.write(value) };
                0
            }
            _ => 0x80004005u32 as i32,
        }
    }
    // SAFETY CONTRACT: the same registration pins State and one writable
    // VARIANT_BOOL result for this synchronous native callback.
    unsafe extern "system" fn boolean_callback(data: *mut c_void, output: *mut i16) -> i32 {
        // SAFETY: registration retains this exact stable State allocation.
        let state = unsafe { &*data.cast::<State>() };
        let value = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            api::with_native_callback_context(|context| {
                let session = state.session.upgrade().ok_or(api::Error::StaleHandle)?;
                let before = context.remaining_blocks;
                let result = session
                    .call_values(&state.boolean_tick, &[], context)
                    .and_then(|value| match value {
                        api::Value::Boolean(value) => Ok(value),
                        _ => Err(api::Error::ArgumentMismatch),
                    });
                assert!(context.remaining_blocks < before);
                state.calls.set(state.calls.get() + 1);
                result
            })
            .and_then(|r| r)
        }));
        match value {
            Ok(Ok(value)) => {
                // SAFETY: exact native result cell; Automation Boolean is canonical -1/0.
                unsafe { output.write(if value { -1 } else { 0 }) };
                0
            }
            _ => 0x80004005u32 as i32,
        }
    }
    struct Registration(Rc<native_identity::Observation>);
    impl Drop for Registration {
        fn drop(&mut self) {
            self.0.double_callback.set(None);
            self.0.boolean_callback.set(None);
            self.0.callback_context.set(std::ptr::null_mut());
        }
    }
    for backend in backends() {
        for dispatch in [false, true] {
            let host = host();
            let (pointer, observed) = native_identity::create();
            {
                // SAFETY: the fixture transfers exactly one owned same-apartment reference.
                let raw = unsafe {
                    host.com()
                        .bind_native_dispatch_object_variant("Owned.DoubleCallback", pointer.cast())
                }
                .unwrap();
                observed.prebound.set(dispatch);
                let mut project = api::Project::module(
                    "Api",
                    format!(
                        "{}\n{}",
                        include_str!(
                            "../../../docs/evidence/programs/ideal-2026-07/core/CORE-7/native-scalar-callback-bodies-run13.bas"
                        ),
                        r#"Public Function RunBoolean(ByVal root As NativeFixture.Calculator) As String
    Dim result As Boolean
    aliasFlag = False
    result = root.CountBoolean(aliasFlag, aliasFlag)
    RunBoolean = result & ":" & aliasFlag
End Function
Public Function Run(ByVal root As NativeFixture.Calculator) As String
    Dim result As Double
    aliasValue = 2#
    result = root.CountDouble(aliasValue, aliasValue)
    Run = result & ":" & aliasValue
End Function
"#
                    ),
                );
                project
                    .references
                    .push(api::TypeLibraryReference::named("NativeFixture"));
                let provider =
                    Provider::with_host(host.clone()).with_type_libraries(Rc::new(Types(dispatch)));
                let module = provider.compile(&project).unwrap();
                let interfaces = module
                    .com_interfaces()
                    .iter()
                    .filter_map(|i| match i {
                        oxvba_oxir::ComInterface::Com(m) => Some(m.clone()),
                        _ => None,
                    })
                    .collect::<Vec<_>>();
                let members = interfaces
                    .iter()
                    .flat_map(|i| &i.members)
                    .cloned()
                    .collect::<Vec<_>>();
                // SAFETY: the exact unchanged fixture descriptors select its native ABI.
                let object = unsafe {
                    host.com().prepare_typed_object_with_interfaces(
                        raw.as_object_ref().unwrap(),
                        &members,
                        &interfaces,
                    )
                }
                .unwrap();
                host.com()
                    .release_object_variant(raw.as_object_ref().unwrap())
                    .unwrap();
                let session: Rc<Box<dyn api::Session>> = Rc::new(
                    provider
                        .load(
                            &project,
                            match backend {
                                Backend::Interpreter => api::Backend::Interpreter,
                                Backend::Jit => api::Backend::Jit,
                            },
                            api::HOST_API_ABI,
                        )
                        .unwrap(),
                );
                let state = Box::new(State {
                    session: Rc::downgrade(&session),
                    tick: session.bind_typed("Api", "CallbackValue").unwrap(),
                    boolean_tick: session
                        .bind_values("Api", "BooleanCallbackValue", &[], api::ValueType::Boolean)
                        .unwrap(),
                    calls: Cell::new(0),
                });
                observed
                    .callback_context
                    .set((&*state as *const State).cast_mut().cast());
                observed.double_callback.set(Some(callback));
                observed.boolean_callback.set(Some(boolean_callback));
                let registration = Registration(observed.clone());
                for (name, expected) in [("Run", "12.5:12.5"), ("RunBoolean", "False:False")] {
                    let bound = session
                        .bind_values(
                            "Api",
                            name,
                            &[api::ValueType::HostObject],
                            api::ValueType::String,
                        )
                        .unwrap();
                    let before = host.com().com_dispatch_transport_counts();
                    let mut invocation = session
                        .begin_values(
                            &bound,
                            &[api::ValueSlot::Object(object.argument())],
                            api::InvocationFrame::new(13, None, Rc::new(())),
                            1000,
                        )
                        .unwrap();
                    loop {
                        if let api::ValueProgress::Returned(value) = invocation.resume(1).unwrap() {
                            let api::Value::String(value) = value else {
                                panic!("String expected")
                            };
                            println!("{backend:?} dispatch={dispatch} {name}: {}", value.as_str());
                            assert_eq!(value.as_str(), expected);
                            break;
                        }
                    }
                    assert_eq!(
                        host.com().com_dispatch_transport_counts(),
                        (
                            before.0 + u64::from(!dispatch),
                            before.1 + u64::from(dispatch)
                        )
                    );
                    drop(invocation);
                }
                let mut context = api::TypedContext {
                    remaining_blocks: 1000,
                };
                for expected in [true, false] {
                    assert!(
                        matches!(session.call_values(&state.boolean_tick,&[],&mut context).unwrap(),api::Value::Boolean(value) if value==expected)
                    );
                }
                assert_eq!(state.calls.get(), 2);
                assert_eq!(observed.dynamic.get(), 0);
                assert_eq!(observed.calls.get(), 2);
                assert_eq!(observed.invokes.get(), u32::from(dispatch) * 2);
                drop(registration);
                drop(state);
                drop(object);
                drop(session);
                drop(module);
            }
            drop(host);
            assert_eq!(observed.references.get(), 0);
        }
    }
}
