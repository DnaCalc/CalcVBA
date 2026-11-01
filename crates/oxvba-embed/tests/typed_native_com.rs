#![cfg(all(windows, target_arch = "x86_64"))]
use calcvba_host_api as api;
use oxvba_embed::{Backend, ProcedureKind, provider::Provider};
use oxvba_hal::{
    HostServices,
    adapters::builder::HostBuilder,
    model::{HalProfileId, HostPolicy},
};
use std::{rc::Rc, sync::Arc};
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
                assert!(oxvba_com::typed_slots::plan::visible_signature(member, 1).is_none());
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
        match call.resume(1).unwrap() {
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
    let setup=oxvba_embed::Module::load("Setup", "Public Function App(ByVal window As Object) As Object\nSet App=window.Application\nEnd Function\nPublic Function Block(ByVal app As Object) As Object\nSet Block=app.ActiveWorkbook.Worksheets(1).Range(\"A1:B2\")\nEnd Function\nPublic Function Cell(ByVal app As Object) As Object\nSet Cell=app.ActiveWorkbook.Worksheets(1).Range(\"A1\")\nEnd Function\n").unwrap().with_host(host.clone()).into_session(Backend::Interpreter).unwrap();
    let bound = setup.bind("Setup", "App", ProcedureKind::Function).unwrap();
    let app = setup
        .call_bound(&bound, std::slice::from_ref(&window), 1000)
        .unwrap();
    let reference = app.as_object_ref().unwrap();
    for backend in backends() {
        let mut project = api::Project::module(
            "Api",
            r#"
Public Function Version(ByVal app As Excel.Application) As String
    Version = app.Version
End Function
Public Function Build(ByVal app As Excel.Application) As Long
    Build = app.Build
End Function
Public Function Height(ByVal app As Excel.Application) As Double
    Height = app.Height
End Function
Public Function Visible(ByVal app As Excel.Application) As Boolean
    Visible = app.Visible
End Function
Public Function ReadRange(ByVal value As Excel.Range) As Variant
    ReadRange = value.Value
End Function
Public Sub WriteRange(ByVal target As Excel.Range, ByVal value As Variant)
    target.Value = value
End Sub
Public Function RuntimeDefault(ByVal value As Object) As Double
    RuntimeDefault = value + 1#
End Function
Public Function GetFunctions(ByVal app As Excel.Application) As Excel.WorksheetFunction
    Set GetFunctions = app.WorksheetFunction
End Function
Public Function SumArray(ByVal functions As Excel.WorksheetFunction, ByVal value As Variant) As Double
    SumArray = functions.Sum(value)
End Function
Public Function SumDifferentArities(ByVal functions As Excel.WorksheetFunction, ByVal x As Variant, ByVal y As Variant) As Double
    SumDifferentArities = functions.Sum(x) + functions.Sum(x, y)
End Function
Public Function SumRange(ByVal app As Excel.Application, ByVal value As Excel.Range) As Double
    SumRange = app.WorksheetFunction.Sum(value)
End Function
"#,
        );
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
                        || (m.is_default_member
                            && m.invoke_kind == oxvba_com::TypeLibMemberInvokeKind::PropertyGet))
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
        let sum = session
            .bind("Api", "SumArray", ProcedureKind::Function)
            .unwrap();
        assert!(
            session
                .bind_values(
                    &sum,
                    &[api::ValueType::HostObject, api::ValueType::Variant],
                    api::ValueType::Double
                )
                .is_err()
        );
        let read = session
            .bind("Api", "ReadRange", ProcedureKind::Function)
            .unwrap();
        assert!(
            session
                .bind_values(
                    &read,
                    &[api::ValueType::HostObject],
                    api::ValueType::Variant
                )
                .is_err()
        );
        assert!(!functions.header().iid.iter().all(|b| *b == 0));
        let counts = host.com().com_dispatch_transport_counts();
        assert_eq!(counts, (baseline.0 + 1, baseline.1));
        println!(
            "{backend:?}: native WorksheetFunction interface result passes with zero Invoke delta; Range.Value and Sum have no native slots"
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
            assert_eq!(observed.references.load(Ordering::SeqCst), 2);
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
            assert_eq!(observed.references.load(Ordering::SeqCst), 1);
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
            assert_eq!(observed.references.load(Ordering::SeqCst), 1);
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
            assert_eq!(observed.references.load(Ordering::SeqCst), 1);
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
