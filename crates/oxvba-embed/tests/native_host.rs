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

/// The existing PID-scoped supervisor holds Excel and its UIA guardian while
/// this case runs. Bind the exact EXCEL7 NativeOM root; never activate/open a
/// second Excel through a file moniker or a global running-object lookup.
#[test]
#[ignore = "requires live OXVBA_OWNED_EXCEL_PID/HWND from the owned worksheet oracle supervisor"]
fn host_injected_owned_excel_range_uses_default_without_losing_identity() {
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
    let mut backends = vec![Backend::Interpreter];
    if cfg!(feature = "jit") {
        backends.push(Backend::Jit);
    }
    for backend in backends {
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
        let hr =
            unsafe { AccessibleObjectFromWindow(search.window, 0xfffffff0, &iid, &mut pointer) };
        assert!(
            hr >= 0 && !pointer.is_null(),
            "owned NativeOM failed: {hr:08x}"
        );
        let host = oxvba_hal::adapters::builder::HostBuilder::new()
            .profile(oxvba_hal::model::HalProfileId::Windows)
            .policy(oxvba_hal::model::HostPolicy::interactive_dev())
            .build();
        // SAFETY: transfer exactly the NativeOM reference to the same-thread HAL.
        let root = unsafe {
            host.com()
                .bind_native_dispatch_object_variant("Owned.ExcelWindow", pointer)
        }
        .unwrap();
        let source = r#"
Private kept As Object
Public Function Probe(ByVal window As Object, ByVal expected As Long) As String
    Dim app As Object, sheet As Object, cell As Object
    Dim argument As Variant
    Dim plus As Double, times As Double, textPlus As Double, blankPlus As Double
    Dim beforeName As String, afterName As String, stillObject As Boolean
    Dim badError As Long, multiError As Long
    Dim stage As Long
    On Error GoTo Diagnose
    stage=1
    Set app=window.Application
    stage=2
    If app.Hwnd <> expected Then Err.Raise 5
    stage=3
    Set sheet=window.ActiveSheet
    stage=4
    Set cell=sheet.Range("D1")
    Set kept=cell
    cell.NumberFormat="General"
    cell.Value2=41
    Set argument=cell
    beforeName=TypeName(argument)
    stage=5
    plus=UntypedPlus(argument)
    stage=6
    times=UntypedTimes(argument)
    afterName=TypeName(argument)
    stillObject=IsObject(argument)
    cell.NumberFormat="@"
    cell.Value2="3"
    stage=7
    textPlus=UntypedPlus(argument)
    cell.ClearContents
    stage=8
    blankPlus=UntypedPlus(argument)
    cell.Value2="bad"
    stage=9
    badError=CapturedError(argument)
    Set argument=sheet.Range("D1:D2")
    stage=10
    multiError=CapturedError(argument)
    Probe=CStr(plus) & "|" & CStr(times) & "|" & beforeName & "|" & afterName & "|" & CStr(stillObject) & "|" & CStr(textPlus) & "|" & CStr(blankPlus) & "|" & CStr(badError) & "|" & CStr(multiError)
    Err.Clear
    Exit Function
Diagnose:
    Probe="stage=" & CStr(stage) & "|name=" & beforeName & "|error=" & CStr(Err.Number) & "|" & Err.Description
End Function
Private Function UntypedPlus(x) As Double
    Dim saved As Object
    Set saved=x
    UntypedPlus=x+1
    If Not IsObject(x) Then Err.Raise 13
    If Not (x Is saved) Then Err.Raise 13
    If TypeName(x) <> "Range" Then Err.Raise 13
End Function
Private Function UntypedTimes(x) As Double
    UntypedTimes=x*2
End Function
Private Function CapturedError(x) As Long
    Dim value As Double
    On Error Resume Next
    Err.Clear
    value=x+1
    CapturedError=Err.Number
End Function
Public Function WriteKept(ByVal value As Double) As String
    Dim current As Variant
    Set current=kept
    kept.Value2=value
    WriteKept=TypeName(current)
    Err.Clear
End Function
Public Function ReadKept() As Double
    ReadKept=kept+1
End Function
"#;
        let module = Module::load("Api", source).unwrap().with_host(host);
        let session = module.prepare(backend).unwrap();
        let handle = session
            .bind("Api", "Probe", ProcedureKind::Function)
            .unwrap();
        let result = session
            .call_bound(
                &handle,
                &[root, oxvba_runtime::Variant::from_i32(hwnd_number)],
                10000,
            )
            .unwrap();
        assert_eq!(
            oxvba_runtime::variant_to_vba_string(&result)
                .unwrap()
                .as_str(),
            "42|82|Range|Range|True|4|1|13|13",
            "{backend:?}"
        );
        let write = session
            .bind("Api", "WriteKept", ProcedureKind::Function)
            .unwrap();
        let name = session
            .call_bound(&write, &[oxvba_runtime::Variant::from_f64(43.0)], 10000)
            .unwrap();
        assert_eq!(
            oxvba_runtime::variant_to_vba_string(&name)
                .unwrap()
                .as_str(),
            "Range"
        );
        let read = session
            .bind("Api", "ReadKept", ProcedureKind::Function)
            .unwrap();
        assert_eq!(
            session.call_bound(&read, &[], 10000).unwrap().as_f64(),
            Some(44.0),
            "module-held Range must survive the preceding host call"
        );
        println!(
            "{backend:?}: real Range identity/default arithmetic passed on owned Excel PID {expected_pid}"
        );
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
            // A library client sees only the API contract. The application
            // composition supplies the real type resolver and native root.
            use calcvba_host_api::{self as api, RuntimeProvider};
            let provider = oxvba_embed::provider::Provider::with_host(host.clone())
                .with_type_libraries(std::rc::Rc::new(HostTypes));
            let mut project = api::Project::module("Api", manifest.modules[0].source.clone());
            project
                .references
                .push(api::TypeLibraryReference::named("TypedHost"));
            let api_backend = match backend {
                Backend::Interpreter => api::Backend::Interpreter,
                Backend::Jit => api::Backend::Jit,
            };
            let session = provider
                .load(&project, api_backend, api::HOST_API_ABI)
                .unwrap();
            let handle = session
                .bind("Api", "ReadCount", api::ProcedureKind::Function)
                .unwrap();
            let (pointer, retained) = oxvba_com::create_oxvba_dual_vtable_object_observed();
            // SAFETY: transfer one owned native reference to the same-thread host.
            let object = unsafe {
                host.com()
                    .bind_native_dispatch_object_variant("Neutral.SecondRoot", pointer)
            }
            .unwrap();
            let before = host.com().com_dispatch_transport_counts();
            assert_eq!(
                session
                    .call_bound(&handle, &[object], 1000)
                    .unwrap()
                    .as_i32(),
                Some(7)
            );
            let after = host.com().com_dispatch_transport_counts();
            assert_eq!(after.0 - before.0, 1);
            assert_eq!(after.1 - before.1, 0);
            assert_eq!(retained.invoke_calls.load(Ordering::SeqCst), 0);
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
