//! Portable (Linux & Windows) end-to-end object-identity and table-driven `IDispatch` tests for
//! VM3 and the JIT (dc-yqq0.23).
//!
//! Exercises the exact scenarios from `docs/design/vba-object-is-one-com-object-2026-11-20.md`
//! through raw `IDispatchVtbl` function pointers (`QueryInterface`, `AddRef`, `Release`,
//! `GetIDsOfNames`, `Invoke`) without requiring Windows COM registration:
//! - One COM identity across `QI(IUnknown)`, `QI(IDispatch)`, `QI(class IID)` -> `&primary`
//! - Foreign store/load round trip: `back Is w` is True and `back.Twice() == 42`
//! - Foreign late-bound call during active run (`s.CallMember(w, "Twice") == 42`)
//! - Foreign late-bound public field Let/Get, named/optional/ByRef/ParamArray args, `Err.Raise` -> `EXCEPINFO`
//! - Foreign raw `AddRef` holder keeping `Class_Terminate` deferred until final `Release`
//! - Off-thread `Invoke` returning `RPC_E_WRONG_THREAD`
//! - Between-runs `ProjectRuntimeSession` `IDispatch::Invoke` and callback during `RaiseEvent` on both VM3 and JIT

use std::collections::HashMap;
use std::ffi::c_void;
use std::sync::{Arc, Mutex};

use calcvba_com::platform::portable::{
    PortableComProjection, PortableDispatch, PortableObjectFactory,
};
use calcvba_engine::{Engine, ExecBackend, HostConfig, RuntimeProfileId};
use calcvba_hal::model::HostPolicy;
use calcvba_oxir::OxImage;
use calcvba_runtime::bstr::BStr;
use calcvba_runtime::com::{
    DISP_E_EXCEPTION, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT,
    DISPID_PROPERTYPUT, DISPID_UNKNOWN, DispParams, ExcepInfo, Guid, IDispatchVtbl, IID_IDISPATCH,
    IID_IUNKNOWN, IID_NULL, RPC_E_WRONG_THREAD,
};
use calcvba_runtime::{ObjectRef, ObjectRefExt, RawRuntimeIUnknown, Variant, VariantCore};
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};

const BACKENDS: [ExecBackend; 2] = [ExecBackend::Vm3, ExecBackend::Jit];

fn utf16_nul(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(core::iter::once(0)).collect()
}

/// Raw COM helper simulating a foreign COM server holding an `IDispatch*` reference and calling
/// `QueryInterface`, `GetIDsOfNames`, `Invoke`, and `Release` through the object's own vtable.
struct RawComPtr(*mut RawRuntimeIUnknown);

// SAFETY: only used on the test's single STA thread (except the explicit wrong-thread probe).
unsafe impl Send for RawComPtr {}

impl RawComPtr {
    fn from_qi(object: &ObjectRef, iid: &Guid) -> Result<Self, i32> {
        let raw = object.raw_iunknown();
        let mut out: *mut c_void = core::ptr::null_mut();
        // SAFETY: `object` retains a live COM pointer whose first word is its vtable.
        let hr = unsafe { ((*(*raw).vtbl).query_interface)(raw.cast(), iid, &mut out) };
        if hr < 0 || out.is_null() {
            return Err(hr);
        }
        Ok(Self(out.cast::<RawRuntimeIUnknown>()))
    }

    fn as_ptr(&self) -> *mut RawRuntimeIUnknown {
        self.0
    }

    fn clone_to_object_ref(&self) -> ObjectRef {
        // SAFETY: `self.0` retains a live COM interface pointer.
        unsafe { ObjectRef::from_raw_iunknown_addref(self.0) }.expect("non-null raw ptr")
    }
}

impl Drop for RawComPtr {
    fn drop(&mut self) {
        if !self.0.is_null() {
            // SAFETY: `self` owns one COM reference obtained via `QueryInterface`.
            unsafe {
                ((*(*self.0).vtbl).release)(self.0.cast());
            }
            self.0 = core::ptr::null_mut();
        }
    }
}

fn raw_get_dispid(ptr: *mut RawRuntimeIUnknown, name: &str) -> Result<i32, i32> {
    let wide = utf16_nul(name);
    let names = [wide.as_ptr()];
    let mut dispid = DISPID_UNKNOWN;
    // SAFETY: `ptr` carries a `ClassDualVtbl` whose prefix is `IDispatchVtbl`.
    let hr = unsafe {
        let vtbl = &*((*ptr).vtbl.cast::<IDispatchVtbl>());
        (vtbl.get_ids_of_names)(ptr.cast(), &IID_NULL, names.as_ptr(), 1, 0, &mut dispid)
    };
    if hr < 0 { Err(hr) } else { Ok(dispid) }
}

fn raw_invoke_simple(
    ptr: *mut RawRuntimeIUnknown,
    name: &str,
    flags: u16,
    args: &[Variant],
) -> Result<Variant, (i32, Option<(i32, String)>)> {
    let dispid = raw_get_dispid(ptr, name).map_err(|hr| (hr, None))?;
    // `DISPPARAMS.rgvarg` stores positional args in reverse order (last first).
    let mut owned: Vec<Variant> = args.iter().rev().cloned().collect();
    let mut named_ids = if flags & DISPATCH_PROPERTYPUT != 0 {
        vec![DISPID_PROPERTYPUT]
    } else {
        Vec::new()
    };
    let mut params = DispParams {
        rgvarg: if owned.is_empty() {
            core::ptr::null_mut()
        } else {
            owned.as_mut_ptr().cast::<VariantCore>()
        },
        rgdispid_named_args: if named_ids.is_empty() {
            core::ptr::null_mut()
        } else {
            named_ids.as_mut_ptr()
        },
        c_args: owned.len() as u32,
        c_named_args: named_ids.len() as u32,
    };
    let mut result = Variant::empty();
    let mut excep = ExcepInfo::default();
    let mut arg_err = 0u32;
    // SAFETY: `ptr` carries a `ClassDualVtbl`; all out-slots and `DispParams` are initialized.
    let hr = unsafe {
        let vtbl = &*((*ptr).vtbl.cast::<IDispatchVtbl>());
        (vtbl.invoke)(
            ptr.cast(),
            dispid,
            &IID_NULL,
            0,
            flags,
            (&raw mut params).cast(),
            (&raw mut result).cast(),
            (&raw mut excep).cast(),
            &raw mut arg_err,
        )
    };
    if hr < 0 {
        let info = if hr == DISP_E_EXCEPTION {
            // SAFETY: `Invoke` transferred ownership of any non-null BSTRs in `excep`.
            let (desc, _src, _help) = unsafe {
                (
                    BStr::from_raw_bstr(excep.bstr_description),
                    BStr::from_raw_bstr(excep.bstr_source),
                    BStr::from_raw_bstr(excep.bstr_help_file),
                )
            };
            Some((excep.scode, desc.as_str().to_string()))
        } else {
            None
        };
        return Err((hr, info));
    }
    Ok(result)
}

/// Portable stand-in for `OxVba.TestEventServer` + `Scripting.Dictionary` that interacts with
/// project objects strictly through their raw `IUnknown`/`IDispatch` vtables.
struct PortableForeignServer {
    held: Mutex<Option<RawComPtr>>,
    bag: Mutex<HashMap<String, RawComPtr>>,
}

impl PortableForeignServer {
    fn new() -> Self {
        Self {
            held: Mutex::new(None),
            bag: Mutex::new(HashMap::new()),
        }
    }
}

impl PortableDispatch for PortableForeignServer {
    fn get(&self, member: &str) -> Result<Variant, String> {
        Err(format!("unsupported get `{member}`"))
    }

    fn invoke(&self, member: &str, args: &[Variant]) -> Result<Variant, String> {
        if member.eq_ignore_ascii_case("Store") && args.len() == 2 {
            let key = args[0]
                .as_bstr()
                .map(|s| s.as_str().to_string())
                .ok_or_else(|| "Store key must be String".to_string())?;
            let object = args[1]
                .as_object_ref()
                .ok_or_else(|| "Store value must be Object".to_string())?;
            let disp = RawComPtr::from_qi(&object, &IID_IDISPATCH)
                .map_err(|hr| format!("QI(IDispatch) failed: {hr:#010X}"))?;
            // Verify COM identity rule: QI(IDispatch) returns the exact same pointer as QI(IUnknown).
            let unk = RawComPtr::from_qi(&object, &IID_IUNKNOWN)
                .map_err(|hr| format!("QI(IUnknown) failed: {hr:#010X}"))?;
            assert_eq!(
                disp.as_ptr(),
                unk.as_ptr(),
                "QI(IDispatch) and QI(IUnknown) must share one COM identity pointer"
            );
            self.bag.lock().expect("bag").insert(key, disp);
            return Ok(Variant::empty());
        }
        if member.eq_ignore_ascii_case("Load") && args.len() == 1 {
            let key = args[0]
                .as_bstr()
                .map(|s| s.as_str().to_string())
                .ok_or_else(|| "Load key must be String".to_string())?;
            let bag = self.bag.lock().expect("bag");
            let held = bag
                .get(&key)
                .ok_or_else(|| format!("missing key `{key}`"))?;
            return Ok(Variant::from_object_ref(held.clone_to_object_ref()));
        }
        if member.eq_ignore_ascii_case("CallMember") && args.len() >= 2 {
            let object = args[0]
                .as_object_ref()
                .ok_or_else(|| "CallMember target must be Object".to_string())?;
            let name = args[1]
                .as_bstr()
                .map(|s| s.as_str().to_string())
                .ok_or_else(|| "CallMember name must be String".to_string())?;
            let disp = RawComPtr::from_qi(&object, &IID_IDISPATCH)
                .map_err(|hr| format!("QI(IDispatch) failed: {hr:#010X}"))?;
            return raw_invoke_simple(
                disp.as_ptr(),
                &name,
                DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                &args[2..],
            )
            .map_err(|(hr, info)| match info {
                Some((scode, desc)) => format!("DISP_E_EXCEPTION ({scode:#010X}): {desc}"),
                None => format!("IDispatch::Invoke failed: {hr:#010X}"),
            });
        }
        if member.eq_ignore_ascii_case("PutField") && args.len() == 3 {
            let object = args[0]
                .as_object_ref()
                .ok_or_else(|| "PutField target must be Object".to_string())?;
            let name = args[1]
                .as_bstr()
                .map(|s| s.as_str().to_string())
                .ok_or_else(|| "PutField name must be String".to_string())?;
            let disp = RawComPtr::from_qi(&object, &IID_IDISPATCH)
                .map_err(|hr| format!("QI(IDispatch) failed: {hr:#010X}"))?;
            return raw_invoke_simple(
                disp.as_ptr(),
                &name,
                DISPATCH_PROPERTYPUT,
                &[args[2].clone()],
            )
            .map_err(|(hr, _)| format!("IDispatch::Invoke(Put) failed: {hr:#010X}"));
        }
        if member.eq_ignore_ascii_case("CallByRef") && args.len() == 2 {
            // Invokes `Bump(ByRef n As Long)` through raw `VT_BYREF | VT_I4` and returns updated `n`.
            let object = args[0]
                .as_object_ref()
                .ok_or_else(|| "CallByRef target must be Object".to_string())?;
            let mut value = args[1]
                .as_i32()
                .or_else(|| args[1].as_i16().map(i32::from))
                .unwrap_or(0);
            let disp = RawComPtr::from_qi(&object, &IID_IDISPATCH)
                .map_err(|hr| format!("QI(IDispatch) failed: {hr:#010X}"))?;
            let dispid = raw_get_dispid(disp.as_ptr(), "Bump")
                .map_err(|hr| format!("GetIDsOfNames(Bump) failed: {hr:#010X}"))?;
            // SAFETY: `value` lives across the synchronous `Invoke` call.
            let byref_arg = unsafe { Variant::by_ref(0x0003, (&raw mut value).cast()) };
            let mut rgvarg = [byref_arg.core()];
            let mut params = DispParams {
                rgvarg: rgvarg.as_mut_ptr(),
                rgdispid_named_args: core::ptr::null_mut(),
                c_args: 1,
                c_named_args: 0,
            };
            let mut result = Variant::empty();
            // SAFETY: `disp` carries a `ClassDualVtbl`; `params` and `result` are initialized.
            let hr = unsafe {
                let vtbl = &*((*disp.as_ptr()).vtbl.cast::<IDispatchVtbl>());
                (vtbl.invoke)(
                    disp.as_ptr().cast(),
                    dispid,
                    &IID_NULL,
                    0,
                    DISPATCH_METHOD,
                    (&raw mut params).cast(),
                    (&raw mut result).cast(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                )
            };
            if hr < 0 {
                return Err(format!("IDispatch::Invoke(Bump) failed: {hr:#010X}"));
            }
            return Ok(Variant::from_i32(value));
        }
        if member.eq_ignore_ascii_case("Hold") && args.len() == 1 {
            let object = args[0]
                .as_object_ref()
                .ok_or_else(|| "Hold target must be Object".to_string())?;
            let unk = RawComPtr::from_qi(&object, &IID_IUNKNOWN)
                .map_err(|hr| format!("QI(IUnknown) failed: {hr:#010X}"))?;
            *self.held.lock().expect("held") = Some(unk);
            return Ok(Variant::empty());
        }
        if member.eq_ignore_ascii_case("Drop") && args.is_empty() {
            *self.held.lock().expect("held") = None;
            return Ok(Variant::empty());
        }
        if member.eq_ignore_ascii_case("DropOnForeignThread") && args.is_empty() {
            let held = self.held.lock().expect("held").take();
            std::thread::spawn(move || {
                // Final COM Release happens strictly on this foreign OS thread.
                drop(held);
            })
            .join()
            .expect("foreign drop thread");
            return Ok(Variant::empty());
        }
        if member.eq_ignore_ascii_case("CurrentThreadId") && args.is_empty() {
            return Ok(Variant::from_i64(
                calcvba_runtime::current_thread_id() as i64
            ));
        }
        if member.eq_ignore_ascii_case("WrongThreadCheck") && args.len() == 1 {
            let object = args[0]
                .as_object_ref()
                .ok_or_else(|| "WrongThreadCheck target must be Object".to_string())?;
            let disp = RawComPtr::from_qi(&object, &IID_IDISPATCH)
                .map_err(|hr| format!("QI(IDispatch) failed: {hr:#010X}"))?;
            let dispid = raw_get_dispid(disp.as_ptr(), "Twice")
                .map_err(|hr| format!("GetIDsOfNames(Twice) failed: {hr:#010X}"))?;
            let ptr_addr = disp.as_ptr() as usize;
            let hr = std::thread::spawn(move || {
                let ptr = ptr_addr as *mut RawRuntimeIUnknown;
                let mut params = DispParams::default();
                // SAFETY: `disp` stays retained on the owning thread until `join()` completes.
                unsafe {
                    let vtbl = &*((*ptr).vtbl.cast::<IDispatchVtbl>());
                    (vtbl.invoke)(
                        ptr.cast(),
                        dispid,
                        &IID_NULL,
                        0,
                        DISPATCH_METHOD,
                        (&raw mut params).cast(),
                        core::ptr::null_mut(),
                        core::ptr::null_mut(),
                        core::ptr::null_mut(),
                    )
                }
            })
            .join()
            .expect("worker thread");
            return Ok(Variant::from_bool(hr == RPC_E_WRONG_THREAD));
        }
        Err(format!("unsupported method `{member}`"))
    }

    fn put(&self, member: &str, _value: Variant) -> Result<(), String> {
        Err(format!("unsupported put `{member}`"))
    }

    fn member_names(&self) -> Vec<String> {
        vec![
            "Store".into(),
            "Load".into(),
            "CallMember".into(),
            "PutField".into(),
            "CallByRef".into(),
            "Hold".into(),
            "Drop".into(),
            "DropOnForeignThread".into(),
            "CurrentThreadId".into(),
            "WrongThreadCheck".into(),
        ]
    }
}

struct PortableServerFactory;
impl PortableObjectFactory for PortableServerFactory {
    fn create(&self) -> Box<dyn PortableDispatch> {
        Box::new(PortableForeignServer::new())
    }
}

const WIDGET_SRC: &str = r#"
Public Value As Long

Public Function Twice() As Long
    Twice = Value * 2
End Function

Public Function AddOpt(ByVal a As Long, Optional ByVal b As Long = 10) As Long
    AddOpt = a + b
End Function

Public Function SumAll(ParamArray items() As Variant) As Long
    Dim i As Long
    Dim acc As Long
    For i = LBound(items) To UBound(items)
        acc = acc + CLng(items(i))
    Next i
    SumAll = acc
End Function

Public Sub Bump(ByRef n As Long)
    n = n + 7
End Sub

Public Function Boom() As Long
    Err.Raise 5000, "Widget", "boom-from-widget"
End Function

Private Sub Class_Terminate()
    gTerm = gTerm + 1
End Sub
"#;

const MAIN_SRC: &str = r#"
Public gTerm As Long
Public gOut As String

Sub Main()
    Dim w As New Widget, s As Object, back As Object
    w.Value = 21
    Set s = CreateObject("Portable.ForeignServer")

    ' 1. Foreign raw IDispatch* round-trip: Is holds and member call returns 42.
    s.Store "k", w
    Set back = s.Load("k")
    gOut = "store-is=" & CStr(back Is w) & ";store-call=" & CStr(back.Twice())

    ' 2. Foreign late-bound call through Widget's own table-driven IDispatch::Invoke -> 42.
    gOut = gOut & ";foreign-call=" & CStr(s.CallMember(w, "Twice"))

    ' 3. Foreign late-bound public field Put + Get + Optional default + ParamArray + ByRef.
    s.PutField w, "Value", 30
    gOut = gOut & ";field-get=" & CStr(s.CallMember(w, "Value"))
    gOut = gOut & ";opt=" & CStr(s.CallMember(w, "AddOpt", 5))
    gOut = gOut & ";paramarray=" & CStr(s.CallMember(w, "SumAll", 2, 4, 6))
    gOut = gOut & ";byref=" & CStr(s.CallByRef(w, 10))

    ' 4. Foreign call raising Err.Raise returns DISP_E_EXCEPTION and is catchable by On Error.
    On Error Resume Next
    Dim ignored As Long
    ignored = s.CallMember(w, "Boom")
    gOut = gOut & ";boom-caught=" & CStr(Err.Number <> 0)
    Err.Clear
    On Error GoTo 0

    ' 5. Off-thread Invoke returns RPC_E_WRONG_THREAD.
    gOut = gOut & ";wrong-thread=" & CStr(s.WrongThreadCheck(w))

    ' 6. Foreign raw AddRef holder keeps Class_Terminate deferred until final Release.
    Dim h As Widget
    Set h = New Widget
    gTerm = 0
    s.Hold h
    Set h = Nothing
    gOut = gOut & ";term-held=" & CStr(gTerm)
    s.Drop
    gOut = gOut & ";term-dropped=" & CStr(gTerm)
End Sub
"#;

#[test]
fn portable_identity_foreign_invoke_and_terminate_match_excel_on_vm3_and_jit() {
    let manifest = SymbolProjectManifest {
        project_name: "PortableObjId".to_string(),
        project_kind: ProjectKind::Source,
        modules: vec![
            ModuleUnit {
                module_name: "Probe".to_string(),
                module_kind: ModuleKind::Procedural,
                attributes: ModuleAttributes::named("Probe"),
                source: MAIN_SRC.to_string(),
            },
            ModuleUnit {
                module_name: "Widget".to_string(),
                module_kind: ModuleKind::Class,
                attributes: ModuleAttributes::named("Widget"),
                source: WIDGET_SRC.to_string(),
            },
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };

    for backend in BACKENDS {
        let projection = Arc::new(PortableComProjection::new());
        projection.register_object("Portable.ForeignServer", Arc::new(PortableServerFactory));
        let engine = Engine::new(HostConfig { backend })
            .with_runtime_profile(RuntimeProfileId::WindowsHeadless)
            .with_portable_com_projection(projection);
        let values = engine
            .execute_project_closure_with_variant_snapshot(std::slice::from_ref(&manifest))
            .unwrap_or_else(|err| panic!("{backend:?}: {err:?}"));
        let output = values
            .iter()
            .filter_map(|v| v.as_bstr().map(|s| s.as_str().to_string()))
            .find(|s| s.contains('='))
            .unwrap_or_else(|| panic!("{backend:?}: missing gOut in {values:?}"));
        assert_eq!(
            output,
            "store-is=True;store-call=42;foreign-call=42;field-get=30;opt=15;paramarray=12;byref=17;boom-caught=True;wrong-thread=True;term-held=0;term-dropped=1",
            "{backend:?}"
        );
    }
}

#[test]
fn session_between_runs_and_raise_event_callback_work_through_raw_idispatch_on_vm3_and_jit() {
    let class_src = r#"
Public Event Ping(ByVal x As Long)
Public Value As Long

Public Function Twice() As Long
    Twice = Value * 2
End Function

Public Function Trigger(ByVal x As Long) As Long
    RaiseEvent Ping(x)
    Trigger = Value
End Function
"#;
    let mut attributes = ModuleAttributes::named("Emitter");
    attributes.vb_exposed = true;
    attributes.vb_creatable = true;
    let manifest = SymbolProjectManifest {
        project_name: "PortableSessionDispatch".to_string(),
        project_kind: ProjectKind::Library,
        modules: vec![ModuleUnit {
            module_name: "Emitter".to_string(),
            module_kind: ModuleKind::Class,
            attributes,
            source: class_src.to_string(),
        }],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let typelibs = calcvba_symbol::CatalogTypeLibResolver;
    let program = calcvba_bind::bind_program(&manifest, &typelibs).expect("bind");
    let oxp = calcvba_oxir::elaborate::elaborate(&program).expect("elaborate");

    for backend in BACKENDS {
        let mut engine = Engine::new(HostConfig { backend });
        engine.set_host_policy(HostPolicy::native());
        let mut session = engine
            .prepare_image_session(OxImage::new(vec![oxp.clone()]))
            .unwrap_or_else(|err| panic!("{backend:?} prepare_image_session: {err}"));

        // During `RaiseEvent`, the host listener calls back into `source`'s raw `IDispatch::Invoke`
        // directly on the active run (closes the JIT callback-during-event re-entry gap!).
        let callback_seen = Arc::new(Mutex::new(Vec::<i32>::new()));
        let callback_seen_sink = Arc::clone(&callback_seen);
        session.set_project_event_sink(move |source, _event_id, args| {
            let disp = RawComPtr::from_qi(&source, &IID_IDISPATCH)
                .map_err(|hr| format!("QI(IDispatch) failed: {hr:#010X}"))?;
            let before = raw_invoke_simple(
                disp.as_ptr(),
                "Twice",
                DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                &[],
            )
            .map_err(|(hr, _)| format!("Twice failed: {hr:#010X}"))?;
            callback_seen_sink
                .lock()
                .expect("seen")
                .push(before.as_i32().unwrap_or(-1));
            let next_val = args[0].as_i32().unwrap_or(0);
            raw_invoke_simple(
                disp.as_ptr(),
                "Value",
                DISPATCH_PROPERTYPUT,
                &[Variant::from_i32(next_val)],
            )
            .map_err(|(hr, _)| format!("Value put failed: {hr:#010X}"))?;
            Ok(())
        });

        let emitter = session
            .create_class_instance("Emitter")
            .unwrap_or_else(|err| panic!("{backend:?} create Emitter: {err}"));

        // Verify class default IID answers QueryInterface with &primary.
        let class_iid = emitter
            .class_descriptor()
            .default_interface_iid
            .expect("Emitter default IID");
        let iface_ptr = RawComPtr::from_qi(&emitter, &class_iid).expect("QI(IEmitter)");
        let disp_ptr = RawComPtr::from_qi(&emitter, &IID_IDISPATCH).expect("QI(IDispatch)");
        assert_eq!(iface_ptr.as_ptr(), disp_ptr.as_ptr());

        // Between-runs `IDispatch::Invoke` via `ProjectRuntimeSession::enter_member_invoker_scope`.
        {
            // SAFETY: `session` is exclusively borrowed by the invoker scope for these calls.
            let _invoker = unsafe { session.enter_member_invoker_scope() };
            raw_invoke_simple(
                disp_ptr.as_ptr(),
                "Value",
                DISPATCH_PROPERTYPUT,
                &[Variant::from_i32(21)],
            )
            .expect("put Value = 21");
            let twice = raw_invoke_simple(
                disp_ptr.as_ptr(),
                "Twice",
                DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                &[],
            )
            .expect("call Twice");
            assert_eq!(twice.as_i32(), Some(42), "{backend:?} between-runs Twice");

            // Trigger RaiseEvent -> host sink calls Twice() (sees 42) and sets Value = 55 -> returns 55!
            let triggered = raw_invoke_simple(
                disp_ptr.as_ptr(),
                "Trigger",
                DISPATCH_METHOD,
                &[Variant::from_i32(55)],
            )
            .expect("call Trigger");
            assert_eq!(
                triggered.as_i32(),
                Some(55),
                "{backend:?} Trigger after re-entrant RaiseEvent callback"
            );
        }
        assert_eq!(
            *callback_seen.lock().expect("seen"),
            vec![42],
            "{backend:?} re-entrant Twice during RaiseEvent"
        );
    }
}

#[test]
fn foreign_thread_final_release_runs_class_terminate_once_on_owner_thread_on_vm3_and_jit() {
    let widget_src = r#"
Public Value As Long
Private Sub Class_Terminate()
    gTermCount = gTermCount + 1
    gTermThread = gServer.CurrentThreadId()
End Sub
"#;
    let wrapper_src = r#"
Public Child As Object
"#;
    let main_src = r#"
Public gServer As Object
Public gOwnerThread As LongLong
Public gTermThread As LongLong
Public gTermCount As Long
Public gOut As String

Sub Main()
    Set gServer = CreateObject("Portable.ForeignServer")
    gOwnerThread = gServer.CurrentThreadId()

    ' 1. Direct Widget held by foreign holder, released on a foreign OS thread.
    Dim w As Widget
    Set w = New Widget
    w.Value = 21
    gTermCount = 0
    gTermThread = 0
    gServer.Hold w
    Set w = Nothing
    gOut = "held=" & CStr(gTermCount)
    gServer.DropOnForeignThread
    gOut = gOut & ";after-foreign-drop=" & CStr(gTermCount) & ";same-thread=" & CStr(gTermThread = gOwnerThread)

    ' 2. Non-terminating Wrapper holding a Widget, released on a foreign OS thread:
    '    Wrapper's fields must drop on the owner thread, running Widget.Class_Terminate once there.
    Dim wrap As Wrapper
    Set wrap = New Wrapper
    Set wrap.Child = New Widget
    gTermCount = 0
    gTermThread = 0
    gServer.Hold wrap
    Set wrap = Nothing
    gOut = gOut & ";wrap-held=" & CStr(gTermCount)
    gServer.DropOnForeignThread
    gOut = gOut & ";wrap-after=" & CStr(gTermCount) & ";wrap-same-thread=" & CStr(gTermThread = gOwnerThread)
End Sub
"#;

    let manifest = SymbolProjectManifest {
        project_name: "ForeignRelease".to_string(),
        project_kind: ProjectKind::Source,
        modules: vec![
            ModuleUnit {
                module_name: "Probe".to_string(),
                module_kind: ModuleKind::Procedural,
                attributes: ModuleAttributes::named("Probe"),
                source: main_src.to_string(),
            },
            ModuleUnit {
                module_name: "Widget".to_string(),
                module_kind: ModuleKind::Class,
                attributes: ModuleAttributes::named("Widget"),
                source: widget_src.to_string(),
            },
            ModuleUnit {
                module_name: "Wrapper".to_string(),
                module_kind: ModuleKind::Class,
                attributes: ModuleAttributes::named("Wrapper"),
                source: wrapper_src.to_string(),
            },
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };

    for backend in BACKENDS {
        let projection = Arc::new(PortableComProjection::new());
        projection.register_object("Portable.ForeignServer", Arc::new(PortableServerFactory));
        let engine = Engine::new(HostConfig { backend })
            .with_runtime_profile(RuntimeProfileId::WindowsHeadless)
            .with_portable_com_projection(projection);
        let values = engine
            .execute_project_closure_with_variant_snapshot(std::slice::from_ref(&manifest))
            .unwrap_or_else(|err| panic!("{backend:?}: {err:?}"));
        let output = values
            .iter()
            .filter_map(|v| v.as_bstr().map(|s| s.as_str().to_string()))
            .find(|s| s.contains('='))
            .unwrap_or_else(|| panic!("{backend:?}: missing gOut in {values:?}"));
        assert_eq!(
            output,
            "held=0;after-foreign-drop=1;same-thread=True;wrap-held=0;wrap-after=1;wrap-same-thread=True",
            "{backend:?}"
        );
    }
}
