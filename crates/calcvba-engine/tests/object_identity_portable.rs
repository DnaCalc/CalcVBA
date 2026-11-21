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

#[test]
fn between_runs_session_invoker_matches_comhost_named_byref_identity_and_terminate_on_vm3_and_jit()
{
    use calcvba_runtime::com::{DISP_E_PARAMNOTFOUND, DISP_E_UNKNOWNNAME};
    use calcvba_runtime::events::{MemberFault, MemberInvoker};
    use std::cell::RefCell;
    use std::rc::Rc;

    let globals_src = r#"
Public gTerm As Long
Public Function ReadTerm() As Long
    ReadTerm = gTerm
End Function
"#;
    let child_src = r#"
Public Tag As Long
Public Function Label() As Long
    Label = Tag + 100
End Function
Private Sub Class_Terminate()
    gTerm = gTerm + 1
End Sub
"#;
    let mixer_src = r#"
Public Function Mix(ByVal a As Long, Optional ByVal b As Long = 10, Optional ByVal c As Long = 100) As Long
    Mix = a * 10000 + b * 100 + c
End Function

Public Sub Bump(ByRef target As Long, ByVal by As Long)
    target = target + by
End Sub

Public Sub AliasPair(ByRef first As Long, ByRef second As Long)
    first = first + 1
    second = second + 10
End Sub

Public Function SelfObj() As Object
    Set SelfObj = Me
End Function

Public Function IsSame(ByVal other As Object) As Boolean
    IsSame = (other Is Me)
End Function

Public Function MakeChild(ByVal tag As Long) As Object
    Dim c As New Child
    c.Tag = tag
    Set MakeChild = c
End Function

Public Function TermCount() As Long
    TermCount = gTerm
End Function

Private Sub Class_Terminate()
    gTerm = gTerm + 1
End Sub
"#;

    let mut mixer_attrs = ModuleAttributes::named("Mixer");
    mixer_attrs.vb_exposed = true;
    mixer_attrs.vb_creatable = true;
    let mut child_attrs = ModuleAttributes::named("Child");
    child_attrs.vb_exposed = true;
    let manifest = SymbolProjectManifest {
        project_name: "DirectComHostPortable".to_string(),
        project_kind: ProjectKind::Library,
        modules: vec![
            ModuleUnit {
                module_name: "State".to_string(),
                module_kind: ModuleKind::Procedural,
                attributes: ModuleAttributes::named("State"),
                source: globals_src.to_string(),
            },
            ModuleUnit {
                module_name: "Child".to_string(),
                module_kind: ModuleKind::Class,
                attributes: child_attrs,
                source: child_src.to_string(),
            },
            ModuleUnit {
                module_name: "Mixer".to_string(),
                module_kind: ModuleKind::Class,
                attributes: mixer_attrs,
                source: mixer_src.to_string(),
            },
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let typelibs = calcvba_symbol::CatalogTypeLibResolver;
    let program = calcvba_bind::bind_program(&manifest, &typelibs).expect("bind");
    let oxp = calcvba_oxir::elaborate::elaborate(&program).expect("elaborate");

    // Mirrors `calcvba-comhost`: a session holder installs a weak between-runs `MemberInvoker`
    // on the session's `RuntimeObjectOwner` so every project object (created via factory OR via
    // `New` inside VBA) crosses as itself and handles raw `IDispatch::Invoke` with no wrapper.
    struct SessionHolderInvoker {
        session: Rc<RefCell<Option<calcvba_engine::ProjectRuntimeSession>>>,
    }
    impl MemberInvoker for SessionHolderInvoker {
        fn invoke_member(
            &mut self,
            object: ObjectRef,
            member: &'static calcvba_runtime::RuntimeMemberDescriptor,
            invoke_kind: calcvba_runtime::RuntimeMemberInvokeKind,
            args: &mut [Variant],
            aliases: &[usize],
        ) -> Result<Variant, MemberFault> {
            let mut session = self
                .session
                .borrow_mut()
                .take()
                .ok_or_else(|| MemberFault::from_number(5))?;
            let result = MemberInvoker::invoke_member(
                &mut session,
                object,
                member,
                invoke_kind,
                args,
                aliases,
            );
            *self.session.borrow_mut() = Some(session);
            result
        }
    }

    let raw_get_ids = |ptr: *mut RawRuntimeIUnknown, names: &[&str]| -> (i32, Vec<i32>) {
        let wide: Vec<Vec<u16>> = names.iter().map(|n| utf16_nul(n)).collect();
        let ptrs: Vec<*const u16> = wide.iter().map(|w| w.as_ptr()).collect();
        let mut ids = vec![0i32; names.len()];
        // SAFETY: `ptr` carries `ClassDualVtbl` whose prefix is `IDispatchVtbl`.
        let hr = unsafe {
            let vtbl = &*((*ptr).vtbl.cast::<IDispatchVtbl>());
            (vtbl.get_ids_of_names)(
                ptr.cast(),
                &IID_NULL,
                ptrs.as_ptr(),
                names.len() as u32,
                0,
                ids.as_mut_ptr(),
            )
        };
        (hr, ids)
    };

    let raw_call_named = |ptr: *mut RawRuntimeIUnknown,
                          member: &str,
                          rgvarg: &mut [VariantCore],
                          named: &[i32]|
     -> (i32, Variant) {
        let (hr, ids) = raw_get_ids(ptr, &[member]);
        assert_eq!(hr, 0, "GetIDsOfNames({member})");
        let mut named = named.to_vec();
        let mut params = DispParams {
            rgvarg: if rgvarg.is_empty() {
                core::ptr::null_mut()
            } else {
                rgvarg.as_mut_ptr()
            },
            rgdispid_named_args: if named.is_empty() {
                core::ptr::null_mut()
            } else {
                named.as_mut_ptr()
            },
            c_args: rgvarg.len() as u32,
            c_named_args: named.len() as u32,
        };
        let mut result = Variant::empty();
        // SAFETY: `ptr` carries `ClassDualVtbl`; `params` and `result` are initialized.
        let hr = unsafe {
            let vtbl = &*((*ptr).vtbl.cast::<IDispatchVtbl>());
            (vtbl.invoke)(
                ptr.cast(),
                ids[0],
                &IID_NULL,
                0,
                DISPATCH_METHOD,
                (&raw mut params).cast(),
                (&raw mut result).cast(),
                core::ptr::null_mut(),
                core::ptr::null_mut(),
            )
        };
        (hr, result)
    };

    for backend in BACKENDS {
        let mut engine = Engine::new(HostConfig { backend });
        engine.set_host_policy(HostPolicy::native());
        let mut session = engine
            .prepare_image_session(OxImage::new(vec![oxp.clone()]))
            .unwrap_or_else(|err| panic!("{backend:?} prepare_image_session: {err}"));

        let session_cell = Rc::new(RefCell::new(None));
        let invoker: Rc<RefCell<dyn MemberInvoker>> = Rc::new(RefCell::new(SessionHolderInvoker {
            session: Rc::clone(&session_cell),
        }));
        session.set_member_invoker(Rc::downgrade(&invoker));

        // Factory creates `Mixer` and hands out only its raw COM `IDispatch*` (`&primary`).
        let mixer_raw = {
            let mixer_obj = session
                .create_class_instance("Mixer")
                .unwrap_or_else(|err| panic!("{backend:?} create Mixer: {err}"));
            let raw = RawComPtr::from_qi(&mixer_obj, &IID_IDISPATCH).expect("QI(IDispatch)");
            // Verify QI(IUnknown) returns the exact same pointer.
            let unk = RawComPtr::from_qi(&mixer_obj, &IID_IUNKNOWN).expect("QI(IUnknown)");
            assert_eq!(raw.as_ptr(), unk.as_ptr());
            raw
        };
        *session_cell.borrow_mut() = Some(session);

        // 1. Named-argument DISPIDs and unknown-parameter handling.
        assert_eq!(
            raw_get_ids(mixer_raw.as_ptr(), &["Mix", "c", "a"]).1[1..],
            [2, 0]
        );
        let (hr, ids) = raw_get_ids(mixer_raw.as_ptr(), &["Mix", "a", "nope"]);
        assert_eq!((hr, &ids[1..]), (DISP_E_UNKNOWNNAME, &[0, -1][..]));

        // 2. Named + Optional default calls with NO `MemberInvokerScope` active on the stack.
        let mut args_named = [Variant::from_i32(3).core(), Variant::from_i32(1).core()];
        let (hr, val) = raw_call_named(mixer_raw.as_ptr(), "Mix", &mut args_named, &[2, 0]);
        assert_eq!(
            (hr, val.as_i32()),
            (0, Some(11003)),
            "{backend:?} Mix(c:=3, a:=1)"
        );

        let mut args_mixed = [Variant::from_i32(5).core(), Variant::from_i32(1).core()];
        let (hr, val) = raw_call_named(mixer_raw.as_ptr(), "Mix", &mut args_mixed, &[2]);
        assert_eq!(
            (hr, val.as_i32()),
            (0, Some(11005)),
            "{backend:?} Mix(1, c:=5)"
        );

        let mut args_pos = [
            Variant::from_i32(3).core(),
            Variant::from_i32(2).core(),
            Variant::from_i32(1).core(),
        ];
        let (hr, val) = raw_call_named(mixer_raw.as_ptr(), "Mix", &mut args_pos, &[]);
        assert_eq!(
            (hr, val.as_i32()),
            (0, Some(10203)),
            "{backend:?} Mix(1, 2, 3)"
        );

        // Duplicate / out-of-range named argument -> DISP_E_PARAMNOTFOUND.
        let mut dup = [Variant::from_i32(1).core(), Variant::from_i32(1).core()];
        assert_eq!(
            raw_call_named(mixer_raw.as_ptr(), "Mix", &mut dup, &[0, 0]).0,
            DISP_E_PARAMNOTFOUND
        );
        let mut bad = [Variant::from_i32(1).core()];
        assert_eq!(
            raw_call_named(mixer_raw.as_ptr(), "Mix", &mut bad, &[7]).0,
            DISP_E_PARAMNOTFOUND
        );

        // 3. Named ByRef + forwarding `VT_BYREF | VT_VARIANT -> VT_BYREF | VT_I4` + ByRef aliasing.
        let mut x = 2i32;
        // SAFETY: `x` stays live across the synchronous `Invoke` call.
        let inner_ref = unsafe { Variant::by_ref(0x0003, (&raw mut x).cast()) };
        let mut inner_cell = inner_ref.core();
        // SAFETY: `inner_cell` stays live across the synchronous `Invoke` call.
        let fwd_ref = unsafe { Variant::by_ref(0x000C, (&raw mut inner_cell).cast()) };
        let mut bump_args = [Variant::from_i32(5).core(), fwd_ref.core()];
        let (hr, _) = raw_call_named(mixer_raw.as_ptr(), "Bump", &mut bump_args, &[1, 0]);
        assert_eq!((hr, x), (0, 7), "{backend:?} forwarding ByRef Bump");

        let mut aliased = 0i32;
        // SAFETY: `aliased` stays live across the synchronous `Invoke` call.
        let r1 = unsafe { Variant::by_ref(0x0003, (&raw mut aliased).cast()) };
        // SAFETY: same live cell passed twice to exercise ByRef aliasing.
        let r2 = unsafe { Variant::by_ref(0x0003, (&raw mut aliased).cast()) };
        let mut pair = [r2.core(), r1.core()];
        let (hr, _) = raw_call_named(mixer_raw.as_ptr(), "AliasPair", &mut pair, &[]);
        assert_eq!((hr, aliased), (0, 11), "{backend:?} ByRef AliasPair(x, x)");

        // 4. Identity after crossing: `SelfObj()` returns `&primary` itself, and passing it back
        //    into `IsSame(other)` evaluates `other Is Me == True` with zero unwrapping.
        let (hr, first) = raw_call_named(mixer_raw.as_ptr(), "SelfObj", &mut [], &[]);
        assert_eq!(hr, 0);
        let (hr, second) = raw_call_named(mixer_raw.as_ptr(), "SelfObj", &mut [], &[]);
        assert_eq!(hr, 0);
        let first_obj = first.as_object_ref().expect("SelfObj returns object");
        let second_obj = second.as_object_ref().expect("SelfObj returns object");
        assert_eq!(first_obj.raw_iunknown(), mixer_raw.as_ptr());
        assert_eq!(second_obj.raw_iunknown(), mixer_raw.as_ptr());
        let (hr, same) = raw_call_named(mixer_raw.as_ptr(), "IsSame", &mut [first.core()], &[]);
        assert_eq!(
            (hr, same.as_bool()),
            (0, Some(true)),
            "{backend:?} IsSame(Self)"
        );
        drop(first_obj);
        drop(second_obj);
        drop(first);
        drop(second);

        // 5. Child object created by VBA `New Child` inherits the session invoker automatically
        //    and runs `Class_Terminate` only after its final COM Release.
        let (hr, child_var) = raw_call_named(
            mixer_raw.as_ptr(),
            "MakeChild",
            &mut [Variant::from_i32(23).core()],
            &[],
        );
        assert_eq!(hr, 0);
        let child_ptr = {
            let child_obj = child_var.as_object_ref().expect("Child object");
            RawComPtr::from_qi(&child_obj, &IID_IDISPATCH).expect("Child QI(IDispatch)")
        };
        drop(child_var);

        let (hr, label) = raw_call_named(child_ptr.as_ptr(), "Label", &mut [], &[]);
        assert_eq!(
            (hr, label.as_i32()),
            (0, Some(123)),
            "{backend:?} Child.Label"
        );

        let (hr, before_drop) = raw_call_named(mixer_raw.as_ptr(), "TermCount", &mut [], &[]);
        assert_eq!(
            (hr, before_drop.as_i32()),
            (0, Some(0)),
            "{backend:?} child still held"
        );

        // Final Release of `child_ptr` queues Child.Class_Terminate; next call / drain runs it once.
        drop(child_ptr);
        session_cell
            .borrow_mut()
            .as_mut()
            .expect("session")
            .drain_object_terminations();
        let (hr, after_drop) = raw_call_named(mixer_raw.as_ptr(), "TermCount", &mut [], &[]);
        assert_eq!(
            (hr, after_drop.as_i32()),
            (0, Some(1)),
            "{backend:?} child terminated once"
        );

        // Final Release of `mixer_raw` terminates Mixer too (verified via drain).
        drop(mixer_raw);
        let mut session = session_cell.borrow_mut().take().expect("session");
        session.drain_object_terminations();
        drop(invoker);
    }
}

/// Portable equivalent of `OxVba.TestEventServer.DescribeDispatch` (dc-yqq0.29): calls
/// `IDispatch::GetTypeInfo(0, 0, &ti)` and dumps both the `TKIND_DISPATCH` half and the
/// `TKIND_INTERFACE` dual half (`GetRefTypeOfImplType(-1)` -> `GetRefTypeInfo`).
fn describe_dispatch_portable(ptr: *mut RawRuntimeIUnknown) -> String {
    use calcvba_runtime::com::{ITypeInfo, TKIND_DISPATCH, TKIND_INTERFACE, TYPEFLAG_FDUAL};
    let mut count = 99u32;
    // SAFETY: `ptr` carries `ClassDualVtbl` whose prefix is `IDispatchVtbl`.
    let hr = unsafe {
        let vtbl = &*((*ptr).vtbl.cast::<IDispatchVtbl>());
        (vtbl.get_type_info_count)(ptr.cast(), &mut count)
    };
    assert_eq!(hr, 0, "GetTypeInfoCount must succeed");
    assert_eq!(count, 1, "dual class must report 1 ITypeInfo");

    let mut ti_ptr: *mut c_void = core::ptr::null_mut();
    // SAFETY: `ptr` carries `ClassDualVtbl`.
    let hr = unsafe {
        let vtbl = &*((*ptr).vtbl.cast::<IDispatchVtbl>());
        (vtbl.get_type_info)(ptr.cast(), 0, 0, &mut ti_ptr)
    };
    assert_eq!(hr, 0, "GetTypeInfo(0) must succeed");
    assert!(!ti_ptr.is_null());

    fn describe_half(ti_ptr: *mut c_void, out: &mut String) -> (i32, u16) {
        use calcvba_runtime::com::{
            ITypeInfo, MEMBERID_NIL, TKIND_DISPATCH, TKIND_INTERFACE, TYPE_E_ELEMENTNOTFOUND,
        };
        // SAFETY: `ti_ptr` is a live `ITypeInfo` pointer returned by `GetTypeInfo`/`GetRefTypeInfo`.
        unsafe {
            let vtbl = &*(*ti_ptr.cast::<ITypeInfo>()).vtbl;
            let mut bstr_name: *mut u16 = core::ptr::null_mut();
            let hr = (vtbl.get_documentation)(
                ti_ptr,
                MEMBERID_NIL,
                &mut bstr_name,
                core::ptr::null_mut(),
                core::ptr::null_mut(),
                core::ptr::null_mut(),
            );
            assert_eq!(hr, 0);
            let type_name = BStr::from_raw_bstr(bstr_name).as_str().to_string();

            let mut pattr = core::ptr::null_mut();
            assert_eq!((vtbl.get_type_attr)(ti_ptr, &mut pattr), 0);
            let attr = *pattr;
            (vtbl.release_type_attr)(ti_ptr, pattr);

            let kind_label = match attr.typekind {
                TKIND_DISPATCH => "TKIND_DISPATCH",
                TKIND_INTERFACE => "TKIND_INTERFACE",
                other => panic!("unexpected typekind {other}"),
            };
            // Normalize vtbl bytes to 64-bit slot width (8 bytes/slot) so expectations match the
            // 64-bit Excel appendix on every pointer width.
            let ptr_size = core::mem::size_of::<*const c_void>() as u16;
            let vtbl_64 = (attr.cb_size_vft / ptr_size) * 8;
            out.push_str(&format!(
                "type {type_name} kind={kind_label} flags=0x{:X} funcs={} vars={} impl={} vtbl={vtbl_64}\n",
                attr.w_type_flags, attr.c_funcs, attr.c_vars, attr.c_impl_types
            ));

            for i in 0..u32::from(attr.c_impl_types) {
                let mut href = 0u32;
                let impl_hr = (vtbl.get_ref_type_of_impl_type)(ti_ptr, i, &mut href);
                assert_eq!(impl_hr, TYPE_E_ELEMENTNOTFOUND);
                let mut impl_flags = 0i32;
                assert_eq!(
                    (vtbl.get_impl_type_flags)(ti_ptr, i, &mut impl_flags),
                    TYPE_E_ELEMENTNOTFOUND
                );
                out.push_str(&format!("  !impl {i} TYPE_E_ELEMENTNOTFOUND\n"));
            }

            for i in 0..u32::from(attr.c_funcs) {
                let mut pfd = core::ptr::null_mut();
                assert_eq!((vtbl.get_func_desc)(ti_ptr, i, &mut pfd), 0);
                let fd = *pfd;
                let max_names = (fd.c_params.max(0) as usize) + 1;
                let mut raw_names = vec![core::ptr::null_mut::<u16>(); max_names];
                let mut got_names = 0u32;
                assert_eq!(
                    (vtbl.get_names)(
                        ti_ptr,
                        fd.memid,
                        raw_names.as_mut_ptr(),
                        max_names as u32,
                        &mut got_names,
                    ),
                    0
                );
                let names: Vec<String> = raw_names[..got_names as usize]
                    .iter()
                    .map(|&raw| BStr::from_raw_bstr(raw).as_str().to_string())
                    .collect();
                let ovft_64 = (fd.o_vft as usize / core::mem::size_of::<*const c_void>()) * 8;
                out.push_str(&format!(
                    "  func memid=0x{:08X} invkind={} oVft={ovft_64} params={} opt={} flags=0x{:X} ret={} {}\n",
                    fd.memid as u32,
                    fd.invkind,
                    fd.c_params,
                    fd.c_params_opt,
                    fd.w_func_flags,
                    fd.elemdesc_func.tdesc.vt,
                    names.join(",")
                ));
                (vtbl.release_func_desc)(ti_ptr, pfd);
            }
            (attr.typekind, attr.w_type_flags)
        }
    }

    let mut text = String::new();
    let (kind, flags) = describe_half(ti_ptr, &mut text);
    assert_eq!(kind, TKIND_DISPATCH);
    assert_ne!(flags & TYPEFLAG_FDUAL, 0);

    // SAFETY: `ti_ptr` is a live `ITypeInfo` for a dual class.
    unsafe {
        let vtbl = &*(*ti_ptr.cast::<ITypeInfo>()).vtbl;
        let mut href = 0u32;
        assert_eq!(
            (vtbl.get_ref_type_of_impl_type)(ti_ptr, u32::MAX, &mut href),
            0
        );
        let mut dual_ptr: *mut c_void = core::ptr::null_mut();
        assert_eq!((vtbl.get_ref_type_info)(ti_ptr, href, &mut dual_ptr), 0);
        assert!(!dual_ptr.is_null());
        text.push_str("--dual\n");
        let (dual_kind, _) = describe_half(dual_ptr, &mut text);
        assert_eq!(dual_kind, TKIND_INTERFACE);

        // Verify round-trip toggle back from TKIND_INTERFACE -> TKIND_DISPATCH and refcount 0.
        let dual_vtbl = &*(*dual_ptr.cast::<ITypeInfo>()).vtbl;
        let mut back_href = 0u32;
        assert_eq!(
            (dual_vtbl.get_ref_type_of_impl_type)(dual_ptr, u32::MAX, &mut back_href),
            0
        );
        let mut back_ptr: *mut c_void = core::ptr::null_mut();
        assert_eq!(
            (dual_vtbl.get_ref_type_info)(dual_ptr, back_href, &mut back_ptr),
            0
        );
        assert!(!back_ptr.is_null());
        let back_vtbl = &*(*back_ptr.cast::<ITypeInfo>()).vtbl;
        assert_eq!((back_vtbl.unknown.release)(back_ptr), 0);
        assert_eq!((dual_vtbl.unknown.release)(dual_ptr), 0);
        assert_eq!((vtbl.unknown.release)(ti_ptr), 0);
    }
    text
}

#[test]
fn class_get_type_info_matches_excel_measured_classes_on_vm3_and_jit() {
    // Exact classes measured in Excel 64-bit (`docs/design/vba-class-dispids-excel-2026-11-21.md`):
    // Widget, W2..W8.
    let classes: Vec<(&str, &str)> = vec![
        (
            "Widget",
            r#"
Public Value As Long
Public Name As String
Public Child As Object
Public Event Changed(ByVal n As Long)
Private mX As Double
Private mObj As Object

Public Function Twice() As Long
    Twice = Value * 2
End Function

Public Sub Reset()
    Value = 0
End Sub

Public Property Get X() As Double
    X = mX
End Property

Public Property Let X(ByVal v As Double)
    mX = v
End Property

Public Property Get Obj() As Object
    Set Obj = mObj
End Property

Public Property Set Obj(ByVal v As Object)
    Set mObj = v
End Property

Public Function WithArgs(ByVal a As Long, Optional ByVal b As Long, Optional ByVal c As Long = 3) As Long
    WithArgs = a + b + c
End Function

Public Function Rest(ParamArray items() As Variant) As Long
    Rest = UBound(items) - LBound(items) + 1
End Function

Friend Sub Hidden()
End Sub

Private Sub Secret()
End Sub

Private Sub Class_Terminate()
End Sub
"#,
        ),
        (
            "W2",
            r#"
Public Function A() As Long
End Function
Public Function B() As Long
End Function
Public Function C() As Long
End Function
"#,
        ),
        (
            "W3",
            r#"
Public Event Ping()
Public Function A() As Long
End Function
Public Function B() As Long
End Function
Private Sub Class_Initialize()
End Sub
Private Sub Class_Terminate()
End Sub
"#,
        ),
        (
            "W4",
            r#"
Public Property Get P1() As Long
End Property
Public Property Get P2() As Long
End Property
Public Property Get P3() As Long
End Property
"#,
        ),
        (
            "W5",
            r#"
Private Sub S()
End Sub
Friend Function F() As Long
End Function
Public Function A() As Long
End Function
"#,
        ),
        (
            "W6",
            r#"
Public Function Zed() As Long
End Function
Public Function Alpha() As Long
End Function
Public Function Mid() As Long
End Function
"#,
        ),
        (
            "W7",
            r#"
Public F1 As Long
Public F2 As Variant
Public Property Get Zed() As Long
End Property
Public Property Get Alpha() As Long
End Property
Public Sub M()
End Sub
"#,
        ),
        (
            "W8",
            r#"
Private mHidden1 As Long
Private mHidden2 As String
Public Function A() As Long
End Function
"#,
        ),
    ];

    let modules: Vec<ModuleUnit> = classes
        .iter()
        .map(|(name, src)| ModuleUnit {
            module_name: (*name).to_string(),
            module_kind: ModuleKind::Class,
            attributes: ModuleAttributes::named(*name),
            source: (*src).to_string(),
        })
        .collect();
    let manifest = SymbolProjectManifest {
        project_name: "TypeInfoExcelProbe".to_string(),
        project_kind: ProjectKind::Library,
        modules,
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    let typelibs = calcvba_symbol::CatalogTypeLibResolver;
    let program = calcvba_bind::bind_program(&manifest, &typelibs).expect("bind");
    let oxp = calcvba_oxir::elaborate::elaborate(&program).expect("elaborate");

    let expected_w2 = "\
type W2 kind=TKIND_DISPATCH flags=0x1C2 funcs=3 vars=0 impl=1 vtbl=80
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x60030000 invkind=1 oVft=56 params=0 opt=0 flags=0x0 ret=3 A
  func memid=0x60030001 invkind=1 oVft=64 params=0 opt=0 flags=0x0 ret=3 B
  func memid=0x60030002 invkind=1 oVft=72 params=0 opt=0 flags=0x0 ret=3 C
--dual
type W2 kind=TKIND_INTERFACE flags=0x1C2 funcs=3 vars=0 impl=1 vtbl=80
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x60030000 invkind=1 oVft=56 params=1 opt=0 flags=0x0 ret=25 A
  func memid=0x60030001 invkind=1 oVft=64 params=1 opt=0 flags=0x0 ret=25 B
  func memid=0x60030002 invkind=1 oVft=72 params=1 opt=0 flags=0x0 ret=25 C
";

    let expected_w3 = "\
type W3 kind=TKIND_DISPATCH flags=0x1C2 funcs=2 vars=0 impl=1 vtbl=72
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x60030000 invkind=1 oVft=56 params=0 opt=0 flags=0x0 ret=3 A
  func memid=0x60030001 invkind=1 oVft=64 params=0 opt=0 flags=0x0 ret=3 B
--dual
type W3 kind=TKIND_INTERFACE flags=0x1C2 funcs=2 vars=0 impl=1 vtbl=72
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x60030000 invkind=1 oVft=56 params=1 opt=0 flags=0x0 ret=25 A
  func memid=0x60030001 invkind=1 oVft=64 params=1 opt=0 flags=0x0 ret=25 B
";

    let expected_w4 = "\
type W4 kind=TKIND_DISPATCH flags=0x1C2 funcs=3 vars=0 impl=1 vtbl=80
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x68030002 invkind=2 oVft=56 params=0 opt=0 flags=0x0 ret=3 P1
  func memid=0x68030001 invkind=2 oVft=64 params=0 opt=0 flags=0x0 ret=3 P2
  func memid=0x68030000 invkind=2 oVft=72 params=0 opt=0 flags=0x0 ret=3 P3
--dual
type W4 kind=TKIND_INTERFACE flags=0x1C2 funcs=3 vars=0 impl=1 vtbl=80
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x68030002 invkind=2 oVft=56 params=1 opt=0 flags=0x0 ret=25 P1
  func memid=0x68030001 invkind=2 oVft=64 params=1 opt=0 flags=0x0 ret=25 P2
  func memid=0x68030000 invkind=2 oVft=72 params=1 opt=0 flags=0x0 ret=25 P3
";

    let expected_w5 = "\
type W5 kind=TKIND_DISPATCH flags=0x1C2 funcs=1 vars=0 impl=1 vtbl=80
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x60030002 invkind=1 oVft=56 params=0 opt=0 flags=0x0 ret=3 A
--dual
type W5 kind=TKIND_INTERFACE flags=0x1C2 funcs=1 vars=0 impl=1 vtbl=80
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x60030002 invkind=1 oVft=56 params=1 opt=0 flags=0x0 ret=25 A
";

    let expected_w6 = "\
type W6 kind=TKIND_DISPATCH flags=0x1C2 funcs=3 vars=0 impl=1 vtbl=80
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x60030000 invkind=1 oVft=56 params=0 opt=0 flags=0x0 ret=3 Zed
  func memid=0x60030001 invkind=1 oVft=64 params=0 opt=0 flags=0x0 ret=3 Alpha
  func memid=0x60030002 invkind=1 oVft=72 params=0 opt=0 flags=0x0 ret=3 Mid
--dual
type W6 kind=TKIND_INTERFACE flags=0x1C2 funcs=3 vars=0 impl=1 vtbl=80
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x60030000 invkind=1 oVft=56 params=1 opt=0 flags=0x0 ret=25 Zed
  func memid=0x60030001 invkind=1 oVft=64 params=1 opt=0 flags=0x0 ret=25 Alpha
  func memid=0x60030002 invkind=1 oVft=72 params=1 opt=0 flags=0x0 ret=25 Mid
";

    let expected_w7 = "\
type W7 kind=TKIND_DISPATCH flags=0x1C2 funcs=8 vars=0 impl=1 vtbl=120
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x40030000 invkind=2 oVft=56 params=0 opt=0 flags=0x0 ret=3 F1
  func memid=0x40030000 invkind=4 oVft=64 params=1 opt=0 flags=0x0 ret=24 F1
  func memid=0x40030001 invkind=2 oVft=72 params=0 opt=0 flags=0x0 ret=12 F2
  func memid=0x40030001 invkind=4 oVft=80 params=1 opt=0 flags=0x0 ret=24 F2
  func memid=0x40030001 invkind=8 oVft=88 params=1 opt=0 flags=0x0 ret=24 F2
  func memid=0x68030001 invkind=2 oVft=96 params=0 opt=0 flags=0x0 ret=3 Zed
  func memid=0x68030000 invkind=2 oVft=104 params=0 opt=0 flags=0x0 ret=3 Alpha
  func memid=0x60030002 invkind=1 oVft=112 params=0 opt=0 flags=0x0 ret=24 M
--dual
type W7 kind=TKIND_INTERFACE flags=0x1C2 funcs=8 vars=0 impl=1 vtbl=120
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x40030000 invkind=2 oVft=56 params=1 opt=0 flags=0x0 ret=25 F1
  func memid=0x40030000 invkind=4 oVft=64 params=1 opt=0 flags=0x0 ret=25 F1
  func memid=0x40030001 invkind=2 oVft=72 params=1 opt=0 flags=0x0 ret=25 F2
  func memid=0x40030001 invkind=4 oVft=80 params=1 opt=0 flags=0x0 ret=25 F2
  func memid=0x40030001 invkind=8 oVft=88 params=1 opt=0 flags=0x0 ret=25 F2
  func memid=0x68030001 invkind=2 oVft=96 params=1 opt=0 flags=0x0 ret=25 Zed
  func memid=0x68030000 invkind=2 oVft=104 params=1 opt=0 flags=0x0 ret=25 Alpha
  func memid=0x60030002 invkind=1 oVft=112 params=0 opt=0 flags=0x0 ret=25 M
";

    let expected_w8 = "\
type W8 kind=TKIND_DISPATCH flags=0x1C2 funcs=1 vars=0 impl=1 vtbl=64
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x60030000 invkind=1 oVft=56 params=0 opt=0 flags=0x0 ret=3 A
--dual
type W8 kind=TKIND_INTERFACE flags=0x1C2 funcs=1 vars=0 impl=1 vtbl=64
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x60030000 invkind=1 oVft=56 params=1 opt=0 flags=0x0 ret=25 A
";

    // Widget's 15 public functions match Excel line-for-line; vtbl has 24 slots (192 bytes)
    // pending vw1's Excel measurement of the unexplained third hidden slot in objid2's Widget.
    let expected_widget = "\
type Widget kind=TKIND_DISPATCH flags=0x1C2 funcs=15 vars=0 impl=1 vtbl=192
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x40030000 invkind=2 oVft=56 params=0 opt=0 flags=0x0 ret=3 Value
  func memid=0x40030000 invkind=4 oVft=64 params=1 opt=0 flags=0x0 ret=24 Value
  func memid=0x40030001 invkind=2 oVft=72 params=0 opt=0 flags=0x0 ret=8 Name
  func memid=0x40030001 invkind=4 oVft=80 params=1 opt=0 flags=0x0 ret=24 Name
  func memid=0x40030002 invkind=2 oVft=88 params=0 opt=0 flags=0x0 ret=9 Child
  func memid=0x40030002 invkind=4 oVft=96 params=1 opt=0 flags=0x1 ret=24 Child
  func memid=0x40030002 invkind=8 oVft=104 params=1 opt=0 flags=0x0 ret=24 Child
  func memid=0x60030002 invkind=1 oVft=112 params=0 opt=0 flags=0x0 ret=3 Twice
  func memid=0x60030003 invkind=1 oVft=120 params=0 opt=0 flags=0x0 ret=24 Reset
  func memid=0x68030001 invkind=2 oVft=128 params=0 opt=0 flags=0x0 ret=5 X
  func memid=0x68030001 invkind=4 oVft=136 params=1 opt=0 flags=0x0 ret=24 X
  func memid=0x68030000 invkind=2 oVft=144 params=0 opt=0 flags=0x0 ret=9 Obj
  func memid=0x68030000 invkind=8 oVft=152 params=1 opt=0 flags=0x0 ret=24 Obj
  func memid=0x60030004 invkind=1 oVft=160 params=3 opt=0 flags=0x0 ret=3 WithArgs,a,b,c
  func memid=0x60030005 invkind=1 oVft=168 params=1 opt=-1 flags=0x0 ret=3 Rest,items
--dual
type Widget kind=TKIND_INTERFACE flags=0x1C2 funcs=15 vars=0 impl=1 vtbl=192
  !impl 0 TYPE_E_ELEMENTNOTFOUND
  func memid=0x40030000 invkind=2 oVft=56 params=1 opt=0 flags=0x0 ret=25 Value
  func memid=0x40030000 invkind=4 oVft=64 params=1 opt=0 flags=0x0 ret=25 Value
  func memid=0x40030001 invkind=2 oVft=72 params=1 opt=0 flags=0x0 ret=25 Name
  func memid=0x40030001 invkind=4 oVft=80 params=1 opt=0 flags=0x0 ret=25 Name
  func memid=0x40030002 invkind=2 oVft=88 params=1 opt=0 flags=0x0 ret=25 Child
  func memid=0x40030002 invkind=4 oVft=96 params=1 opt=0 flags=0x1 ret=25 Child
  func memid=0x40030002 invkind=8 oVft=104 params=1 opt=0 flags=0x0 ret=25 Child
  func memid=0x60030002 invkind=1 oVft=112 params=1 opt=0 flags=0x0 ret=25 Twice
  func memid=0x60030003 invkind=1 oVft=120 params=0 opt=0 flags=0x0 ret=25 Reset
  func memid=0x68030001 invkind=2 oVft=128 params=1 opt=0 flags=0x0 ret=25 X
  func memid=0x68030001 invkind=4 oVft=136 params=1 opt=0 flags=0x0 ret=25 X
  func memid=0x68030000 invkind=2 oVft=144 params=1 opt=0 flags=0x0 ret=25 Obj
  func memid=0x68030000 invkind=8 oVft=152 params=1 opt=0 flags=0x0 ret=25 Obj
  func memid=0x60030004 invkind=1 oVft=160 params=4 opt=0 flags=0x0 ret=25 WithArgs,a,b,c
  func memid=0x60030005 invkind=1 oVft=168 params=2 opt=-1 flags=0x0 ret=25 Rest,items
";

    let expectations = [
        ("Widget", expected_widget),
        ("W2", expected_w2),
        ("W3", expected_w3),
        ("W4", expected_w4),
        ("W5", expected_w5),
        ("W6", expected_w6),
        ("W7", expected_w7),
        ("W8", expected_w8),
    ];

    for backend in BACKENDS {
        let mut engine = Engine::new(HostConfig { backend });
        engine.set_host_policy(HostPolicy::native());
        let mut session = engine
            .prepare_image_session(OxImage::new(vec![oxp.clone()]))
            .unwrap_or_else(|err| panic!("{backend:?} prepare_image_session: {err}"));

        for (class_name, expected) in &expectations {
            let instance = session
                .create_class_instance(class_name)
                .unwrap_or_else(|err| panic!("{backend:?} create {class_name}: {err}"));
            let actual = describe_dispatch_portable(instance.raw_iunknown());
            assert_eq!(
                actual, *expected,
                "{backend:?} GetTypeInfo mismatch for {class_name}"
            );
        }

        // Also verify ITypeInfo::GetIDsOfNames and ITypeInfo::Invoke on Widget.
        let widget = session
            .create_class_instance("Widget")
            .expect("create Widget");
        let raw = widget.raw_iunknown();
        // SAFETY: `session` is exclusively borrowed by the invoker scope for these calls.
        let _invoker = unsafe { session.enter_member_invoker_scope() };
        // SAFETY: `raw` is a live project object `IDispatch` pointer and `ti_ptr` is
        // released before `_invoker` and `widget` drop.
        unsafe {
            use calcvba_runtime::com::ITypeInfo;
            let disp_vtbl = &*((*raw).vtbl.cast::<IDispatchVtbl>());
            let mut ti_ptr: *mut c_void = core::ptr::null_mut();
            assert_eq!((disp_vtbl.get_type_info)(raw.cast(), 0, 0, &mut ti_ptr), 0);
            let ti_vtbl = &*(*ti_ptr.cast::<ITypeInfo>()).vtbl;

            let n_value = utf16_nul("Value");
            let n_twice = utf16_nul("Twice");
            let mut value_id = DISPID_UNKNOWN;
            let mut twice_id = DISPID_UNKNOWN;
            assert_eq!(
                (ti_vtbl.get_ids_of_names)(ti_ptr, [n_value.as_ptr()].as_ptr(), 1, &mut value_id),
                0
            );
            assert_eq!(
                (ti_vtbl.get_ids_of_names)(ti_ptr, [n_twice.as_ptr()].as_ptr(), 1, &mut twice_id),
                0
            );
            assert_eq!(value_id, 0x4003_0000);
            assert_eq!(twice_id, 0x6003_0002);

            // Put Value = 21 via ITypeInfo::Invoke, then call Twice() -> 42.
            let mut put_val = Variant::from_i32(21);
            let mut put_named = [DISPID_PROPERTYPUT];
            let mut put_params = DispParams {
                rgvarg: (&raw mut put_val).cast::<VariantCore>(),
                rgdispid_named_args: put_named.as_mut_ptr(),
                c_args: 1,
                c_named_args: 1,
            };
            assert_eq!(
                (ti_vtbl.invoke)(
                    ti_ptr,
                    raw.cast(),
                    value_id,
                    DISPATCH_PROPERTYPUT,
                    &raw mut put_params,
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                ),
                0
            );

            let mut empty_params = DispParams::default();
            let mut result = Variant::empty();
            assert_eq!(
                (ti_vtbl.invoke)(
                    ti_ptr,
                    raw.cast(),
                    twice_id,
                    DISPATCH_METHOD,
                    &raw mut empty_params,
                    (&raw mut result).cast(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                ),
                0
            );
            assert_eq!(
                result.as_i32(),
                Some(42),
                "{backend:?} ITypeInfo::Invoke Twice"
            );
            assert_eq!((ti_vtbl.unknown.release)(ti_ptr), 0);
        }
    }
}

/// dc-yqq0.27: End-to-end VM3 + JIT portable proof of:
/// 1. `Implements IShape` as an extra COM vtable on `CompatObjectBase` sharing one COM identity
///    (`QI(IUnknown) == &primary`, `Is` holds when the secondary interface pointer crosses back
///    into VBA, and `Invoke` on the secondary face dispatches `IShape_*` with `Me = &primary`).
/// 2. Connection points (`IConnectionPointContainer` / `IConnectionPoint` / `Advise` / `Unadvise`)
///    receiving `RaiseEvent` from both VM3 and JIT, including a re-entrant callback into
///    `IShape.Area()` during the event and filtering out `Private Event`.
/// 3. `ISupportErrorInfo` and `IErrorInfo` populated on VBA `Err.Raise` for both late-bound and
///    early-bound calls on both the primary default interface and the implemented interface.
#[test]
fn implements_connection_points_and_error_info_end_to_end_on_vm3_and_jit() {
    use calcvba_runtime::com::{
        DISP_E_EXCEPTION, DISPATCH_METHOD, DispParams, E_NOINTERFACE, E_NOTIMPL, ExcepInfo, Guid,
        HResult, IConnectionPoint, IConnectionPointContainer, IDispatchVtbl, IErrorInfo,
        IID_ICONNECTIONPOINTCONTAINER, IID_IDISPATCH, IID_ISUPPORTERRORINFO, IID_IUNKNOWN,
        IID_NULL, ISupportErrorInfo, IUnknownVtbl, S_FALSE, S_OK,
    };
    use calcvba_runtime::{
        ObjectRefExt, default_class_interface_iid, default_class_source_interface_iid,
        invoke_early_bound_slot, oleaut,
    };
    use core::sync::atomic::{AtomicU32, Ordering};
    use std::cell::RefCell;
    use std::rc::Rc;

    let ishape_src = "\
Public Function Area() As Long
End Function
Public Sub Boom()
End Sub
";
    let square_src = "\
Implements IShape
Public Side As Long
Public Event Resized(ByVal OldSide As Long, ByVal NewSide As Long)
Private Event InternalPing()
Public Sub Resize(ByVal NextSide As Long)
    Dim prev As Long
    prev = Side
    Side = NextSide
    RaiseEvent Resized(prev, NextSide)
    RaiseEvent InternalPing
End Sub
Public Sub PrimaryBoom()
    Err.Raise 1051, \"Square.Primary\", \"primary boom\"
End Sub
Public Function IsSameAs(ByVal other As Object) As Boolean
    IsSameAs = (Me Is other)
End Function
Private Function IShape_Area() As Long
    IShape_Area = Side * Side
End Function
Private Sub IShape_Boom()
    Err.Raise 1050, \"Square.IShape\", \"shape boom\"
End Sub
";

    let mut manifest = SymbolProjectManifest {
        project_name: "Dc27Proj".to_string(),
        project_kind: ProjectKind::Library,
        modules: vec![
            ModuleUnit {
                module_name: "IShape".to_string(),
                module_kind: ModuleKind::Class,
                attributes: ModuleAttributes::named("IShape"),
                source: ishape_src.to_string(),
            },
            ModuleUnit {
                module_name: "Square".to_string(),
                module_kind: ModuleKind::Class,
                attributes: ModuleAttributes::named("Square"),
                source: square_src.to_string(),
            },
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    };
    manifest.modules[0].attributes.vb_exposed = true;
    manifest.modules[1].attributes.vb_exposed = true;

    let typelibs = calcvba_symbol::CatalogTypeLibResolver;
    let core = calcvba_bind::bind_program(&manifest, &typelibs).expect("bind");
    let oxp = calcvba_oxir::elaborate::elaborate(&core).expect("elaborate");

    let shape_iid = default_class_interface_iid("Dc27Proj", "IShape");
    let square_iid = default_class_interface_iid("Dc27Proj", "Square");
    let source_iid = default_class_source_interface_iid("Dc27Proj", "Square");

    type ReentrantLog = Rc<RefCell<Vec<(i32, Vec<i32>, i32)>>>;

    #[repr(C)]
    struct ReentrantSink {
        vtbl: *const IDispatchVtbl,
        ref_count: AtomicU32,
        shape_face: *mut c_void,
        area_dispid: i32,
        log: ReentrantLog,
    }
    unsafe extern "system" fn sink_qi(
        this: *mut c_void,
        riid: *const Guid,
        ppv: *mut *mut c_void,
    ) -> HResult {
        // SAFETY: caller supplies writable `ppv` and readable `riid`.
        unsafe {
            *ppv = core::ptr::null_mut();
            let iid = *riid;
            if iid == IID_IUNKNOWN || iid == IID_IDISPATCH {
                sink_add_ref(this);
                *ppv = this;
                S_OK
            } else {
                E_NOINTERFACE
            }
        }
    }
    unsafe extern "system" fn sink_add_ref(this: *mut c_void) -> u32 {
        // SAFETY: live `ReentrantSink`.
        unsafe {
            (*this.cast::<ReentrantSink>())
                .ref_count
                .fetch_add(1, Ordering::AcqRel)
                + 1
        }
    }
    unsafe extern "system" fn sink_release(this: *mut c_void) -> u32 {
        // SAFETY: live `ReentrantSink`.
        unsafe {
            let prev = (*this.cast::<ReentrantSink>())
                .ref_count
                .fetch_sub(1, Ordering::AcqRel);
            let rem = prev - 1;
            if rem == 0 {
                drop(Box::from_raw(this.cast::<ReentrantSink>()));
            }
            rem
        }
    }
    unsafe extern "system" fn sink_tic(_: *mut c_void, c: *mut u32) -> HResult {
        // SAFETY: writable out-slot.
        unsafe {
            *c = 0;
        }
        S_OK
    }
    unsafe extern "system" fn sink_ti(
        _: *mut c_void,
        _: u32,
        _: u32,
        out: *mut *mut c_void,
    ) -> HResult {
        // SAFETY: writable out-slot.
        unsafe {
            *out = core::ptr::null_mut();
        }
        E_NOTIMPL
    }
    unsafe extern "system" fn sink_ids(
        _: *mut c_void,
        _: *const Guid,
        _: *const *const u16,
        _: u32,
        _: u32,
        _: *mut i32,
    ) -> HResult {
        E_NOTIMPL
    }
    unsafe extern "system" fn sink_invoke(
        this: *mut c_void,
        dispid: i32,
        _: *const Guid,
        _: u32,
        _: u16,
        params: *mut c_void,
        _: *mut c_void,
        _: *mut c_void,
        _: *mut u32,
    ) -> HResult {
        // SAFETY: `this` is a live `ReentrantSink` and `shape_face` is a live `IShape` face.
        unsafe {
            let sink = &*this.cast::<ReentrantSink>();
            let mut logical = Vec::new();
            if !params.is_null() {
                let dp = &*params.cast::<DispParams>();
                for i in (0..dp.c_args as usize).rev() {
                    let cell = &*dp.rgvarg.add(i).cast::<Variant>();
                    logical.push(cell.as_i32().unwrap_or(-1));
                }
            }
            // Re-entrantly invoke `IShape.Area()` back into the active VM3/JIT run!
            let shape_vtbl = &*(*sink
                .shape_face
                .cast::<calcvba_runtime::RawRuntimeIUnknown>())
            .vtbl
            .cast::<IDispatchVtbl>();
            let mut dp_empty = DispParams::default();
            let mut area_val = Variant::empty();
            let hr = (shape_vtbl.invoke)(
                sink.shape_face,
                sink.area_dispid,
                &IID_NULL,
                0,
                DISPATCH_METHOD,
                (&raw mut dp_empty).cast(),
                (&raw mut area_val).cast(),
                core::ptr::null_mut(),
                core::ptr::null_mut(),
            );
            assert_eq!(hr, S_OK, "re-entrant IShape.Area during RaiseEvent");
            sink.log
                .borrow_mut()
                .push((dispid, logical, area_val.as_i32().unwrap_or(-1)));
        }
        S_OK
    }
    static REENTRANT_SINK_VTBL: IDispatchVtbl = IDispatchVtbl {
        unknown: IUnknownVtbl {
            query_interface: sink_qi,
            add_ref: sink_add_ref,
            release: sink_release,
        },
        get_type_info_count: sink_tic,
        get_type_info: sink_ti,
        get_ids_of_names: sink_ids,
        invoke: sink_invoke,
    };

    let read_thread_error_info = || -> (Guid, String, String) {
        let mut err_ptr: *mut c_void = core::ptr::null_mut();
        // SAFETY: `oleaut::get_error_info` retrieves and clears the thread error object.
        unsafe {
            assert_eq!(oleaut::get_error_info(0, &mut err_ptr), S_OK);
            assert!(!err_ptr.is_null());
            let vtbl = &*(*err_ptr.cast::<IErrorInfo>()).vtbl;
            let mut g = Guid::default();
            let mut s = core::ptr::null_mut();
            let mut d = core::ptr::null_mut();
            assert_eq!((vtbl.get_guid)(err_ptr, &mut g), S_OK);
            assert_eq!((vtbl.get_source)(err_ptr, &mut s), S_OK);
            assert_eq!((vtbl.get_description)(err_ptr, &mut d), S_OK);
            let src = calcvba_runtime::bstr::BStr::from_raw_bstr(s)
                .as_str()
                .to_string();
            let desc = calcvba_runtime::bstr::BStr::from_raw_bstr(d)
                .as_str()
                .to_string();
            (vtbl.unknown.release)(err_ptr);
            (g, src, desc)
        }
    };

    for backend in BACKENDS {
        let mut engine = Engine::new(HostConfig { backend });
        engine.set_host_policy(HostPolicy::native());
        let mut session = engine
            .prepare_image_session(OxImage::new(vec![oxp.clone()]))
            .unwrap_or_else(|err| panic!("{backend:?} prepare_image_session: {err}"));

        let square = session
            .create_class_instance("Square")
            .unwrap_or_else(|err| panic!("{backend:?} create Square: {err}"));
        let primary = square.raw_iunknown();

        // SAFETY: `primary` is a live project object; secondary pointers are released before drop.
        let shape_ptr = unsafe {
            let vtbl = &*(*primary).vtbl;
            let mut out = core::ptr::null_mut();
            assert_eq!(
                (vtbl.query_interface)(primary.cast(), &shape_iid, &mut out),
                S_OK,
                "{backend:?} QI(IShape)"
            );
            assert!(!out.is_null());
            assert_ne!(out, primary.cast());
            out
        };

        // Pass `shape_ptr` back into VBA and check `Me Is other`!
        // SAFETY: `shape_ptr` is a live `IShape` face on `square`.
        let shape_obj = unsafe { ObjectRef::from_raw_iunknown_addref(shape_ptr.cast()) }
            .expect("adopt IShape face");
        let same = session
            .invoke_member_values(
                square.clone(),
                "IsSameAs",
                None,
                vec![Variant::from_object_ref(shape_obj)],
            )
            .unwrap_or_else(|err| panic!("{backend:?} IsSameAs: {err}"));
        assert_eq!(
            same.as_bool(),
            Some(true),
            "{backend:?} `Is` must treat Implements face as same COM identity"
        );

        // Set Side = 3, then test Implements call, connection point + re-entry, and error info.
        square.project_field_set(0, Variant::from_i32(3));
        // SAFETY: `session` is exclusively borrowed by `_invoker` for the COM calls below.
        let _invoker = unsafe { session.enter_member_invoker_scope() };
        // SAFETY: all COM interface pointers belong to live `square` and are released before scope end.
        unsafe {
            let primary_vtbl = &*((*primary).vtbl.cast::<IDispatchVtbl>());
            let shape_vtbl = &*(*shape_ptr.cast::<calcvba_runtime::RawRuntimeIUnknown>())
                .vtbl
                .cast::<IDispatchVtbl>();

            // 1. IShape.Area() -> 9 via IDispatch on secondary face.
            let area_name = utf16_nul("Area");
            let mut area_id = DISPID_UNKNOWN;
            assert_eq!(
                (shape_vtbl.get_ids_of_names)(
                    shape_ptr,
                    &IID_NULL,
                    [area_name.as_ptr()].as_ptr(),
                    1,
                    0,
                    &mut area_id
                ),
                S_OK
            );
            let mut empty_dp = DispParams::default();
            let mut area_res = Variant::empty();
            assert_eq!(
                (shape_vtbl.invoke)(
                    shape_ptr,
                    area_id,
                    &IID_NULL,
                    0,
                    DISPATCH_METHOD,
                    (&raw mut empty_dp).cast(),
                    (&raw mut area_res).cast(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                ),
                S_OK
            );
            assert_eq!(area_res.as_i32(), Some(9), "{backend:?} IShape.Area");

            // 2. Connection point Advise + RaiseEvent("Resized") with re-entrant IShape.Area() -> 49.
            let mut cpc_ptr = core::ptr::null_mut();
            assert_eq!(
                (primary_vtbl.unknown.query_interface)(
                    primary.cast(),
                    &IID_ICONNECTIONPOINTCONTAINER,
                    &mut cpc_ptr
                ),
                S_OK,
                "{backend:?} QI(IConnectionPointContainer)"
            );
            let cpc_vtbl = &*(*cpc_ptr.cast::<IConnectionPointContainer>()).vtbl;
            let mut cp_ptr = core::ptr::null_mut();
            assert_eq!(
                (cpc_vtbl.find_connection_point)(cpc_ptr, &source_iid, &mut cp_ptr),
                S_OK,
                "{backend:?} FindConnectionPoint"
            );
            let cp_vtbl = &*(*cp_ptr.cast::<IConnectionPoint>()).vtbl;

            let log = Rc::new(RefCell::new(Vec::new()));
            let sink = Box::into_raw(Box::new(ReentrantSink {
                vtbl: &REENTRANT_SINK_VTBL,
                ref_count: AtomicU32::new(1),
                shape_face: shape_ptr,
                area_dispid: area_id,
                log: log.clone(),
            }))
            .cast::<c_void>();
            let mut cookie = 0u32;
            assert_eq!((cp_vtbl.advise)(cp_ptr, sink, &mut cookie), S_OK);
            assert_ne!(cookie, 0);

            let resize_name = utf16_nul("Resize");
            let mut resize_id = DISPID_UNKNOWN;
            assert_eq!(
                (primary_vtbl.get_ids_of_names)(
                    primary.cast(),
                    &IID_NULL,
                    [resize_name.as_ptr()].as_ptr(),
                    1,
                    0,
                    &mut resize_id,
                ),
                S_OK
            );
            let mut next_side = Variant::from_i32(7);
            let mut resize_dp = DispParams {
                rgvarg: (&raw mut next_side).cast::<VariantCore>(),
                rgdispid_named_args: core::ptr::null_mut(),
                c_args: 1,
                c_named_args: 0,
            };
            assert_eq!(
                (primary_vtbl.invoke)(
                    primary.cast(),
                    resize_id,
                    &IID_NULL,
                    0,
                    DISPATCH_METHOD,
                    (&raw mut resize_dp).cast(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                    core::ptr::null_mut(),
                ),
                S_OK,
                "{backend:?} Resize(7) with advised connection-point sink"
            );
            // Public event Resized(3, 7) fired with re-entrant Area() == 49; Private InternalPing did not fire.
            assert_eq!(
                *log.borrow(),
                vec![(0, vec![3, 7], 49)],
                "{backend:?} connection-point event payload + re-entrant IShape.Area"
            );
            assert_eq!((cp_vtbl.unadvise)(cp_ptr, cookie), S_OK);
            sink_release(sink);
            (cp_vtbl.unknown.release)(cp_ptr);
            (cpc_vtbl.unknown.release)(cpc_ptr);

            // 3. ISupportErrorInfo + IErrorInfo on both PrimaryBoom and IShape.Boom (late & early).
            let mut sei_ptr = core::ptr::null_mut();
            assert_eq!(
                (primary_vtbl.unknown.query_interface)(
                    primary.cast(),
                    &IID_ISUPPORTERRORINFO,
                    &mut sei_ptr
                ),
                S_OK
            );
            let sei_vtbl = &*(*sei_ptr.cast::<ISupportErrorInfo>()).vtbl;
            assert_eq!(
                (sei_vtbl.interface_supports_error_info)(sei_ptr, &square_iid),
                S_OK
            );
            assert_eq!(
                (sei_vtbl.interface_supports_error_info)(sei_ptr, &shape_iid),
                S_OK
            );
            assert_eq!(
                (sei_vtbl.interface_supports_error_info)(sei_ptr, &source_iid),
                S_FALSE
            );
            (sei_vtbl.unknown.release)(sei_ptr);

            // Late-bound IShape.Boom -> DISP_E_EXCEPTION + IErrorInfo(shape_iid).
            let boom_name = utf16_nul("Boom");
            let mut boom_id = DISPID_UNKNOWN;
            assert_eq!(
                (shape_vtbl.get_ids_of_names)(
                    shape_ptr,
                    &IID_NULL,
                    [boom_name.as_ptr()].as_ptr(),
                    1,
                    0,
                    &mut boom_id,
                ),
                S_OK
            );
            let mut excep = ExcepInfo::default();
            assert_eq!(
                (shape_vtbl.invoke)(
                    shape_ptr,
                    boom_id,
                    &IID_NULL,
                    0,
                    DISPATCH_METHOD,
                    (&raw mut empty_dp).cast(),
                    core::ptr::null_mut(),
                    (&raw mut excep).cast(),
                    core::ptr::null_mut(),
                ),
                DISP_E_EXCEPTION,
                "{backend:?} late-bound IShape.Boom"
            );
            assert_eq!(excep.scode, (0x800A_0000u32 | 1050u32) as i32);
            oleaut::sys_free_string(excep.bstr_source);
            oleaut::sys_free_string(excep.bstr_description);
            oleaut::sys_free_string(excep.bstr_help_file);
            assert_eq!(
                read_thread_error_info(),
                (
                    shape_iid,
                    "Square.IShape".to_string(),
                    "shape boom".to_string()
                ),
                "{backend:?} late-bound IShape.Boom IErrorInfo"
            );

            // Early-bound PrimaryBoom -> 0x800A041B + IErrorInfo(square_iid).
            let primary_boom_slot = square
                .class_descriptor()
                .dispatch_interface()
                .unwrap()
                .members
                .iter()
                .find(|m| m.name == "PrimaryBoom")
                .unwrap()
                .vtable_slot
                .unwrap();
            assert_eq!(
                invoke_early_bound_slot(&square, None, primary_boom_slot, &mut []),
                Err((0x800A_0000u32 | 1051u32) as i32),
                "{backend:?} early-bound PrimaryBoom HRESULT"
            );
            assert_eq!(
                read_thread_error_info(),
                (
                    square_iid,
                    "Square.Primary".to_string(),
                    "primary boom".to_string()
                ),
                "{backend:?} early-bound PrimaryBoom IErrorInfo"
            );

            // Early-bound IShape.Boom -> 0x800A041A + IErrorInfo(shape_iid).
            let shape_boom_slot = square
                .class_descriptor()
                .interfaces
                .iter()
                .find(|i| i.identity.guid == shape_iid)
                .unwrap()
                .members
                .iter()
                .find(|m| m.name == "Boom")
                .unwrap()
                .vtable_slot
                .unwrap();
            assert_eq!(
                invoke_early_bound_slot(&square, Some(shape_iid), shape_boom_slot, &mut []),
                Err((0x800A_0000u32 | 1050u32) as i32),
                "{backend:?} early-bound IShape.Boom HRESULT"
            );
            assert_eq!(
                read_thread_error_info(),
                (
                    shape_iid,
                    "Square.IShape".to_string(),
                    "shape boom".to_string()
                ),
                "{backend:?} early-bound IShape.Boom IErrorInfo"
            );

            (shape_vtbl.unknown.release)(shape_ptr);
        }
    }
}
