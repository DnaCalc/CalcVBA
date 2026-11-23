//! dc-yqq0.20: a project object is one COM object (design note
//! docs/design/vba-object-is-one-com-object-2026-11-20.md). Each case runs on VM3 and the JIT
//! and expects the answer Excel's VBA gave on Windows (2026-11-21, `temp/objid2` probe):
//!
//! `coll-is=True;coll-call=42;dict-is=True;dict-call=42;coll-term-held=0;coll-term-removed=1;
//! foreign-call=42;term-held=0;term-dropped=1`
//!
//! The early-bound foreign call (dc-yqq0.24, `temp/objid4` probe) calls vtable slots directly:
//! `early-get=21;early-call=42;after-put=5;early-call2=10`. The final release on a foreign thread
//! (dc-yqq0.25) has no Excel answer (a raw Release of a VBA object off its thread is undefined
//! there); CalcVBA runs `Class_Terminate` on the owner thread at the next statement.
//!
//! The live-COM cases need OxVba.TestEventServer (tools/OxVba.TestEventServer/register.ps1).

use calcvba_engine::{Engine, ExecBackend, HostConfig};
use calcvba_hal::model::HostPolicy;
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};

const BACKENDS: [ExecBackend; 2] = [ExecBackend::Vm3, ExecBackend::Jit];

const WIDGET: &str = r#"
Public Value As Long
Public Function Twice() As Long
    Twice = Value * 2
End Function
Private Sub Class_Terminate()
    gTerm = gTerm + 1
End Sub
"#;

fn module(name: &str, kind: ModuleKind, source: &str) -> ModuleUnit {
    ModuleUnit {
        module_name: name.to_string(),
        module_kind: kind,
        attributes: ModuleAttributes::named(name),
        source: source.to_string(),
    }
}

/// Runs `body` as `Sub Main` (with `w` a `Widget` whose Value is 21) and returns `gOut`.
fn run(backend: ExecBackend, body: &str) -> String {
    let main = format!(
        "Public gTerm As Long\nPublic gOut As String\nSub Main()\n    Dim w As New Widget\n    w.Value = 21\n{body}\nEnd Sub\n"
    );
    let manifest = SymbolProjectManifest {
        project_name: "ObjId".to_string(),
        project_kind: ProjectKind::Source,
        modules: vec![
            module("Probe", ModuleKind::Procedural, &main),
            module("Widget", ModuleKind::Class, WIDGET),
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
        codepage: None,
    };
    let mut engine = Engine::new(HostConfig { backend });
    engine.set_host_policy(HostPolicy::native());
    let values = engine
        .execute_manifest_with_variant_snapshot(&manifest)
        .unwrap_or_else(|err| panic!("{backend:?}: {err}"));
    values
        .iter()
        .filter_map(|value| value.as_bstr().map(|text| text.as_str().to_string()))
        .find(|text| text.contains('='))
        .unwrap_or_else(|| panic!("{backend:?}: no gOut in {values:?}"))
}

#[test]
fn collection_round_trip_and_terminate_match_excel() {
    let body = r#"
    Dim c As Collection, h As Widget
    Set c = New Collection
    c.Add w, "k"
    gOut = "coll-is=" & CStr(c("k") Is w) & ";coll-call=" & CStr(c("k").Twice())
    Set h = New Widget
    gTerm = 0
    c.Add h, "h"
    Set h = Nothing
    gOut = gOut & ";coll-term-held=" & CStr(gTerm)
    c.Remove "h"
    gOut = gOut & ";coll-term-removed=" & CStr(gTerm)"#;
    for backend in BACKENDS {
        assert_eq!(
            run(backend, body),
            "coll-is=True;coll-call=42;coll-term-held=0;coll-term-removed=1",
            "{backend:?}"
        );
    }
}

#[cfg(target_os = "windows")]
#[test]
fn dictionary_round_trip_matches_excel() {
    // Scripting.Dictionary is a foreign COM object: the instance crosses out and back.
    let body = r#"
    Dim d As Object, back As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "k", w
    Set back = d("k")
    gOut = "dict-is=" & CStr(back Is w) & ";dict-call=" & CStr(back.Twice())"#;
    for backend in BACKENDS {
        assert_eq!(
            run(backend, body),
            "dict-is=True;dict-call=42",
            "{backend:?}"
        );
    }
}

#[cfg(target_os = "windows")]
#[test]
#[ignore = "live COM: needs OxVba.TestEventServer (tools/OxVba.TestEventServer/register.ps1)"]
fn foreign_holder_keeps_the_object_alive_as_excel_does() {
    let body = r#"
    Dim s As Object, h As Widget
    Set s = CreateObject("OxVba.TestEventServer")
    Set h = New Widget
    gTerm = 0
    s.Hold h
    Set h = Nothing
    gOut = "term-held=" & CStr(gTerm)
    s.Drop
    gOut = gOut & ";term-dropped=" & CStr(gTerm)"#;
    for backend in BACKENDS {
        assert_eq!(
            run(backend, body),
            "term-held=0;term-dropped=1",
            "{backend:?}"
        );
    }
}

#[cfg(target_os = "windows")]
#[test]
#[ignore = "live COM: needs OxVba.TestEventServer (tools/OxVba.TestEventServer/register.ps1)"]
fn foreign_late_bound_call_matches_excel() {
    // .NET `InvokeMember` calls the object through its own IDispatch.
    let body = r#"
    Dim s As Object
    Set s = CreateObject("OxVba.TestEventServer")
    On Error Resume Next
    gOut = "foreign-call=" & CStr(s.CallMember(w, "Twice"))
    If Err.Number <> 0 Then gOut = "foreign-call=ERR" & Err.Number"#;
    for backend in BACKENDS {
        assert_eq!(run(backend, body), "foreign-call=42", "{backend:?}");
    }
}

#[cfg(target_os = "windows")]
#[test]
#[ignore = "live COM: needs OxVba.TestEventServer (tools/OxVba.TestEventServer/register.ps1)"]
fn foreign_early_bound_slot_calls_match_excel() {
    // Widget's dual vtable, as in Excel: Value get = slot 7, Value let = 8, Twice = 9.
    let body = r#"
    Dim s As Object
    Set s = CreateObject("OxVba.TestEventServer")
    On Error Resume Next
    gOut = "early-get=" & CStr(s.CallSlotGetLong(w, 7))
    If Err.Number <> 0 Then gOut = gOut & "ERR" & Err.Number: Err.Clear
    gOut = gOut & ";early-call=" & CStr(s.CallSlotGetLong(w, 9))
    If Err.Number <> 0 Then gOut = gOut & "ERR" & Err.Number: Err.Clear
    s.CallSlotPutLong w, 8, 5
    If Err.Number <> 0 Then gOut = gOut & ";early-put=ERR" & Err.Number: Err.Clear
    gOut = gOut & ";after-put=" & CStr(w.Value) & ";early-call2=" & CStr(s.CallSlotGetLong(w, 9))
    If Err.Number <> 0 Then gOut = gOut & "ERR" & Err.Number: Err.Clear"#;
    for backend in BACKENDS {
        assert_eq!(
            run(backend, body),
            "early-get=21;early-call=42;after-put=5;early-call2=10",
            "{backend:?}"
        );
    }
}

#[cfg(target_os = "windows")]
#[test]
#[ignore = "live COM: needs OxVba.TestEventServer (tools/OxVba.TestEventServer/register.ps1)"]
fn foreign_thread_final_release_terminates_on_the_owner_thread() {
    let body = r#"
    Dim s As Object, h As Widget
    Set s = CreateObject("OxVba.TestEventServer")
    Set h = New Widget
    gTerm = 0
    s.Hold h
    Set h = Nothing
    gOut = "term-held=" & CStr(gTerm)
    s.DropOnOtherThread
    gOut = gOut & ";term-dropped=" & CStr(gTerm)"#;
    for backend in BACKENDS {
        assert_eq!(
            run(backend, body),
            "term-held=0;term-dropped=1",
            "{backend:?}"
        );
    }
}

#[cfg(target_os = "windows")]
#[test]
#[ignore = "live COM: needs OxVba.TestEventServer (tools/OxVba.TestEventServer/register.ps1)"]
fn foreign_get_type_info_describe_dispatch_matches_excel() {
    // dc-yqq0.29: OxVba.TestEventServer.DescribeDispatch inspects GetTypeInfo(0) and its dual
    // TKIND_INTERFACE half via COM interop. This extends Excel's W7 fixture with the
    // leading private field and nonempty lifecycle hook measured in the follow-up probes.
    let w7 = r#"
Private Hidden As Long
Public F1 As Long
Public F2 As Variant
Public Property Get Zed() As Long
End Property
Public Property Get Alpha() As Long
End Property
Public Sub M()
End Sub
Private Sub Class_Terminate()
    Hidden = 1
End Sub
"#;
    let main = r#"
Public gOut As String
Sub Main()
    Dim x As W7, s As Object
    Set x = New W7
    Set s = CreateObject("OxVba.TestEventServer")
    gOut = s.DescribeDispatch(x)
End Sub
"#;
    let manifest = SymbolProjectManifest {
        project_name: "ObjIdTypeInfo".to_string(),
        project_kind: ProjectKind::Source,
        modules: vec![
            module("Probe", ModuleKind::Procedural, main),
            module("W7", ModuleKind::Class, w7),
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
        codepage: None,
    };
    for backend in BACKENDS {
        let mut engine = Engine::new(HostConfig { backend });
        engine.set_host_policy(HostPolicy::native());
        let values = engine
            .execute_manifest_with_variant_snapshot(&manifest)
            .unwrap_or_else(|err| panic!("{backend:?}: {err}"));
        let desc = values
            .iter()
            .filter_map(|v| v.as_bstr().map(|text| text.as_str().to_string()))
            .find(|text| text.contains("TKIND_DISPATCH"))
            .unwrap_or_else(|| panic!("{backend:?}: no DescribeDispatch output in {values:?}"));
        assert!(
            desc.contains("type W7 kind=TKIND_DISPATCH flags=0x1C2"),
            "{desc}"
        );
        assert!(desc.contains("funcs=8 vars=0 impl=1 vtbl=128"), "{desc}");
        assert!(desc.contains("TYPE_E_ELEMENTNOTFOUND"), "{desc}");
        assert!(
            desc.contains(
                "func memid=0x40030001 invkind=2 oVft=56 params=0 opt=0 flags=0x0 ret=3 F1"
            ),
            "{desc}"
        );
        assert!(
            desc.contains(
                "func memid=0x40030002 invkind=8 oVft=88 params=1 opt=0 flags=0x0 ret=24 F2"
            ),
            "{desc}"
        );
        assert!(
            desc.contains(
                "func memid=0x68030001 invkind=2 oVft=96 params=0 opt=0 flags=0x0 ret=3 Zed"
            ),
            "{desc}"
        );
        assert!(
            desc.contains(
                "func memid=0x68030000 invkind=2 oVft=104 params=0 opt=0 flags=0x0 ret=3 Alpha"
            ),
            "{desc}"
        );
        assert!(
            desc.contains(
                "func memid=0x60030002 invkind=1 oVft=112 params=0 opt=0 flags=0x0 ret=24 M"
            ),
            "{desc}"
        );
        assert!(
            desc.contains("--dual\ntype W7 kind=TKIND_INTERFACE flags=0x1C2"),
            "{desc}"
        );
        assert!(
            desc.contains(
                "func memid=0x60030002 invkind=1 oVft=112 params=0 opt=0 flags=0x0 ret=25 M"
            ),
            "{desc}"
        );
    }
}

#[cfg(all(windows, target_arch = "x86_64"))]
#[test]
fn foreign_aggregate_slot_calls_match_excel() {
    use calcvba_oxir::OxImage;
    use calcvba_runtime::bstr::BStr;
    use calcvba_runtime::safe_array::{SafeArray, SafeArrayBound};
    use calcvba_runtime::{Decimal96, Variant};
    use std::ffi::c_void;

    #[repr(C)]
    struct Pair {
        number: i32,
        tag: *mut u16,
    }

    let mut widget = module(
        "SlotArrays",
        ModuleKind::Class,
        r#"
Public Sub Change(ByRef values() As Long)
    values(LBound(values)) = values(LBound(values)) + 7
End Sub
Public Function CopyOf(ByRef values() As Long) As Long()
    CopyOf = values
End Function
Public Function Count(ParamArray values() As Variant) As Long
    Count = UBound(values) + 1
End Function
Public Function DecimalValue(ByVal value As Variant) As Variant
    DecimalValue = CDec(value)
End Function
Public Function ChangeRecord(ByRef value As SlotPair) As SlotPair
    value.Number = value.Number + 7
    value.Tag = value.Tag & "!"
    ChangeRecord = value
End Function
"#,
    );
    widget.attributes.vb_exposed = true;
    widget.attributes.vb_creatable = true;
    let manifest = SymbolProjectManifest {
        project_name: "SlotProbe".into(),
        project_kind: ProjectKind::Library,
        modules: vec![
            module(
                "Types",
                ModuleKind::Procedural,
                "Public Type SlotPair\n    Number As Long\n    Tag As String\nEnd Type",
            ),
            widget,
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
        codepage: None,
    };
    let bound =
        calcvba_bind::bind_program(&manifest, &calcvba_symbol::CatalogTypeLibResolver).unwrap();
    let program = calcvba_oxir::elaborate::elaborate(&bound).unwrap();
    for backend in BACKENDS {
        let engine = Engine::new(HostConfig { backend });
        let mut session = engine
            .prepare_image_session(OxImage::new(vec![program.clone()]))
            .unwrap();
        let object = session.create_class_instance("SlotArrays").unwrap();
        let original = SafeArray::from_typed_variants_nd(
            vec![SafeArrayBound { lower: 3, count: 2 }],
            3,
            vec![Variant::from_i32(10), Variant::from_i32(0)],
        )
        .unwrap();
        let mut values = original.into_raw_safearray_ptr();
        let before = values;
        let mut copied = std::ptr::null_mut();
        let mut params =
            SafeArray::from_variants(vec![Variant::from_i32(7), Variant::from_string("x")])
                .into_raw_safearray_ptr();
        let mut count = -1;
        let value = Variant::from_decimal96(Decimal96::from_parts(45125, 0, 0, 3, false));
        let mut decimal = Variant::empty();
        let mut record = Pair {
            number: 10,
            tag: BStr::from("x").into_raw_bstr(),
        };
        let mut record_result = Pair {
            number: 0,
            tag: std::ptr::null_mut(),
        };
        // SAFETY: the test owns this session, object and the ABI cells for the whole call.
        // GetTypeInfo in Excel gives SAFEARRAY** for the typed array and ParamArray
        // arguments, including the array retval. Win64 passes a ByVal VARIANT by address.
        let (changed, copied) = unsafe {
            let _invoker = session.enter_member_invoker_scope();
            let this = object.raw_iunknown().cast::<c_void>();
            let slots = *this.cast::<*const *const c_void>();
            let change: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> i32 =
                std::mem::transmute(*slots.add(7));
            let copy: unsafe extern "system" fn(
                *mut c_void,
                *mut *mut c_void,
                *mut *mut c_void,
            ) -> i32 = std::mem::transmute(*slots.add(8));
            let count_params: unsafe extern "system" fn(
                *mut c_void,
                *mut *mut c_void,
                *mut i32,
            ) -> i32 = std::mem::transmute(*slots.add(9));
            let decimal_value: unsafe extern "system" fn(
                *mut c_void,
                *const Variant,
                *mut Variant,
            ) -> i32 = std::mem::transmute(*slots.add(10));
            let change_record: unsafe extern "system" fn(*mut c_void, *mut Pair, *mut Pair) -> i32 =
                std::mem::transmute(*slots.add(11));
            let change_hr = change(this, &mut values);
            let copy_hr = copy(this, &mut values, &mut copied);
            let count_hr = count_params(this, &mut params, &mut count);
            let decimal_hr = decimal_value(this, &value, &mut decimal);
            let record_hr = change_record(this, &mut record, &mut record_result);
            let changed = SafeArray::from_raw_safearray_owned(values, None);
            let copied = SafeArray::from_raw_safearray_owned(copied, None);
            drop(SafeArray::from_raw_safearray_owned(params, None));
            let argument_tag = BStr::from_raw_bstr(record.tag);
            let result_tag = BStr::from_raw_bstr(record_result.tag);
            assert_eq!(
                [change_hr, copy_hr, count_hr, decimal_hr, record_hr],
                [0; 5],
                "{backend:?}"
            );
            assert_eq!((record.number, record_result.number), (17, 17));
            assert_eq!(argument_tag.as_str(), "x!");
            assert_eq!(result_tag.as_str(), "x!");
            (changed.unwrap(), copied.unwrap())
        };
        assert_eq!(changed.raw_safearray_ptr(), before);
        assert_eq!(changed.i32_element(0).unwrap(), Some(17));
        assert_eq!(copied.i32_element(0).unwrap(), Some(17));
        assert_eq!(copied.bounds(), changed.bounds());
        assert_eq!(count, 2);
        assert_eq!(decimal, value);
    }
}

/// dc-yqq0.28: real Win64 vtable slots + real `IDispatch::Invoke` + real `oleaut32`
/// `SAFEARRAY(VT_RECORD)` with nominal `IRecordInfo` on VM3 and the JIT.
#[cfg(all(windows, target_arch = "x86_64"))]
#[test]
fn vtable_slots_and_idispatch_marshal_nominal_record_safearrays_on_windows() {
    use calcvba_oxir::OxImage;
    use calcvba_runtime::automation_record::{
        AutomationVariant, automation_safearray_to_vba_record_array, create_record_info_for_layout,
        record_info_get_guid, record_info_get_name, record_info_get_size,
        record_info_is_matching_type, release_record_info,
        vba_record_array_to_automation_safearray,
    };
    use calcvba_runtime::com::{
        DISP_E_TYPEMISMATCH, DISPATCH_METHOD, DispParams, ExcepInfo, IDispatchVtbl, IID_NULL, S_OK,
    };
    use calcvba_runtime::safe_array::{SafeArray, SafeArrayBound};
    use calcvba_runtime::{
        Variant, VariantCore, VbaRecord, VbaRecordFieldKind, VbaRecordFieldSpec, VbaRecordLayout,
        default_record_type_guid, oleaut,
    };
    use std::ffi::c_void;
    use std::sync::Arc;

    let mut host_class = module(
        "PairHost",
        ModuleKind::Class,
        r#"
Public Sub MutateInPlace(ByRef items() As SlotPair)
    Dim i As Long
    For i = LBound(items) To UBound(items)
        items(i).Number = items(i).Number + 10
        items(i).Tag = items(i).Tag & "!"
    Next i
End Sub
Public Sub GrowPreserve(ByRef items() As SlotPair)
    Dim lb As Long
    Dim ub As Long
    lb = LBound(items)
    ub = UBound(items) + 1
    ReDim Preserve items(lb To ub)
    items(ub).Number = 99
    items(ub).Tag = "grown"
End Sub
Public Function CopyPairs(ByRef items() As SlotPair) As SlotPair()
    CopyPairs = items
End Function
Public Function RoundTripVariant(ByRef v As Variant) As Variant
    Dim lb As Long
    Dim ub As Long
    Dim arr() As SlotPair
    lb = LBound(v)
    ub = UBound(v)
    ReDim Preserve v(lb To ub + 1)
    arr = v
    arr(lb).Number = arr(lb).Number + 1
    arr(lb).Tag = arr(lb).Tag & "*"
    arr(ub + 1).Number = 55
    arr(ub + 1).Tag = "var-added"
    v = arr
    RoundTripVariant = v
End Function
"#,
    );
    host_class.attributes.vb_exposed = true;
    host_class.attributes.vb_creatable = true;
    let manifest = SymbolProjectManifest {
        project_name: "WinRecordArrayProbe".into(),
        project_kind: ProjectKind::Library,
        modules: vec![
            module(
                "Types",
                ModuleKind::Procedural,
                "Public Type SlotPair\n    Number As Long\n    Tag As String\nEnd Type\nPublic Type OtherPair\n    Number As Long\n    Tag As String\nEnd Type",
            ),
            host_class,
        ],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
        codepage: None,
    };
    let bound =
        calcvba_bind::bind_program(&manifest, &calcvba_symbol::CatalogTypeLibResolver).unwrap();
    let program = calcvba_oxir::elaborate::elaborate(&bound).unwrap();

    let pair_guid = default_record_type_guid("WinRecordArrayProbe", "SlotPair");
    let other_guid = default_record_type_guid("WinRecordArrayProbe", "OtherPair");
    let base_layout = VbaRecordLayout::new(vec![
        VbaRecordFieldSpec::named("Number", VbaRecordFieldKind::Long),
        VbaRecordFieldSpec::named("Tag", VbaRecordFieldKind::String),
    ])
    .unwrap();
    let pair_layout = Arc::new(base_layout.clone().with_identity("SlotPair", pair_guid));
    let other_layout = Arc::new(base_layout.with_identity("OtherPair", other_guid));

    let make_pair = |layout: &Arc<VbaRecordLayout>, number: i32, tag: &str| {
        let mut rec = VbaRecord::new_default(Arc::clone(layout)).unwrap();
        rec.write_field_variant(0, &Variant::from_i32(number))
            .unwrap();
        rec.write_field_variant(1, &Variant::from_string(tag))
            .unwrap();
        rec
    };

    for backend in BACKENDS {
        let engine = Engine::new(HostConfig { backend });
        let mut session = engine
            .prepare_image_session(OxImage::new(vec![program.clone()]))
            .unwrap();
        let object = session.create_class_instance("PairHost").unwrap();
        let initial = Variant::from_safearray(
            SafeArray::from_vba_records_nd(
                vec![SafeArrayBound { count: 2, lower: 1 }],
                Arc::clone(&pair_layout),
                vec![
                    make_pair(&pair_layout, 10, "x"),
                    make_pair(&pair_layout, 20, "y"),
                ],
            )
            .unwrap(),
        );
        let wrong = Variant::from_safearray(
            SafeArray::from_vba_records_nd(
                vec![SafeArrayBound { count: 1, lower: 1 }],
                Arc::clone(&other_layout),
                vec![make_pair(&other_layout, 10, "bad")],
            )
            .unwrap(),
        );
        // SAFETY: real Win64 vtable stubs and oleaut32 SAFEARRAY(VT_RECORD) storage.
        unsafe {
            let _invoker = session.enter_member_invoker_scope();
            let this = object.raw_iunknown().cast::<c_void>();
            let slots = *this.cast::<*const *const c_void>();
            let disp_vtbl = &*(*this.cast::<*const IDispatchVtbl>());

            let mutate_in_place: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> i32 =
                std::mem::transmute(*slots.add(7));
            let grow_preserve: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> i32 =
                std::mem::transmute(*slots.add(8));
            let copy_pairs: unsafe extern "system" fn(
                *mut c_void,
                *mut *mut c_void,
                *mut *mut c_void,
            ) -> i32 = std::mem::transmute(*slots.add(9));
            let round_trip_variant: unsafe extern "system" fn(
                *mut c_void,
                *mut c_void,
                *mut c_void,
            ) -> i32 = std::mem::transmute(*slots.add(10));

            // Nominal mismatch rejected BEFORE copying data.
            let mut wrong_psa =
                vba_record_array_to_automation_safearray(&wrong, Some(&other_layout)).unwrap();
            assert_eq!(
                mutate_in_place(this, &mut wrong_psa),
                DISP_E_TYPEMISMATCH,
                "{backend:?} Win64 slot must reject OtherPair"
            );
            assert_eq!(oleaut::safe_array_destroy(wrong_psa), S_OK);

            // Real Win64 vtable slot calls with oleaut32 SAFEARRAY(VT_RECORD).
            let mut psa =
                vba_record_array_to_automation_safearray(&initial, Some(&pair_layout)).unwrap();
            let before_psa = psa;
            assert_eq!(mutate_in_place(this, &mut psa), S_OK, "{backend:?}");
            assert_eq!(psa, before_psa, "{backend:?} in-place preserves descriptor");
            assert_eq!(grow_preserve(this, &mut psa), S_OK, "{backend:?}");

            let mut copied = std::ptr::null_mut();
            assert_eq!(copy_pairs(this, &mut psa, &mut copied), S_OK, "{backend:?}");
            assert!(oleaut::safe_array_is_record_array(copied));
            let mut info = std::ptr::null_mut();
            assert_eq!(oleaut::safe_array_get_record_info(copied, &mut info), S_OK);
            assert_eq!(record_info_get_guid(info), Ok(pair_guid));
            assert_eq!(record_info_get_name(info).as_deref(), Ok("SlotPair"));
            assert_eq!(record_info_get_size(info), Ok(pair_layout.size() as u32));
            let other_info =
                create_record_info_for_layout(&other_layout, "OtherPair", other_guid).unwrap();
            assert!(!record_info_is_matching_type(info, other_info));
            release_record_info(other_info);
            release_record_info(info);

            let back =
                automation_safearray_to_vba_record_array(copied, &pair_layout, false, &|_, _| None)
                    .unwrap();
            assert_eq!(back.safearray_bounds_len().unwrap().1, 3);
            assert_eq!(
                back.safearray_element(0)
                    .unwrap()
                    .unwrap()
                    .read_record_field_variant(0)
                    .unwrap()
                    .as_i32(),
                Some(20)
            );
            assert_eq!(
                back.safearray_element(2)
                    .unwrap()
                    .unwrap()
                    .read_record_field_variant(1)
                    .unwrap()
                    .as_bstr()
                    .unwrap()
                    .as_str(),
                "grown"
            );
            assert_eq!(oleaut::safe_array_destroy(copied), S_OK);

            // Real Win64 early-bound Variant holding UDT array + copy-back.
            let mut var_in = AutomationVariant::from_runtime(&initial).unwrap();
            let mut var_out = AutomationVariant::empty();
            assert_eq!(
                round_trip_variant(this, var_in.as_mut_ptr(), var_out.as_mut_ptr()),
                S_OK,
                "{backend:?}"
            );
            assert_eq!(var_in.raw_vt(), 0x2024);
            assert_eq!(var_out.raw_vt(), 0x2024);
            let var_back = var_out.to_runtime(None, &|_, _| None).unwrap();
            assert_eq!(var_back.safearray_bounds_len().unwrap().1, 3);

            // Real Win64 late-bound IDispatch::Invoke with VT_BYREF | VT_ARRAY | VT_RECORD.
            let mut_name: Vec<u16> = "MutateInPlace\0".encode_utf16().collect();
            let mut dispid = -1i32;
            assert_eq!(
                (disp_vtbl.get_ids_of_names)(
                    this,
                    &IID_NULL,
                    [mut_name.as_ptr()].as_ptr(),
                    1,
                    0,
                    &mut dispid
                ),
                S_OK
            );
            let mut byref_arg = VariantCore::from_pointer(0x4000 | 0x2000 | 0x0024, &raw mut psa);
            let mut dp = DispParams {
                rgvarg: &raw mut byref_arg,
                rgdispid_named_args: std::ptr::null_mut(),
                c_args: 1,
                c_named_args: 0,
            };
            let mut excep = ExcepInfo::default();
            let mut arg_err = 0u32;
            assert_eq!(
                (disp_vtbl.invoke)(
                    this,
                    dispid,
                    &IID_NULL,
                    0,
                    DISPATCH_METHOD,
                    (&raw mut dp).cast(),
                    std::ptr::null_mut(),
                    (&raw mut excep).cast(),
                    &mut arg_err
                ),
                S_OK,
                "{backend:?} late-bound MutateInPlace"
            );
            assert_eq!(oleaut::safe_array_destroy(psa), S_OK);
        }
    }
}
