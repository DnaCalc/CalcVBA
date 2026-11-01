use oxvba_embed::{Backend, Error, Module, ProcedureKind};
use oxvba_runtime::Variant;

fn backends() -> Vec<Backend> {
    let mut all = vec![Backend::Interpreter];
    if cfg!(feature = "jit") {
        all.push(Backend::Jit);
    }
    all
}
const ADD: &str = "Public Function Add2(ByVal x As Double, ByVal y As Double) As Double\nAdd2 = x + y\nEnd Function\n";

#[test]
fn dynamic_raise_preserves_runtime_number_and_explicit_text_and_help_fields() {
    let source = r#"
Public Function DescribeFault(ByVal code As Long) As String
    On Error GoTo Handler
    Err.Raise code, "host-source", "host-description", "host-help", 73
    Exit Function
Handler:
    DescribeFault = CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & CStr(Err.HelpContext)
End Function
Public Function UnhandledFault(ByVal code As Long) As Long
    Err.Raise code, "callee", "explicit message"
End Function
"#;
    for backend in backends() {
        let module = Module::load("Api", source).unwrap();
        let session = module.prepare(backend).unwrap();
        let handled = session
            .bind("Api", "DescribeFault", ProcedureKind::Function)
            .unwrap();
        let result = session
            .call_bound(&handled, &[Variant::from_i32(9)], 1000)
            .unwrap();
        assert_eq!(
            result.as_bstr().unwrap().as_str(),
            "9|host-source|host-description|host-help|73"
        );
        let unhandled = session
            .bind("Api", "UnhandledFault", ProcedureKind::Function)
            .unwrap();
        assert!(
            matches!(session.call_bound(&unhandled, &[Variant::from_i32(11)], 1000), Err(Error::Runtime { number: 11, description, vba_source }) if description == "explicit message" && vba_source == "callee")
        );
    }
}

#[test]
fn typed_and_dynamic_numeric_calls_match() {
    for backend in backends() {
        let module = Module::load("Math", ADD).unwrap();
        let session = module.prepare(backend).unwrap();
        let typed = session
            .bind_typed::<(f64, f64), f64>("Math", "Add2")
            .unwrap();
        let dynamic = session
            .bind("Math", "Add2", ProcedureKind::Function)
            .unwrap();
        for (a, b) in [(19.0, 23.0), (-1.0, 3.0), (-0.0, -0.0), (1e-200, 1e-200)] {
            let expected = session
                .call_bound(
                    &dynamic,
                    &[Variant::from_f64(a), Variant::from_f64(b)],
                    1000,
                )
                .unwrap()
                .as_f64()
                .unwrap();
            assert_eq!(
                session.call_typed(&typed, (a, b), 1000).unwrap().to_bits(),
                expected.to_bits()
            );
        }
        assert!(matches!(
            session.call_typed(&typed, (1.0, 2.0), 0),
            Err(Error::BudgetExceeded)
        ));
        assert!(matches!(
            session.call_typed(&typed, (f64::MAX, f64::MAX), 1000),
            Err(Error::Runtime { number: 6, .. })
        ));
        assert_eq!(
            session.call_typed(&typed, (19.0, 23.0), 1000).unwrap(),
            42.0
        );
    }
}

#[test]
fn stale_cross_session_and_failed_reload_handles_are_safe() {
    for backend in backends() {
        let mut session = Module::load("Math", ADD)
            .unwrap()
            .into_session(backend)
            .unwrap();
        let handle = session
            .bind_typed::<(f64, f64), f64>("Math", "Add2")
            .unwrap();
        let dynamic = session
            .bind("Math", "Add2", ProcedureKind::Function)
            .unwrap();
        let other = Module::load("Math", ADD)
            .unwrap()
            .into_session(backend)
            .unwrap();
        assert!(matches!(
            other.call_typed(&handle, (1.0, 2.0), 100),
            Err(Error::StaleHandle)
        ));
        // The old code is freed by successful reload; using either old handle
        // must fail before reading or calling its address.
        session
            .reload(Module::load("Math", &ADD.replace("x + y", "x * y")).unwrap())
            .unwrap();
        assert!(matches!(
            session.call_typed(&handle, (1.0, 2.0), 100),
            Err(Error::StaleHandle)
        ));
        assert!(matches!(
            session.call_bound(&dynamic, &[], 100),
            Err(Error::StaleHandle)
        ));
        let fresh = session
            .bind_typed::<(f64, f64), f64>("Math", "Add2")
            .unwrap();
        assert_eq!(session.call_typed(&fresh, (6.0, 7.0), 100).unwrap(), 42.0);
    }
}

#[test]
fn catalogue_retains_option_private_defaults_subs_and_accessors() {
    let source = "Option Private Module\nPublic Function Answer(Optional ByVal x As Long = 42) As Long\nAnswer = x\nEnd Function\nPrivate Function Hidden() As Long\nHidden = 7\nEnd Function\nPublic Sub Touch()\nEnd Sub\nPublic Property Get Size() As Long\nSize = 9\nEnd Property\nPublic Property Let Size(ByVal value As Long)\nEnd Property\n";
    let module = Module::load("Api", source).unwrap();
    let catalogue = module.procedures().collect::<Vec<_>>();
    assert_eq!(catalogue.len(), 5);
    assert!(catalogue.iter().all(|p| p.option_private_module));
    assert_eq!(
        catalogue[0].parameters[0].default,
        Some(oxvba_bundle::coreir::CoreConst::I32(42))
    );
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        assert!(matches!(
            session.bind("Api", "Hidden", ProcedureKind::Function),
            Err(Error::UnknownProcedure(_))
        ));
        let sub = session.bind("Api", "Touch", ProcedureKind::Sub).unwrap();
        assert_eq!(
            session.call_bound(&sub, &[], 1000).unwrap().vtype(),
            oxvba_runtime::VarType::Empty
        );
        let getter = session
            .bind("Api", "Size", ProcedureKind::PropertyGet)
            .unwrap();
        let typed_getter = session.bind_typed_handle::<(), i32>(&getter).unwrap();
        assert_eq!(session.call_typed(&typed_getter, (), 1000).unwrap(), 9);
        assert_eq!(
            session.call_bound(&getter, &[], 1000).unwrap().as_i32(),
            Some(9)
        );
        let setter = session
            .bind("Api", "Size", ProcedureKind::PropertyLet)
            .unwrap();
        session
            .call_bound(&setter, &[Variant::from_i32(20)], 1000)
            .unwrap();
    }
}

#[test]
fn long_overflow_and_division_faults_match() {
    let module = Module::load("Math", "Public Function Sum(ByVal a As Long, ByVal b As Long) As Long\nSum = a + b\nEnd Function\nPublic Function Divide(ByVal a As Double, ByVal b As Double) As Double\nDivide = a / b\nEnd Function\n").unwrap();
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        let sum = session
            .bind_typed::<(i32, i32), i32>("Math", "Sum")
            .unwrap();
        assert_eq!(session.call_typed(&sum, (19, 23), 1000).unwrap(), 42);
        assert!(matches!(
            session.call_typed(&sum, (i32::MAX, 1), 1000),
            Err(Error::Runtime { number: 6, .. })
        ));
        let divide = session
            .bind_typed::<(f64, f64), f64>("Math", "Divide")
            .unwrap();
        assert!(matches!(
            session.call_typed(&divide, (1.0, 0.0), 1000),
            Err(Error::Runtime { number: 11, .. })
        ));
    }
}

#[test]
fn typed_binding_refuses_byref_variant_and_wrong_signature() {
    let module = Module::load(
        "Math",
        &(ADD.to_owned() + "Public Function Dynamic(x)\nDynamic = x\nEnd Function\n"),
    )
    .unwrap();
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        assert!(matches!(
            session.bind_typed::<(i32, i32), f64>("Math", "Add2"),
            Err(Error::ArgumentMismatch)
        ));
        assert!(matches!(
            session.bind_typed::<(f64,), f64>("Math", "Dynamic"),
            Err(Error::UnsupportedSignature(_))
        ));
    }
}

#[test]
fn typed_native_argument_positions_control_flow_and_err_are_preserved() {
    let source = "Public Function Zero() As Long\nZero = 42\nEnd Function\nPublic Function Square(ByVal x As Double) As Double\nSquare = x * x\nEnd Function\nPublic Function Three(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Double\nThree = x + y + z\nEnd Function\nPublic Function Mixed(ByVal a As Long, ByVal b As Double, ByVal c As Double, ByVal d As Long) As Double\nMixed = b + c\nEnd Function\nPublic Function MixedLong(ByVal a As Long, ByVal b As Double, ByVal c As Double, ByVal d As Long) As Long\nMixedLong = a + d\nEnd Function\nPublic Function Greatest(ByVal a As Double, ByVal b As Double) As Double\nIf a > b Then\nGreatest = a\nElse\nGreatest = b\nEnd If\nEnd Function\nPublic Function LastError() As Long\nLastError = Err.Number\nEnd Function\n";
    for backend in backends() {
        let session = Module::load("Abi", source)
            .unwrap()
            .into_session(backend)
            .unwrap();
        assert_eq!(
            session
                .call_typed(
                    &session.bind_typed::<(), i32>("Abi", "Zero").unwrap(),
                    (),
                    100
                )
                .unwrap(),
            42
        );
        assert_eq!(
            session
                .call_typed(
                    &session.bind_typed::<(f64,), f64>("Abi", "Square").unwrap(),
                    (3.0,),
                    100
                )
                .unwrap(),
            9.0
        );
        assert_eq!(
            session
                .call_typed(
                    &session
                        .bind_typed::<(f64, f64, f64), f64>("Abi", "Three")
                        .unwrap(),
                    (1.0, 2.0, 3.0),
                    100
                )
                .unwrap(),
            6.0
        );
        let mixed = session
            .bind_typed::<(i32, f64, f64, i32), f64>("Abi", "Mixed")
            .unwrap();
        assert_eq!(
            session
                .call_typed(&mixed, (1000, 19.0, 23.0, -9), 100)
                .unwrap(),
            42.0
        );
        let long = session
            .bind_typed::<(i32, f64, f64, i32), i32>("Abi", "MixedLong")
            .unwrap();
        assert_eq!(
            session
                .call_typed(&long, (19, 1000.0, -500.0, 23), 100)
                .unwrap(),
            42
        );
        let greatest = session
            .bind_typed::<(f64, f64), f64>("Abi", "Greatest")
            .unwrap();
        assert_eq!(session.call_typed(&greatest, (1.0, 2.0), 100).unwrap(), 2.0);
        assert_eq!(session.call_typed(&greatest, (3.0, 2.0), 100).unwrap(), 3.0);
        for invalid in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            assert!(matches!(
                session.call_typed(&greatest, (invalid, 2.0), 100),
                Err(Error::Runtime { number: 6, .. })
            ));
        }
        assert!(matches!(
            session.call_typed(&long, (i32::MAX, 0.0, 0.0, 1), 100),
            Err(Error::Runtime { number: 6, .. })
        ));
        let err = session
            .bind("Abi", "LastError", ProcedureKind::Function)
            .unwrap();
        assert_eq!(
            session.call_bound(&err, &[], 100).unwrap().as_i32(),
            Some(6)
        );
        assert!(matches!(
            session.call_typed(&long, (1, 0.0, 0.0, 1), 0),
            Err(Error::BudgetExceeded)
        ));
        assert_eq!(
            session.call_bound(&err, &[], 100).unwrap().as_i32(),
            Some(6)
        );
        session.reset().unwrap();
        assert_eq!(
            session.call_bound(&err, &[], 100).unwrap().as_i32(),
            Some(0)
        );
        assert_eq!(
            session.call_typed(&long, (19, 0.0, 0.0, 23), 100).unwrap(),
            42
        );
    }
}

#[test]
fn cached_sessions_have_distinct_handles_and_keep_code_live_after_eviction() {
    for backend in backends() {
        let mut cache = oxvba_embed::CodeCache::new(1, 1024 * 1024);
        let first = cache
            .prepare(Module::load("Math", ADD).unwrap(), backend)
            .unwrap();
        let second = cache
            .prepare(Module::load("Math", ADD).unwrap(), backend)
            .unwrap();
        assert_eq!(first.code_identity(), second.code_identity());
        let handle = first.bind_typed::<(f64, f64), f64>("Math", "Add2").unwrap();
        assert!(matches!(
            second.call_typed(&handle, (19.0, 23.0), 100),
            Err(Error::StaleHandle)
        ));
        cache.clear();
        assert_eq!(first.call_typed(&handle, (19.0, 23.0), 100).unwrap(), 42.0);
        drop(first);
        assert!(matches!(
            second.call_typed(&handle, (19.0, 23.0), 100),
            Err(Error::StaleHandle)
        ));
    }
}

#[cfg(feature = "jit")]
#[test]
fn failed_jit_reload_preserves_the_old_generation() {
    let mut session = Module::load("Math", ADD)
        .unwrap()
        .into_session(Backend::Jit)
        .unwrap();
    let handle = session
        .bind_typed::<(f64, f64), f64>("Math", "Add2")
        .unwrap();
    let unsupported = Module::load(
        "Bad",
        "Public Function F()\nDim value As Long\nF = VarPtr(value)\nEnd Function\n",
    )
    .unwrap();
    assert!(matches!(
        session.reload(unsupported),
        Err(Error::UnsupportedBackend(_))
    ));
    assert_eq!(
        session.call_typed(&handle, (19.0, 23.0), 100).unwrap(),
        42.0
    );
}

#[test]
fn exact_carriers_preserve_null_cverr_and_unpaired_utf16() {
    let source = "Public Function Echo(ByVal value As Variant) As Variant\nEcho = value\nEnd Function\nPublic Function Fail()\nErr.Raise 13\nEnd Function\n";
    for backend in backends() {
        let session = Module::load("Values", source)
            .unwrap()
            .into_session(backend)
            .unwrap();
        let echo = session
            .bind("Values", "Echo", ProcedureKind::Function)
            .unwrap();
        for value in [
            Variant::null(),
            Variant::from_error_code(2042),
            Variant::from_i16(7),
            Variant::from_currency_scaled_i64(42_000),
            Variant::from_utf16_units(&[0xD800, 0, 0x03B1]),
        ] {
            let result = session
                .call_bound(&echo, std::slice::from_ref(&value), 100)
                .unwrap();
            assert_eq!(result.vtype(), value.vtype());
            if let Some(text) = value.as_bstr() {
                assert_eq!(
                    result.as_bstr().unwrap().to_utf16_units(),
                    text.to_utf16_units()
                );
            }
            if let Some(error) = value.as_error_code() {
                assert_eq!(result.as_error_code(), Some(error));
            }
        }
        let fail = session
            .bind("Values", "Fail", ProcedureKind::Function)
            .unwrap();
        assert!(matches!(
            session.call_bound(&fail, &[], 100),
            Err(Error::Runtime { number: 13, .. })
        ));
    }
}

#[test]
fn typed_loop_uses_unboxed_locals_and_consumes_its_block_budget() {
    let source = "Public Function SumTo(ByVal count As Long) As Long\nDim total As Long\nDo While count > 0\ntotal = total + count\ncount = count - 1\nLoop\nSumTo = total\nEnd Function\n";
    for backend in backends() {
        let session = Module::load("Looping", source)
            .unwrap()
            .into_session(backend)
            .unwrap();
        let typed = session
            .bind_typed::<(i32,), i32>("Looping", "SumTo")
            .unwrap();
        assert_eq!(session.call_typed(&typed, (6,), 1000).unwrap(), 21);
        assert!(matches!(
            session.call_typed(&typed, (1000,), 10),
            Err(Error::BudgetExceeded)
        ));
        assert_eq!(session.call_typed(&typed, (6,), 1000).unwrap(), 21);
        let floating = Module::load("Looping", &source.replace("As Long", "As Double"))
            .unwrap()
            .into_session(backend)
            .unwrap();
        let typed = floating
            .bind_typed::<(f64,), f64>("Looping", "SumTo")
            .unwrap();
        assert_eq!(floating.call_typed(&typed, (2.5,), 1000).unwrap(), 4.5);
        assert!(matches!(
            floating.call_typed(&typed, (1000.0,), 10),
            Err(Error::BudgetExceeded)
        ));
    }
}

#[test]
fn finite_typed_host_inputs_are_a_boundary_check_not_a_different_vm_body() {
    let source = "Public Function IgnoreInput(ByVal value As Double) As Long\nIgnoreInput = 42\nEnd Function\n";
    for backend in backends() {
        let session = Module::load("Api", source)
            .unwrap()
            .into_session(backend)
            .unwrap();
        let bound = session
            .bind("Api", "IgnoreInput", ProcedureKind::Function)
            .unwrap();
        let typed = session.bind_typed_handle::<(f64,), i32>(&bound).unwrap();
        for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            assert!(matches!(
                session.call_typed(&typed, (value,), 100),
                Err(Error::Runtime { number: 6, .. })
            ));
            // The explicit exact-Variant route preserves engine carriers. An
            // unused value must not give VM3 a different body from the JIT.
            assert_eq!(
                session
                    .call_bound(&bound, &[Variant::from_f64(value)], 100)
                    .unwrap()
                    .as_i32(),
                Some(42)
            );
        }
    }
}

#[test]
fn module_qualified_handles_keep_compiler_ids_and_distinct_signatures() {
    use oxvba_symbol::manifest as sym;
    let manifest = sym::SymbolProjectManifest {
        project_name: "TwoModules".into(),
        project_kind: sym::ProjectKind::Library,
        modules: [("First", "Private Function Hidden() As Long\nEnd Function\nPublic Function Same(ByVal x As Long) As Long\nSame = x + 1\nEnd Function\n"), ("Second", "Public Function Same(ByVal x As Double) As Double\nSame = x + 2\nEnd Function\nPublic Sub Touch()\nEnd Sub\n")].into_iter().map(|(name, source)| sym::ModuleUnit {
            module_name: name.into(), module_kind: sym::ModuleKind::Procedural,
            attributes: sym::ModuleAttributes::named(name), source: source.into(),
        }).collect(),
        references: vec![], reference_projects: vec![],
        conditional_constants: Default::default(), conditional_compilation_target: Default::default(),
    };
    let module = Module::load_project(&manifest, &oxvba_symbol::CatalogTypeLibResolver).unwrap();
    for backend in backends() {
        let session = module.clone().into_session(backend).unwrap();
        let first = session.bind_typed::<(i32,), i32>("FIRST", "same").unwrap();
        let second = session.bind_typed::<(f64,), f64>("Second", "Same").unwrap();
        assert_eq!(session.call_typed(&first, (41,), 100).unwrap(), 42);
        assert_eq!(session.call_typed(&second, (41.0,), 100).unwrap(), 43.0);
        assert_eq!(session.procedures().len(), 4);
        assert!(session.com_interfaces().is_empty());
    }
}
