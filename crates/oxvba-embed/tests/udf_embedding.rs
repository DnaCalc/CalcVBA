use oxvba_embed::{Backend, Error, Module, Value};

fn backends() -> Vec<Backend> {
    let mut backends = vec![Backend::Interpreter];
    if cfg!(feature = "jit") {
        backends.push(Backend::Jit);
    }
    backends
}

#[test]
fn add2_lists_compiler_signatures_and_calls_by_name() {
    let module = Module::load("Math", "Function Add2(a, b): Add2 = a + b: End Function\nPrivate Function Hidden(): Hidden = 9: End Function\nPublic Sub Main(): Err.Raise 99: End Sub\n").unwrap();
    let functions: Vec<_> = module.functions().collect();
    assert_eq!(functions.len(), 1);
    assert_eq!(functions[0].qualified_name, "Math.Add2");
    assert_eq!(
        functions[0]
            .parameters
            .iter()
            .map(|param| param.name.as_str())
            .collect::<Vec<_>>(),
        ["a", "b"]
    );
    assert!(functions[0].parameters.iter().all(|param| param.by_ref));
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        for name in ["Add2", "mAtH.aDd2"] {
            assert_eq!(
                session
                    .call(name, &[Value::Number(19.0), Value::Number(23.0)], 1000)
                    .unwrap(),
                Value::Number(42.0)
            );
        }
        assert!(matches!(
            session.call("Hidden", &[], 1000),
            Err(Error::UnknownFunction(_))
        ));
        assert!(matches!(
            session.call("Add2", &[], 1000),
            Err(Error::ArgumentCount {
                expected: 2,
                received: 0
            })
        ));
    }
}

#[test]
fn spreadsheet_values_round_trip_structurally() {
    let module = Module::load("Values", "Public Function Echo(value)\nEcho = value\nEnd Function\nPublic Function Pick(value)\nPick = value(2, 1)\nEnd Function\n").unwrap();
    let array = Value::Array {
        rows: 2,
        columns: 3,
        cells: vec![
            Value::Number(1.0),
            Value::Text("α\0😀".into()),
            Value::Boolean(true),
            Value::Error(2042),
            Value::Empty,
            Value::Number(-0.0),
        ],
    };
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        for value in [
            Value::Empty,
            Value::Number(42.5),
            Value::Text("hello\0😀".into()),
            Value::Boolean(false),
            Value::Error(2042),
            array.clone(),
        ] {
            let result = session
                .call("Echo", std::slice::from_ref(&value), 1000)
                .unwrap();
            assert_eq!(result, value, "{backend:?}");
        }
        assert_eq!(
            session
                .call("Pick", std::slice::from_ref(&array), 1000)
                .unwrap(),
            Value::Error(2042)
        );
    }
}

#[test]
fn runtime_errors_remain_distinct_from_worksheet_error_values() {
    let module = Module::load("Errors", "Public Function Boom()\nErr.Raise 13\nEnd Function\nPublic Function NAValue()\nNAValue = CVErr(2042)\nEnd Function\n").unwrap();
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        match session.call("Boom", &[], 1000) {
            Err(Error::Runtime {
                number,
                description,
                vba_source,
            }) => {
                assert_eq!(number, 13);
                assert_eq!(description, "Type mismatch");
                assert_eq!(vba_source, "UdfProject");
            }
            other => panic!("{backend:?}: {other:?}"),
        }
        assert_eq!(
            session.call("NAValue", &[], 1000).unwrap(),
            Value::Error(2042)
        );
    }
}

#[cfg(feature = "jit")]
#[test]
fn unsupported_jit_code_is_rejected_before_execution_without_fallback() {
    let module = Module::load(
        "Errors",
        "Public Function Boom()\nErr.Raise 13, \"UdfTest\", \"bad value\"\nEnd Function\n",
    )
    .unwrap();
    assert!(matches!(
        module.prepare(Backend::Jit),
        Err(Error::UnsupportedBackend(_))
    ));
    let session = module.prepare(Backend::Interpreter).unwrap();
    assert!(matches!(
        session.call("Boom", &[], 100),
        Err(Error::Runtime { number: 13, .. })
    ));
}

#[test]
fn budget_cannot_be_caught_and_resets_for_each_call() {
    let module = Module::load("Budget", "Public Function Forever()\nOn Error Resume Next\nDo\nLoop\nEnd Function\nPublic Function Good()\nGood = 42\nEnd Function\nPublic Function Recurse()\nRecurse = Recurse()\nEnd Function\n").unwrap();
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        assert!(matches!(
            session.call("Good", &[], 0),
            Err(Error::BudgetExceeded)
        ));
        assert!(matches!(
            session.call("Forever", &[], 100),
            Err(Error::BudgetExceeded)
        ));
        assert!(matches!(
            session.call("Recurse", &[], 20),
            Err(Error::BudgetExceeded)
        ));
        assert_eq!(session.call("Good", &[], 100).unwrap(), Value::Number(42.0));
    }
}

#[test]
fn private_byref_copies_and_persistent_globals_are_the_call_contract() {
    let module = Module::load("State", "Private count As Long\nPublic Function Change(ByRef value As Double) As Double\nvalue = value + 1\ncount = count + 1\nChange = value + count\nEnd Function\n").unwrap();
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        let args = [Value::Number(40.0)];
        for index in 0..100 {
            assert_eq!(
                session.call("Change", &args, 1000).unwrap(),
                Value::Number(42.0 + index as f64)
            );
            assert_eq!(args, [Value::Number(40.0)]);
        }
    }
}

#[test]
fn invalid_values_and_compile_errors_are_typed() {
    for source in [
        "Function Broken(\n",
        "Function Broken()\n",
        "Sub Broken(): End Function\n",
    ] {
        assert!(matches!(
            Module::load("Bad", source),
            Err(Error::Compile(_))
        ));
    }
    let module = Module::load(
        "Values",
        "Public Function Echo(value)\nEcho = value\nEnd Function\n",
    )
    .unwrap();
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        for number in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            assert!(matches!(
                session.call("Echo", &[Value::Number(number)], 100),
                Err(Error::ValueConversion(_))
            ));
        }
        assert!(matches!(
            session.call(
                "Echo",
                &[Value::Array {
                    rows: 2,
                    columns: 3,
                    cells: vec![]
                }],
                100
            ),
            Err(Error::ValueConversion(_))
        ));
    }
}

#[test]
fn variant_array_parameters_and_unrepresentable_text_are_explicit() {
    let module = Module::load("Arrays", "Public Function Pick(values() As Variant)\nPick = values(2, 1)\nEnd Function\nPublic Function OddBytes()\nOddBytes = LeftB(\"A\", 1)\nEnd Function\n").unwrap();
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        let values = Value::Array {
            rows: 2,
            columns: 1,
            cells: vec![Value::Number(1.0), Value::Number(42.0)],
        };
        assert_eq!(
            session.call("Pick", &[values], 100).unwrap(),
            Value::Number(42.0)
        );
        assert!(matches!(
            session.call("OddBytes", &[], 100),
            Err(Error::ValueConversion(_))
        ));
    }
}

#[test]
fn repeated_load_prepare_call_drop_releases_runtime_values() {
    let before = oxvba_runtime::current_thread_live_handle_counts();
    for _ in 0..16 {
        let module = Module::load(
            "Values",
            "Public Function Echo(value)\nEcho = value\nEnd Function\n",
        )
        .unwrap();
        for backend in backends() {
            let session = module.prepare(backend).unwrap();
            assert_eq!(
                session
                    .call("Echo", &[Value::Text("owned".into())], 100)
                    .unwrap(),
                Value::Text("owned".into())
            );
        }
    }
    assert!(
        before
            .balance_to(oxvba_runtime::current_thread_live_handle_counts())
            .is_zero()
    );
}

#[test]
fn function_created_arrays_have_explicit_shape_and_empty_arrays_are_rejected() {
    let module = Module::load("Arrays", "Public Function Row()\nRow = Array(19, 23)\nEnd Function\nPublic Function NoCells()\nNoCells = Array()\nEnd Function\n").unwrap();
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        assert_eq!(
            session.call("Row", &[], 1000).unwrap(),
            Value::Array {
                rows: 1,
                columns: 2,
                cells: vec![Value::Number(19.0), Value::Number(23.0)]
            }
        );
        assert!(matches!(
            session.call("NoCells", &[], 1000),
            Err(Error::ValueConversion(_))
        ));
    }
}

#[cfg(not(feature = "jit"))]
#[test]
fn interpreter_only_build_reports_jit_as_unavailable() {
    let module = Module::load(
        "Math",
        "Public Function Answer()\nAnswer = 42\nEnd Function\n",
    )
    .unwrap();
    assert!(matches!(
        module.prepare(Backend::Jit),
        Err(Error::BackendUnavailable)
    ));
}

#[test]
fn optional_defaults_missing_and_paramarray_are_compiler_owned() {
    let module = Module::load("Args", "Public Function Defaults(Optional x As Long = 7, Optional text As String = \"hi\") As String\n Defaults = CStr(x) & text\nEnd Function\nPublic Function MissingArg(Optional x) As Boolean\n MissingArg = IsMissing(x)\nEnd Function\nPublic Function Packed(ParamArray x() As Variant) As Long\n Packed = UBound(x) - LBound(x) + 1\nEnd Function\nPublic Function PackedItem(ParamArray x() As Variant)\n PackedItem = x(1)\nEnd Function\n").unwrap();
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        assert_eq!(
            session.call("Defaults", &[], 1000).unwrap(),
            Value::Text("7hi".into())
        );
        assert_eq!(
            session
                .call(
                    "Defaults",
                    &[Value::Missing, Value::Text("there".into())],
                    1000
                )
                .unwrap(),
            Value::Text("7there".into())
        );
        assert_eq!(
            session.call("MissingArg", &[], 1000).unwrap(),
            Value::Boolean(true)
        );
        assert_eq!(
            session.call("MissingArg", &[Value::Empty], 1000).unwrap(),
            Value::Boolean(false)
        );
        assert_eq!(
            session.call("Packed", &[], 1000).unwrap(),
            Value::Number(0.0)
        );
        assert_eq!(
            session
                .call("Packed", &[Value::Number(1.0), Value::Number(2.0)], 1000)
                .unwrap(),
            Value::Number(2.0)
        );
        assert_eq!(
            session
                .call(
                    "PackedItem",
                    &[Value::Number(1.0), Value::Text("item".into())],
                    1000
                )
                .unwrap(),
            Value::Text("item".into())
        );
    }
}

#[test]
fn sessions_isolate_static_storage_and_reset_after_errors_and_budget_abort() {
    let module = Module::load("State", "Public Function Counter() As Long\n Static n As Long\n n = n + 1\n Counter = n\nEnd Function\nPublic Function Fault()\n Err.Raise 5\nEnd Function\nPublic Function Spin()\n Do\n Loop\nEnd Function\n").unwrap();
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        let other = module.prepare(backend).unwrap();
        assert_eq!(
            session.call("Counter", &[], 100).unwrap(),
            Value::Number(1.0)
        );
        assert!(matches!(
            session.call("Fault", &[], 100),
            Err(Error::Runtime { number: 5, .. })
        ));
        assert!(matches!(
            session.call("Spin", &[], 20),
            Err(Error::BudgetExceeded)
        ));
        assert_eq!(
            session.call("Counter", &[], 100).unwrap(),
            Value::Number(2.0)
        );
        assert_eq!(other.call("Counter", &[], 100).unwrap(), Value::Number(1.0));
        session.reset().unwrap();
        assert_eq!(
            session.call("Counter", &[], 100).unwrap(),
            Value::Number(1.0)
        );
    }
}

#[test]
fn global_as_new_and_initialization_abort_match_between_sessions() {
    let module = Module::load("Init", "Private c As New Collection\nPrivate a(1 To 2) As Long\nPublic Function Good()\n Good = 42\nEnd Function\nPublic Function CountIt()\n c.Add 1\n CountIt = c.Count\nEnd Function\nPublic Function ClearIt()\n Set c = Nothing\n ClearIt = CountIt()\nEnd Function\n").unwrap();
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        assert!(matches!(
            session.call("Good", &[], 0),
            Err(Error::BudgetExceeded)
        ));
        assert_eq!(
            session.call("Good", &[], 1000).unwrap(),
            Value::Number(42.0)
        );
        assert_eq!(
            session.call("CountIt", &[], 1000).unwrap(),
            Value::Number(1.0)
        );
        assert_eq!(
            session.call("CountIt", &[], 1000).unwrap(),
            Value::Number(2.0)
        );
        assert_eq!(
            session.call("ClearIt", &[], 1000).unwrap(),
            Value::Number(1.0)
        );
    }
}

#[test]
fn end_is_declined_instead_of_retaining_state_after_project_termination() {
    assert!(
        matches!(Module::load("Ender", "Public Function StopIt()\n End\nEnd Function\n"), Err(Error::UnsupportedSignature(reason)) if reason.contains("reset"))
    );
}
