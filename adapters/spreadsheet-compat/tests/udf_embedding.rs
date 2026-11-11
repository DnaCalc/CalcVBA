use calcvba_spreadsheet_compat::{Backend, Error, Module, Value};

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
        "Public Function Boom()\nIf False Then Boom = UBound(Filter(Array(\"a\"), \"a\"))\nErr.Raise 13, \"UdfTest\", \"bad value\"\nEnd Function\n",
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
    let before = calcvba_runtime::current_thread_live_handle_counts();
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
            .balance_to(calcvba_runtime::current_thread_live_handle_counts())
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

#[test]
fn worksheet_doubles_use_checked_vba_parameter_coercion() {
    let module = Module::load(
        "Coercion",
        "Public Function Typed(x As Long) As Long\n Typed = x\nEnd Function\n",
    )
    .unwrap();
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        for (input, output) in [(2.5, 2.0), (3.5, 4.0), (-1.5, -2.0)] {
            assert_eq!(
                session.call("Typed", &[Value::Number(input)], 100).unwrap(),
                Value::Number(output)
            );
        }
        assert_eq!(
            session
                .call("Typed", &[Value::Text("12".into())], 100)
                .unwrap(),
            Value::Number(12.0)
        );
        assert_eq!(
            session.call("Typed", &[Value::Boolean(true)], 100).unwrap(),
            Value::Number(-1.0)
        );
        assert!(matches!(
            session.call("Typed", &[Value::Number(2147483648.0)], 100),
            Err(Error::Runtime { number: 6, .. })
        ));
    }
}

#[test]
fn spreadsheet_objects_read_live_ranges_caller_and_host_functions_on_both_backends() {
    use calcvba_spreadsheet_compat::{Range, SpreadsheetContext};
    use std::cell::Cell;
    struct Context {
        first: Cell<f64>,
        callbacks: Cell<u32>,
        volatile: Cell<bool>,
        invalid_caller: Cell<bool>,
    }
    impl SpreadsheetContext for Context {
        fn caller(&self) -> Option<Range> {
            Some(Range {
                sheet: 1,
                row: if self.invalid_caller.get() { 0 } else { 1 },
                column: 8,
                rows: 1,
                columns: 1,
            })
        }
        fn read_range(&self, range: Range) -> Result<Value, String> {
            if range.sheet != 1 || range.column != 1 || range.columns != 1 {
                return Err("missing host input".into());
            }
            let cells = (0..range.rows)
                .map(|i| {
                    Value::Number(if range.row + i == 1 {
                        self.first.get()
                    } else {
                        2.0
                    })
                })
                .collect::<Vec<_>>();
            if cells.len() == 1 {
                Ok(cells[0].clone())
            } else {
                Ok(Value::Array {
                    rows: range.rows,
                    columns: 1,
                    cells,
                })
            }
        }
        fn worksheet_function(&self, name: &str, args: &[Value]) -> Result<Value, String> {
            self.callbacks.set(self.callbacks.get() + 1);
            if name == "MATCH" {
                return Ok(Value::Error(2042));
            }
            if name != "SUM" {
                return Err("unknown host function".into());
            }
            let mut sum = 0.0;
            for arg in args {
                let value = match arg {
                    Value::Range(range) => self.read_range(*range)?,
                    value => value.clone(),
                };
                match value {
                    Value::Number(n) => sum += n,
                    Value::Array { cells, .. } => {
                        for cell in cells {
                            if let Value::Number(n) = cell {
                                sum += n;
                            }
                        }
                    }
                    _ => return Err("invalid sum argument".into()),
                }
            }
            Ok(Value::Number(sum))
        }
        fn set_volatile(&self, value: bool) -> Result<(), String> {
            self.volatile.set(value);
            Ok(())
        }
    }
    let source = r#"Private saved As Object
Public Function RangeInfo(r As Range) As String
 RangeInfo = CStr(r.Count) & ":" & CStr(r.Cells(1, 1).Value2) & ":" & Application.Caller.Address
End Function
Public Function WfSum(r As Range) As Double
 WfSum = Application.WorksheetFunction.Sum(r)
End Function
Public Function SaveRange(r As Range)
 Set saved = r
 SaveRange = 1
End Function
Public Function SavedValue()
 SavedValue = saved.Cells(1, 1).Value2
End Function
Public Function WfError()
 On Error GoTo Failed
 WfError = Application.WorksheetFunction.Match(1, Array(2, 3), 0)
 Exit Function
Failed:
 WfError = Err.Number
End Function
Public Function Vol()
 Application.Volatile
 Vol = 42
End Function
Public Function NumericVol()
 Application.Volatile 1
 NumericVol = 42
End Function
Public Function DoubleIndex(r As Range, i As Double)
 DoubleIndex = r.Cells(i, 1).Row
End Function
Public Function CallerOnly()
 CallerOnly = Application.Caller.Address
End Function
Public Function CatchOffset(r As Object)
 On Error Resume Next
 CatchOffset = r.Offset(1, 0).Value2
 CatchOffset = 7
End Function
Public Function CatchRead(r As Range)
 On Error Resume Next
 CatchRead = r.Value2
 CatchRead = 7
End Function
Public Function Echo(r)
 Echo = r
End Function
Public Function ArgType(r)
 ArgType = TypeName(r)
End Function
Public Function ArgVarType(r)
 ArgVarType = VarType(r)
End Function
Public Function NothingVarType()
 Dim r As Variant
 Set r = Nothing
 NothingVarType = VarType(r)
End Function
"#;
    let module = Module::load_spreadsheet("Grid", source).unwrap();
    let range = Value::Range(Range {
        sheet: 1,
        row: 1,
        column: 1,
        rows: 2,
        columns: 1,
    });
    for backend in backends() {
        let session = module.prepare(backend).unwrap();
        assert!(matches!(
            session.call("CallerOnly", &[], 1000),
            Err(Error::HostRefused(_))
        ));
        let mut context = Context {
            first: Cell::new(40.0),
            callbacks: Cell::new(0),
            volatile: Cell::new(false),
            invalid_caller: Cell::new(false),
        };
        assert_eq!(
            session
                .call_with_context(
                    "RangeInfo",
                    std::slice::from_ref(&range),
                    1000,
                    &mut context
                )
                .unwrap(),
            Value::Text("2:40:$H$1".into())
        );
        assert_eq!(
            session
                .call_with_context("WfSum", std::slice::from_ref(&range), 1000, &mut context)
                .unwrap(),
            Value::Number(42.0)
        );
        assert_eq!(context.callbacks.get(), 1);
        assert_eq!(
            session
                .call_with_context("ArgType", std::slice::from_ref(&range), 1000, &mut context)
                .unwrap(),
            Value::Text("Range".into())
        );
        assert_eq!(
            session
                .call_with_context(
                    "ArgVarType",
                    std::slice::from_ref(&range),
                    1000,
                    &mut context
                )
                .unwrap(),
            Value::Number(8204.0)
        );
        let scalar = Value::Range(Range {
            sheet: 1,
            row: 1,
            column: 1,
            rows: 1,
            columns: 1,
        });
        assert_eq!(
            session
                .call_with_context("ArgVarType", &[scalar], 1000, &mut context)
                .unwrap(),
            Value::Number(5.0)
        );
        assert_eq!(
            session
                .call_with_context("NothingVarType", &[], 1000, &mut context)
                .unwrap(),
            Value::Number(9.0)
        );
        assert_eq!(
            session
                .call_with_context(
                    "SaveRange",
                    std::slice::from_ref(&range),
                    1000,
                    &mut context
                )
                .unwrap(),
            Value::Number(1.0)
        );
        context.first.set(41.0);
        assert_eq!(
            session
                .call_with_context("SavedValue", &[], 1000, &mut context)
                .unwrap(),
            Value::Number(41.0)
        );
        assert_eq!(
            session
                .call_with_context("WfError", &[], 1000, &mut context)
                .unwrap(),
            Value::Number(1004.0)
        );
        assert_eq!(
            session
                .call_with_context("Vol", &[], 1000, &mut context)
                .unwrap(),
            Value::Number(42.0)
        );
        assert!(context.volatile.get());
        context.volatile.set(false);
        assert_eq!(
            session
                .call_with_context("NumericVol", &[], 1000, &mut context)
                .unwrap(),
            Value::Number(42.0)
        );
        assert!(context.volatile.get());
        for index in [1.5, 2.5] {
            assert_eq!(
                session
                    .call_with_context(
                        "DoubleIndex",
                        &[range.clone(), Value::Number(index)],
                        1000,
                        &mut context
                    )
                    .unwrap(),
                Value::Number(2.0)
            );
        }
        assert!(matches!(
            session.call_with_context(
                "CatchOffset",
                std::slice::from_ref(&range),
                1000,
                &mut context
            ),
            Err(Error::HostRefused(_))
        ));
        context.invalid_caller.set(true);
        assert!(matches!(
            session.call_with_context("CallerOnly", &[], 1000, &mut context),
            Err(Error::Runtime { number: 1004, .. })
        ));
        context.invalid_caller.set(false);
        let missing = Value::Range(Range {
            sheet: 99,
            row: 1,
            column: 1,
            rows: 1,
            columns: 1,
        });
        assert!(matches!(
            session.call_with_context("CatchRead", &[missing], 1000, &mut context),
            Err(Error::HostRefused(_))
        ));

        assert_eq!(
            session
                .call_with_context("Echo", std::slice::from_ref(&range), 1000, &mut context)
                .unwrap(),
            Value::Array {
                rows: 2,
                columns: 1,
                cells: vec![Value::Number(41.0), Value::Number(2.0)]
            }
        );
    }
}

#[test]
fn bounded_cache_shares_code_without_sharing_vba_state_and_live_sessions_survive_eviction() {
    use calcvba_spreadsheet_compat::CodeCache;
    let before = calcvba_runtime::current_thread_live_handle_counts();
    let source = "Private n As Long\nPublic Function CountIt() As Long\n n = n + 1\n CountIt = n\nEnd Function\nPublic Function Defaults(Optional x As Long = 7) As Long\n Defaults = x\nEnd Function\n";
    for backend in backends() {
        let mut cache = CodeCache::new(1, 1024 * 1024);
        let first = cache
            .prepare(Module::load("State", source).unwrap(), backend)
            .unwrap();
        let warm = cache
            .prepare(Module::load("State", source).unwrap(), backend)
            .unwrap();
        assert_eq!(first.backend(), backend);
        assert_eq!(first.code_identity(), warm.code_identity());
        assert_eq!(first.call("CountIt", &[], 100).unwrap(), Value::Number(1.0));
        assert_eq!(first.call("CountIt", &[], 100).unwrap(), Value::Number(2.0));
        assert_eq!(warm.call("CountIt", &[], 100).unwrap(), Value::Number(1.0));
        let changed = cache
            .prepare(
                Module::load("State", &source.replace("= 7)", "= 8)")).unwrap(),
                backend,
            )
            .unwrap();
        assert_ne!(changed.code_identity(), first.code_identity());
        assert_eq!(
            changed.call("Defaults", &[], 100).unwrap(),
            Value::Number(8.0)
        );
        assert_eq!(
            first.call("Defaults", &[], 100).unwrap(),
            Value::Number(7.0)
        );
        assert_eq!(first.call("CountIt", &[], 100).unwrap(), Value::Number(3.0));
        let cold_again = cache
            .prepare(Module::load("State", source).unwrap(), backend)
            .unwrap();
        assert_ne!(first.code_identity(), cold_again.code_identity());
        assert_eq!(
            cold_again.call("CountIt", &[], 100).unwrap(),
            Value::Number(1.0)
        );
        assert_eq!(cache.stats().hits, 1);
        assert_eq!(cache.stats().misses, 3);
        assert_eq!(cache.stats().entries, 1);
        assert_eq!(
            cache.stats().compilations,
            if backend == Backend::Jit { 3 } else { 0 }
        );
        cache.clear();
        assert_eq!(first.call("CountIt", &[], 100).unwrap(), Value::Number(4.0));
        first.reset().unwrap();
        assert_eq!(first.call("CountIt", &[], 100).unwrap(), Value::Number(1.0));
    }
    assert!(
        before
            .balance_to(calcvba_runtime::current_thread_live_handle_counts())
            .is_zero()
    );
}

#[test]
fn cache_limits_bypass_oversized_entries_and_do_not_retain_session_hosts() {
    use calcvba_hal::adapters::builder::HostBuilder;
    use calcvba_spreadsheet_compat::CodeCache;
    use std::sync::Arc;
    for backend in backends() {
        let host = HostBuilder::new().build();
        let weak = Arc::downgrade(&host);
        let module = Module::load("Pure", "Function F(): F = 42: End Function\n")
            .unwrap()
            .with_host(host);
        let mut cache = CodeCache::new(2, 1024 * 1024);
        let session = cache.prepare(module, backend).unwrap();
        assert!(weak.upgrade().is_some());
        drop(session);
        assert!(
            weak.upgrade().is_none(),
            "cached code must not own a disposed session's HAL"
        );
        assert_eq!(cache.stats().entries, 1);
        let mut tiny = CodeCache::new(1, 1);
        let session = tiny
            .prepare(
                Module::load("Pure", "Function F(): F = 42: End Function\n").unwrap(),
                backend,
            )
            .unwrap();
        assert_eq!(tiny.stats().entries, 0);
        assert_eq!(tiny.stats().bypasses, 1);
        assert_eq!(session.call("F", &[], 100).unwrap(), Value::Number(42.0));
    }
}

#[cfg(feature = "jit")]
#[test]
fn cache_uses_the_compiled_configuration_and_host_capability_profile() {
    use calcvba_hal::{adapters::builder::HostBuilder, model::HostPolicy};
    use calcvba_jit::JitConfiguration;
    use calcvba_spreadsheet_compat::CodeCache;
    let source = "Function F(): F = 42: End Function\n";
    let fast = JitConfiguration::native_with_optimization(true).unwrap();
    let plain = JitConfiguration::native_with_optimization(false).unwrap();
    assert_ne!(fast.identity(), plain.identity());
    let mut cache = CodeCache::new(4, 1024 * 1024);
    let first = cache
        .prepare_with_configuration(Module::load("Pure", source).unwrap(), &fast)
        .unwrap();
    let other = cache
        .prepare_with_configuration(Module::load("Pure", source).unwrap(), &plain)
        .unwrap();
    assert_ne!(first.code_identity(), other.code_identity());
    let denied = HostBuilder::new()
        .policy(HostPolicy {
            allow_process_spawn: false,
            allow_com_activation: false,
            ui_virtualization: calcvba_hal::model::UiVirtualizationMode::FailOnPrompt,
            unsupported_feature_mode: calcvba_hal::model::UnsupportedFeatureMode::CompileTime,
            ..HostPolicy::default()
        })
        .build();
    let policy = cache
        .prepare_with_configuration(
            Module::load("Pure", source).unwrap().with_host(denied),
            &fast,
        )
        .unwrap();
    assert_ne!(first.code_identity(), policy.code_identity());
    assert_eq!(cache.stats().compilations, 3);
    for session in [&first, &other, &policy] {
        assert_eq!(session.call("F", &[], 100).unwrap(), Value::Number(42.0));
    }
}
