use crate::{BuildError, artifact_stem, build_artifacts, load_single_project, write_bytes};
use calcvba_bundle::{ExportTarget, ExportToken, ProjectMemberKind};
use calcvba_project::OutputType;
use cranelift_codegen::ir::{AbiParam, types};
use cranelift_module::Module;
use std::{
    fs,
    path::{Path, PathBuf},
};

#[derive(Debug, Clone)]
pub struct NativeLibraryBuildOptions {
    pub project_path: PathBuf,
    pub out_dir: PathBuf,
    /// Override a shipped template; ABI and image checks still apply.
    pub template_path: Option<PathBuf>,
}
#[derive(Debug, Clone)]
pub struct NativeLibraryBuildOutput {
    pub dll_path: PathBuf,
    pub map_path: PathBuf,
}
/// A standalone native build has the same template override and path options.
pub type NativeExecutableBuildOptions = NativeLibraryBuildOptions;
#[derive(Debug, Clone)]
pub struct NativeExecutableBuildOutput {
    pub exe_path: PathBuf,
    pub map_path: PathBuf,
}
struct NativeBuildOutput {
    image_path: PathBuf,
    map_path: PathBuf,
}

fn native_error(message: impl ToString) -> BuildError {
    BuildError::NativeExport(message.to_string())
}

struct NativeBodyImports {
    library_helpers:
        std::collections::HashMap<calcvba_bundle::NativeImplId, cranelift_module::FuncId>,
    long_from_f64: cranelift_module::FuncId,
    value_helpers: std::collections::HashMap<&'static str, cranelift_module::FuncId>,
    foreign_states:
        std::collections::HashMap<calcvba_jit::native::NativeCallSite, cranelift_module::DataId>,
}

fn native_body_imports(
    module: &mut calcvba_aot::PeModule,
    program: &calcvba_oxir::OxProgram,
) -> Result<NativeBodyImports, BuildError> {
    let ptr = module.target_config().pointer_type();
    let mut math_signature = module.make_signature();
    math_signature.params = vec![
        AbiParam::new(ptr),
        AbiParam::new(types::F64),
        AbiParam::new(ptr),
    ];
    math_signature.returns.push(AbiParam::new(types::I32));
    let long_from_f64 = module
        .declare_helper("cvb_long_from_f64", 1, &math_signature)
        .map_err(native_error)?;
    let mut library_helpers = std::collections::HashMap::new();
    for (native, name) in [
        (calcvba_bundle::NativeImplId::Abs, "cvb_abs_f64"),
        (calcvba_bundle::NativeImplId::Sqr, "cvb_sqr_f64"),
    ] {
        library_helpers.insert(
            native,
            module
                .declare_helper(name, 1, &math_signature)
                .map_err(native_error)?,
        );
    }
    let mut value_helpers = std::collections::HashMap::new();
    for (name, params) in calcvba_jit::native::value_helper_signatures() {
        let mut signature = module.make_signature();
        signature.params = params.iter().copied().map(AbiParam::new).collect();
        signature.returns.push(AbiParam::new(types::I32));
        value_helpers.insert(
            *name,
            module
                .declare_helper(
                    name,
                    calcvba_jit::native::value_helper_version(name),
                    &signature,
                )
                .map_err(native_error)?,
        );
    }
    let mut foreign_states = std::collections::HashMap::new();
    for (procedure, function) in program.funcs.iter().enumerate() {
        for block in &function.blocks {
            for (index, instruction) in block.instrs.iter().enumerate() {
                if !matches!(
                    instruction,
                    calcvba_oxir::OxInst::CallNative {
                        callee: calcvba_oxir::OxNativeCallee::Declare { .. },
                        ..
                    }
                ) {
                    continue;
                }
                let id = module
                    .declare_anonymous_data(true, false)
                    .map_err(native_error)?;
                let mut data = cranelift_module::DataDescription::new();
                data.define(vec![0; 16].into_boxed_slice());
                data.set_align(8);
                module.define_data(id, &data).map_err(native_error)?;
                foreign_states.insert((procedure, block.id.0, index), id);
            }
        }
    }
    Ok(NativeBodyImports {
        library_helpers,
        long_from_f64,
        value_helpers,
        foreign_states,
    })
}

/// Compile procedure bodies without publishing a DLL. The conformance census
/// uses exactly the same compiler and runtime imports as NativeLibrary builds.
pub fn check_native_program(
    program: &calcvba_oxir::OxProgram,
    roots: &[calcvba_oxir::FuncId],
) -> Result<(), BuildError> {
    let mut module = calcvba_aot::PeModule::new().map_err(native_error)?;
    let imports = native_body_imports(&mut module, program)?;
    calcvba_jit::native::compile_native_scalars(
        &mut module,
        program,
        roots,
        &imports.library_helpers,
        imports.long_from_f64,
        &imports.value_helpers,
        &imports.foreign_states,
    )
    .map_err(native_error)?;
    Ok(())
}

/// Compile and patch entirely in this process. There is deliberately no cargo,
/// linker invocation, runtime compiler or interpreter in the resulting DLL.
pub fn build_native_library(
    options: &NativeLibraryBuildOptions,
) -> Result<NativeLibraryBuildOutput, BuildError> {
    let output = build_native_artifact(options, OutputType::NativeLibrary)?;
    Ok(NativeLibraryBuildOutput {
        dll_path: output.image_path,
        map_path: output.map_path,
    })
}

/// Compile Sub Main and its reachable closure into the shipped executable
/// template. No compiler, linker or interpreter runs inside the resulting EXE.
pub fn build_native_executable(
    options: &NativeExecutableBuildOptions,
) -> Result<NativeExecutableBuildOutput, BuildError> {
    let output = build_native_artifact(options, OutputType::Exe)?;
    Ok(NativeExecutableBuildOutput {
        exe_path: output.image_path,
        map_path: output.map_path,
    })
}

/// Build an out-of-process Automation server containing compiled VBA entries.
pub fn build_native_com_executable(
    options: &NativeExecutableBuildOptions,
) -> Result<NativeExecutableBuildOutput, BuildError> {
    let output = build_native_artifact(options, OutputType::ComExe)?;
    Ok(NativeExecutableBuildOutput {
        exe_path: output.image_path,
        map_path: output.map_path,
    })
}

fn build_native_artifact(
    options: &NativeLibraryBuildOptions,
    required: OutputType,
) -> Result<NativeBuildOutput, BuildError> {
    let mut loaded = load_single_project(&options.project_path)?;
    let executable = required != OutputType::NativeLibrary;
    let com_executable = required == OutputType::ComExe;
    if loaded.output_type != required {
        return Err(native_error(format!(
            "{required:?} requires OutputType={required:?}"
        )));
    }
    if !executable && loaded.native_exports.is_empty() {
        return Err(native_error(
            "NativeLibrary requires at least one NativeExport",
        ));
    }
    let mut closure = calcvba_project::load_project_closure(&options.project_path)?;
    let com = if com_executable {
        Some(crate::native_com::prepare(
            &mut loaded,
            closure.last_mut().expect("root project"),
        )?)
    } else {
        None
    };
    let image = build_artifacts(&closure)?;
    let linked = crate::native_link::link(&image.programs).map_err(native_error)?;
    let program = &linked;
    let mut roots = Vec::new();
    if required == OutputType::Exe {
        if !loaded.native_exports.is_empty() {
            return Err(native_error(
                "Exe selects Sub Main; NativeExport declarations require NativeLibrary",
            ));
        }
        let entry = program
            .entry
            .ok_or_else(|| native_error("Exe requires Sub Main"))?;
        let function = &program.funcs[entry.0];
        if function.param_count != 0
            || function.return_local.is_some()
            || function.kind != calcvba_bundle::ProcedureKind::Sub
        {
            return Err(native_error("Exe entry must be a parameterless Sub"));
        }
        let (module_name, procedure_name) = loaded
            .entry_point
            .as_deref()
            .and_then(|name| name.rsplit_once('.'))
            .map(|(module, procedure)| (module.to_owned(), procedure.to_owned()))
            .unwrap_or_else(|| ("Main".into(), "Main".into()));
        loaded
            .native_exports
            .push(calcvba_project::NativeExportDescriptor {
                exported_name: "CalcVbaExeMain".into(),
                module_name,
                procedure_name,
                calling_convention: calcvba_project::CallingConvention::Stdcall,
                ordinal: None,
                string_encoding: calcvba_project::StringEncoding::Bstr,
                error_policy: calcvba_project::ErrorPolicy::HResult,
                kind: None,
                param_types: None,
                return_type: None,
            });
        roots.push(entry);
    }
    if required != OutputType::Exe {
        for selected in &loaded.native_exports {
            if [
                "CalcVbaGetLastError",
                "CalcVbaGetLastErrorHelp",
                "CalcVbaShutdown",
            ]
            .iter()
            .any(|name| selected.exported_name.eq_ignore_ascii_case(name))
            {
                return Err(native_error(
                    "CalcVbaGetLastError, CalcVbaGetLastErrorHelp and CalcVbaShutdown are reserved runtime exports",
                ));
            }
            let module = loaded
                .manifest
                .modules
                .iter()
                .find(|m| m.module_name.eq_ignore_ascii_case(&selected.module_name))
                .ok_or_else(|| {
                    native_error(format!(
                        "export {}: absent module {}",
                        selected.exported_name, selected.module_name
                    ))
                })?;
            if module.module_kind != calcvba_project::manifest::ModuleKind::Procedural {
                return Err(native_error(format!(
                    "export {} must select a Public procedure in a standard module",
                    selected.exported_name
                )));
            }
            let token = ExportToken::ModuleFunc {
                module: selected.module_name.clone(),
                member: selected.procedure_name.clone(),
                kind: ProjectMemberKind::Method,
            };
            let target = program
                .exports
                .iter()
                .find(|e| e.token.matches(&token))
                .and_then(|e| {
                    if let ExportTarget::Proc(id) = e.target {
                        Some(calcvba_oxir::FuncId(id))
                    } else {
                        None
                    }
                })
                .ok_or_else(|| {
                    native_error(format!(
                        "export {}: no Public procedure {}.{}",
                        selected.exported_name, selected.module_name, selected.procedure_name
                    ))
                })?;
            let func = &program.funcs[target.0];
            if let Some(parameter) = func.locals[..func.param_count]
                .iter()
                .filter_map(|local| local.param)
                .find(|parameter| parameter.variadic || parameter.optional)
            {
                let kind = if parameter.variadic {
                    "ParamArray signature"
                } else {
                    "optional signature"
                };
                let line = module
                    .source
                    .lines()
                    .position(|line| {
                        line.split_whitespace().any(|word| {
                            word.eq_ignore_ascii_case("Function")
                                || word.eq_ignore_ascii_case("Sub")
                        }) && line
                            .to_ascii_lowercase()
                            .contains(&selected.procedure_name.to_ascii_lowercase())
                    })
                    .unwrap_or(0)
                    + 1;
                return Err(native_error(format!(
                    "{}:{line}:1: native export procedure {}: {kind} requires a fixed foreign ABI",
                    module.module_name, func.name
                )));
            }
            roots.push(target);
        }
    }
    let template_bytes = if let Some(path) = &options.template_path {
        fs::read(path).map_err(|source| BuildError::Io {
            path: path.display().to_string(),
            source,
        })?
    } else if com_executable {
        include_bytes!("../../../assets/native/calcvba-native-com-x64.exe").to_vec()
    } else if executable {
        include_bytes!("../../../assets/native/calcvba-native-x64.exe").to_vec()
    } else {
        include_bytes!("../../../assets/native/calcvba-native-x64.dll").to_vec()
    };
    let template = calcvba_pe::Template::parse(&template_bytes, calcvba_pe::HELPER_ABI)
        .map_err(native_error)?;
    if template.is_dll() == executable {
        return Err(native_error(if executable {
            "Exe requires an EXE native template"
        } else {
            "NativeLibrary requires a DLL native template"
        }));
    }
    let mut module = calcvba_aot::PeModule::new().map_err(native_error)?;
    let mut enter_signature = module.make_signature();
    enter_signature.returns.push(AbiParam::new(types::I32));
    let enter = module
        .declare_helper("cvb_enter_scalar", 2, &enter_signature)
        .map_err(native_error)?;
    let ptr = module.target_config().pointer_type();
    let NativeBodyImports {
        library_helpers,
        long_from_f64,
        value_helpers,
        foreign_states,
    } = native_body_imports(&mut module, program)?;
    let plans = calcvba_jit::native::compile_native_scalars(
        &mut module,
        program,
        &roots,
        &library_helpers,
        long_from_f64,
        &value_helpers,
        &foreign_states,
    )
    .map_err(|error| {
        let message = error.to_string();
        for source in &loaded.manifest.modules {
            for (index, line) in source.source.lines().enumerate() {
                let words: Vec<_> = line
                    .split(|c: char| !c.is_alphanumeric() && c != '_')
                    .filter(|w| !w.is_empty())
                    .collect();
                for triple in words.windows(3) {
                    if triple[0].eq_ignore_ascii_case("Property")
                        && message.contains(&format!("procedure {}:", triple[2]))
                    {
                        return native_error(format!(
                            "{}:{}:1: {message}",
                            source.module_name,
                            index + 1
                        ));
                    }
                }
                for pair in words.windows(2) {
                    if (pair[0].eq_ignore_ascii_case("Function")
                        || pair[0].eq_ignore_ascii_case("Sub"))
                        && message.contains(&format!("procedure {}:", pair[1]))
                    {
                        return native_error(format!(
                            "{}:{}:1: {message}",
                            source.module_name,
                            index + 1
                        ));
                    }
                }
            }
        }
        let origin = loaded
            .manifest
            .modules
            .first()
            .map(|source| source.module_name.as_str())
            .unwrap_or("project");
        native_error(format!("{origin}:1:1: {message}"))
    })?;
    let state = module
        .declare_anonymous_data(true, false)
        .map_err(native_error)?;
    let mut bytes = vec![0; calcvba_jit::native::native_state_size(program)];
    bytes[..8].copy_from_slice(&u64::MAX.to_le_bytes());
    let mut data = cranelift_module::DataDescription::new();
    data.define(bytes.into_boxed_slice());
    data.set_align(8);
    module.define_data(state, &data).map_err(native_error)?;
    let mut record_signature = module.make_signature();
    record_signature.params = vec![
        AbiParam::new(ptr),
        AbiParam::new(types::I32),
        AbiParam::new(ptr),
        AbiParam::new(ptr),
    ];
    record_signature.returns.push(AbiParam::new(types::I32));
    let record_error = module
        .declare_helper("cvb_record_error", 1, &record_signature)
        .map_err(native_error)?;
    let mut error_signature = module.make_signature();
    error_signature.params = vec![AbiParam::new(ptr); 3];
    error_signature.returns.push(AbiParam::new(types::I32));
    let get_last_error = module
        .declare_helper("cvb_get_last_error", 1, &error_signature)
        .map_err(native_error)?;
    let runtime = calcvba_jit::native::NativeRuntime {
        foreign_states: foreign_states.values().copied().collect(),
        values: value_helpers,
        context: state,
        enter,
        record_error,
        get_last_error,
        initializer: program
            .global_initializer
            .and_then(|id| plans.iter().find(|p| p.procedure == id).map(|p| p.function)),
    };
    let mut exports = Vec::new();
    for (index, (selected, root)) in loaded.native_exports.iter().zip(&roots).enumerate() {
        let body = plans
            .iter()
            .find(|p| p.procedure == *root)
            .expect("compiled reachable export");
        let wrapper = calcvba_jit::native::compile_native_wrapper(
            &mut module,
            &format!("cvb$export${index}"),
            body,
            &runtime,
            calcvba_jit::native::NativeWrapperOptions {
                hresult: selected.error_policy == calcvba_project::ErrorPolicy::HResult,
                string_encoding: match selected.string_encoding {
                    calcvba_project::StringEncoding::Bstr => 0,
                    calcvba_project::StringEncoding::Utf16 => 1,
                    calcvba_project::StringEncoding::Ansi => 2,
                },
                source: &loaded.manifest.project_name,
                com_receiver: com
                    .as_ref()
                    .is_some_and(|com| com.receiver_exports.contains(&selected.exported_name)),
            },
        )
        .map_err(native_error)?;
        exports.push((selected.exported_name.clone(), selected.ordinal, wrapper));
    }
    let error_accessor = calcvba_jit::native::compile_last_error_wrapper(&mut module, &runtime)
        .map_err(native_error)?;
    exports.push(("CalcVbaGetLastError".into(), None, error_accessor));
    let help_accessor = calcvba_jit::native::compile_last_error_help_wrapper(&mut module, &runtime)
        .map_err(native_error)?;
    exports.push(("CalcVbaGetLastErrorHelp".into(), None, help_accessor));
    let shutdown = calcvba_jit::native::compile_shutdown_wrapper(&mut module, program, &runtime)
        .map_err(native_error)?;
    exports.push(("CalcVbaShutdown".into(), None, shutdown));
    let mut native = module.finish(&template, &exports).map_err(native_error)?;
    native.image = image.to_bytes().map_err(native_error)?;
    let name = artifact_stem(&loaded.manifest.project_name);
    let extension = if executable { "exe" } else { "dll" };
    let mut patched = template
        .patch(&format!("{name}.{extension}"), &native)
        .map_err(native_error)?;
    if required == OutputType::Exe {
        calcvba_pe::patch_executable_entries(&mut patched.bytes, &patched.map, "CalcVbaExeMain")
            .map_err(native_error)?;
    }
    let mut map = String::from("Export\tOrdinal\tProcedure\tWrapperRva\n");
    for (name, ordinal, rva) in &patched.map {
        let source = loaded
            .native_exports
            .iter()
            .find(|e| e.exported_name == *name);
        let procedure = source
            .map(|e| format!("{}.{}", e.module_name, e.procedure_name))
            .unwrap_or_else(|| "runtime error accessor".into());
        map.push_str(&format!("{name}\t{ordinal}\t{procedure}\t{rva:#010x}\n"));
    }
    // All compilation, relocation and PE checks succeeded before publishing.
    fs::create_dir_all(&options.out_dir).map_err(|source| BuildError::Io {
        path: options.out_dir.display().to_string(),
        source,
    })?;
    let image_path = options.out_dir.join(format!("{name}.{extension}"));
    let map_path = options.out_dir.join(format!("{name}.map"));
    write_bytes(&image_path, &patched.bytes)?;
    write_bytes(&map_path, map.as_bytes())?;
    if let Some(com) = com {
        crate::native_com::write(&options.out_dir, &name, &com)?;
    }
    Ok(NativeBuildOutput {
        image_path,
        map_path,
    })
}

/// The compiled-in template is also available to packaging tools.
pub fn write_native_template(path: &Path) -> Result<(), BuildError> {
    write_bytes(
        path,
        include_bytes!("../../../assets/native/calcvba-native-x64.dll"),
    )
}
