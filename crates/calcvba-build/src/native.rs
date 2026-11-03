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

fn native_error(message: impl ToString) -> BuildError {
    BuildError::NativeExport(message.to_string())
}

/// Compile and patch entirely in this process. There is deliberately no cargo,
/// linker invocation, runtime compiler or interpreter in the resulting DLL.
pub fn build_native_library(
    options: &NativeLibraryBuildOptions,
) -> Result<NativeLibraryBuildOutput, BuildError> {
    let loaded = load_single_project(&options.project_path)?;
    if loaded.output_type != OutputType::NativeLibrary {
        return Err(native_error(
            "NativeLibrary requires OutputType=NativeLibrary",
        ));
    }
    if loaded.native_exports.is_empty() {
        return Err(native_error(
            "NativeLibrary requires at least one NativeExport",
        ));
    }
    let closure = calcvba_project::load_project_closure(&options.project_path)?;
    if closure.len() != 1 {
        return Err(native_error(
            "native cross-project calls are not yet implemented",
        ));
    }
    let image = build_artifacts(&closure)?;
    let program = image
        .programs
        .last()
        .ok_or_else(|| native_error("empty native project"))?;
    let mut roots = Vec::new();
    for selected in &loaded.native_exports {
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
        roots.push(target);
    }
    let template_bytes = if let Some(path) = &options.template_path {
        fs::read(path).map_err(|source| BuildError::Io {
            path: path.display().to_string(),
            source,
        })?
    } else {
        include_bytes!("../../../assets/native/calcvba-native-x64.dll").to_vec()
    };
    let template = calcvba_pe::Template::parse(&template_bytes, calcvba_pe::HELPER_ABI)
        .map_err(native_error)?;
    let mut module = calcvba_aot::PeModule::new().map_err(native_error)?;
    let mut enter_signature = module.make_signature();
    enter_signature.returns.push(AbiParam::new(types::I32));
    let enter = module
        .declare_helper("cvb_enter_scalar", 2, &enter_signature)
        .map_err(native_error)?;
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
    let plans = calcvba_jit::native::compile_native_scalars(
        &mut module,
        program,
        &roots,
        &library_helpers,
        long_from_f64,
    )
    .map_err(native_error)?;
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
            enter,
        )
        .map_err(native_error)?;
        exports.push((selected.exported_name.clone(), selected.ordinal, wrapper));
    }
    let mut native = module.finish(&template, &exports).map_err(native_error)?;
    native.image = image.to_bytes().map_err(native_error)?;
    let name = artifact_stem(&loaded.manifest.project_name);
    let patched = template
        .patch(&format!("{name}.dll"), &native)
        .map_err(native_error)?;
    let mut map = String::from("Export\tOrdinal\tProcedure\tWrapperRva\n");
    for (name, ordinal, rva) in &patched.map {
        let source = loaded
            .native_exports
            .iter()
            .find(|e| e.exported_name == *name)
            .expect("selected export");
        map.push_str(&format!(
            "{name}\t{ordinal}\t{}.{}\t{rva:#010x}\n",
            source.module_name, source.procedure_name
        ));
    }
    // All compilation, relocation and PE checks succeeded before publishing.
    fs::create_dir_all(&options.out_dir).map_err(|source| BuildError::Io {
        path: options.out_dir.display().to_string(),
        source,
    })?;
    let dll_path = options.out_dir.join(format!("{name}.dll"));
    let map_path = options.out_dir.join(format!("{name}.map"));
    write_bytes(&dll_path, &patched.bytes)?;
    write_bytes(&map_path, map.as_bytes())?;
    Ok(NativeLibraryBuildOutput { dll_path, map_path })
}

/// The compiled-in template is also available to packaging tools.
pub fn write_native_template(path: &Path) -> Result<(), BuildError> {
    write_bytes(
        path,
        include_bytes!("../../../assets/native/calcvba-native-x64.dll"),
    )
}
