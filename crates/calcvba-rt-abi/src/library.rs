//! Library identity shared by the JIT and direct native compiler.
use calcvba_bundle::{ExportTarget, NativeBody, NativeImplId, vba_library_bundle};
use calcvba_oxir::{ImportId, OxProgram};

pub struct Import {
    pub index: u32,
    pub implementation: NativeImplId,
    pub string_typed_alias: bool,
}

pub fn resolve_import(program: &OxProgram, import: ImportId) -> Result<Import, String> {
    let request = program
        .imports
        .get(import.0)
        .ok_or_else(|| format!("CallExtern import {} out of range", import.0))?;
    if !request.unit.eq_ignore_ascii_case("VBA") {
        return Err(format!(
            "CallExtern could not resolve non-VBA import unit '{}'",
            request.unit
        ));
    }
    let library = vba_library_bundle();
    let export = library
        .exports
        .iter()
        .find(|export| export.token.matches(&request.token))
        .ok_or_else(|| {
            format!(
                "CallExtern could not resolve VBA library import {}",
                import.0
            )
        })?;
    let ExportTarget::Proc(procedure) = export.target else {
        return Err("VBA library import is not a procedure".into());
    };
    let Some(NativeBody::Library(implementation)) = library
        .procedures
        .get(procedure)
        .and_then(|procedure| procedure.native)
    else {
        return Err("VBA library procedure has no native body".into());
    };
    let string_typed_alias = matches!(&request.token, calcvba_bundle::ExportToken::ModuleFunc { member, .. } if implementation.is_string_typed_library_alias(member));
    let index = NativeImplId::ALL
        .iter()
        .position(|id| *id == implementation)
        .and_then(|index| u32::try_from(index).ok())
        .ok_or("library implementation index overflow")?;
    Ok(Import {
        index,
        implementation,
        string_typed_alias,
    })
}
