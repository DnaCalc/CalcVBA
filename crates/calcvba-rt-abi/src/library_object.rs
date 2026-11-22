//! Calls on the owning runtime's ErrObject, shared by VM3 and compiled code.
use crate::{ExecState, Fault};
use calcvba_bundle::ProjectMemberKind;
use calcvba_runtime::{ObjectRef, Variant, variant::VarType};

/// Built-in object dispatch. This uses the active runtime's error state directly;
/// aliases of Err never hold a stale copy of Number or Description.
pub fn invoke_library_object(
    exec: &mut ExecState<'_>,
    object: &ObjectRef,
    name: &str,
    kind: ProjectMemberKind,
    args: &[Variant],
    names: &[Option<String>],
) -> Result<Variant, Fault> {
    if !calcvba_lib::objects::is_err(object) {
        return calcvba_lib::objects::dispatch(object, name, kind, args, names)
            .map_err(Fault::from_lib);
    }
    let name = if name.is_empty() {
        "number".to_string()
    } else {
        name.to_ascii_lowercase()
    };
    let field = match name.as_str() {
        "number" => Some(crate::RT_ERR_FIELD_NUMBER),
        "description" => Some(crate::RT_ERR_FIELD_DESCRIPTION),
        "source" => Some(crate::RT_ERR_FIELD_SOURCE),
        "helpfile" => Some(crate::RT_ERR_FIELD_HELP_FILE),
        "helpcontext" => Some(crate::RT_ERR_FIELD_HELP_CONTEXT),
        "lastdllerror" => Some(crate::RT_ERR_FIELD_LAST_DLL_ERROR),
        _ => None,
    };
    let error = |number| Fault::new(number, crate::default_error_message(number));
    if let Some(field) = field {
        if matches!(
            kind,
            ProjectMemberKind::PropertyLet | ProjectMemberKind::PropertySet
        ) {
            if args.len() != 1 {
                return Err(error(450));
            }
            if kind == ProjectMemberKind::PropertySet || field == crate::RT_ERR_FIELD_LAST_DLL_ERROR
            {
                return Err(error(438));
            }
            crate::apply_err_field_set(exec, field, &args[0])?;
            return Ok(Variant::empty());
        }
        if !args.is_empty() {
            return Err(error(450));
        }
        return Ok(match field {
            crate::RT_ERR_FIELD_NUMBER => Variant::from_i32(exec.err_engine.err.number),
            crate::RT_ERR_FIELD_DESCRIPTION => {
                Variant::from_string(&exec.err_engine.err.description)
            }
            crate::RT_ERR_FIELD_SOURCE => Variant::from_string(&exec.err_engine.err.source),
            crate::RT_ERR_FIELD_HELP_FILE => Variant::from_string(&exec.err_engine.err.help_file),
            crate::RT_ERR_FIELD_HELP_CONTEXT => Variant::from_i32(exec.err_engine.err.help_context),
            crate::RT_ERR_FIELD_LAST_DLL_ERROR => Variant::from_i32(exec.err_engine.last_dll_error),
            _ => unreachable!(),
        });
    }
    if kind != ProjectMemberKind::Method {
        return Err(error(438));
    }
    match name.as_str() {
        "clear" => {
            if !args.is_empty() {
                return Err(error(450));
            }
            exec.err_engine.clear_err();
            Ok(Variant::empty())
        }
        "raise" => {
            let params = ["Number", "Source", "Description", "HelpFile", "HelpContext"];
            if args.is_empty() || args.len() > params.len() {
                return Err(error(450));
            }
            let mut ordered = [None; 5];
            for (i, arg) in args.iter().enumerate() {
                let index = match names.get(i).and_then(Option::as_deref) {
                    Some(name) => params
                        .iter()
                        .position(|p| p.eq_ignore_ascii_case(name))
                        .ok_or_else(|| error(448))?,
                    None => i,
                };
                if ordered[index].is_some() {
                    return Err(error(450));
                }
                if !(arg.vtype() == VarType::Error
                    && arg.as_error_code() == Some(0x8002_0004u32 as i32))
                {
                    ordered[index] = Some(arg);
                }
            }
            Err(exec.err_engine.fault_from_raise_args(ordered, true))
        }
        _ => Err(error(438)),
    }
}
