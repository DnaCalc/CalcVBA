//! Calls on the owning runtime's ErrObject, shared by VM3 and compiled code.
use crate::{ExecState, Fault};
use calcvba_bundle::ProjectMemberKind;
use calcvba_oxir::value::ErrField;
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
        "number" => Some(ErrField::Number),
        "description" => Some(ErrField::Description),
        "source" => Some(ErrField::Source),
        "helpfile" => Some(ErrField::HelpFile),
        "helpcontext" => Some(ErrField::HelpContext),
        "lastdllerror" => Some(ErrField::LastDllError),
        _ => None,
    };
    let error = Fault::standard;
    if let Some(field) = field {
        if matches!(
            kind,
            ProjectMemberKind::PropertyLet | ProjectMemberKind::PropertySet
        ) {
            if args.len() != 1 {
                return Err(error(450));
            }
            if kind == ProjectMemberKind::PropertySet || field == ErrField::LastDllError {
                return Err(error(438));
            }
            exec.err_engine.set_field(field, &args[0])?;
            return Ok(Variant::empty());
        }
        if !args.is_empty() {
            return Err(error(450));
        }
        return Ok(exec.err_engine.field_value(field));
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

/// Normalize a fault raised by `CallByName` on `ErrObject.Raise`: Excel's
/// `CallByName` does not preserve custom `Source`/`Description`/help fields
/// from `ErrObject.Raise` and surfaces the standard default error for `code`.
pub fn normalize_call_by_name_err_fault(
    object: &ObjectRef,
    member_name: &str,
    fault: Fault,
) -> Fault {
    if calcvba_lib::objects::is_err(object) && member_name.eq_ignore_ascii_case("raise") {
        Fault::standard(fault.code)
    } else {
        fault
    }
}
