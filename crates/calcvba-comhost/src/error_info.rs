use super::*;

/// Preserve a VBA fault across class-factory HRESULT boundaries (`Class_Initialize` errors).
pub(super) fn with_project_session<R>(
    f: impl FnOnce(&mut ProjectRuntimeSession) -> Result<R, calcvba_engine::PhaseDiagnostic>,
) -> Result<R, i32> {
    // SAFETY: the documented null form clears this thread's previous error.
    unsafe {
        SetErrorInfo(0, ptr::null_mut());
    }
    let mut error = None;
    with_session(|session| {
        f(session).map_err(|fault| {
            let message = fault.to_string();
            error = Some(fault);
            message
        })
    })
    .map_err(|_| error.as_ref().map(publish_project_error).unwrap_or(E_FAIL))
}

fn publish_project_error(error: &calcvba_engine::PhaseDiagnostic) -> i32 {
    let Some(fault) = error.vba_fault() else {
        return E_FAIL;
    };
    let hr = if (1..=65535).contains(&fault.code) {
        (0x800a_0000u32 | fault.code as u32) as i32
    } else {
        fault.code
    };
    let default_source = Utf16String::from(
        descriptor()
            .map(|d| d.project_name.as_str())
            .unwrap_or("CalcVBA"),
    );
    let source = fault.source.as_ref().unwrap_or(&default_source);
    let empty_help = Utf16String::new();
    let help_file = fault.help_file.as_ref().unwrap_or(&empty_help);
    let wide = |text: &Utf16String| {
        text.as_units()
            .iter()
            .copied()
            .chain(Some(0))
            .collect::<Vec<_>>()
    };
    // SAFETY: CreateErrorInfo supplies an owned ICreateErrorInfo with the
    // documented Automation vtable. Each input stays live through its setter;
    // QI and SetErrorInfo retain their own references before ours are released.
    unsafe {
        let mut create = ptr::null_mut();
        if CreateErrorInfo(&mut create) >= 0 && !create.is_null() {
            let table = &**create.cast::<*const CreateErrorInfoVtbl>();
            (table.set_guid)(create, &IID_IDISPATCH);
            (table.set_source)(create, wide(source).as_ptr());
            (table.set_description)(create, wide(&fault.message).as_ptr());
            (table.set_help_file)(create, wide(help_file).as_ptr());
            (table.set_help_context)(create, fault.help_context.unwrap_or(0) as u32);
            let mut info = ptr::null_mut();
            if (table.query_interface)(create, &IID_IERRORINFO, &mut info) >= 0 && !info.is_null() {
                SetErrorInfo(0, info);
                release_unknown(info);
            }
            (table.release)(create);
        }
    }
    hr
}
