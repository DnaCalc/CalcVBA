//! Native ParamArray aliases are attached to hidden caller array cells. Ordinary
//! array assignment copies values, so these borrowed caller addresses cannot escape.
use calcvba_rt_abi::native_export::Context;
use std::collections::HashMap;
use windows_sys::Win32::System::{Com::SAFEARRAY, Ole::*};

#[derive(Clone)]
pub(super) struct Aliases {
    pub depth: u32,
    targets: Vec<Option<(usize, u32)>>,
}
pub(super) type Arrays = HashMap<usize, Aliases>;

/// # Safety
/// Compiler metadata is count/result/lower/input kinds/alias kinds; both pointer
/// vectors contain count live cells (null alias pointers mean value arguments).
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_paramarray_literal(
    context: *mut Context,
    metadata: *const u32,
    arguments: *const *const u8,
    aliases: *const *mut u8,
    out: *mut u8,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: the generated frame retains all cells and immutable descriptors.
        unsafe {
            let status = super::values::cvb_array_literal(metadata, arguments, out);
            if status != 0 {
                return Err(status);
            }
            let count = metadata.read() as usize;
            let targets = (0..count)
                .map(|index| {
                    let pointer = aliases.add(index).read();
                    (!pointer.is_null())
                        .then(|| (pointer.addr(), metadata.add(3 + count + index).read()))
                })
                .collect();
            (*super::services::runtime(context)).param_arrays.insert(
                out.addr(),
                Aliases {
                    depth: (*context).depth,
                    targets,
                },
            );
            Ok(())
        }
    })
}

/// # Safety
/// Exact native array/value carriers and indices, retained by the active frame.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_paramarray_set(
    context: *mut Context,
    array: *const *mut SAFEARRAY,
    indices: *const i32,
    rank: u32,
    kind: u32,
    value: *const u8,
    from: u32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: alias metadata belongs to an active caller frame, never an
        // escaping array copy. No runtime borrow spans an owning value operation.
        unsafe {
            let aliases = (*super::services::runtime(context))
                .param_arrays
                .get(&array.addr())
                .cloned();
            let Some(aliases) = aliases else {
                let status = super::values::cvb_array_set(array, indices, rank, kind, value, from);
                return if status == 0 { Ok(()) } else { Err(status) };
            };
            if rank != 1 || array.read().is_null() {
                return Err(9);
            }
            let mut element = std::ptr::null_mut();
            super::values::status(SafeArrayPtrOfIndex(array.read(), indices, &mut element))?;
            let mut lower = 0;
            super::values::status(SafeArrayGetLBound(array.read(), 1, &mut lower))?;
            let flat = indices.read().checked_sub(lower).ok_or(9)? as usize;
            if let Some(Some((target, target_kind))) = aliases.targets.get(flat).copied() {
                super::values::copy(target as *mut u8, value, target_kind, from)?;
                for (index, alias) in aliases.targets.iter().enumerate() {
                    if *alias == Some((target, target_kind)) {
                        let coordinate = lower.checked_add(index as i32).ok_or(6)?;
                        let status = super::values::cvb_array_set(
                            array,
                            &coordinate,
                            1,
                            kind,
                            target as *const u8,
                            target_kind,
                        );
                        if status != 0 {
                            return Err(status);
                        }
                    }
                }
                Ok(())
            } else {
                let status = super::values::cvb_array_set(array, indices, rank, kind, value, from);
                if status == 0 { Ok(()) } else { Err(status) }
            }
        }
    })
}
