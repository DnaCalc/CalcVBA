//! Native entries for the member slots of a class's dual vtable (dc-yqq0.24).
//!
//! An early-bound foreign caller (compiled C++, a .NET delegate over the vtable, Excel's own
//! VBA) calls vtable slot `7 + n` with the dual-interface signature
//! `HRESULT member(this, [in] args..., [out, retval] result*)`. Every slot entry is a small
//! stub (`mov eax, slot; jmp common`); the common part spills the argument registers and calls
//! [`slot_dispatch`], which finds the member by its `vtable_slot`, reads each argument by its
//! declared type, invokes the member through the same path as `IDispatch::Invoke`, writes
//! ByRef arguments back and stores the return value.
//!
//! Windows x64 only for now. Elsewhere the slots keep the `E_NOTIMPL` placeholder.
//!
//! Win64 argument layout at the common entry: `this` and arguments 1..=3 are in
//! rcx/rdx/r8/r9 (spilled to the caller's home area), later arguments follow on the stack, so
//! the home area and the stack arguments form one array of 8-byte slots indexed by position. A
//! Double, Single or Date in positions 1..=3 arrives in xmm1..xmm3 instead, which the common
//! part spills to a second array. A ByVal Variant is passed as a pointer to the caller's copy.

use core::ffi::c_void;

/// Number of slot entries (vtable slots `0..SLOT_THUNK_COUNT`; slots below 7 are unused).
#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
const SLOT_THUNK_COUNT: usize = 4096;

/// Each stub is padded to 16 bytes (`mov eax, imm32` + `jmp rel32` is at most 10).
#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
const SLOT_THUNK_STRIDE: usize = 16;

#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
core::arch::global_asm!(
    ".text",
    ".balign 16",
    ".globl calcvba_class_slot_thunks",
    "calcvba_class_slot_thunks:",
    ".set calcvba_class_slot_index, 0",
    ".rept {count}",
    ".balign 16",
    "movl $calcvba_class_slot_index, %eax",
    "jmp calcvba_class_slot_common",
    ".set calcvba_class_slot_index, calcvba_class_slot_index + 1",
    ".endr",
    ".balign 16",
    "calcvba_class_slot_common:",
    ".seh_proc calcvba_class_slot_common",
    "movq %rcx, 8(%rsp)",
    "movq %rdx, 16(%rsp)",
    "movq %r8, 24(%rsp)",
    "movq %r9, 32(%rsp)",
    "subq $72, %rsp",
    ".seh_stackalloc 72",
    ".seh_endprologue",
    "movsd %xmm1, 32(%rsp)",
    "movsd %xmm2, 40(%rsp)",
    "movsd %xmm3, 48(%rsp)",
    "movl %eax, %ecx",
    "leaq 80(%rsp), %rdx",
    "leaq 32(%rsp), %r8",
    "call {dispatch}",
    "addq $72, %rsp",
    "ret",
    ".seh_endproc",
    count = const SLOT_THUNK_COUNT,
    dispatch = sym slot_dispatch,
    options(att_syntax),
);

#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
unsafe extern "C" {
    static calcvba_class_slot_thunks: u8;
}

/// The native entry for vtable slot `slot`, when this platform has one.
#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
pub(super) fn slot_entry(slot: usize) -> Option<*const c_void> {
    if !(7..SLOT_THUNK_COUNT).contains(&slot) {
        return None;
    }
    // Only the address of the stub table is taken; it is never read as data.
    let base = &raw const calcvba_class_slot_thunks;
    Some(base.wrapping_add(slot * SLOT_THUNK_STRIDE).cast::<c_void>())
}

/// The native entry for vtable slot `slot`, when this platform has one.
#[cfg(not(all(target_os = "windows", target_arch = "x86_64")))]
pub(super) fn slot_entry(_slot: usize) -> Option<*const c_void> {
    None
}

#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
use win64::slot_dispatch;

#[cfg(all(target_os = "windows", target_arch = "x86_64"))]
mod win64 {
    use core::ffi::c_void;
    use std::panic::{AssertUnwindSafe, catch_unwind};

    use calcvba_automation::com::{
        DISP_E_TYPEMISMATCH, E_FAIL, E_NOTIMPL, E_POINTER, HResult, IID_IDISPATCH,
        RPC_E_WRONG_THREAD, S_OK,
    };
    use calcvba_automation::oleaut;
    use calcvba_automation::safe_array::SafeArray;

    use super::super::error_info::{clear_thread_error_info, publish_member_fault_error_info};
    use super::super::implements::{ImplementedInterfaceFace, is_impl_query_interface};
    use super::super::{
        coerce_for_byref_writeback, copy_dispatch_arg, current_thread_id, invoke_object_member,
        vba_error_to_hresult,
    };
    use crate::object_ref::{
        COMPAT_OBJECT_VTBL, CompatObjectBase, ObjectRef, RawRuntimeIUnknown,
        RuntimeInterfaceDescriptor, RuntimeMemberDescriptor, RuntimeMemberInvokeKind,
        RuntimeValueType,
    };
    use crate::{VarType, Variant, coerce};

    /// Called by the common stub. `args[0]` is `this`, `args[n]` the n-th argument's 8-byte
    /// slot; `xmm[n - 1]` holds a floating-point argument in position `n` (1..=3).
    pub(super) unsafe extern "system" fn slot_dispatch(
        slot: u32,
        args: *const u64,
        xmm: *const u64,
    ) -> HResult {
        // A panic must not unwind into the foreign caller.
        catch_unwind(AssertUnwindSafe(|| {
            // SAFETY: the stub passes the spilled argument arrays of a live call.
            unsafe { dispatch(slot as usize, args, xmm) }
        }))
        .unwrap_or(E_FAIL)
    }

    /// The `VARTYPE` an argument or return value of `ty` has in the dual interface, or `None`
    /// when slot calls do not carry that type yet.
    fn slot_vt(ty: RuntimeValueType) -> Option<u16> {
        Some(match ty {
            RuntimeValueType::Integer => 0x02,
            RuntimeValueType::Long => 0x03,
            RuntimeValueType::Single => 0x04,
            RuntimeValueType::Double => 0x05,
            RuntimeValueType::Currency => 0x06,
            RuntimeValueType::Date => 0x07,
            RuntimeValueType::String => 0x08,
            RuntimeValueType::Object => 0x09,
            RuntimeValueType::Boolean => 0x0B,
            RuntimeValueType::Variant => 0x0C,
            RuntimeValueType::Byte => 0x11,
            RuntimeValueType::LongLong | RuntimeValueType::LongPtr => 0x14,
            RuntimeValueType::Decimal | RuntimeValueType::Record | RuntimeValueType::Array => {
                return None;
            }
        })
    }

    fn is_float(vt: u16) -> bool {
        matches!(vt, 0x04 | 0x05 | 0x07)
    }

    /// `value` as the declared type `vt`, by VBA's conversion rules.
    fn as_slot_type(value: &Variant, vt: u16) -> Result<Variant, HResult> {
        match vt {
            0x0C | 0x09 => Ok(value.clone()),
            _ => match VarType::from_u16(vt) {
                Some(target) if value.core().raw_vt() != vt => coerce::coerce_to(value, target)
                    .or_else(|_| coerce_for_byref_writeback(value, vt)),
                _ => Ok(value.clone()),
            },
        }
        .map_err(|_| DISP_E_TYPEMISMATCH)
    }

    fn member_for_slot(
        interface: &'static RuntimeInterfaceDescriptor,
        slot: usize,
    ) -> Option<&'static RuntimeMemberDescriptor> {
        interface
            .members
            .iter()
            .find(|member| member.vtable_slot.map(usize::from) == Some(slot))
    }

    enum SlotParamCell {
        VariantCell {
            by_ref: bool,
            vt: u16,
            target: *mut u8,
            cell: Variant,
        },
        SafeArrayRef {
            by_ref: bool,
            target: *mut *mut c_void,
        },
    }

    enum SlotReturnTarget {
        VariantCell { vt: u16, out: *mut u8 },
        SafeArrayOut { out: *mut *mut c_void },
    }

    unsafe fn clone_raw_safearray_to_variant(psa: *mut c_void) -> Result<Variant, HResult> {
        if psa.is_null() {
            return Ok(Variant::empty());
        }
        let mut copy = core::ptr::null_mut();
        // SAFETY: `psa` is a live caller-owned `SAFEARRAY*`; `safe_array_copy` deep-copies it.
        let hr = unsafe { oleaut::safe_array_copy(psa, &mut copy) };
        if hr < 0 {
            return Err(hr);
        }
        if copy.is_null() {
            return Ok(Variant::empty());
        }
        // SAFETY: `copy` is an owned `SAFEARRAY*` returned by `safe_array_copy`.
        match unsafe { SafeArray::from_raw_safearray_owned(copy, None) } {
            Some(arr) => Ok(Variant::from_safearray(arr)),
            None => {
                // SAFETY: destroy unadopted copy.
                unsafe {
                    let _ = oleaut::safe_array_destroy(copy);
                }
                Err(DISP_E_TYPEMISMATCH)
            }
        }
    }

    unsafe fn dispatch(slot: usize, args: *const u64, xmm: *const u64) -> HResult {
        // SAFETY: position 0 is `this`.
        let this = unsafe { *args } as usize as *mut c_void;
        if this.is_null() {
            return E_POINTER;
        }
        // Resolve primary `CompatObjectBase` or secondary `ImplementedInterfaceFace`.
        // SAFETY: `this` is a live COM interface pointer kept alive by the caller's reference.
        let (owner, interface, interface_iid) = unsafe {
            let vtbl = (*this.cast::<RawRuntimeIUnknown>()).vtbl;
            let qi = (*vtbl).query_interface;
            if core::ptr::fn_addr_eq(qi, COMPAT_OBJECT_VTBL.query_interface) {
                let owner = this.cast::<CompatObjectBase>();
                let desc = (*owner).class_descriptor;
                (
                    owner,
                    desc.dispatch_interface(),
                    desc.default_interface_iid.unwrap_or(IID_IDISPATCH),
                )
            } else if is_impl_query_interface(qi) {
                let face = &*this.cast::<ImplementedInterfaceFace>();
                (
                    face.owner,
                    Some(face.interface),
                    face.interface.identity.guid,
                )
            } else {
                return E_POINTER;
            }
        };
        // SAFETY: `owner` is a live `CompatObjectBase`, kept alive by the caller's reference.
        if current_thread_id() != unsafe { (*owner).owner_thread } {
            return RPC_E_WRONG_THREAD;
        }
        clear_thread_error_info();
        let Some(member) = interface.and_then(|iface| member_for_slot(iface, slot)) else {
            return E_NOTIMPL;
        };

        // Each argument as a VT_BYREF cell (or SAFEARRAY* / SAFEARRAY**) over the caller's storage.
        let mut cells = Vec::with_capacity(member.params.len());
        let mut values = Vec::with_capacity(member.params.len());
        for (index, param) in member.params.iter().enumerate() {
            let position = index + 1;
            if param.param_array {
                return E_NOTIMPL;
            }
            if param.value_type == RuntimeValueType::Array {
                // SAFETY: caller passed at least `params.len()` arguments after `this`.
                let raw_word = unsafe { *args.add(position) } as usize;
                if param.by_ref {
                    let pp = raw_word as *mut *mut c_void;
                    let value = if pp.is_null() {
                        Variant::empty()
                    } else {
                        // SAFETY: caller supplied a readable `SAFEARRAY**`.
                        let psa = unsafe { *pp };
                        // SAFETY: `psa` is null or a live `SAFEARRAY*`.
                        match unsafe { clone_raw_safearray_to_variant(psa) } {
                            Ok(val) => val,
                            Err(hr) => return hr,
                        }
                    };
                    values.push(value);
                    cells.push(SlotParamCell::SafeArrayRef {
                        by_ref: !pp.is_null(),
                        target: pp,
                    });
                } else {
                    let psa = raw_word as *mut c_void;
                    // SAFETY: `psa` is null or a live `SAFEARRAY*`.
                    let value = match unsafe { clone_raw_safearray_to_variant(psa) } {
                        Ok(val) => val,
                        Err(hr) => return hr,
                    };
                    values.push(value);
                    cells.push(SlotParamCell::SafeArrayRef {
                        by_ref: false,
                        target: core::ptr::null_mut(),
                    });
                }
                continue;
            }
            let Some(vt) = slot_vt(param.value_type) else {
                return E_NOTIMPL;
            };
            let raw = if !param.by_ref && is_float(vt) && position <= 3 {
                // SAFETY: the stub spilled xmm1..=xmm3.
                unsafe { xmm.add(position - 1) }
            } else {
                // SAFETY: the caller passed at least `params.len()` arguments after `this`.
                unsafe { args.add(position) }
            };
            let target = if param.by_ref || vt == 0x0C {
                // SAFETY: the slot holds the caller's pointer to the value.
                unsafe { *raw as usize as *mut u8 }
            } else {
                raw.cast_mut().cast::<u8>()
            };
            if target.is_null() {
                return E_POINTER;
            }
            // SAFETY: `target` is the caller's live storage of type `vt` for this call.
            let cell = unsafe { Variant::by_ref(vt, target) };
            let core = cell.core();
            // SAFETY: `core` is a VT_BYREF cell over live storage.
            match unsafe { copy_dispatch_arg(&raw const core) } {
                Ok(value) => values.push(value),
                Err(hr) => return hr,
            }
            cells.push(SlotParamCell::VariantCell {
                by_ref: param.by_ref,
                vt,
                target,
                cell,
            });
        }

        let returns = matches!(
            member.invoke_kind,
            RuntimeMemberInvokeKind::Method | RuntimeMemberInvokeKind::PropertyGet
        ) && member.return_type.is_some();
        let result = if returns {
            if member.return_type == Some(RuntimeValueType::Array) {
                // SAFETY: the retval pointer follows the declared arguments.
                let out =
                    unsafe { *args.add(member.params.len() + 1) } as usize as *mut *mut c_void;
                if out.is_null() {
                    return E_POINTER;
                }
                // SAFETY: `out` is the caller's writable `SAFEARRAY**` out-slot.
                unsafe {
                    *out = core::ptr::null_mut();
                }
                Some(SlotReturnTarget::SafeArrayOut { out })
            } else {
                let Some(vt) = member.return_type.and_then(slot_vt) else {
                    return E_NOTIMPL;
                };
                // SAFETY: the retval pointer follows the declared arguments.
                let out = unsafe { *args.add(member.params.len() + 1) } as usize as *mut u8;
                if out.is_null() {
                    return E_POINTER;
                }
                // An [out] slot holds nothing yet: give owned types an empty value to replace.
                // SAFETY: `out` is the caller's writable storage of type `vt`.
                unsafe {
                    match vt {
                        0x08 | 0x09 => out.cast::<*mut c_void>().write(core::ptr::null_mut()),
                        0x0C => out.cast::<Variant>().write(Variant::empty()),
                        _ => {}
                    }
                }
                Some(SlotReturnTarget::VariantCell { vt, out })
            }
        } else {
            None
        };

        // ByRef arguments over the same storage are one variable.
        let aliases: Vec<usize> = (0..cells.len())
            .map(|index| match &cells[index] {
                SlotParamCell::VariantCell {
                    by_ref: true,
                    vt,
                    target,
                    ..
                } => cells[..index]
                    .iter()
                    .position(|other| {
                        matches!(
                            other,
                            SlotParamCell::VariantCell {
                                by_ref: true,
                                vt: other_vt,
                                target: other_target,
                                ..
                            } if other_vt == vt && other_target == target
                        )
                    })
                    .unwrap_or(index),
                _ => index,
            })
            .collect();

        // SAFETY: `owner` is live; always take one primary reference (`Me = owner`) for the call.
        let Some(object) = (unsafe { ObjectRef::from_raw_iunknown_addref(owner.cast()) }) else {
            return E_POINTER;
        };
        let outcome = invoke_object_member(
            owner,
            object,
            member,
            member.invoke_kind,
            &mut values,
            &aliases,
        );

        for (cell_entry, value) in cells.iter().zip(&values) {
            match cell_entry {
                SlotParamCell::VariantCell {
                    by_ref: true,
                    vt,
                    cell,
                    ..
                } => {
                    let prepared = match as_slot_type(value, *vt) {
                        Ok(prepared) => prepared,
                        Err(hr) => return hr,
                    };
                    // SAFETY: the caller's ByRef storage is live and writable for this call.
                    if unsafe { cell.write_through(prepared) }.is_err() {
                        return DISP_E_TYPEMISMATCH;
                    }
                }
                SlotParamCell::SafeArrayRef {
                    by_ref: true,
                    target,
                } => {
                    if let Some(arr) = value.as_safearray() {
                        let new_ptr = arr.clone().into_raw_safearray_ptr();
                        // SAFETY: caller supplied a live writable `SAFEARRAY**`.
                        unsafe {
                            let old = *(*target);
                            *(*target) = new_ptr;
                            if !old.is_null() {
                                let _ = oleaut::safe_array_destroy(old);
                            }
                        }
                    }
                }
                _ => {}
            }
        }

        match outcome {
            Ok(value) => {
                if let Some(target) = result {
                    match target {
                        SlotReturnTarget::VariantCell { vt, out } => {
                            let prepared = match as_slot_type(&value, vt) {
                                Ok(prepared) => prepared,
                                Err(hr) => return hr,
                            };
                            // SAFETY: `out` is writable storage of type `vt`, initialised above.
                            let cell = unsafe { Variant::by_ref(vt, out) };
                            // SAFETY: same storage.
                            if unsafe { cell.write_through(prepared) }.is_err() {
                                return DISP_E_TYPEMISMATCH;
                            }
                        }
                        SlotReturnTarget::SafeArrayOut { out } => {
                            let Some(arr) = value.as_safearray() else {
                                return DISP_E_TYPEMISMATCH;
                            };
                            // SAFETY: `out` is a writable `SAFEARRAY**` out-slot.
                            unsafe {
                                *out = arr.clone().into_raw_safearray_ptr();
                            }
                        }
                    }
                }
                S_OK
            }
            Err(fault) => {
                publish_member_fault_error_info(interface_iid, &fault);
                vba_error_to_hresult(fault.number)
            }
        }
    }
}
