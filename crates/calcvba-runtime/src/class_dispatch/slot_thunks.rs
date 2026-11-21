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
//! The Win64 stub table (`global_asm!`) is compiled on `windows && x86_64`. The marshalling and
//! dispatch engine below is portable across every platform (and runs under Miri via
//! [`invoke_early_bound_slot_raw`]), using [`calcvba_automation::oleaut`] and
//! [`calcvba_automation::automation_record`].
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
use dispatch_impl::slot_dispatch;

pub use dispatch_impl::invoke_early_bound_slot_raw;

mod dispatch_impl {
    use core::ffi::c_void;
    use std::panic::{AssertUnwindSafe, catch_unwind};
    use std::sync::Arc;

    use calcvba_automation::automation_record::{
        automation_safearray_to_vba_record_array, automation_variant_to_runtime_variant,
        write_runtime_variant_to_automation_cell, write_vba_record_array_to_automation_slot,
    };
    use calcvba_automation::com::{
        DISP_E_TYPEMISMATCH, E_FAIL, E_NOTIMPL, E_POINTER, HResult, IID_IDISPATCH,
        RPC_E_WRONG_THREAD, S_OK,
    };
    use calcvba_automation::oleaut::{
        safe_array_copy_data, safe_array_destroy, safe_array_features, safe_array_is_record_array,
    };
    use calcvba_automation::safe_array::{
        FADF_FIXEDSIZE_VALUE, FADF_STATIC_VALUE, SafeArray, VT_VARIANT_VALUE as VT_VARIANT,
    };

    use super::super::error_info::{clear_thread_error_info, publish_member_fault_error_info};
    use super::super::implements::{ImplementedInterfaceFace, is_impl_query_interface};
    use super::super::{
        coerce_for_byref_writeback, copy_dispatch_arg, current_thread_id, invoke_object_member,
        vba_error_to_hresult,
    };
    use crate::object_ref::{
        COMPAT_OBJECT_VTBL, CompatObjectBase, ObjectRef, RawRuntimeIUnknown,
        RuntimeClassDescriptor, RuntimeInterfaceDescriptor, RuntimeMemberBacking,
        RuntimeMemberDescriptor, RuntimeMemberInvokeKind, RuntimeValueType,
    };
    use crate::{VarType, Variant, VbaRecord, VbaRecordLayout, coerce};

    const VT_DECIMAL: u16 = VarType::Decimal as u16;
    const VT_RECORD: u16 = VarType::Record as u16;
    const VT_ARRAY: u16 = 0x2000;
    const FADF_AUTO_VALUE: u16 = 0x0001;
    const FADF_EMBEDDED_VALUE: u16 = 0x0004;
    const FADF_FIXED_STORAGE: u16 =
        FADF_AUTO_VALUE | FADF_STATIC_VALUE | FADF_EMBEDDED_VALUE | FADF_FIXEDSIZE_VALUE;

    #[repr(C, align(8))]
    struct WireDecimal {
        reserved: u16,
        scale_sign: u16,
        hi: u32,
        lo: u32,
        mid: u32,
    }

    /// Called by the Win64 common stub. `args[0]` is `this`, `args[n]` the n-th argument's 8-byte
    /// slot; `xmm[n - 1]` holds a floating-point argument in position `n` (1..=3).
    #[cfg(all(target_os = "windows", target_arch = "x86_64"))]
    pub(super) unsafe extern "system" fn slot_dispatch(
        slot: u32,
        args: *const u64,
        xmm: *const u64,
    ) -> HResult {
        // SAFETY: forwarded from Win64 stub.
        unsafe { invoke_early_bound_slot_raw(slot as usize, args, xmm) }
    }

    /// Portable entry point into the early-bound vtable slot marshaller used by Win64 slot thunks.
    ///
    /// `args[0]` holds `this` (interface pointer exposed as `u64`); `args[1..=arity]` hold the
    /// ABI argument slots (`ByRef` pointers, `SAFEARRAY**`, `VARIANT*`, or scalar values);
    /// `args[arity + 1]` holds the `[out, retval]` pointer when the member returns a value.
    /// `xmm` points to at least 3 `u64` slots for ByVal float arguments in positions 1..=3.
    ///
    /// # Safety
    /// `args` and `xmm` must point to live, properly laid-out call arguments matching the member
    /// at `slot`.
    pub unsafe fn invoke_early_bound_slot_raw(
        slot: usize,
        args: *const u64,
        xmm: *const u64,
    ) -> HResult {
        // A panic must not unwind into the foreign caller.
        catch_unwind(AssertUnwindSafe(|| {
            // SAFETY: the caller passes the spilled argument arrays of a live call.
            unsafe { dispatch(slot, args, xmm) }
        }))
        .unwrap_or(E_FAIL)
    }

    /// Automation storage tag for a slot's declared type.
    fn slot_vt(ty: RuntimeValueType) -> u16 {
        match ty {
            RuntimeValueType::Integer => 0x02,
            RuntimeValueType::Long => 0x03,
            RuntimeValueType::Single => 0x04,
            RuntimeValueType::Double => 0x05,
            RuntimeValueType::Currency => 0x06,
            RuntimeValueType::Date => 0x07,
            RuntimeValueType::String => 0x08,
            RuntimeValueType::Object => 0x09,
            RuntimeValueType::Boolean => 0x0B,
            RuntimeValueType::Variant => VT_VARIANT,
            RuntimeValueType::Decimal => VT_DECIMAL,
            RuntimeValueType::Byte => 0x11,
            RuntimeValueType::LongLong | RuntimeValueType::LongPtr => 0x14,
            RuntimeValueType::Array => VT_ARRAY,
            RuntimeValueType::Record => VT_RECORD,
        }
    }

    fn slot_record_layout(
        descriptor: &'static RuntimeClassDescriptor,
        member: &RuntimeMemberDescriptor,
        vt: u16,
        parameter: Option<usize>,
    ) -> Result<Option<&'static Arc<VbaRecordLayout>>, HResult> {
        match vt {
            VT_RECORD => {
                let RuntimeMemberBacking::Proc(procedure) = member.backing else {
                    return Err(E_NOTIMPL);
                };
                descriptor
                    .record_member_layout(procedure, parameter, false)
                    .map(Some)
                    .ok_or(E_NOTIMPL)
            }
            VT_ARRAY => {
                let RuntimeMemberBacking::Proc(procedure) = member.backing else {
                    return Ok(None);
                };
                Ok(descriptor.record_member_layout(procedure, parameter, true))
            }
            _ => Ok(None),
        }
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

    unsafe fn read_slot(
        descriptor: &'static RuntimeClassDescriptor,
        target: *mut u8,
        vt: u16,
        record: Option<&Arc<VbaRecordLayout>>,
    ) -> Result<Variant, HResult> {
        let resolve = |guid: &crate::Guid, size: u32| descriptor.resolve_project_record(guid, size);
        if vt == VT_RECORD {
            let layout = record.ok_or(DISP_E_TYPEMISMATCH)?;
            // SAFETY: the supplied record layout describes the caller's live record payload.
            let rec = unsafe { VbaRecord::clone_from_raw(target, Arc::clone(layout)) }
                .map_err(|_| DISP_E_TYPEMISMATCH)?;
            return Ok(Variant::from_vba_record(rec));
        }
        if vt == VT_ARRAY {
            // SAFETY: a ByRef array or a ParamArray has one live SAFEARRAY* slot.
            let array = unsafe { target.cast::<*mut c_void>().read() };
            if let Some(layout) = record {
                // Nominal UDT array (`items() As MyType`): validates FADF_RECORD, IRecordInfo,
                // GetSize, and nominal GetGuid BEFORE copying any data.
                // SAFETY: `array` is null or a caller-owned live `SAFEARRAY*`.
                return unsafe {
                    automation_safearray_to_vba_record_array(array, layout, true, &resolve)
                };
            }
            if array.is_null() {
                return Ok(Variant::unallocated_array(VT_VARIANT));
            }
            // Reject `SAFEARRAY(VT_RECORD)` passed to an intrinsic array parameter before copying.
            // SAFETY: `array` is a live `SAFEARRAY*`.
            if unsafe { safe_array_is_record_array(array) } {
                return Err(DISP_E_TYPEMISMATCH);
            }
            // SAFETY: the caller retains the array for the synchronous copy below.
            let view = unsafe { SafeArray::borrow_raw(array, None) }.ok_or(E_POINTER)?;
            let element = view.element_vartype();
            if !SafeArray::supports_intrinsic_element_vartype(element) {
                return Err(DISP_E_TYPEMISMATCH);
            }
            // SAFETY: target is the caller's live `SAFEARRAY*` slot.
            let core = unsafe { Variant::by_ref(VT_ARRAY | element, target) }.core();
            // SAFETY: `core` holds `VT_BYREF | VT_ARRAY | element` pointing to `target`.
            return unsafe {
                automation_variant_to_runtime_variant((&raw const core).cast(), None, &resolve)
            };
        }
        if vt == VT_VARIANT {
            // SAFETY: `target` is a live Automation `VARIANT` cell.
            return unsafe { automation_variant_to_runtime_variant(target.cast(), None, &resolve) };
        }
        if vt == VT_DECIMAL {
            // SAFETY: standalone DECIMAL storage matches `WireDecimal`.
            let wire = unsafe { target.cast::<WireDecimal>().read() };
            return Ok(Variant::from_decimal96(crate::Decimal96::from_scale_sign(
                wire.lo,
                wire.mid,
                wire.hi,
                wire.scale_sign,
            )));
        }
        // SAFETY: target is the caller's live storage with the ABI shape described by vt.
        let core = unsafe { Variant::by_ref(vt, target) }.core();
        // SAFETY: core holds a reference to that same storage.
        unsafe { copy_dispatch_arg(&raw const core) }
    }

    unsafe fn write_slot(
        target: *mut u8,
        vt: u16,
        record: Option<&Arc<VbaRecordLayout>>,
        value: &Variant,
    ) -> Result<(), HResult> {
        if vt == VT_RECORD {
            let layout = record.ok_or(DISP_E_TYPEMISMATCH)?;
            let new = value.as_vba_record().ok_or(DISP_E_TYPEMISMATCH)?;
            if !new.layout().matches_nominal(layout) {
                return Err(DISP_E_TYPEMISMATCH);
            }
            // SAFETY: the target has the declared layout. Dereferenced arguments are live and
            // the retval storage was initialized to an empty record before entering VBA.
            unsafe {
                VbaRecord::drop_raw(target, layout);
                new.move_into_raw(target);
            }
            return Ok(());
        }
        if vt == VT_ARRAY {
            if let Some(layout) = record {
                // SAFETY: caller owns this initialized `SAFEARRAY*` slot (`ByRef` or `[out, retval]`).
                return unsafe {
                    write_vba_record_array_to_automation_slot(
                        target.cast::<*mut c_void>(),
                        value,
                        layout,
                    )
                };
            }
            if value.vtype() != VarType::ArrayVariant
                || value
                    .array_element_vartype()
                    .is_some_and(|element| !SafeArray::supports_intrinsic_element_vartype(element))
            {
                return Err(DISP_E_TYPEMISMATCH);
            }
            let array = value.try_as_safearray().map_err(|_| DISP_E_TYPEMISMATCH)?;
            let slot = target.cast::<*mut c_void>();
            // SAFETY: the caller owns this initialized SAFEARRAY* slot.
            let previous = unsafe { slot.read() };
            // SAFETY: a non-null previous pointer is live until replaced.
            let old = unsafe { SafeArray::borrow_raw(previous, None) };
            if let (Some(old), Some(new)) = (old, array.as_ref())
                && old.element_vartype() == new.element_vartype()
                && old.bounds() == new.bounds()
            {
                // Update the original descriptor's data, including for a fixed VBA array.
                // SAFETY: both descriptors are live, type and bounds match, and the caller
                // lent the destination for writes for this call.
                let status = unsafe { safe_array_copy_data(new.raw_safearray_ptr(), previous) };
                return if status < 0 { Err(status) } else { Ok(()) };
            }
            if !previous.is_null() {
                // SAFETY: a changed dynamic array transfers its old descriptor to OleAut.
                // Fixed arrays cannot change shape and must keep the original descriptor.
                if unsafe { safe_array_features(previous) } & FADF_FIXED_STORAGE != 0 {
                    return Err(DISP_E_TYPEMISMATCH);
                }
                // SAFETY: the caller's slot owns the old descriptor. On failure it stays live.
                let status = unsafe { safe_array_destroy(previous) };
                if status < 0 {
                    return Err(status);
                }
            }
            let raw = array.map_or(core::ptr::null_mut(), SafeArray::into_raw_safearray_ptr);
            // SAFETY: the slot is writable and takes ownership of the replacement.
            unsafe { slot.write(raw) };
            return Ok(());
        }
        if vt == VT_VARIANT {
            // SAFETY: `target` is a live initialized `VARIANT` cell (`ByRef` or `[out, retval]`).
            return unsafe { write_runtime_variant_to_automation_cell(target.cast(), value) };
        }

        let prepared = as_slot_type(value, vt)?;
        if vt == VT_DECIMAL {
            let decimal = prepared.as_decimal96().ok_or(DISP_E_TYPEMISMATCH)?;
            // SAFETY: standalone DECIMAL storage is writable and has no owned payload.
            unsafe {
                target.cast::<WireDecimal>().write(WireDecimal {
                    reserved: 0,
                    scale_sign: decimal.scale_sign,
                    hi: decimal.hi,
                    lo: decimal.lo,
                    mid: decimal.mid,
                });
            }
            return Ok(());
        }
        // SAFETY: target is a writable initialized cell of the declared ABI type.
        unsafe { Variant::by_ref(vt, target).write_through(prepared) }
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

    unsafe fn dispatch(slot: usize, args: *const u64, xmm: *const u64) -> HResult {
        if args.is_null() {
            return E_POINTER;
        }
        // SAFETY: position 0 is `this`.
        let this = core::ptr::with_exposed_provenance_mut::<c_void>(unsafe { *args } as usize);
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
        // SAFETY: the primary owner remains live through this synchronous call.
        let descriptor = unsafe { (*owner).class_descriptor };
        clear_thread_error_info();
        let Some(member) = interface.and_then(|iface| member_for_slot(iface, slot)) else {
            return E_NOTIMPL;
        };

        // Each argument as a caller-storage pointer.
        let mut cells = Vec::with_capacity(member.params.len());
        let mut values = Vec::with_capacity(member.params.len());
        for (index, param) in member.params.iter().enumerate() {
            let position = index + 1;
            let vt = if param.param_array {
                VT_ARRAY
            } else {
                slot_vt(param.value_type)
            };
            let record = match slot_record_layout(descriptor, member, vt, Some(index)) {
                Ok(record) => record,
                Err(hr) => return hr,
            };
            let raw = if !param.by_ref && is_float(vt) && position <= 3 {
                if xmm.is_null() {
                    return E_POINTER;
                }
                // SAFETY: the stub spilled xmm1..=xmm3.
                unsafe { xmm.add(position - 1) }
            } else {
                // SAFETY: the caller passed at least `params.len()` arguments after `this`.
                unsafe { args.add(position) }
            };
            let indirect = matches!(vt, VT_VARIANT | VT_DECIMAL)
                || param.param_array
                || (vt == VT_RECORD
                    && record.is_some_and(|layout| !matches!(layout.size(), 1 | 2 | 4 | 8)));
            let target = if param.by_ref || indirect {
                // SAFETY: the slot holds the caller's pointer to the value.
                core::ptr::with_exposed_provenance_mut::<u8>(unsafe { *raw } as usize)
            } else {
                raw.cast_mut().cast::<u8>()
            };
            if target.is_null() {
                return E_POINTER;
            }
            // SAFETY: target is the caller's live storage of type vt for this call.
            match unsafe { read_slot(descriptor, target, vt, record) } {
                Ok(value) => values.push(value),
                Err(hr) => return hr,
            }
            cells.push((param.by_ref, vt, target, record));
        }

        let returns = matches!(
            member.invoke_kind,
            RuntimeMemberInvokeKind::Method | RuntimeMemberInvokeKind::PropertyGet
        ) && member.return_type.is_some();
        let result = if returns {
            let vt = slot_vt(member.return_type.expect("member has a return type"));
            let record = match slot_record_layout(descriptor, member, vt, None) {
                Ok(record) => record,
                Err(hr) => return hr,
            };
            // SAFETY: the retval pointer follows the declared arguments.
            let out = core::ptr::with_exposed_provenance_mut::<u8>(unsafe {
                *args.add(member.params.len() + 1)
            } as usize);
            if out.is_null() {
                return E_POINTER;
            }
            // An [out] slot holds nothing yet: give owned types an empty value to replace.
            // SAFETY: `out` is the caller's writable storage of type `vt`.
            unsafe {
                match vt {
                    0x08 | 0x09 | VT_ARRAY => {
                        out.cast::<*mut c_void>().write(core::ptr::null_mut())
                    }
                    VT_VARIANT => out.cast::<Variant>().write(Variant::empty()),
                    VT_RECORD => out.write_bytes(0, record.expect("record has a layout").size()),
                    _ => {}
                }
            }
            Some((vt, out, record))
        } else {
            None
        };

        // ByRef arguments over the same storage are one variable.
        let aliases: Vec<usize> = (0..cells.len())
            .map(|index| {
                let (by_ref, vt, target, record) = &cells[index];
                if !*by_ref {
                    return index;
                }
                cells[..index]
                    .iter()
                    .position(|(other_ref, other_vt, other_target, other_record)| {
                        *other_ref
                            && other_vt == vt
                            && other_target == target
                            && other_record == record
                    })
                    .unwrap_or(index)
            })
            .collect();

        // SAFETY: the owner is live. The method receives the canonical object as Me.
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

        for (index, (&(by_ref, vt, target, record), value)) in cells.iter().zip(&values).enumerate()
        {
            if !by_ref || aliases[index] != index {
                continue;
            }
            // SAFETY: the caller's ByRef storage is live and writable for this call.
            if let Err(hr) = unsafe { write_slot(target, vt, record, value) } {
                return hr;
            }
        }

        match outcome {
            Ok(value) => {
                if let Some((vt, out, record)) = result {
                    // SAFETY: `out` is writable storage of type `vt`, initialised above.
                    if let Err(hr) = unsafe { write_slot(out, vt, record, &value) } {
                        return hr;
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

    #[cfg(test)]
    #[test]
    fn standalone_decimal_uses_the_decimal_header() {
        let value = Variant::from_decimal96(crate::Decimal96::from_parts(1, 2, 3, 4, true));
        let mut wire = core::mem::MaybeUninit::<WireDecimal>::uninit();
        let target = wire.as_mut_ptr().cast();
        // SAFETY: target points to writable DECIMAL storage and write_slot initializes it.
        unsafe {
            write_slot(target, VT_DECIMAL, None, &value).unwrap();
            assert_eq!((*wire.as_ptr()).reserved, 0);
            assert_eq!(
                read_slot(
                    &crate::object_ref::COMPAT_OBJECT_CLASS_DESCRIPTOR,
                    target,
                    VT_DECIMAL,
                    None
                )
                .unwrap(),
                value
            );
        }
    }
}
