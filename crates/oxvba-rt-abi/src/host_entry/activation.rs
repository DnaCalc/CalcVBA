//! Owning physical storage and exact continuation position for typed execution.
//! Global storage is the session's canonical Variant store; there is no mirror.
use super::{MAX_TYPED_SLOTS, ScalarBits, ScalarLayout, ScalarType};
use calcvba_host_api::{HostObject, Slot, Value, ValueSlot};
use oxvba_runtime::Variant;
use std::{ops::Deref, rc::Rc};

pub const SUSPENDED: i32 = -2;
/// Return to the shared activation-stack driver for a call/error-state operation.
pub const CONTROL: i32 = -3;
pub const FINISHED_PC: u32 = u32::MAX;
const SCRATCH_CAPACITY: usize = 16;
/// Private generated-code ABI. Every nonnull object slot owns one reference.
#[repr(C)]
pub struct RawActivation {
    pub slots: [ScalarBits; MAX_TYPED_SLOTS],
    pub pc: u32,
    pub global_base: usize,
    pub remaining_operations: u64,
    aliases: [Alias; calcvba_host_api::MAX_TYPED_ARGS],
    globals: *mut Variant,
    global_len: usize,
    pub(crate) layout: *const ScalarLayout,
    variants: *mut Variant,
    variant_len: usize,
    scratch: [Scratch; SCRATCH_CAPACITY],
    scratch_len: usize,
    pub(super) param_arrays: std::collections::HashMap<usize, Vec<Option<Alias>>>,
}
#[derive(Clone, Copy)]
pub(super) struct Alias {
    pub(super) frame: *mut RawActivation,
    pub(super) index: usize,
}
impl Alias {
    const NONE: Self = Self {
        frame: std::ptr::null_mut(),
        index: 0,
    };
}
pub struct Activation<'a> {
    raw: RawActivation,
    layout: FrameLayout<'a>,
}
enum FrameLayout<'a> {
    Borrowed(&'a ScalarLayout),
    Owned(Rc<ScalarLayout>),
}
impl Deref for FrameLayout<'_> {
    type Target = ScalarLayout;
    fn deref(&self) -> &ScalarLayout {
        match self {
            Self::Borrowed(layout) => layout,
            Self::Owned(layout) => layout,
        }
    }
}
#[derive(Clone, Copy)]
struct Scratch {
    ty: ScalarType,
    bits: ScalarBits,
}
impl Scratch {
    const EMPTY: Self = Self {
        ty: ScalarType::Void,
        bits: ScalarBits::ZERO,
    };
}
impl<'a> Activation<'a> {
    /// # Safety
    /// globals names global_len initialized canonical session cells. The owner
    /// must keep them stable and live until this activation drops; reset and
    /// competing execution must be excluded. No Rust reference spans a callback.
    /// layout must retain the checked admission facts, unmodified, for the
    /// function that will execute this frame (including parameter/slot kinds).
    pub unsafe fn new(
        layout: &ScalarLayout,
        args: &[Slot<'_>],
        globals: *mut Variant,
        global_len: usize,
        entry_pc: u32,
    ) -> Result<Self, i32> {
        // SAFETY: forwards the complete retained-storage contract.
        unsafe {
            Self::build(
                std::borrow::Cow::Owned(layout.clone()),
                args,
                globals,
                global_len,
                entry_pc,
                true,
            )
        }
    }
    /// # Safety
    /// Same contract as new; layout additionally spans the returned frame.
    pub unsafe fn borrowed(
        layout: &'a ScalarLayout,
        args: &[Slot<'_>],
        globals: *mut Variant,
        global_len: usize,
        entry_pc: u32,
    ) -> Result<Self, i32> {
        // SAFETY: forwards the complete retained-storage contract.
        unsafe {
            Self::build(
                std::borrow::Cow::Borrowed(layout),
                args,
                globals,
                global_len,
                entry_pc,
                true,
            )
        }
    }
    /// # Safety
    /// Same contract as borrowed. This internal exact-carrier adapter permits
    /// nonfinite numeric payloads already present in the universal VBA carrier;
    /// it must not be used to bypass typed host argument preparation.
    pub unsafe fn borrowed_exact(
        layout: &'a ScalarLayout,
        args: &[Slot<'_>],
        entry_pc: u32,
    ) -> Result<Self, i32> {
        // SAFETY: this route admits no global storage and retains any arguments.
        unsafe {
            Self::build(
                std::borrow::Cow::Borrowed(layout),
                args,
                std::ptr::null_mut(),
                0,
                entry_pc,
                false,
            )
        }
    }
    /// # Safety
    /// Same retained canonical-store contract as new, with extended typed values.
    pub unsafe fn new_values(
        layout: &ScalarLayout,
        args: &[ValueSlot<'_>],
        globals: *mut Variant,
        global_len: usize,
        entry_pc: u32,
    ) -> Result<Activation<'static>, i32> {
        // SAFETY: forwards the complete owning frame contract.
        unsafe {
            Activation::build_values(
                std::borrow::Cow::Owned(layout.clone()),
                args,
                globals,
                global_len,
                entry_pc,
                true,
            )
        }
    }
    unsafe fn build(
        layout: std::borrow::Cow<'a, ScalarLayout>,
        args: &[Slot<'_>],
        globals: *mut Variant,
        global_len: usize,
        entry_pc: u32,
        finite_inputs: bool,
    ) -> Result<Self, i32> {
        if args.len() > calcvba_host_api::MAX_TYPED_ARGS {
            return Err(13);
        }
        let mut values = [ValueSlot::Long(0); calcvba_host_api::MAX_TYPED_ARGS];
        for (i, arg) in args.iter().enumerate() {
            values[i] = (*arg).into();
        }
        // SAFETY: converted slots borrow the same retained caller arguments.
        unsafe {
            Self::build_values(
                layout,
                &values[..args.len()],
                globals,
                global_len,
                entry_pc,
                finite_inputs,
            )
        }
    }
    unsafe fn build_values(
        layout: std::borrow::Cow<'a, ScalarLayout>,
        args: &[ValueSlot<'_>],
        globals: *mut Variant,
        global_len: usize,
        entry_pc: u32,
        finite_inputs: bool,
    ) -> Result<Self, i32> {
        let layout = match layout {
            std::borrow::Cow::Borrowed(layout) => FrameLayout::Borrowed(layout),
            std::borrow::Cow::Owned(layout) => FrameLayout::Owned(Rc::new(layout)),
        };
        if args.len() != layout.parameters.len()
            || args
                .iter()
                .zip(&layout.parameters)
                .any(|(a, t)| a.ty() != *t)
        {
            return Err(13);
        }
        if finite_inputs && args.iter().any(|a| !a.is_valid()) {
            return Err(6);
        }
        for (arg, iid) in args.iter().zip(&layout.parameter_iids) {
            if let Some(iid) = iid {
                let pointer = arg.bits().as_object();
                // SAFETY: checked object carrier; the input borrow owns a live
                // factory table, or is Nothing. Reject before retaining/entry.
                // SAFETY: input is an exact retained runtime IUnknown carrier.
                let owner = unsafe {
                    calcvba_host_api::value::retain_bits(ScalarType::HostObject, arg.bits())
                }?;
                if !pointer.is_null() && !owner.has_interface(*iid) {
                    return Err(13);
                }
            }
        }
        for (index, arg) in args.iter().enumerate() {
            if let Some(required) = &layout.parameter_classes[index] {
                // SAFETY: checked exact object input, retained through validation.
                let value = unsafe {
                    calcvba_host_api::value::retain_bits(ScalarType::HostObject, arg.bits())
                }?;
                match value {
                    Value::Object(None) => {}
                    Value::RuntimeObject(object) if super::class_matches(&object, required) => {}
                    _ => return Err(13),
                }
            }
            if let ValueSlot::Record(record) = arg
                && layout
                    .record_slots
                    .get(&index)
                    .is_none_or(|expected| expected.as_ref() != record.layout().as_ref())
            {
                return Err(13);
            }
        }
        if layout.slots.len() - layout.global_base > global_len {
            return Err(5);
        }
        let (variants, variant_len) = if layout.slots.contains(&ScalarType::Variant) {
            let count = layout.slots[..layout.global_base]
                .iter()
                .filter(|ty| **ty == ScalarType::Variant)
                .count();
            let cells = (0..count + SCRATCH_CAPACITY)
                .map(|_| Variant::empty())
                .collect::<Vec<_>>()
                .into_boxed_slice();
            let len = cells.len();
            (Box::into_raw(cells).cast::<Variant>(), len)
        } else {
            (std::ptr::null_mut(), 0)
        };
        let layout_pointer: *const ScalarLayout = &*layout;
        let mut frame = Self {
            raw: RawActivation {
                slots: [ScalarBits::ZERO; MAX_TYPED_SLOTS],
                pc: entry_pc,
                global_base: layout.global_base,
                remaining_operations: 0,
                aliases: [Alias::NONE; calcvba_host_api::MAX_TYPED_ARGS],
                globals,
                global_len,
                layout: layout_pointer,
                variants,
                variant_len,
                scratch: [Scratch::EMPTY; SCRATCH_CAPACITY],
                scratch_len: 0,
                param_arrays: Default::default(),
            },
            layout,
        };
        let mut variant_index = 0;
        for i in 0..frame.layout.global_base {
            if frame.layout.slots[i] == ScalarType::Variant {
                // SAFETY: the frame owns initialized stable cells for every index.
                frame.raw.slots[i] =
                    ScalarBits::object(unsafe { variants.add(variant_index) }.cast());
                variant_index += 1;
            }
        }
        for (i, arg) in args.iter().enumerate() {
            let bits = arg.bits();
            // SAFETY: args retain all object borrows and this frame owns storage.
            unsafe {
                write_slot(&mut frame.raw, i, arg.ty(), &bits)?;
            }
        }
        for (&index, record_layout) in &frame.layout.record_slots {
            if index < frame.layout.global_base && index >= args.len() {
                let record =
                    oxvba_runtime::vba_record::VbaRecord::new_default(record_layout.clone())
                        .map_err(|_| 7)?;
                frame.raw.slots[index] = Value::Record(record).into_owned_bits().map_err(|_| 13)?;
            }
        }
        Ok(frame)
    }
    pub fn layout(&self) -> &ScalarLayout {
        &self.layout
    }
    pub fn raw(&self) -> &RawActivation {
        &self.raw
    }
    /// # Safety
    /// Preserve slot types and exactly one retained reference per object slot;
    /// globals and scratch must continue to satisfy the activation lease.
    pub unsafe fn raw_mut(&mut self) -> &mut RawActivation {
        &mut self.raw
    }
    pub fn pc(&self) -> u32 {
        self.raw.pc
    }
    pub fn set_pc(&mut self, pc: u32) {
        self.raw.pc = pc;
    }
    pub fn result(&self) -> Value {
        if !self.layout.has_result {
            return Value::Empty;
        }
        let bits = self.raw.slots[self.layout.return_slot];
        // SAFETY: this admitted result cell retains the exact carrier.
        unsafe { calcvba_host_api::value::retain_bits(self.layout.result, bits) }
            .expect("admitted result")
    }
    pub fn take_result(&mut self) -> Value {
        if !self.layout.has_result {
            return Value::Empty;
        }
        let index = self.layout.return_slot;
        let ty = self.layout.result;
        if ty == ScalarType::Variant {
            let cell = self.raw.slots[index]
                .as_object()
                .cast_mut()
                .cast::<Variant>();
            // SAFETY: this exact initialized cell belongs to the returning frame.
            return Value::Variant(unsafe { cell.replace(Variant::empty()) });
        }
        let mut bits = std::mem::replace(&mut self.raw.slots[index], ScalarBits::ZERO);
        // SAFETY: the result slot transfers its exact owned native carrier.
        unsafe { calcvba_host_api::value::take_result(ty, &mut bits) }.expect("admitted result")
    }
    pub fn read(
        &mut self,
        operand: &oxvba_oxir::OxOperand,
        to: ScalarType,
    ) -> Result<ScalarBits, i32> {
        // SAFETY: this method uniquely owns the admitted frame during the read.
        unsafe { super::value::read_operand(&mut self.raw, &self.layout, operand, to) }
    }
    pub fn write_value(&mut self, dst: oxvba_oxir::OxPlace, value: &Value) -> Result<(), i32> {
        let index = self.layout.slot(dst).map_err(|_| 5)?;
        if self.layout.slots[index] != value.ty() {
            return Err(13);
        }
        let bits = value.borrow_slot().ok_or(13)?.bits();
        // SAFETY: the owned input value retains the exact admitted carrier.
        unsafe { write_slot(&mut self.raw, index, self.layout.slots[index], &bits) }
    }
    pub fn write_owned_value(&mut self, dst: oxvba_oxir::OxPlace, value: Value) -> Result<(), i32> {
        let index = self.layout.slot(dst).map_err(|_| 5)?;
        // SAFETY: the owned value spans this exact admitted destination transfer.
        unsafe { write_owned(&mut self.raw, index, self.layout.slots[index], value) }
    }
    /// # Safety
    /// value has the admitted destination carrier; object pointers are null or
    /// live factory tables retained through this synchronous assignment.
    pub unsafe fn write(&mut self, dst: oxvba_oxir::OxPlace, value: ScalarBits) -> Result<(), i32> {
        let index = self.layout.slot(dst).expect("admitted destination");
        let ty = self.layout.slots[index];
        // SAFETY: the checked instruction proves the exact physical carrier.
        unsafe { write_slot(&mut self.raw, index, ty, &value) }
    }
    pub fn clear_scratch(&mut self) {
        // SAFETY: raw owns every retained scratch reference and is not running.
        unsafe {
            clear_scratch(&mut self.raw);
        }
    }
    /// Release this frame's own locals without writing through caller aliases.
    /// The result remains retained while Class_Terminate callbacks run.
    pub fn release_locals_except_result(&mut self) {
        self.clear_scratch();
        self.raw.aliases.fill(Alias::NONE);
        self.raw.param_arrays.clear();
        for index in 0..self.layout.global_base {
            if self.layout.has_result && index == self.layout.return_slot {
                continue;
            }
            if self.layout.slots[index] == ScalarType::Record {
                let bits = std::mem::replace(&mut self.raw.slots[index], ScalarBits::ZERO);
                // SAFETY: this exact physical record slot transfers its owned box.
                unsafe {
                    calcvba_host_api::value::release_owned_bits(ScalarType::Record, bits);
                }
            } else {
                // SAFETY: aliases are detached; these are only this frame's owned cells.
                unsafe {
                    clear_slot(&mut self.raw, index, self.layout.slots[index])
                        .expect("retained local");
                }
            }
        }
    }
    pub fn clear_temps(&mut self, from: usize) {
        for i in self.layout.local_count + from..self.layout.global_base {
            // SAFETY: initialized admitted local slots, with no overlapping entry.
            unsafe {
                clear_slot(&mut self.raw, i, self.layout.slots[i]).expect("frame slot");
            }
        }
    }
    /// # Safety
    /// parent has a stable address and outlives this child and all descendants.
    /// The caller keeps the entire stack exclusively running on this thread.
    pub unsafe fn child(
        parent: &mut Activation<'_>,
        layout: &ScalarLayout,
        args: &[oxvba_oxir::OxArg],
        entry_pc: u32,
    ) -> Result<Activation<'static>, i32> {
        // SAFETY: share the parent's leased unit store for an internal call.
        unsafe {
            Self::child_with_globals(
                parent,
                layout,
                args,
                entry_pc,
                parent.raw.globals,
                parent.raw.global_len,
            )
        }
    }
    /// # Safety
    /// Parent and canonical target unit store stay leased through all descendants.
    /// The boxed parent remains stable and uniquely runs under the stack guards.
    pub unsafe fn child_with_globals(
        parent: &mut Activation<'_>,
        layout: &ScalarLayout,
        args: &[oxvba_oxir::OxArg],
        entry_pc: u32,
        globals: *mut Variant,
        global_len: usize,
    ) -> Result<Activation<'static>, i32> {
        use oxvba_oxir::{OxArg, OxOperand};
        if args.len() > layout.parameters.len() {
            return Err(13);
        }
        let mut owners = Vec::with_capacity(args.len());
        for (i, ty) in layout.parameters.iter().enumerate() {
            let arg = args.get(i).unwrap_or(&OxArg::Omitted);
            if matches!(arg, OxArg::Omitted)
                && layout.optional_parameters[i]
                && *ty == ScalarType::Variant
            {
                owners.push(Value::Variant(Variant::from_error_code(
                    0x8002_0004u32 as i32,
                )));
                continue;
            }
            let operand = match arg {
                OxArg::ByVal(op) => op.clone(),
                OxArg::ByRef(p) => OxOperand::Use(*p),
                _ => return Err(13),
            };
            let bits = parent.read(&operand, *ty)?;
            // SAFETY: the parent retains every exact source carrier through cloning.
            owners.push(unsafe { calcvba_host_api::value::retain_bits(*ty, bits) }?);
        }
        let slots = owners
            .iter()
            .map(|value| value.borrow_slot().ok_or(13))
            .collect::<Result<Vec<_>, _>>()?;
        // SAFETY: the child shares the parent's retained canonical-store lease.
        let mut child =
            unsafe { Activation::new_values(layout, &slots, globals, global_len, entry_pc) }?;
        parent.clear_scratch();
        // Mint one ancestor authority after the last mutable parent method.
        // Multiple aliases must reuse it rather than retagging the ancestor.
        let parent_raw = std::ptr::addr_of_mut!(parent.raw);
        for (i, arg) in args.iter().enumerate() {
            if let OxArg::ByVal(OxOperand::Use(place)) = arg {
                let index = parent.layout.slot(*place).map_err(|_| 5)?;
                // SAFETY: all recorded aliases refer to retained ancestors.
                if let Some(aliases) = unsafe { array_aliases(parent_raw, index) } {
                    child.raw.param_arrays.insert(i, aliases);
                }
            }
            if let OxArg::ByRef(place) = arg {
                if !layout.by_ref_parameters[i] {
                    return Err(13);
                }
                // Release the private copy before replacing it with a nonowning alias.
                // SAFETY: exact initialized parameter carrier.
                unsafe { clear_slot(&mut child.raw, i, layout.parameters[i]) }?;
                let index = parent.layout.slot(*place).map_err(|_| 5)?;
                child.raw.aliases[i] = Alias {
                    frame: parent_raw,
                    index,
                };
            }
        }
        Ok(child)
    }
}
impl Drop for Activation<'_> {
    fn drop(&mut self) {
        self.clear_scratch();
        for i in 0..self.layout.global_base {
            let bits = self.raw.slots[i];
            match self.layout.slots[i] {
                ScalarType::Record | ScalarType::Array => {
                    // SAFETY: transfer this slot's exact aggregate owner.
                    unsafe {
                        calcvba_host_api::value::release_owned_bits(self.layout.slots[i], bits)
                    };
                }
                ScalarType::HostObject => {
                    // SAFETY: transfer this slot's one retained reference.
                    unsafe {
                        calcvba_host_api::value::release_owned_bits(ScalarType::HostObject, bits);
                    }
                }
                ScalarType::String => {
                    // SAFETY: transfer this slot's one owned raw BSTR.
                    drop(unsafe {
                        calcvba_host_api::BStr::from_raw_bstr(bits.as_object().cast_mut().cast())
                    });
                }
                _ => {}
            }
        }
        if !self.raw.variants.is_null() {
            let slice = std::ptr::slice_from_raw_parts_mut(self.raw.variants, self.raw.variant_len);
            // SAFETY: consume the exact boxed initialized Variant-cell allocation.
            drop(unsafe { Box::from_raw(slice) });
        }
    }
}
/// # Safety
/// raw is the unique retained frame; value transfers an exact typed carrier.
pub(super) unsafe fn pin_value(raw: *mut RawActivation, value: Value) -> Result<ScalarBits, i32> {
    let ty = value.ty();
    let bits = match value {
        Value::Long(v) => return Ok(ScalarBits::long(v)),
        Value::Double(v) => return Ok(ScalarBits::double(v)),
        Value::Boolean(v) => return Ok(ScalarBits::long(if v { -1 } else { 0 })),
        Value::Empty => return Ok(ScalarBits::ZERO),
        value => {
            // SAFETY: short field access; no borrow spans a host callback.
            let index = unsafe { (*raw).scratch_len };
            if index >= SCRATCH_CAPACITY {
                return Err(5);
            }
            match value {
                value @ (Value::Record(_) | Value::Array(_) | Value::RuntimeObject(_)) => {
                    value.into_owned_bits().map_err(|_| 13)?
                }
                Value::Object(value) => {
                    ScalarBits::object(value.map_or(std::ptr::null(), HostObject::into_header))
                }
                Value::String(value) => ScalarBits::object(value.into_raw_bstr().cast()),
                Value::Variant(value) => {
                    // Literal Empty/Null arguments can require Variant scratch
                    // even when the caller declares no Variant physical slots.
                    // Allocate only when such a value actually needs pinning;
                    // existing physical Variant cells are never moved.
                    // SAFETY: this uniquely executing frame owns these fields.
                    let mut cells = unsafe { (*raw).variants };
                    if cells.is_null() {
                        let storage = (0..SCRATCH_CAPACITY)
                            .map(|_| Variant::empty())
                            .collect::<Vec<_>>()
                            .into_boxed_slice();
                        cells = Box::into_raw(storage).cast::<Variant>();
                        // SAFETY: publish one initialized allocation to its owner.
                        unsafe {
                            (*raw).variants = cells;
                            (*raw).variant_len = SCRATCH_CAPACITY;
                        }
                    }
                    // SAFETY: fixed scratch index is inside the retained allocation.
                    let cell = unsafe { cells.add((*raw).variant_len - SCRATCH_CAPACITY + index) };
                    // SAFETY: install before dropping the previous owner.
                    drop(unsafe { cell.replace(value) });
                    ScalarBits::object(cell.cast())
                }
                _ => unreachable!(),
            }
        }
    };
    // SAFETY: unique frame and checked initialized scratch extent.
    unsafe {
        (*raw).scratch[(*raw).scratch_len] = Scratch { ty, bits };
        (*raw).scratch_len += 1;
    }
    Ok(bits)
}
/// # Safety
/// raw follows Activation::new's lease and exclusive running contract.
pub unsafe fn clear_scratch(raw: *mut RawActivation) {
    loop {
        // SAFETY: brief access ends before releasing any removed owner.
        let entry = unsafe {
            if (*raw).scratch_len == 0 {
                break;
            }
            (*raw).scratch_len -= 1;
            std::ptr::replace(
                std::ptr::addr_of_mut!((*raw).scratch[(*raw).scratch_len]),
                Scratch::EMPTY,
            )
        };
        match entry.ty {
            ScalarType::Record | ScalarType::Array => {
                // SAFETY: transfer one removed exact aggregate scratch owner.
                unsafe { calcvba_host_api::value::release_owned_bits(entry.ty, entry.bits) };
            }
            ScalarType::HostObject => {
                // SAFETY: transfer one removed scratch reference.
                unsafe {
                    calcvba_host_api::value::release_owned_bits(ScalarType::HostObject, entry.bits);
                }
            }
            ScalarType::String => {
                // SAFETY: transfer one removed scratch BSTR.
                drop(unsafe {
                    calcvba_host_api::BStr::from_raw_bstr(entry.bits.as_object().cast_mut().cast())
                });
            }
            ScalarType::Variant => {
                let cell = entry.bits.as_object().cast_mut().cast::<Variant>();
                // SAFETY: clear the initialized owned inline scratch cell.
                drop(unsafe { cell.replace(Variant::empty()) });
            }
            _ => {}
        }
    }
}
/// # Safety
/// raw follows Activation::new; ty is the declared type at global index.
pub unsafe fn read_global(
    raw: *mut RawActivation,
    index: usize,
    ty: ScalarType,
) -> Result<ScalarBits, i32> {
    // SAFETY: read-only pointers under the retained canonical-store lease.
    let (globals, len) = unsafe { ((*raw).globals, (*raw).global_len) };
    if index >= len || globals.is_null() {
        return Err(5);
    }
    let value = {
        // SAFETY: no reference into canonical storage survives this scope.
        let value = unsafe { &*globals.add(index) };
        match ty {
            ScalarType::Long => return value.as_i32().map(ScalarBits::long).ok_or(13),
            ScalarType::Double => return value.as_f64().map(ScalarBits::double).ok_or(13),
            ScalarType::Boolean => {
                return value
                    .as_bool()
                    .map(|v| ScalarBits::long(if v { -1 } else { 0 }))
                    .ok_or(13);
            }
            ScalarType::String => Value::String(value.as_bstr().ok_or(13)?),
            ScalarType::Variant => Value::Variant(value.clone()),
            ScalarType::Record => Value::Record(value.as_vba_record().ok_or(13)?),
            ScalarType::Array => {
                if value.vtype() != oxvba_runtime::VarType::ArrayVariant {
                    return Err(13);
                }
                Value::Array(value.as_safearray())
            }
            ScalarType::HostObject => {
                let object = value.as_object_ref();
                if object.is_none() && value.vtype() != oxvba_runtime::VarType::Object {
                    return Err(13);
                }
                super::value::from_variant(ScalarType::HostObject, value.clone())?
            }
            _ => return Err(13),
        }
    };
    // SAFETY: move the snapshot into stable frame scratch before any callback.
    unsafe { pin_value(raw, value) }
}
/// # Safety
/// raw follows the retained activation lease; index is an admitted slot. The
/// returned address is an identity only and grants no dereference authority.
pub(super) unsafe fn slot_identity(raw: *mut RawActivation, index: usize) -> Result<usize, i32> {
    // SAFETY: brief read of the admitted frame; aliases retain live ancestors.
    let frame = unsafe { &*raw };
    if let Some(alias) = frame.aliases.get(index)
        && !alias.frame.is_null()
    {
        // SAFETY: boxed ancestor activations outlive descendants.
        return unsafe { slot_identity(alias.frame, alias.index) };
    }
    if index >= frame.global_base {
        let index = index - frame.global_base;
        if index >= frame.global_len || frame.globals.is_null() {
            return Err(5);
        }
        // SAFETY: in-bounds canonical storage, used only for identity comparison.
        return Ok(unsafe { frame.globals.add(index) }.addr());
    }
    frame
        .slots
        .get(index)
        .map(|slot| std::ptr::from_ref(slot).addr())
        .ok_or(5)
}
/// # Safety
/// raw follows Activation::new; index/ty are checked physical layout facts;
/// bits contains that carrier, with any object borrowed through the operation.
pub unsafe fn write_slot(
    raw: *mut RawActivation,
    index: usize,
    ty: ScalarType,
    bits: *const ScalarBits,
) -> Result<(), i32> {
    // SAFETY: aliases point only to live, stable ancestor frames and carry the
    // same admitted type. No Rust reference into an ancestor spans recursion.
    let alias = unsafe { (&*raw).aliases.get(index).copied() }.unwrap_or(Alias::NONE);
    if !alias.frame.is_null() {
        // SAFETY: the stack owner retains the ancestor until this frame drops.
        return unsafe { write_slot(alias.frame, alias.index, ty, bits) };
    }
    // SAFETY: caller provides one initialized readable physical carrier.
    let bits = unsafe { *bits };
    if ty == ScalarType::Array || (ty == ScalarType::Record && !bits.as_object().is_null()) {
        // SAFETY: retain the exact borrowed aggregate before validation/store.
        let value = unsafe { calcvba_host_api::value::retain_bits(ty, bits) }?;
        // SAFETY: write_owned validates the target source layout before mutation.
        return unsafe { write_owned(raw, index, ty, value) };
    }
    let (base, globals, len) = {
        // SAFETY: brief frame borrow ends before any replaced owner is released.
        let frame = unsafe { &*raw };
        (frame.global_base, frame.globals, frame.global_len)
    };
    if index >= base {
        let index = index - base;
        if index >= len || globals.is_null() {
            return Err(5);
        }
        let value = match ty {
            ScalarType::Long => Variant::from_i32(bits.as_long()),
            ScalarType::Double => Variant::from_f64(bits.as_double()),
            ScalarType::HostObject => {
                let pointer = bits.as_object();
                if pointer.is_null() {
                    Variant::nothing()
                } else {
                    // SAFETY: the admitted object operand retains this table.
                    super::value::to_variant(unsafe {
                        calcvba_host_api::value::retain_bits(ScalarType::HostObject, bits)
                    }?)
                }
            }
            ScalarType::Boolean => Variant::from_bool(bits.as_long() != 0),
            ScalarType::String | ScalarType::Variant | ScalarType::Record | ScalarType::Array => {
                // SAFETY: the input owner spans this retaining store.
                super::value::to_variant(unsafe { calcvba_host_api::value::retain_bits(ty, bits) }?)
            }
            _ => return Err(13),
        };
        // SAFETY: canonical cell is live; install before dropping the old owner.
        let old = unsafe { std::ptr::replace(globals.add(index), value) };
        drop(old);
    } else {
        if index >= MAX_TYPED_SLOTS {
            return Err(5);
        }
        if ty == ScalarType::Variant {
            // SAFETY: every Variant physical slot points into initialized owned storage.
            let cell = unsafe { (*raw).slots[index] }
                .as_object()
                .cast_mut()
                .cast::<Variant>();
            if cell.is_null() {
                return Err(5);
            }
            // SAFETY: clone before replacing, including self-assignment.
            let Value::Variant(value) = unsafe { calcvba_host_api::value::retain_bits(ty, bits) }?
            else {
                return Err(13);
            };
            // SAFETY: cell belongs to this frame and retains its initialized extent.
            drop(unsafe { cell.replace(value) });
            return Ok(());
        }
        let next = if matches!(ty, ScalarType::Record | ScalarType::Array)
            && !bits.as_object().is_null()
        {
            // SAFETY: clone the exact aggregate before releasing the old slot.
            unsafe { calcvba_host_api::value::retain_bits(ty, bits) }?
                .into_owned_bits()
                .map_err(|_| 13)?
        } else if ty == ScalarType::String {
            // SAFETY: clone the exact borrowed BSTR before releasing the old cell.
            let clone = unsafe {
                calcvba_host_api::BStr::clone_from_raw_bstr(bits.as_object().cast_mut().cast())
            }
            .map_err(|_| 7)?;
            ScalarBits::object(clone.into_raw_bstr().cast())
        } else if ty == ScalarType::HostObject && !bits.as_object().is_null() {
            // SAFETY: retain before releasing the destination, including self-assignment.
            unsafe { calcvba_host_api::value::retain_bits(ScalarType::HostObject, bits) }?
                .into_owned_bits()
                .map_err(|_| 13)?
        } else {
            bits
        };
        let old = {
            // SAFETY: checked initialized frame cell; borrow ends before release.
            let frame = unsafe { &mut *raw };
            std::mem::replace(&mut frame.slots[index], next)
        };
        if ty == ScalarType::String {
            // SAFETY: the previous String cell transfers its raw BSTR allocation.
            drop(unsafe {
                calcvba_host_api::BStr::from_raw_bstr(old.as_object().cast_mut().cast())
            });
        }
        if ty == ScalarType::HostObject {
            // SAFETY: the previous cell transfers its one retained reference.
            unsafe {
                calcvba_host_api::value::release_owned_bits(ScalarType::HostObject, old);
            }
        }
        if matches!(ty, ScalarType::Record | ScalarType::Array) {
            // SAFETY: old transfers this slot's previous exact aggregate owner.
            unsafe { calcvba_host_api::value::release_owned_bits(ty, old) };
        }
    }
    Ok(())
}

/// # Safety
/// Same canonical-store/stable-alias contract as write_slot. value transfers its
/// exact carrier; install occurs before releasing the destination's old owner.
pub unsafe fn write_owned(
    raw: *mut RawActivation,
    index: usize,
    ty: ScalarType,
    value: Value,
) -> Result<(), i32> {
    if value.ty() != ty {
        return Err(13);
    }
    // SAFETY: aliases name retained stable ancestors with the same checked type.
    let alias = unsafe { (&*raw).aliases.get(index).copied() }.unwrap_or(Alias::NONE);
    if !alias.frame.is_null() {
        // SAFETY: the complete ancestor chain is retained.
        return unsafe { write_owned(alias.frame, alias.index, ty, value) };
    }
    // Validate aggregates against retained source metadata before touching the
    // canonical destination, including internal calls and host ByRef copies.
    // SAFETY: immutable admission facts are leased by this activation.
    let layout = unsafe { &*(*raw).layout };
    match &value {
        Value::Record(record)
            if layout
                .record_slots
                .get(&index)
                .is_none_or(|expected| expected.as_ref() != record.layout().as_ref()) =>
        {
            return Err(13);
        }
        Value::Array(Some(array))
            if layout.array_slots.get(&index) != Some(&array.element_vartype()) =>
        {
            return Err(13);
        }
        _ => {}
    }
    // SAFETY: short field reads; no reference spans dropping a replaced owner.
    let (base, globals, len) = unsafe { ((*raw).global_base, (*raw).globals, (*raw).global_len) };
    if index >= base {
        let index = index - base;
        if index >= len || globals.is_null() {
            return Err(5);
        }
        let value = if matches!(value, Value::Array(None)) {
            Variant::unallocated_array(*layout.array_slots.get(&(index + base)).ok_or(5)?)
        } else {
            super::value::to_variant(value)
        };
        // SAFETY: the retained canonical global cell transfers its old owner.
        drop(unsafe { globals.add(index).replace(value) });
        return Ok(());
    }
    if index >= MAX_TYPED_SLOTS {
        return Err(5);
    }
    if let Value::Variant(value) = value {
        // SAFETY: each Variant slot points into initialized frame-owned storage.
        let cell = unsafe { (*raw).slots[index] }
            .as_object()
            .cast_mut()
            .cast::<Variant>();
        if cell.is_null() {
            return Err(5);
        }
        // SAFETY: cell is the checked, initialized destination owned by this frame.
        drop(unsafe { cell.replace(value) });
        return Ok(());
    }
    let bits = value.into_owned_bits().map_err(|_| 13)?;
    // SAFETY: exact initialized physical slot; retain moved owner before release.
    let old = unsafe { std::ptr::replace(std::ptr::addr_of_mut!((*raw).slots[index]), bits) };
    match ty {
        ScalarType::Record | ScalarType::Array => {
            // SAFETY: transfer one replaced exact aggregate owner.
            unsafe { calcvba_host_api::value::release_owned_bits(ty, old) };
        }
        ScalarType::HostObject => {
            // SAFETY: transfer one replaced factory owner.
            unsafe {
                calcvba_host_api::value::release_owned_bits(ScalarType::HostObject, old);
            }
        }
        ScalarType::String => {
            // SAFETY: transfer one replaced raw BSTR owner.
            drop(unsafe {
                calcvba_host_api::BStr::from_raw_bstr(old.as_object().cast_mut().cast())
            });
        }
        _ => {}
    }
    Ok(())
}

/// # Safety
/// Same stable stack/slot contract as write_slot. Aliases refer to ancestors only.
pub unsafe fn read_slot(
    raw: *mut RawActivation,
    index: usize,
    ty: ScalarType,
) -> Result<ScalarBits, i32> {
    let mut target = raw;
    let mut index = index;
    loop {
        // SAFETY: each alias moves strictly toward a retained stable ancestor.
        let (base, alias) = unsafe {
            (
                (*target).global_base,
                (&*target)
                    .aliases
                    .get(index)
                    .copied()
                    .unwrap_or(Alias::NONE),
            )
        };
        if !alias.frame.is_null() {
            target = alias.frame;
            index = alias.index;
            continue;
        }
        return if index >= base {
            // The alias may cross source units. Read the target unit's store,
            // retaining scratch on the requesting frame rather than its parked
            // ancestor. No reference into either store spans the snapshot.
            // SAFETY: alias resolution retains the target canonical store.
            let value = unsafe {
                let offset = index - base;
                if offset >= (*target).global_len || (*target).globals.is_null() {
                    return Err(5);
                }
                (*(*target).globals.add(offset)).clone()
            };
            let value = super::value::from_variant(ty, value)?;
            // SAFETY: snapshot transfers into this executing frame's scratch.
            unsafe { pin_value(raw, value) }
        } else if index < MAX_TYPED_SLOTS {
            // SAFETY: initialized ancestor slot with its admitted carrier.
            Ok(unsafe { (*target).slots[index] })
        } else {
            Err(5)
        };
    }
}

// SAFETY CONTRACT: aliases name this stack's boxed, leased ancestor frames.
pub(super) unsafe fn resolve_alias(
    mut raw: *mut RawActivation,
    mut index: usize,
) -> (*mut RawActivation, usize) {
    loop {
        // SAFETY: each step follows an admitted retained ancestor authority.
        let alias = unsafe { (&*raw).aliases.get(index).copied() }.unwrap_or(Alias::NONE);
        if alias.frame.is_null() {
            return (raw, index);
        }
        raw = alias.frame;
        index = alias.index;
    }
}
pub(super) unsafe fn array_aliases(
    raw: *mut RawActivation,
    index: usize,
) -> Option<Vec<Option<Alias>>> {
    // SAFETY: forwarded stable ancestor contract; clone only inert authorities.
    let (raw, index) = unsafe { resolve_alias(raw, index) };
    // SAFETY: brief metadata borrow, ending before any carrier operation.
    unsafe { (&*raw).param_arrays.get(&index).cloned() }
}
pub(super) unsafe fn alias_value(alias: Alias) -> Result<Variant, i32> {
    // SAFETY: inherited ancestor lease; no borrowed storage leaves this helper.
    unsafe {
        let (raw, index) = resolve_alias(alias.frame, alias.index);
        let layout = &*(*raw).layout;
        let ty = *layout.slots.get(index).ok_or(5)?;
        if index >= (*raw).global_base {
            let offset = index - (*raw).global_base;
            if offset >= (*raw).global_len || (*raw).globals.is_null() {
                return Err(5);
            }
            Ok((*(*raw).globals.add(offset)).clone())
        } else {
            if ty == ScalarType::Array && (*raw).slots[index].as_object().is_null() {
                return Ok(Variant::unallocated_array(
                    *layout.array_slots.get(&index).ok_or(5)?,
                ));
            }
            Ok(super::value::to_variant(
                calcvba_host_api::value::retain_bits(ty, (*raw).slots[index])?,
            ))
        }
    }
}
pub(super) unsafe fn write_alias(alias: Alias, value: Variant) -> Result<(), i32> {
    // SAFETY: short metadata reads and an exact owned write through this stack's
    // retained ancestor authority; no frame/global borrow spans a release.
    unsafe {
        let (raw, index) = resolve_alias(alias.frame, alias.index);
        let ty = *(&*(*raw).layout).slots.get(index).ok_or(5)?;
        write_owned(raw, index, ty, super::value::from_variant(ty, value)?)
    }
}

/// # Safety
/// Generated caller supplies a leased stable frame and exact slot kind/output.
pub unsafe extern "C" fn slot_read(
    raw: *mut RawActivation,
    index: usize,
    kind: u32,
    out: *mut ScalarBits,
) -> i32 {
    let Some(ty) = decode_kind(kind) else {
        return 5;
    };
    // SAFETY: inherited checked native-entry contract.
    match unsafe { read_slot(raw, index, ty) } {
        Ok(bits) => {
            // SAFETY: caller supplies one aligned writable output cell.
            unsafe {
                out.write(bits);
            }
            0
        }
        Err(error) => error,
    }
}

/// # Safety
/// Same lease/slot/output contracts as read_global. This is a native typed helper.
pub unsafe extern "C" fn global_read(
    raw: *mut RawActivation,
    index: usize,
    kind: u32,
    out: *mut ScalarBits,
) -> i32 {
    let Some(ty) = decode_kind(kind) else {
        return 5;
    };
    // SAFETY: inherited generated entry contract.
    match unsafe { read_global(raw, index, ty) } {
        Ok(bits) => {
            // SAFETY: generated entry supplies one aligned writable output cell.
            unsafe {
                out.write(bits);
            }
            0
        }
        Err(error) => error,
    }
}
/// # Safety
/// Same lease/slot/input contracts as write_slot. No panic may cross this entry.
pub unsafe extern "C" fn slot_write(
    raw: *mut RawActivation,
    index: usize,
    kind: u32,
    value: *const ScalarBits,
) -> i32 {
    let Some(ty) = decode_kind(kind) else {
        return 5;
    };
    // SAFETY: inherited generated entry contract.
    unsafe { write_slot(raw, index, ty, value) }.map_or_else(|e| e, |()| 0)
}
/// # Safety
/// raw is a live unique activation outside an instruction's callback.
pub unsafe extern "C" fn scratch_clear(raw: *mut RawActivation) {
    // SAFETY: inherited generated entry contract.
    unsafe {
        clear_scratch(raw);
    }
}
/// # Safety
/// raw/index name a live initialized slot of this exact physical carrier.
unsafe fn clear_slot(raw: *mut RawActivation, index: usize, ty: ScalarType) -> Result<(), i32> {
    if ty == ScalarType::Variant {
        // SAFETY: retain the actual cell's pointer while replacing its owned value.
        unsafe { write_owned(raw, index, ty, Value::Variant(Variant::empty())) }
    } else {
        // SAFETY: zero is the empty/null primitive, BSTR or object carrier.
        unsafe { write_slot(raw, index, ty, &ScalarBits::ZERO) }
    }
}
/// # Safety
/// raw/index name a retained initialized actual Variant slot; no entry overlaps.
pub unsafe extern "C" fn clear_variant(raw: *mut RawActivation, index: usize) -> i32 {
    // SAFETY: generated caller supplies the checked slot's exact carrier.
    unsafe { clear_slot(raw, index, ScalarType::Variant) }.map_or_else(|e| e, |()| 0)
}
pub fn kind(ty: ScalarType) -> u32 {
    match ty {
        ScalarType::Long => 1,
        ScalarType::Double => 2,
        ScalarType::HostObject => 3,
        ScalarType::Boolean => 4,
        ScalarType::String => 5,
        ScalarType::Variant => 6,
        ScalarType::Void => 7,
        ScalarType::Interface(_) => 3,
        ScalarType::Record => 8,
        ScalarType::Array => 9,
    }
}
fn decode_kind(kind: u32) -> Option<ScalarType> {
    match kind {
        1 => Some(ScalarType::Long),
        2 => Some(ScalarType::Double),
        3 => Some(ScalarType::HostObject),
        4 => Some(ScalarType::Boolean),
        5 => Some(ScalarType::String),
        6 => Some(ScalarType::Variant),
        7 => Some(ScalarType::Void),
        8 => Some(ScalarType::Record),
        9 => Some(ScalarType::Array),
        _ => None,
    }
}

/// # Safety
/// Generated caller transfers a successful typed native result at output. Its
/// exact extent is determined by kind; raw/index are an admitted destination.
pub unsafe extern "C" fn result_write(
    raw: *mut RawActivation,
    index: usize,
    kind: u32,
    output: *mut ScalarBits,
) -> i32 {
    let Some(ty) = decode_kind(kind) else {
        return 5;
    };
    // SAFETY: success initialized the advertised native output carrier.
    let value = match unsafe { calcvba_host_api::value::take_result(ty, output) } {
        Ok(value) => value,
        Err(e) => return e,
    };
    if ty == ScalarType::Void {
        return 0;
    }
    // SAFETY: the successful exact result transfers into this admitted slot.
    unsafe { write_owned(raw, index, ty, value) }.map_or_else(|e| e, |()| 0)
}
/// # Safety
/// Same result-transfer contract as result_write; the IID words are a checked
/// nominal return constraint (all zero means no constraint). Failure still
/// consumes and releases the successful native result exactly once.
pub unsafe extern "C" fn result_write_checked(
    raw: *mut RawActivation,
    index: usize,
    kind: u32,
    output: *mut ScalarBits,
    iid_low: u64,
    iid_high: u64,
) -> i32 {
    let Some(ty) = decode_kind(kind) else {
        return 5;
    };
    // SAFETY: generated success initialized this exact output carrier.
    let value = match unsafe { calcvba_host_api::value::take_result(ty, output) } {
        Ok(value) => value,
        Err(e) => return e,
    };
    if iid_low != 0 || iid_high != 0 {
        let mut iid = [0u8; 16];
        iid[..8].copy_from_slice(&iid_low.to_le_bytes());
        iid[8..].copy_from_slice(&iid_high.to_le_bytes());
        if !value.has_interface(iid) {
            return 13;
        }
    }
    if ty == ScalarType::Void {
        return 0;
    }
    // SAFETY: this exact checked result transfers into the admitted destination.
    unsafe { write_owned(raw, index, ty, value) }.map_or_else(|e| e, |()| 0)
}

/// Release an owned native slot result when VBA ignores a Function's value.
/// # Safety
/// output transfers one initialized owner of the exact successful result kind.
pub unsafe extern "C" fn result_drop(kind: u32, output: *mut ScalarBits) -> i32 {
    crate::with_status(|| {
        let Some(ty) = decode_kind(kind) else {
            return 5;
        };
        // SAFETY: the generated caller transfers the successful slot result once.
        match unsafe { calcvba_host_api::value::take_result(ty, output) } {
            Ok(value) => {
                drop(value);
                0
            }
            Err(error) => error,
        }
    })
}

/// # Safety
/// raw is a retained unique frame; bytes is a live UTF-8 literal and output is
/// one aligned writable carrier. The original OxFunc retains the literal.
pub unsafe extern "C" fn literal_string(
    raw: *mut RawActivation,
    bytes: *const u8,
    length: usize,
    output: *mut ScalarBits,
) -> i32 {
    // SAFETY: generated caller retains the complete original literal allocation.
    let text = match std::str::from_utf8(unsafe { std::slice::from_raw_parts(bytes, length) }) {
        Ok(text) => text,
        Err(_) => return 5,
    };
    // SAFETY: frame owns the new BSTR until the instruction completes.
    match unsafe { pin_value(raw, Value::String(text.into())) } {
        Ok(bits) => {
            // SAFETY: checked output extent.
            unsafe {
                output.write(bits);
            }
            0
        }
        Err(e) => e,
    }
}
/// # Safety
/// raw retains Variant scratch storage; output is one aligned writable carrier.
pub unsafe extern "C" fn constant_variant(
    raw: *mut RawActivation,
    null: u32,
    output: *mut ScalarBits,
) -> i32 {
    let value = if null != 0 {
        Variant::null()
    } else {
        Variant::empty()
    };
    // SAFETY: frame owns the actual inline Variant until instruction completion.
    match unsafe { pin_value(raw, Value::Variant(value)) } {
        Ok(bits) => {
            // SAFETY: checked output extent.
            unsafe {
                output.write(bits);
            }
            0
        }
        Err(e) => e,
    }
}

/// # Safety
/// pointer is null or a retained canonical object carrier on its owning thread.
/// IID halves contain the exact little-endian descriptor identity.
pub unsafe extern "C" fn check_object_interface(
    pointer: *const calcvba_host_api::object::ObjectHeader,
    low: u64,
    high: u64,
) -> i32 {
    let mut iid = [0; 16];
    iid[..8].copy_from_slice(&low.to_le_bytes());
    iid[8..].copy_from_slice(&high.to_le_bytes());
    // SAFETY: inherited canonical IUnknown carrier contract; header inspection
    // happens only after the factory vtable proof in retain_bits.
    match unsafe {
        calcvba_host_api::value::retain_bits(ScalarType::HostObject, ScalarBits::object(pointer))
    } {
        Ok(value) if value.has_interface(iid) => 0,
        _ => 13,
    }
}
