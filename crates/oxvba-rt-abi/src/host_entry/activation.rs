//! Owning physical storage and exact continuation position for typed execution.
//! Global storage is the session's canonical Variant store; there is no mirror.
use super::{MAX_TYPED_SLOTS, ScalarBits, ScalarLayout, ScalarType};
use calcvba_host_api::{HostObject, OwnedValue, Slot, object::ObjectHeader};
use oxvba_runtime::Variant;

pub const SUSPENDED: i32 = -2;
pub const FINISHED_PC: u32 = u32::MAX;
/// Private generated-code ABI. Every nonnull object slot owns one reference.
#[repr(C)]
pub struct RawActivation {
    pub slots: [ScalarBits; MAX_TYPED_SLOTS],
    pub pc: u32,
    pub global_base: usize,
    globals: *mut Variant,
    global_len: usize,
    scratch: [*const ObjectHeader; 4],
    scratch_len: usize,
}
pub struct Activation<'a> {
    raw: RawActivation,
    layout: std::borrow::Cow<'a, ScalarLayout>,
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
    unsafe fn build(
        layout: std::borrow::Cow<'a, ScalarLayout>,
        args: &[Slot<'_>],
        globals: *mut Variant,
        global_len: usize,
        entry_pc: u32,
        finite_inputs: bool,
    ) -> Result<Self, i32> {
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
                if !pointer.is_null() && unsafe { (*pointer).iid != *iid } {
                    return Err(13);
                }
            }
        }
        if layout.slots.len() - layout.global_base > global_len {
            return Err(5);
        }
        let mut frame = Self {
            raw: RawActivation {
                slots: [ScalarBits::ZERO; MAX_TYPED_SLOTS],
                pc: entry_pc,
                global_base: layout.global_base,
                globals,
                global_len,
                scratch: [std::ptr::null(); 4],
                scratch_len: 0,
            },
            layout,
        };
        for (i, arg) in args.iter().enumerate() {
            let bits = arg.bits();
            // SAFETY: args retain all object borrows and this frame owns storage.
            unsafe {
                write_slot(&mut frame.raw, i, arg.ty(), &bits)?;
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
    pub fn result(&self) -> OwnedValue {
        let bits = self.raw.slots[self.layout.return_slot];
        match self.layout.result {
            ScalarType::Long => OwnedValue::Long(bits.as_long()),
            ScalarType::Double => OwnedValue::Double(bits.as_double()),
            ScalarType::HostObject => {
                let pointer = bits.as_object();
                // SAFETY: this admitted slot is null or owns a live table.
                OwnedValue::Object(
                    (!pointer.is_null()).then(|| unsafe { HostObject::retain_header(pointer) }),
                )
            }
            ScalarType::Boolean => unreachable!("Boolean result ABI is refused"),
        }
    }
    pub fn read(
        &mut self,
        operand: &oxvba_oxir::OxOperand,
        to: ScalarType,
    ) -> Result<ScalarBits, i32> {
        use oxvba_oxir::{OxOperand, OxPlace};
        let from = self.layout.operand_type(operand).expect("admitted operand");
        let bits = if let OxOperand::Use(OxPlace::Global(id)) = operand {
            // SAFETY: owning session lease spans the activation; no reference
            // into the global store survives this bounded read.
            unsafe { read_global(&mut self.raw, id.0, from)? }
        } else {
            self.layout.read(&self.raw.slots, operand, from)
        };
        Ok(if from == ScalarType::Long && to == ScalarType::Double {
            ScalarBits::double(f64::from(bits.as_long()))
        } else {
            bits
        })
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
    pub fn clear_temps(&mut self, from: usize) {
        for i in self.layout.local_count + from..self.layout.global_base {
            // SAFETY: initialized admitted local slots, with no overlapping entry.
            unsafe {
                write_slot(&mut self.raw, i, self.layout.slots[i], &ScalarBits::ZERO)
                    .expect("frame slot");
            }
        }
    }
}
impl Drop for Activation<'_> {
    fn drop(&mut self) {
        self.clear_scratch();
        for i in 0..self.layout.global_base {
            if self.layout.slots[i] == ScalarType::HostObject {
                let old = std::mem::replace(&mut self.raw.slots[i], ScalarBits::ZERO);
                // SAFETY: each admitted object slot transfers its one reference.
                unsafe {
                    HostObject::release_header(old.as_object());
                }
            }
        }
    }
}
/// # Safety
/// raw follows Activation::new's lease and exclusive running contract.
pub unsafe fn clear_scratch(raw: *mut RawActivation) {
    // Release after removing each pointer, so a host destructor cannot observe
    // a still-owned reference or an exclusive Rust borrow into this frame.
    loop {
        let pointer = {
            // SAFETY: the running owner supplies the unique frame pointer.
            let frame = unsafe { &mut *raw };
            if frame.scratch_len == 0 {
                break;
            }
            frame.scratch_len -= 1;
            std::mem::replace(&mut frame.scratch[frame.scratch_len], std::ptr::null())
        };
        // SAFETY: removed exactly one retained factory reference.
        unsafe {
            HostObject::release_header(pointer);
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
    // SAFETY: this non-callback read temporarily borrows the owned frame.
    let frame = unsafe { &mut *raw };
    if index >= frame.global_len || frame.globals.is_null() {
        return Err(5);
    }
    // SAFETY: checked cell extent; only clone/scalar access occurs in this scope.
    let value = unsafe { &*frame.globals.add(index) };
    match ty {
        ScalarType::Long => value.as_i32().map(ScalarBits::long).ok_or(13),
        ScalarType::Double => value.as_f64().map(ScalarBits::double).ok_or(13),
        ScalarType::HostObject => {
            let Some(object) = value.as_object_ref() else {
                return if value.vtype() == oxvba_runtime::VarType::Object {
                    Ok(ScalarBits::ZERO)
                } else {
                    Err(13)
                };
            };
            let object = HostObject::from_object_ref(&object).ok_or(13)?;
            if frame.scratch_len == frame.scratch.len() {
                return Err(5);
            }
            let pointer = object.into_header();
            frame.scratch[frame.scratch_len] = pointer;
            frame.scratch_len += 1;
            Ok(ScalarBits::object(pointer))
        }
        ScalarType::Boolean => Err(13),
    }
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
    // SAFETY: caller provides one initialized readable physical carrier.
    let bits = unsafe { *bits };
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
                    Variant::from_object_ref(
                        unsafe { HostObject::retain_header(pointer) }.object_ref(),
                    )
                }
            }
            ScalarType::Boolean => return Err(13),
        };
        // SAFETY: canonical cell is live; install before dropping the old owner.
        let old = unsafe { std::ptr::replace(globals.add(index), value) };
        drop(old);
    } else {
        if index >= MAX_TYPED_SLOTS {
            return Err(5);
        }
        let next = if ty == ScalarType::HostObject && !bits.as_object().is_null() {
            // SAFETY: retain before releasing the destination, including self-assignment.
            ScalarBits::object(unsafe { HostObject::retain_header(bits.as_object()) }.into_header())
        } else {
            bits
        };
        let old = {
            // SAFETY: checked initialized frame cell; borrow ends before release.
            let frame = unsafe { &mut *raw };
            std::mem::replace(&mut frame.slots[index], next)
        };
        if ty == ScalarType::HostObject {
            // SAFETY: the previous cell transfers its one retained reference.
            unsafe {
                HostObject::release_header(old.as_object());
            }
        }
    }
    Ok(())
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
pub fn kind(ty: ScalarType) -> u32 {
    match ty {
        ScalarType::Long => 1,
        ScalarType::Double => 2,
        ScalarType::HostObject => 3,
        ScalarType::Boolean => 4,
    }
}
fn decode_kind(kind: u32) -> Option<ScalarType> {
    match kind {
        1 => Some(ScalarType::Long),
        2 => Some(ScalarType::Double),
        3 => Some(ScalarType::HostObject),
        4 => Some(ScalarType::Boolean),
        _ => None,
    }
}
