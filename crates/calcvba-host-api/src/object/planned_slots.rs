//! Descriptor-planned member calls with typed owning cells and explicit aliases.
//! This linked Rust ABI is separate from Automation and the original slot ABIs.
use super::*;
use crate::{Value, ValueType};

pub const SLOT_PLANNED_MEMBER: u32 = 0x4000_0001;
pub(super) type PlannedCallback =
    dyn Fn(&mut TypedContext, &mut MemberArguments<'_>) -> Result<Value, i32>;

#[derive(Clone, Copy)]
pub enum CallerSyncPhase {
    WriteCaller,
    ReadCaller,
}
pub type CallerSync<'a> = dyn FnMut(CallerSyncPhase, &mut [Value]) -> Result<(), i32> + 'a;

/// A bounded typed argument list. Repeated ByRef arguments address the same cell;
/// no simultaneous mutable Rust references to an aliased value are exposed.
pub struct MemberArguments<'a> {
    values: &'a mut [Value],
    aliases: &'a [usize],
    caller_sync: Option<&'a mut CallerSync<'a>>,
}
impl<'a> MemberArguments<'a> {
    pub fn new(values: &'a mut [Value], aliases: &'a [usize]) -> Result<Self, i32> {
        if values.len() > crate::MAX_TYPED_ARGS
            || values.len() != aliases.len()
            || aliases.iter().enumerate().any(|(i, &alias)| {
                alias > i || aliases[alias] != alias || values[alias].ty() != values[i].ty()
            })
        {
            return Err(13);
        }
        Ok(Self {
            values,
            aliases,
            caller_sync: None,
        })
    }
    /// Lend the owning caller's lvalue copies for this exact member operation.
    /// The callback must read/write only its canonical alias positions.
    pub fn with_caller_sync(mut self, sync: &'a mut CallerSync<'a>) -> Self {
        self.caller_sync = Some(sync);
        self
    }
    fn sync_caller(&mut self, phase: CallerSyncPhase) -> Result<(), i32> {
        let Some(sync) = &mut self.caller_sync else {
            return Ok(());
        };
        match catch_unwind(AssertUnwindSafe(|| sync(phase, self.values))) {
            Ok(result) => result,
            Err(payload) => {
                discard_panic(payload);
                Err(5)
            }
        }
    }
    pub fn write_caller_values(&mut self) -> Result<(), i32> {
        self.sync_caller(CallerSyncPhase::WriteCaller)
    }
    pub fn read_caller_values(&mut self) -> Result<(), i32> {
        self.sync_caller(CallerSyncPhase::ReadCaller)
    }
    pub fn len(&self) -> usize {
        self.values.len()
    }
    pub fn is_empty(&self) -> bool {
        self.values.is_empty()
    }
    pub fn alias(&self, index: usize) -> Option<usize> {
        self.aliases.get(index).copied()
    }
    pub fn get(&self, index: usize) -> Option<&Value> {
        self.alias(index).and_then(|index| self.values.get(index))
    }
    pub fn set(&mut self, index: usize, value: Value) -> Result<(), i32> {
        let index = self.alias(index).ok_or(13)?;
        let target = self.values.get_mut(index).ok_or(13)?;
        if target.ty() != value.ty() {
            return Err(13);
        }
        *target = value;
        Ok(())
    }
}
impl HostSlot {
    /// A planned callback validates its retained declared signature before entry.
    /// Values stay in their actual carriers; only declared Variants contain Variant.
    pub fn planned(
        callback: impl Fn(&mut TypedContext, &mut MemberArguments<'_>) -> Result<Value, i32> + 'static,
    ) -> Self {
        Self {
            callback: Box::new(Callback::Planned(Box::new(callback))),
        }
    }
}

// SAFETY CONTRACT: data is the retained factory callback; context/arguments are
// unique, aligned live objects. output is writable uninitialized Value storage.
pub(super) unsafe extern "C" fn planned_entry(
    data: *const (),
    context: *mut TypedContext,
    arguments: *mut MemberArguments<'_>,
    output: *mut Value,
) -> i32 {
    let result = catch_unwind(AssertUnwindSafe(|| {
        // SAFETY: the factory pairs this entry with Callback::Planned.
        let Callback::Planned(callback) = (unsafe { &*data.cast::<Callback>() }) else {
            return Err(5);
        };
        // SAFETY: inherited exclusive synchronous call contract.
        let context = unsafe { &mut *context };
        drop(crate::take_typed_fault(context, 0));
        // SAFETY: caller lends the complete checked argument object.
        let value = callback(context, unsafe { &mut *arguments })?;
        drop(crate::take_typed_fault(context, 0));
        Ok(value)
    }));
    match result {
        Ok(Ok(value)) => {
            // SAFETY: success transfers exactly one owning typed result.
            unsafe { output.write(value) };
            0
        }
        Ok(Err(status)) if status == -1 || status > 0 => status,
        Ok(Err(_)) => 5,
        Err(payload) => {
            discard_panic(payload);
            5
        }
    }
}

/// # Safety
/// header is null or a retained factory table. arguments owns each exact carrier
/// through this call and result is the compiler-retained result carrier.
pub unsafe fn invoke_planned_member(
    header: *const ObjectHeader,
    iid: InterfaceId,
    slot: usize,
    result: ValueType,
    arguments: &mut MemberArguments<'_>,
    context: &mut TypedContext,
) -> Result<Value, i32> {
    // SAFETY: inherited factory-header lifetime contract.
    let header = unsafe { header.as_ref() }.ok_or(91)?;
    if header.iid != iid || slot >= header.slot_count {
        return Err(438);
    }
    // SAFETY: retained factory header spans the complete immutable slot table.
    let native = unsafe { &*header.slots.add(slot) };
    if native.signature != SLOT_PLANNED_MEMBER {
        return Err(13);
    }
    // SAFETY: the factory signature proves this exact linked entry ABI.
    let call: unsafe extern "C" fn(
        *const (),
        *mut TypedContext,
        *mut MemberArguments<'_>,
        *mut Value,
    ) -> i32 = unsafe { std::mem::transmute(native.entry) };
    let mut output = std::mem::MaybeUninit::uninit();
    // SAFETY: unique arguments/context and uninitialized typed result storage.
    let status = unsafe { call(native.data, context, arguments, output.as_mut_ptr()) };
    if status != 0 {
        return Err(status);
    }
    // SAFETY: successful entry initialized one exact owned Value.
    let value = unsafe { output.assume_init() };
    if value.ty().carrier() != result.carrier() {
        return Err(13);
    }
    if let Value::Double(v) = value
        && !v.is_finite()
    {
        return Err(6);
    }
    if let Some(iid) = result.interface()
        && !value.has_interface(iid)
    {
        return Err(13);
    }
    Ok(value)
}
