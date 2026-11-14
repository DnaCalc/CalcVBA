//! Additional typed carriers. The original scalar/slot API remains unchanged.
use crate::{
    Error, HostObject, InterfaceId, ObjectArg, OwnedValue, ScalarBits, ScalarType, Slot,
    TypedContext, Variant,
};
pub use calcvba_runtime::bstr::BStr;
pub use calcvba_runtime::{safe_array::SafeArray, vba_record::VbaRecord};
pub const TYPED_VALUE_ABI: u32 = 1;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ValueType {
    Long,
    Double,
    Boolean,
    Byte,
    Integer,
    LongLong,
    Single,
    Currency,
    Date,
    HostObject,
    String,
    Variant,
    Record,
    Array,
    Void,
    /// Nominal metadata for a host-library interface; its physical carrier is Object.
    Interface(InterfaceId),
}
impl ValueType {
    pub fn carrier(self) -> Self {
        match self {
            Self::Interface(_) => Self::HostObject,
            other => other,
        }
    }
    pub fn interface(self) -> Option<InterfaceId> {
        match self {
            Self::Interface(iid) => Some(iid),
            _ => None,
        }
    }
    pub fn legacy(self) -> Option<ScalarType> {
        match self.carrier() {
            Self::Long => Some(ScalarType::Long),
            Self::Double => Some(ScalarType::Double),
            Self::Boolean => Some(ScalarType::Boolean),
            Self::HostObject => Some(ScalarType::HostObject),
            _ => None,
        }
    }
    pub fn native_code(self) -> u32 {
        match self.carrier() {
            Self::Void => 0,
            Self::Long => 1,
            Self::Double => 2,
            Self::Boolean => 3,
            Self::HostObject => 4,
            Self::String => 5,
            Self::Variant => 6,
            Self::Record => 7,
            Self::Array => 8,
            Self::Byte => 9,
            Self::Integer => 10,
            Self::LongLong => 11,
            Self::Single => 12,
            Self::Currency => 13,
            Self::Date => 14,
            _ => unreachable!(),
        }
    }
}
impl From<ScalarType> for ValueType {
    fn from(ty: ScalarType) -> Self {
        match ty {
            ScalarType::Long => Self::Long,
            ScalarType::Double => Self::Double,
            ScalarType::Boolean => Self::Boolean,
            ScalarType::HostObject => Self::HostObject,
        }
    }
}
impl PartialEq<ScalarType> for ValueType {
    fn eq(&self, other: &ScalarType) -> bool {
        self.legacy().is_some_and(|ty| ty == *other)
    }
}
impl PartialEq<ValueType> for ScalarType {
    fn eq(&self, other: &ValueType) -> bool {
        other == self
    }
}

#[derive(Clone, Copy)]
pub enum ValueSlot<'a> {
    Long(i32),
    Double(f64),
    Boolean(bool),
    Object(ObjectArg<'a>),
    RuntimeObject(&'a calcvba_runtime::ObjectRef),
    Nothing,
    String(&'a BStr),
    Variant(&'a Variant),
    Record(&'a VbaRecord),
    Array(&'a SafeArray),
    UnallocatedArray,
}
impl<'a> From<Slot<'a>> for ValueSlot<'a> {
    fn from(slot: Slot<'a>) -> Self {
        match slot {
            Slot::Long(v) => Self::Long(v),
            Slot::Double(v) => Self::Double(v),
            Slot::Object(v) => Self::Object(v),
            Slot::Nothing => Self::Nothing,
        }
    }
}
impl ValueSlot<'_> {
    pub fn ty(self) -> ValueType {
        match self {
            Self::Long(_) => ValueType::Long,
            Self::Double(_) => ValueType::Double,
            Self::Boolean(_) => ValueType::Boolean,
            Self::Object(_) | Self::RuntimeObject(_) | Self::Nothing => ValueType::HostObject,
            Self::String(_) => ValueType::String,
            Self::Variant(_) => ValueType::Variant,
            Self::Record(_) => ValueType::Record,
            Self::Array(_) | Self::UnallocatedArray => ValueType::Array,
        }
    }
    pub fn bits(self) -> ScalarBits {
        match self {
            Self::Long(v) => ScalarBits::long(v),
            Self::Double(v) => ScalarBits::double(v),
            Self::Boolean(v) => ScalarBits::long(if v { -1 } else { 0 }),
            Self::Object(v) => ScalarBits::object(v.as_ptr()),
            Self::RuntimeObject(v) => ScalarBits::object(v.raw_iunknown().cast()),
            Self::Nothing => ScalarBits::ZERO,
            Self::String(v) => ScalarBits::object(v.raw_bstr().cast()),
            Self::Variant(v) => ScalarBits::object(std::ptr::from_ref(v).cast()),
            Self::Record(v) => ScalarBits::object(std::ptr::from_ref(v).cast()),
            Self::Array(v) => ScalarBits::object(v.raw_safearray_ptr().cast()),
            Self::UnallocatedArray => ScalarBits::ZERO,
        }
    }
    pub fn is_valid(self) -> bool {
        !matches!(self, Self::Double(v) if !v.is_finite())
    }
}

#[derive(Debug)]
pub enum Value {
    Empty,
    Long(i32),
    Double(f64),
    Boolean(bool),
    String(BStr),
    Variant(Variant),
    Object(Option<HostObject>),
    RuntimeObject(calcvba_runtime::ObjectRef),
    Record(VbaRecord),
    Array(Option<SafeArray>),
}
impl Value {
    pub fn has_interface(&self, iid: InterfaceId) -> bool {
        match self {
            Self::Object(None) => true,
            Self::Object(Some(object)) => object.header().iid == iid,
            _ => false,
        }
    }
    /// Transfer an owned non-Variant carrier. Its eventual consumer must release
    /// the exact owner with release_owned_bits or take_result on this thread.
    pub fn into_owned_bits(self) -> Result<ScalarBits, Self> {
        Ok(match self {
            Self::Empty => ScalarBits::ZERO,
            Self::Long(v) => ScalarBits::long(v),
            Self::Double(v) => ScalarBits::double(v),
            Self::Boolean(v) => ScalarBits::long(if v { -1 } else { 0 }),
            Self::Object(v) => {
                ScalarBits::object(v.map_or(std::ptr::null(), HostObject::into_header))
            }
            Self::RuntimeObject(v) => ScalarBits::object(v.into_raw_iunknown().cast()),
            Self::String(v) => ScalarBits::object(v.into_raw_bstr().cast()),
            Self::Record(v) => ScalarBits::object(Box::into_raw(Box::new(v)).cast()),
            Self::Array(v) => ScalarBits::object(
                v.map_or(std::ptr::null_mut(), SafeArray::into_raw_safearray_ptr)
                    .cast(),
            ),
            v @ Self::Variant(_) => return Err(v),
        })
    }
    pub fn ty(&self) -> ValueType {
        match self {
            Self::Empty => ValueType::Void,
            Self::Long(_) => ValueType::Long,
            Self::Double(_) => ValueType::Double,
            Self::Boolean(_) => ValueType::Boolean,
            Self::String(_) => ValueType::String,
            Self::Variant(_) => ValueType::Variant,
            Self::Object(_) | Self::RuntimeObject(_) => ValueType::HostObject,
            Self::Record(_) => ValueType::Record,
            Self::Array(_) => ValueType::Array,
        }
    }
    pub fn borrow_slot(&self) -> Option<ValueSlot<'_>> {
        Some(match self {
            Self::Empty => return None,
            Self::Long(v) => ValueSlot::Long(*v),
            Self::Double(v) => ValueSlot::Double(*v),
            Self::Boolean(v) => ValueSlot::Boolean(*v),
            Self::String(v) => ValueSlot::String(v),
            Self::Variant(v) => ValueSlot::Variant(v),
            Self::RuntimeObject(v) => ValueSlot::RuntimeObject(v),
            Self::Record(v) => ValueSlot::Record(v),
            Self::Array(v) => v
                .as_ref()
                .map_or(ValueSlot::UnallocatedArray, ValueSlot::Array),
            Self::Object(v) => v
                .as_ref()
                .map_or(ValueSlot::Nothing, |o| ValueSlot::Object(o.argument())),
        })
    }
    pub fn into_legacy(self) -> Result<OwnedValue, Error> {
        match self {
            Self::Long(v) => Ok(OwnedValue::Long(v)),
            Self::Double(v) => Ok(OwnedValue::Double(v)),
            Self::Object(v) => Ok(OwnedValue::Object(v)),
            _ => Err(Error::UnsupportedSignature(
                "this result requires the extended value API".into(),
            )),
        }
    }
}
impl From<OwnedValue> for Value {
    fn from(value: OwnedValue) -> Self {
        match value {
            OwnedValue::Long(v) => Self::Long(v),
            OwnedValue::Double(v) => Self::Double(v),
            OwnedValue::Object(v) => Self::Object(v),
        }
    }
}

pub enum ValueProgress {
    Suspended,
    Returned(Value),
}
pub trait ValueInvocation {
    fn frame(&self) -> &crate::InvocationFrame;
    fn remaining_blocks(&self) -> u64;
    fn resume(&mut self, operations: u64) -> Result<ValueProgress, Error>;
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TypedMemberKind {
    Method,
    PropertyGet,
    PropertyLet,
    PropertySet,
}
#[derive(Clone)]
pub struct TypedObjectMember {
    pub name: String,
    pub slot: usize,
    pub kind: TypedMemberKind,
    pub parameters: Vec<ValueType>,
    pub result: ValueType,
    pub is_default: bool,
}
#[derive(Clone)]
pub struct TypedObjectInterface {
    pub name: String,
    pub iid: InterfaceId,
    pub members: Vec<TypedObjectMember>,
}
#[derive(Clone)]
pub struct TypedObjectLibrary {
    pub name: String,
    pub interfaces: Vec<TypedObjectInterface>,
}

/// A signature tag for the existing NativeSlot layout. Original numeric tags
/// 1/2 stay unchanged, and original double constructors keep their exact C ABI.
pub fn value_slot_signature(parameters: &[ValueType], result: ValueType) -> Option<u32> {
    // These carriers currently belong to source procedure frames. The native
    // host-member table has no aggregate signature extension yet.
    if parameters
        .iter()
        .chain(std::iter::once(&result))
        .any(|p| matches!(p.carrier(), ValueType::Record | ValueType::Array))
    {
        return None;
    }
    if parameters.len() > crate::MAX_TYPED_ARGS
        || parameters.iter().any(|p| p.carrier() == ValueType::Void)
    {
        return None;
    }
    if result.carrier() == ValueType::Double {
        if parameters.is_empty() {
            return Some(crate::object::SLOT_DOUBLE_0);
        }
        if parameters == [ValueType::Long] {
            return Some(crate::object::SLOT_DOUBLE_LONG);
        }
    }
    let mut signature = 0x8000_0000 | result.native_code() | ((parameters.len() as u32) << 4);
    for (i, ty) in parameters.iter().enumerate() {
        signature |= ty.native_code() << (8 + i * 3);
    }
    Some(signature)
}

pub(crate) type TypedCallback = dyn Fn(&mut TypedContext, &[ValueSlot<'_>]) -> Result<Value, i32>;

/// # Safety
/// bits borrows the exact carrier ty; its owner remains live until return.
pub unsafe fn retain_bits(ty: ValueType, bits: ScalarBits) -> Result<Value, i32> {
    Ok(match ty.carrier() {
        ValueType::Void => Value::Empty,
        ValueType::Long => Value::Long(bits.as_long()),
        ValueType::Double => Value::Double(bits.as_double()),
        ValueType::Boolean => Value::Boolean(bits.as_long() != 0),
        ValueType::HostObject => {
            // SAFETY: the input retains an exact runtime IUnknown prefix.
            let object = unsafe {
                calcvba_runtime::ObjectRef::from_raw_iunknown_addref(
                    bits.as_object().cast_mut().cast(),
                )
            };
            match object {
                None => Value::Object(None),
                Some(object) => match HostObject::from_object_ref(&object) {
                    Some(host) => Value::Object(Some(host)),
                    None => Value::RuntimeObject(object),
                },
            }
        }
        ValueType::String => {
            // SAFETY: the input owner retains this exact BSTR.
            Value::String(
                unsafe { BStr::clone_from_raw_bstr(bits.as_object().cast_mut().cast()) }
                    .map_err(|_| 7)?,
            )
        }
        ValueType::Record => {
            // SAFETY: caller retains one exact VbaRecord through this copy.
            Value::Record(
                unsafe { bits.as_object().cast::<VbaRecord>().as_ref() }
                    .ok_or(13)?
                    .clone(),
            )
        }
        ValueType::Array => {
            let raw = bits.as_object().cast_mut().cast::<core::ffi::c_void>();
            Value::Array(if raw.is_null() {
                None
            } else {
                // SAFETY: caller retains a live canonical SAFEARRAY descriptor.
                Some(unsafe { SafeArray::clone_from_raw_safearray(raw, None) }.ok_or(13)?)
            })
        }
        ValueType::Variant => {
            // SAFETY: a nonnull pointer borrows a live exact Variant.
            Value::Variant(
                unsafe { bits.as_object().cast::<Variant>().as_ref() }
                    .cloned()
                    .unwrap_or_else(Variant::empty),
            )
        }
        _ => return Err(13),
    })
}
/// # Safety
/// bits transfers one owned non-Variant carrier of exactly ty, on its owning
/// thread. Null aggregate/object/string pointers carry no allocation.
pub unsafe fn release_owned_bits(ty: ValueType, bits: ScalarBits) {
    let raw = bits.as_object().cast_mut();
    if raw.is_null() {
        return;
    }
    match ty.carrier() {
        ValueType::HostObject => {
            // SAFETY: inherited exact runtime IUnknown ownership transfer.
            drop(unsafe { calcvba_runtime::ObjectRef::from_raw_iunknown_owned(raw.cast()) });
        }
        ValueType::String => {
            // SAFETY: inherited exact BSTR ownership transfer.
            drop(unsafe { BStr::from_raw_bstr(raw.cast()) });
        }
        ValueType::Record => {
            // SAFETY: inherited complete Box<VbaRecord> ownership transfer.
            drop(unsafe { Box::from_raw(raw.cast::<VbaRecord>()) });
        }
        ValueType::Array => {
            // SAFETY: inherited canonical descriptor ownership transfer.
            drop(unsafe { SafeArray::from_raw_safearray_owned(raw.cast(), None) });
        }
        _ => {}
    }
}
/// # Safety
/// output has the advertised native result extent and alignment: eight bytes
/// for scalar/BSTR/object carriers, or one exact Variant. It is uninitialized
/// output storage; success transfers the value owner to that storage.
pub unsafe fn write_result(value: Value, output: *mut ScalarBits) {
    let bits = match value {
        Value::Variant(value) => {
            // SAFETY: the checked signature supplies the exact Variant extent.
            unsafe {
                output.cast::<Variant>().write(value);
            }
            return;
        }
        Value::Empty => return,
        Value::Long(v) => ScalarBits::long(v),
        Value::Double(v) => ScalarBits::double(v),
        Value::Boolean(v) => ScalarBits::long(if v { -1 } else { 0 }),
        Value::Object(v) => ScalarBits::object(v.map_or(std::ptr::null(), HostObject::into_header)),
        Value::RuntimeObject(v) => ScalarBits::object(v.into_raw_iunknown().cast()),
        Value::String(v) => ScalarBits::object(v.into_raw_bstr().cast()),
        Value::Record(v) => ScalarBits::object(Box::into_raw(Box::new(v)).cast()),
        Value::Array(v) => ScalarBits::object(
            v.map_or(std::ptr::null_mut(), SafeArray::into_raw_safearray_ptr)
                .cast(),
        ),
    };
    // SAFETY: the checked signature supplies one writable eight-byte carrier.
    unsafe {
        output.write(bits);
    }
}
/// # Safety
/// output contains one successful native result with the exact advertised
/// extent/alignment. This consumes that owner, once; do not clear it afterward.
pub unsafe fn take_result(ty: ValueType, output: *mut ScalarBits) -> Result<Value, i32> {
    if ty.carrier() == ValueType::Void {
        return Ok(Value::Empty);
    }
    if ty.carrier() == ValueType::Variant {
        // SAFETY: this ABI returns the actual inline canonical Variant.
        return Ok(Value::Variant(unsafe { output.cast::<Variant>().read() }));
    }
    // SAFETY: every other result initializes all eight carrier bytes.
    let bits = unsafe { output.read() };
    Ok(match ty.carrier() {
        ValueType::Long => Value::Long(bits.as_long()),
        ValueType::Double => Value::Double(bits.as_double()),
        ValueType::Boolean => Value::Boolean(bits.as_long() != 0),
        ValueType::HostObject => {
            let pointer = bits.as_object();
            if pointer.is_null() {
                Value::Object(None)
            } else {
                // SAFETY: successful output transfers this exact IUnknown once.
                let object = unsafe {
                    calcvba_runtime::ObjectRef::from_raw_iunknown_owned(pointer.cast_mut().cast())
                }
                .ok_or(13)?;
                match HostObject::from_object_ref(&object) {
                    Some(host) => Value::Object(Some(host)),
                    None => Value::RuntimeObject(object),
                }
            }
        }
        ValueType::String => {
            // SAFETY: transfer the exact owned raw BSTR.
            Value::String(unsafe { BStr::from_raw_bstr(bits.as_object().cast_mut().cast()) })
        }
        ValueType::Record => {
            let raw = bits.as_object().cast_mut().cast::<VbaRecord>();
            if raw.is_null() {
                return Err(13);
            }
            // SAFETY: successful record output transfers this exact Box once.
            Value::Record(*unsafe { Box::from_raw(raw) })
        }
        ValueType::Array => {
            let raw = bits.as_object().cast_mut().cast::<core::ffi::c_void>();
            Value::Array(if raw.is_null() {
                None
            } else {
                // SAFETY: successful array output transfers its descriptor once.
                Some(unsafe { SafeArray::from_raw_safearray_owned(raw, None) }.ok_or(13)?)
            })
        }
        _ => return Err(13),
    })
}

/// Rich fault published by a synchronous typed callback. The provider consumes it
/// with the returned error number before another VBA operation can execute.
#[derive(Debug)]
pub struct TypedFault {
    pub number: i32,
    pub description: String,
    pub source: Option<String>,
    pub help_file: Option<String>,
    pub help_context: Option<i32>,
}
thread_local! {
    static PENDING_FAULT: std::cell::RefCell<Option<(usize, TypedFault)>> = const { std::cell::RefCell::new(None) };
}
impl TypedFault {
    pub fn publish(self, context: &mut TypedContext) -> i32 {
        let number = self.number;
        PENDING_FAULT.with(|slot| slot.replace(Some((std::ptr::from_ref(context).addr(), self))));
        number
    }
}
/// Consume the fault belonging to this exact synchronous callback context.
/// A context address is compared only; it is never reconstructed or dereferenced.
pub fn take_typed_fault(context: *const TypedContext, number: i32) -> Option<TypedFault> {
    PENDING_FAULT
        .with(|slot| slot.take())
        .filter(|(key, fault)| *key == context.addr() && fault.number == number)
        .map(|(_, fault)| fault)
}
