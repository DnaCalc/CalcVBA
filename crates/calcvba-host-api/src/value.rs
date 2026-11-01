//! Additional typed carriers. The original scalar/slot API remains unchanged.
use crate::{
    Error, HostObject, InterfaceId, ObjectArg, OwnedValue, ScalarBits, ScalarType, Slot,
    TypedContext, Variant,
};
pub use oxvba_runtime::bstr::BStr;
pub const TYPED_VALUE_ABI: u32 = 1;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ValueType {
    Long,
    Double,
    Boolean,
    HostObject,
    String,
    Variant,
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
    Nothing,
    String(&'a BStr),
    Variant(&'a Variant),
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
            Self::Object(_) | Self::Nothing => ValueType::HostObject,
            Self::String(_) => ValueType::String,
            Self::Variant(_) => ValueType::Variant,
        }
    }
    pub fn bits(self) -> ScalarBits {
        match self {
            Self::Long(v) => ScalarBits::long(v),
            Self::Double(v) => ScalarBits::double(v),
            Self::Boolean(v) => ScalarBits::long(if v { -1 } else { 0 }),
            Self::Object(v) => ScalarBits::object(v.as_ptr()),
            Self::Nothing => ScalarBits::ZERO,
            Self::String(v) => ScalarBits::object(v.raw_bstr().cast()),
            Self::Variant(v) => ScalarBits::object(std::ptr::from_ref(v).cast()),
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
}
impl Value {
    pub fn ty(&self) -> ValueType {
        match self {
            Self::Empty => ValueType::Void,
            Self::Long(_) => ValueType::Long,
            Self::Double(_) => ValueType::Double,
            Self::Boolean(_) => ValueType::Boolean,
            Self::String(_) => ValueType::String,
            Self::Variant(_) => ValueType::Variant,
            Self::Object(_) => ValueType::HostObject,
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
        ValueType::HostObject => Value::Object((!bits.as_object().is_null()).then(|| {
            // SAFETY: the input owner retains this exact factory allocation.
            unsafe { HostObject::retain_header(bits.as_object()) }
        })),
        ValueType::String => {
            // SAFETY: the input owner retains this exact BSTR.
            Value::String(
                unsafe { BStr::clone_from_raw_bstr(bits.as_object().cast_mut().cast()) }
                    .map_err(|_| 7)?,
            )
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
        Value::String(v) => ScalarBits::object(v.into_raw_bstr().cast()),
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
                // SAFETY: retain the transferred factory owner, then consume it.
                let object = unsafe { HostObject::retain_header(pointer) };
                // SAFETY: consume the one transferred raw reference.
                unsafe {
                    HostObject::release_header(pointer);
                }
                Value::Object(Some(object))
            }
        }
        ValueType::String => {
            // SAFETY: transfer the exact owned raw BSTR.
            Value::String(unsafe { BStr::from_raw_bstr(bits.as_object().cast_mut().cast()) })
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
