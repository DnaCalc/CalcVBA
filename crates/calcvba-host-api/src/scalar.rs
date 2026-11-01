//! Versioned unboxed entry carriers. All entry pointers remain provider-owned.
use crate::object::ObjectArg;

pub const TYPED_ENTRY_ABI: u32 = 2;
pub const MAX_TYPED_ARGS: usize = 4;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ScalarType {
    Long,
    Double,
    Boolean,
    HostObject,
}

/// Borrowed during synchronous execution. Offset zero is the block budget.
/// Positive statuses are VBA errors; -1 is an uncatchable budget interruption.
#[repr(C)]
pub struct TypedContext {
    pub remaining_blocks: u64,
}
impl TypedContext {
    pub fn step(&mut self) -> Result<(), i32> {
        if self.remaining_blocks == 0 {
            return Err(-1);
        }
        self.remaining_blocks -= 1;
        Ok(())
    }
}

/// One physical eight-byte slot. Object pointers retain provenance while moving
/// through interpreter slots; layout admission determines how a slot is read.
#[derive(Clone, Copy)]
#[repr(C)]
pub union ScalarBits {
    bits: u64,
    pointer: *const crate::object::ObjectHeader,
}
impl Default for ScalarBits {
    fn default() -> Self {
        Self::ZERO
    }
}
impl std::fmt::Debug for ScalarBits {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // SAFETY: a raw pointer permits every initialized address bit pattern;
        // reading its address never dereferences it or exposes provenance.
        f.debug_tuple("ScalarBits")
            .field(&unsafe { self.pointer }.addr())
            .finish()
    }
}
impl ScalarBits {
    pub const ZERO: Self = Self { bits: 0 };
    pub const fn from_bits(bits: u64) -> Self {
        Self { bits }
    }
    pub fn object(pointer: *const crate::object::ObjectHeader) -> Self {
        let mut slot = Self::ZERO;
        slot.pointer = pointer;
        slot
    }
    pub fn as_object(self) -> *const crate::object::ObjectHeader {
        // SAFETY: all address bit patterns are valid raw pointers; callers must
        // prove the object layout and lifetime before any actual dereference.
        unsafe { self.pointer }
    }
    fn numeric_bits(self) -> u64 {
        #[cfg(target_pointer_width = "64")]
        {
            self.as_object().addr() as u64
        }
        #[cfg(not(target_pointer_width = "64"))]
        {
            // SAFETY: numeric constructors initialize all eight bytes. Object
            // construction starts with ZERO before storing the native pointer.
            unsafe { self.bits }
        }
    }
    pub fn long(value: i32) -> Self {
        Self::from_bits(value as u32 as u64)
    }
    pub fn double(value: f64) -> Self {
        Self::from_bits(value.to_bits())
    }
    pub fn as_long(self) -> i32 {
        self.numeric_bits() as i32
    }
    pub fn as_double(self) -> f64 {
        f64::from_bits(self.numeric_bits())
    }
}

/// Safe provider-facing slots preserve the object borrow for the whole call.
#[derive(Clone, Copy)]
pub enum Slot<'a> {
    Long(i32),
    Double(f64),
    Object(ObjectArg<'a>),
    /// A null object reference, with no borrowed owner.
    Nothing,
}
impl Slot<'_> {
    pub fn ty(self) -> ScalarType {
        match self {
            Self::Long(_) => ScalarType::Long,
            Self::Double(_) => ScalarType::Double,
            Self::Object(_) | Self::Nothing => ScalarType::HostObject,
        }
    }
    pub fn bits(self) -> ScalarBits {
        match self {
            Self::Long(v) => ScalarBits::long(v),
            Self::Double(v) => ScalarBits::double(v),
            Self::Object(v) => ScalarBits::object(v.as_ptr()),
            Self::Nothing => ScalarBits::object(std::ptr::null()),
        }
    }
    pub fn is_valid(self) -> bool {
        !matches!(self, Self::Double(v) if !v.is_finite())
    }
}

mod sealed {
    pub trait Sealed {}
}
/// Sealed primitives and a lifetime-bound host object borrow have a known C ABI.
pub trait Argument: sealed::Sealed + Copy {
    const TYPE: ScalarType;
    fn bits(self) -> ScalarBits;
    fn slot(&self) -> Slot<'_>;
    fn is_valid(self) -> bool {
        true
    }
}
pub trait Scalar: Argument + Default {
    fn from_bits(value: ScalarBits) -> Self;
}
impl sealed::Sealed for i32 {}
impl Argument for i32 {
    const TYPE: ScalarType = ScalarType::Long;
    fn bits(self) -> ScalarBits {
        ScalarBits::long(self)
    }
    fn slot(&self) -> Slot<'_> {
        Slot::Long(*self)
    }
}
impl Scalar for i32 {
    fn from_bits(value: ScalarBits) -> Self {
        value.as_long()
    }
}
impl sealed::Sealed for f64 {}
impl Argument for f64 {
    const TYPE: ScalarType = ScalarType::Double;
    fn bits(self) -> ScalarBits {
        ScalarBits::double(self)
    }
    fn slot(&self) -> Slot<'_> {
        Slot::Double(*self)
    }
    fn is_valid(self) -> bool {
        self.is_finite()
    }
}
impl Scalar for f64 {
    fn from_bits(value: ScalarBits) -> Self {
        value.as_double()
    }
}
impl sealed::Sealed for ObjectArg<'_> {}
impl Argument for ObjectArg<'_> {
    const TYPE: ScalarType = ScalarType::HostObject;
    fn bits(self) -> ScalarBits {
        Slot::Object(self).bits()
    }
    fn slot(&self) -> Slot<'_> {
        Slot::Object(*self)
    }
}

pub trait Arguments: sealed::Sealed + Copy {
    const COUNT: usize;
    fn types() -> Vec<ScalarType>;
    fn bits(self) -> [ScalarBits; MAX_TYPED_ARGS];
    fn slots(&self) -> [Slot<'_>; MAX_TYPED_ARGS];
    fn is_valid(self) -> bool;
    /// # Safety
    /// Entry has exactly this platform C signature and is live throughout the
    /// call. The provider checked its signature, object contracts and session.
    unsafe fn invoke<R: Scalar>(
        self,
        entry: *const u8,
        context: &mut TypedContext,
        result: &mut R,
    ) -> i32;
}
macro_rules! arguments {
    ($($ty:ident:$idx:tt),*) => {
        impl<$($ty: Argument),*> sealed::Sealed for ($($ty,)*) {}
        impl<$($ty: Argument),*> Arguments for ($($ty,)*) {
            const COUNT: usize = 0 $(+ { let _ = stringify!($ty); 1 })*;
            fn types() -> Vec<ScalarType> { vec![$($ty::TYPE),*] }
            fn is_valid(self) -> bool { true $(&& self.$idx.is_valid())* }
            fn bits(self) -> [ScalarBits; MAX_TYPED_ARGS] {
                #[allow(unused_mut)]
                let mut bits = [ScalarBits::ZERO; MAX_TYPED_ARGS];
                $(bits[$idx] = self.$idx.bits();)* bits
            }
            fn slots(&self) -> [Slot<'_>; MAX_TYPED_ARGS] {
                #[allow(unused_mut)]
                let mut slots = [Slot::Long(0); MAX_TYPED_ARGS];
                $(slots[$idx] = self.$idx.slot();)* slots
            }
            unsafe fn invoke<R: Scalar>(self, entry: *const u8, context: &mut TypedContext, result: &mut R) -> i32 {
                // SAFETY: provider proves the exact signature. Sealed arguments
                // are i32/f64 or repr(transparent) borrowed object pointers.
                let entry: unsafe extern "C" fn(*mut TypedContext, $($ty,)* *mut R) -> i32 = unsafe { std::mem::transmute(entry) };
                // SAFETY: provider and Rust borrows retain code, context and objects.
                unsafe { entry(context, $(self.$idx,)* result) }
            }
        }
    }
}
arguments!();
arguments!(A:0);
arguments!(A:0, B:1);
arguments!(A:0, B:1, C:2);
arguments!(A:0, B:1, C:2, D:3);
