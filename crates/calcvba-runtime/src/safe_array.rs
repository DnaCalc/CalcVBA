use crate::{
    Decimal96, VarType, Variant, VbaRecord, VbaRecordLayout,
    bstr::{BStr, borrow_raw_bstr},
    object_ref::{ObjectRef, RawRuntimeIUnknown},
};
use core::ptr::NonNull;
use std::sync::Arc;

pub const ARRAY_TAG_BASE: i32 = -1_000_000_000;
pub const ARRAY_TAG_LIMIT: i32 = ARRAY_TAG_BASE + 1_000_000;
const DISPATCH_ARRAY_PAYLOAD_BASE: i32 = 20_000;

pub const VT_I2_VALUE: u16 = 0x0002;
pub const VT_I4_VALUE: u16 = 0x0003;
pub const VT_R4_VALUE: u16 = 0x0004;
pub const VT_R8_VALUE: u16 = 0x0005;
pub const VT_CY_VALUE: u16 = 0x0006;
pub const VT_DATE_VALUE: u16 = 0x0007;
pub const VT_BSTR_VALUE: u16 = 0x0008;
pub const VT_DISPATCH_VALUE: u16 = 0x0009;
pub const VT_BOOL_VALUE: u16 = 0x000B;
pub const VT_VARIANT_VALUE: u16 = 0x000C;
pub const VT_UNKNOWN_VALUE: u16 = 0x000D;
pub const VT_DECIMAL_VALUE: u16 = 0x000E;
pub const VT_I1_VALUE: u16 = 0x0010;
pub const VT_UI1_VALUE: u16 = 0x0011;
pub const VT_UI2_VALUE: u16 = 0x0012;
pub const VT_UI4_VALUE: u16 = 0x0013;
pub const VT_I8_VALUE: u16 = 0x0014;
pub const VT_UI8_VALUE: u16 = 0x0015;
pub const VT_INT_VALUE: u16 = 0x0016;
pub const VT_UINT_VALUE: u16 = 0x0017;
pub const VT_RECORD_VALUE: u16 = 0x0024;

/// `FADF_STATIC` (0x0002): the payload is not the allocator's to free. VBA sets it, with
/// `FADF_FIXEDSIZE`, on every fixed-size array.
pub const FADF_STATIC_VALUE: u16 = 0x0002;
/// `FADF_FIXEDSIZE` (0x0010): the SAFEARRAY may not be resized or reallocated.
/// VBA sets this on a fixed-size array (`Dim a(1 To 3)`), where `Erase`
/// reinitializes each element to its type default and keeps the array, in
/// contrast to a dynamic array (`Dim a()` + `ReDim`) whose `Erase` deallocates.
pub const FADF_FIXEDSIZE_VALUE: u16 = 0x0010;
pub const FADF_RECORD_VALUE: u16 = 0x0020;
/// `FADF_HAVEIID` (0x0040): the interface IID is in the 16 bytes before the descriptor.
pub const FADF_HAVEIID_VALUE: u16 = 0x0040;
pub const FADF_HAVEVARTYPE_VALUE: u16 = 0x0080;
pub const FADF_BSTR_VALUE: u16 = 0x0100;
pub const FADF_UNKNOWN_VALUE: u16 = 0x0200;
pub const FADF_DISPATCH_VALUE: u16 = 0x0400;
pub const FADF_VARIANT_VALUE: u16 = 0x0800;
/// The bits VBA sets on a fixed-size array (`&H12`, measured in Excel 2026-11-14).
const FADF_FIXED_ARRAY: u16 = FADF_STATIC_VALUE | FADF_FIXEDSIZE_VALUE;

/// The bytes oleaut32 keeps before every descriptor: the IID (`FADF_HAVEIID`), the
/// `IRecordInfo*` in the last pointer (`FADF_RECORD`), or the `VARTYPE` in the last four bytes
/// (`FADF_HAVEVARTYPE`).
const HIDDEN_SIZE: usize = 16;
const IID_IUNKNOWN: [u8; 16] = [0, 0, 0, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46];
const IID_IDISPATCH: [u8; 16] = [0, 4, 2, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46];

#[cfg(test)]
std::thread_local! {
    static PAYLOAD_ALLOCATION_EVENTS: core::cell::Cell<(usize, usize)> = const {
        core::cell::Cell::new((0, 0))
    };
}

#[cfg(test)]
fn note_payload_allocation() {
    PAYLOAD_ALLOCATION_EVENTS.with(|events| {
        let (allocated, freed) = events.get();
        events.set((allocated + 1, freed));
    });
}

#[cfg(test)]
fn note_payload_free() {
    PAYLOAD_ALLOCATION_EVENTS.with(|events| {
        let (allocated, freed) = events.get();
        events.set((allocated, freed + 1));
    });
}

#[cfg(test)]
fn payload_allocation_events() -> (usize, usize) {
    PAYLOAD_ALLOCATION_EVENTS.with(core::cell::Cell::get)
}

#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SafeArrayBound {
    pub count: u32,
    pub lower: i32,
}

#[repr(C)]
struct RawSafeArray {
    c_dims: u16,
    f_features: u16,
    cb_elements: u32,
    c_locks: u32,
    pv_data: *mut core::ffi::c_void,
    rgsabound: [SafeArrayBound; 1],
}

#[repr(C)]
#[derive(Clone, Copy)]
struct RawDecimalArrayElement {
    reserved: u16,
    scale: u8,
    sign: u8,
    hi32: u32,
    lo32: u32,
    mid32: u32,
}

impl RawDecimalArrayElement {
    fn from_decimal96(value: Decimal96) -> Self {
        Self {
            reserved: 0,
            scale: value.scale(),
            sign: if value.is_negative() { 0x80 } else { 0 },
            hi32: value.hi,
            lo32: value.lo,
            mid32: value.mid,
        }
    }

    fn to_decimal96(self) -> Decimal96 {
        Decimal96::from_scale_sign(
            self.lo32,
            self.mid32,
            self.hi32,
            (u16::from(self.sign) << 8) | u16::from(self.scale),
        )
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SafeArrayElementKind {
    Variant,
    I1,
    Ui1,
    I2,
    Ui2,
    I4,
    Ui4,
    I8,
    Ui8,
    Int,
    UInt,
    R4,
    R8,
    Currency,
    Date,
    Bool,
    BStr,
    Dispatch,
    Unknown,
    Decimal,
}

impl SafeArrayElementKind {
    fn from_vartype(value: u16) -> Option<Self> {
        match value {
            VT_VARIANT_VALUE => Some(Self::Variant),
            VT_I1_VALUE => Some(Self::I1),
            VT_UI1_VALUE => Some(Self::Ui1),
            VT_I2_VALUE => Some(Self::I2),
            VT_UI2_VALUE => Some(Self::Ui2),
            VT_I4_VALUE => Some(Self::I4),
            VT_UI4_VALUE => Some(Self::Ui4),
            VT_I8_VALUE => Some(Self::I8),
            VT_UI8_VALUE => Some(Self::Ui8),
            VT_INT_VALUE => Some(Self::Int),
            VT_UINT_VALUE => Some(Self::UInt),
            VT_R4_VALUE => Some(Self::R4),
            VT_R8_VALUE => Some(Self::R8),
            VT_CY_VALUE => Some(Self::Currency),
            VT_DATE_VALUE => Some(Self::Date),
            VT_BOOL_VALUE => Some(Self::Bool),
            VT_BSTR_VALUE => Some(Self::BStr),
            VT_DISPATCH_VALUE => Some(Self::Dispatch),
            VT_UNKNOWN_VALUE => Some(Self::Unknown),
            VT_DECIMAL_VALUE => Some(Self::Decimal),
            _ => None,
        }
    }

    fn element_size(self) -> usize {
        match self {
            Self::Variant => core::mem::size_of::<Variant>(),
            Self::I1 | Self::Ui1 => 1,
            Self::I2 | Self::Ui2 | Self::Bool => 2,
            Self::I4 | Self::Ui4 | Self::Int | Self::UInt | Self::R4 => 4,
            Self::I8 | Self::Ui8 | Self::R8 | Self::Currency | Self::Date => 8,
            Self::BStr | Self::Dispatch | Self::Unknown => core::mem::size_of::<*mut u8>(),
            Self::Decimal => core::mem::size_of::<RawDecimalArrayElement>(),
        }
    }

    /// Whether elements own resources (strings, objects, Variants) that must be released
    /// before the payload is freed.
    fn owns_resources(self) -> bool {
        matches!(
            self,
            Self::Variant | Self::BStr | Self::Dispatch | Self::Unknown
        )
    }

    /// `fFeatures` as `SafeArrayAllocDescriptorEx` and VBA set them for this element type
    /// (Excel 2026-11-14: Long `&H80`, String `&H180`, Variant `&H880`, Object `&H440`).
    fn feature_flags(self) -> u16 {
        match self {
            Self::Variant => FADF_HAVEVARTYPE_VALUE | FADF_VARIANT_VALUE,
            Self::BStr => FADF_HAVEVARTYPE_VALUE | FADF_BSTR_VALUE,
            Self::Dispatch => FADF_HAVEIID_VALUE | FADF_DISPATCH_VALUE,
            Self::Unknown => FADF_HAVEIID_VALUE | FADF_UNKNOWN_VALUE,
            _ => FADF_HAVEVARTYPE_VALUE,
        }
    }
}

/// What a descriptor's elements are: an Automation type, or a VBA `Type` whose layout the
/// declaration supplies (VBA keeps no layout, vt or `IRecordInfo` in a UDT array's descriptor).
#[derive(Clone)]
enum ElementType {
    Intrinsic(SafeArrayElementKind),
    Record(Arc<VbaRecordLayout>),
}

impl ElementType {
    fn element_size(&self) -> usize {
        match self {
            Self::Intrinsic(kind) => kind.element_size(),
            Self::Record(layout) => layout.size(),
        }
    }
}

/// An array descriptor as oleaut32 lays it out: the header, `rgsabound` with the last
/// dimension first, and the 16 hidden bytes before it. On Windows oleaut32 allocates and
/// frees descriptors and payloads, so foreign code may `SafeArrayRedim` or destroy them
/// exactly as it may VBA's; elsewhere [`sys`] provides the same layout.
///
/// A UDT array's record layout travels with the handle (and with the cell holding it), not in
/// the descriptor: VBA knows it from the declaration.
pub struct SafeArray {
    header: NonNull<RawSafeArray>,
    record: Option<Arc<VbaRecordLayout>>,
}

// SAFETY: ASSUMPTION — `SafeArray` exclusively owns its descriptor and payload
// (allocated in `allocate`, freed only in `Drop`), and scalar/BStr elements are plain data or
// independently owned allocations, so those move freely between threads. Object
// (VT_DISPATCH/VT_UNKNOWN) elements hold a retained reference whose
// AddRef/Release counter is an `AtomicU32`, but the FINAL release parks the
// object in a `thread_local!` termination queue and touches `Cell` state
// (`compat_release` in object_ref.rs), so soundness additionally requires that
// arrays holding object elements are dropped on the VM thread that manages
// those objects' lifecycle.
unsafe impl Send for SafeArray {}
// SAFETY: no `&self` method mutates the descriptor or payload (this module has
// no interior mutability; `c_locks` is never toggled after construction), so
// shared access performs only reads plus atomic AddRefs when object elements
// are cloned out — data-race-free.
unsafe impl Sync for SafeArray {}

unsafe fn free_raw_bstr(ptr: *mut u16) {
    // SAFETY: the caller transfers ownership of `ptr` (null or the raw BSTR a
    // SAFEARRAY slot owned, written there by `encode_element_variant`), which
    // is exactly `BStr::from_raw_bstr`'s contract; dropping the wrapper frees
    // the allocation exactly once.
    let _ = unsafe { BStr::from_raw_bstr(ptr) };
}

/// Largest SAFEARRAY rank: the `cDims` field of the Automation SAFEARRAY is 16 bits, so the
/// rank is stored exactly and a larger rank is rejected before any allocation. VBA source
/// limits (60 dimensions) are enforced above this layer.
pub const MAX_SAFEARRAY_RANK: usize = u16::MAX as usize;

/// The size of a header with `dimensions` bounds, as `SafeArrayAllocDescriptor` allocates it.
#[cfg(not(windows))]
fn header_size(dimensions: usize) -> Result<usize, String> {
    let header = std::alloc::Layout::new::<RawSafeArray>();
    let extra = std::alloc::Layout::array::<SafeArrayBound>(dimensions.saturating_sub(1))
        .map_err(|_| "SAFEARRAY bounds layout overflow".to_string())?;
    header
        .extend(extra)
        .map(|(layout, _)| layout.pad_to_align().size())
        .map_err(|_| "SAFEARRAY header layout overflow".to_string())
}

/// Descriptor and payload allocation. On Windows this is oleaut32 itself; elsewhere the same
/// layout (16 hidden bytes, then the header) from the Rust allocator, with a size word in front
/// of each block so a free needs nothing from the (mutable) descriptor.
#[cfg(windows)]
mod sys {
    use super::RawSafeArray;
    use windows_sys::Win32::System::Ole::{
        SafeArrayAllocData, SafeArrayAllocDescriptorEx, SafeArrayDestroy,
    };

    /// A descriptor of `dims` dimensions with its hidden bytes, `cDims` set. A UDT array (`vt`
    /// `None`) is allocated as a Byte array, which reserves the same hidden bytes; the caller
    /// rewrites `fFeatures`, `cbElements` and the hidden bytes either way.
    pub(super) fn alloc_descriptor(dims: u16, vt: Option<u16>) -> Option<*mut RawSafeArray> {
        let mut psa = core::ptr::null_mut();
        // SAFETY: `psa` is a valid out pointer; the call allocates a descriptor with `dims`
        // bounds and its hidden prefix, or fails without writing it.
        let hr = unsafe {
            SafeArrayAllocDescriptorEx(vt.unwrap_or(super::VT_UI1_VALUE), u32::from(dims), &mut psa)
        };
        (hr >= 0 && !psa.is_null()).then_some(psa.cast())
    }

    /// Allocates `cbElements` × the cell count for a descriptor whose bounds are set.
    ///
    /// # Safety
    /// `header` is a live descriptor from [`alloc_descriptor`] with null `pvData`.
    pub(super) unsafe fn alloc_data(header: *mut RawSafeArray) -> bool {
        // SAFETY: forwarded from the caller.
        unsafe { SafeArrayAllocData(header.cast()) >= 0 }
    }

    /// Frees the payload and the descriptor.
    ///
    /// # Safety
    /// `header` is a live, unlocked descriptor whose owning elements were already released
    /// (and zeroed), with `FADF_STATIC` clear so oleaut32 frees the payload.
    pub(super) unsafe fn destroy(header: *mut RawSafeArray) {
        // SAFETY: forwarded from the caller; nothing touches the descriptor afterwards.
        let _ = unsafe { SafeArrayDestroy(header.cast()) };
    }
}

#[cfg(not(windows))]
mod sys {
    use super::{HIDDEN_SIZE, RawSafeArray, header_size};

    /// Every block starts with its size, then pads to this alignment.
    const BLOCK_ALIGN: usize = 16;

    fn alloc_block(size: usize) -> Option<*mut u8> {
        let total = size.checked_add(BLOCK_ALIGN)?;
        let layout = std::alloc::Layout::from_size_align(total, BLOCK_ALIGN).ok()?;
        // SAFETY: `layout` has non-zero size.
        let base = unsafe { std::alloc::alloc_zeroed(layout) };
        if base.is_null() {
            return None;
        }
        // SAFETY: `base` has at least `BLOCK_ALIGN` bytes, aligned for `usize`.
        unsafe {
            base.cast::<usize>().write(total);
            Some(base.add(BLOCK_ALIGN))
        }
    }

    /// # Safety
    /// `block` came from [`alloc_block`] and is freed once.
    unsafe fn free_block(block: *mut u8) {
        // SAFETY: `alloc_block` returned `base + BLOCK_ALIGN` and stored the total at `base`.
        unsafe {
            let base = block.sub(BLOCK_ALIGN);
            let total = base.cast::<usize>().read();
            std::alloc::dealloc(
                base,
                std::alloc::Layout::from_size_align_unchecked(total, BLOCK_ALIGN),
            );
        }
    }

    pub(super) fn alloc_descriptor(dims: u16, _vt: Option<u16>) -> Option<*mut RawSafeArray> {
        let size = HIDDEN_SIZE.checked_add(header_size(usize::from(dims)).ok()?)?;
        let block = alloc_block(size)?;
        // SAFETY: the block holds the hidden bytes and then the header, 16-byte aligned.
        unsafe {
            let header = block.add(HIDDEN_SIZE).cast::<RawSafeArray>();
            (*header).c_dims = dims;
            Some(header)
        }
    }

    /// # Safety
    /// `header` is a live descriptor from [`alloc_descriptor`] with its bounds set and null
    /// `pvData`.
    pub(super) unsafe fn alloc_data(header: *mut RawSafeArray) -> bool {
        // SAFETY: forwarded from the caller.
        let (count, size) = unsafe {
            let bounds = super::raw_bounds_slice(header);
            (
                super::bounds_total_len(bounds).ok(),
                (*header).cb_elements as usize,
            )
        };
        let Some(bytes) = count.and_then(|count| count.checked_mul(size)) else {
            return false;
        };
        let Some(data) = alloc_block(bytes) else {
            return false;
        };
        // SAFETY: forwarded from the caller.
        unsafe { (*header).pv_data = data.cast() };
        true
    }

    /// # Safety
    /// As the Windows `destroy`: a live descriptor whose owning elements were released.
    pub(super) unsafe fn destroy(header: *mut RawSafeArray) {
        // SAFETY: both blocks came from `alloc_block` and are freed once here.
        unsafe {
            let data = (*header).pv_data.cast::<u8>();
            if !data.is_null() {
                free_block(data);
            }
            free_block(header.cast::<u8>().sub(HIDDEN_SIZE));
        }
    }
}

fn default_bounds_for_len(len: usize) -> Result<Vec<SafeArrayBound>, String> {
    Ok(vec![SafeArrayBound {
        count: u32::try_from(len).map_err(|_| {
            format!("SAFEARRAY length {len} exceeds supported u32 element capacity")
        })?,
        lower: 0,
    }])
}

fn bounds_total_len(bounds: &[SafeArrayBound]) -> Result<usize, String> {
    // Reject a hostile rank before any payload or descriptor is allocated.
    if bounds.len() > MAX_SAFEARRAY_RANK {
        return Err(format!(
            "SAFEARRAY rank {} exceeds the {MAX_SAFEARRAY_RANK}-dimension limit",
            bounds.len()
        ));
    }
    if bounds.is_empty() {
        return Ok(0);
    }
    let mut total = 1usize;
    for bound in bounds {
        total = total
            .checked_mul(bound.count as usize)
            .ok_or_else(|| "SAFEARRAY total element count overflowed".to_string())?;
    }
    Ok(total)
}

/// The descriptor's `rgsabound` as stored: the last dimension first.
unsafe fn raw_bounds_slice<'a>(header: *const RawSafeArray) -> &'a [SafeArrayBound] {
    // SAFETY: callers pass a live descriptor. Reading c_dims and forming the slice over the
    // inline rgsabound tail mirrors `SafeArray::raw_bounds` without allocating.
    let dims = unsafe { (*header).c_dims as usize };
    if dims == 0 {
        return &[];
    }
    // SAFETY: a descriptor has exactly `dims` bounds at rgsabound.
    let ptr = unsafe { core::ptr::addr_of!((*header).rgsabound).cast::<SafeArrayBound>() };
    // SAFETY: `ptr` addresses the first of `dims` contiguous bounds in the live descriptor.
    unsafe { core::slice::from_raw_parts(ptr, dims) }
}

/// The hidden bytes before a descriptor.
fn hidden_bytes(header: *const RawSafeArray) -> *mut u8 {
    // SAFETY: every descriptor has `HIDDEN_SIZE` bytes before it in the same allocation
    // (`SafeArrayAllocDescriptor[Ex]`, or `sys::alloc_descriptor`).
    unsafe { header.cast::<u8>().cast_mut().sub(HIDDEN_SIZE) }
}

/// The element `VARTYPE` as `SafeArrayGetVartype` derives it, or `None` for a descriptor that
/// names none (a VBA UDT array, or a foreign array of an unknown kind).
///
/// # Safety
/// `header` is a live descriptor.
unsafe fn descriptor_vartype(header: *const RawSafeArray) -> Option<u16> {
    // SAFETY: forwarded from the caller.
    let features = unsafe { (*header).f_features };
    if features & FADF_RECORD_VALUE != 0 {
        Some(VT_RECORD_VALUE)
    } else if features & FADF_HAVEIID_VALUE != 0 || features & FADF_DISPATCH_VALUE != 0 {
        Some(if features & FADF_DISPATCH_VALUE != 0 {
            VT_DISPATCH_VALUE
        } else {
            VT_UNKNOWN_VALUE
        })
    } else if features & FADF_VARIANT_VALUE != 0 {
        Some(VT_VARIANT_VALUE)
    } else if features & FADF_BSTR_VALUE != 0 {
        Some(VT_BSTR_VALUE)
    } else if features & FADF_UNKNOWN_VALUE != 0 {
        Some(VT_UNKNOWN_VALUE)
    } else if features & FADF_HAVEVARTYPE_VALUE != 0 {
        // SAFETY: a `FADF_HAVEVARTYPE` descriptor keeps its VARTYPE in the last four hidden
        // bytes.
        Some(unsafe {
            hidden_bytes(header)
                .add(HIDDEN_SIZE - 4)
                .cast::<u32>()
                .read_unaligned()
        } as u16)
    } else {
        None
    }
}

fn payload_offset(kind: SafeArrayElementKind, index: usize) -> usize {
    kind.element_size() * index
}

fn record_payload_offset(layout: &VbaRecordLayout, index: usize) -> usize {
    layout.size() * index
}

/// The owned interface pointer an object element stores: null for `Nothing`.
fn object_element(value: &Variant) -> Result<*mut RawRuntimeIUnknown, String> {
    if value.vtype() != VarType::Object {
        return Err(format!("expected Object SAFEARRAY element, got {value:?}"));
    }
    Ok(value
        .as_object_ref()
        .map_or(core::ptr::null_mut(), ObjectRef::into_raw_iunknown))
}

fn variant_i64(value: &Variant) -> Result<i64, String> {
    if matches!(value.vtype(), VarType::Empty) {
        return Ok(0);
    }
    value
        .as_i64()
        .or_else(|| value.as_i32().map(i64::from))
        .or_else(|| value.as_i16().map(i64::from))
        .or_else(|| value.as_u8().map(i64::from))
        .ok_or_else(|| format!("expected integer-compatible SAFEARRAY element, got {value:?}"))
}

fn variant_f64(value: &Variant) -> Result<f64, String> {
    value
        .as_f64()
        .or_else(|| value.as_f32().map(f64::from))
        .or_else(|| value.as_date_f64())
        .ok_or_else(|| format!("expected floating-point SAFEARRAY element, got {value:?}"))
}

unsafe fn decode_element_variant(
    kind: SafeArrayElementKind,
    payload: *const u8,
    index: usize,
) -> Result<Variant, String> {
    // SAFETY: the caller guarantees `payload` is the live element buffer built
    // by `prepare_payload_from_variants` for this same `kind`, holding more than
    // `index` fully initialized elements, so the offset stays inside that
    // allocation and `ptr` is aligned for the element type matched below.
    let ptr = unsafe { payload.add(payload_offset(kind, index)) };
    Ok(match kind {
        // SAFETY: kind == Variant, so this slot was initialized as a `Variant`
        // at encode time; it is only borrowed here long enough to clone.
        SafeArrayElementKind::Variant => unsafe { &*ptr.cast::<Variant>() }.try_clone()?,
        // SAFETY: kind == I1, so encode initialized this slot as an i8.
        SafeArrayElementKind::I1 => Variant::from_i16(i16::from(unsafe { *ptr.cast::<i8>() })),
        // SAFETY: kind == Ui1, so encode initialized this slot as a u8.
        SafeArrayElementKind::Ui1 => Variant::from_u8(unsafe { *ptr.cast::<u8>() }),
        // SAFETY: kind == I2, so encode initialized this slot as an i16.
        SafeArrayElementKind::I2 => Variant::from_i16(unsafe { *ptr.cast::<i16>() }),
        // SAFETY: kind == Ui2, so encode initialized this slot as a u16.
        SafeArrayElementKind::Ui2 => Variant::from_i32(i32::from(unsafe { *ptr.cast::<u16>() })),
        SafeArrayElementKind::I4 | SafeArrayElementKind::Int => {
            // SAFETY: kind == I4/Int, so encode initialized this slot as an i32.
            Variant::from_i32(unsafe { *ptr.cast::<i32>() })
        }
        SafeArrayElementKind::Ui4 | SafeArrayElementKind::UInt => {
            // SAFETY: kind == Ui4/UInt, so encode initialized this slot as a u32.
            Variant::from_i64(i64::from(unsafe { *ptr.cast::<u32>() }))
        }
        // SAFETY: kind == I8, so encode initialized this slot as an i64.
        SafeArrayElementKind::I8 => Variant::from_i64(unsafe { *ptr.cast::<i64>() }),
        SafeArrayElementKind::Ui8 => {
            // SAFETY: kind == Ui8, so encode initialized this slot as a u64.
            let value = unsafe { *ptr.cast::<u64>() };
            Variant::from_i64(i64::try_from(value).map_err(|_| {
                format!("VT_UI8 SAFEARRAY element {value} exceeds i64 carrier range")
            })?)
        }
        // SAFETY: kind == R4, so encode initialized this slot as an f32.
        SafeArrayElementKind::R4 => Variant::from_f32(unsafe { *ptr.cast::<f32>() }),
        // SAFETY: kind == R8, so encode initialized this slot as an f64.
        SafeArrayElementKind::R8 => Variant::from_f64(unsafe { *ptr.cast::<f64>() }),
        SafeArrayElementKind::Currency => {
            // SAFETY: kind == Currency, so encode initialized this slot as the
            // scaled-i64 currency carrier.
            Variant::from_currency_scaled_i64(unsafe { *ptr.cast::<i64>() })
        }
        // SAFETY: kind == Date, so encode initialized this slot as an f64 date serial.
        SafeArrayElementKind::Date => Variant::from_date_f64(unsafe { *ptr.cast::<f64>() }),
        // SAFETY: kind == Bool, so encode initialized this slot as a VARIANT_BOOL i16.
        SafeArrayElementKind::Bool => Variant::from_bool(unsafe { *ptr.cast::<i16>() } != 0),
        SafeArrayElementKind::BStr => {
            // SAFETY: kind == BStr, so this slot holds the raw BSTR pointer the
            // array owns. BorrowedBStr has no destructor, so cloning cannot
            // consume the slot's allocation on error or unwind.
            Variant::from_string(
                unsafe { borrow_raw_bstr(*ptr.cast::<*mut u16>()) }.try_to_owned()?,
            )
        }
        SafeArrayElementKind::Dispatch | SafeArrayElementKind::Unknown => {
            // SAFETY: kind == Dispatch/Unknown, so the slot holds null (Nothing) or an
            // interface pointer the array holds one reference on, live here;
            // `from_raw_iunknown_addref` takes an additional independent reference.
            match unsafe {
                ObjectRef::from_raw_iunknown_addref(*ptr.cast::<*mut RawRuntimeIUnknown>())
            } {
                Some(object) => Variant::from_object_ref(object),
                None => Variant::nothing(),
            }
        }
        SafeArrayElementKind::Decimal => {
            // SAFETY: kind == Decimal, so encode initialized this slot as a
            // plain-data RawDecimalArrayElement.
            Variant::from_decimal96(unsafe { *ptr.cast::<RawDecimalArrayElement>() }.to_decimal96())
        }
    })
}

unsafe fn encode_element_variant(
    kind: SafeArrayElementKind,
    payload: *mut u8,
    index: usize,
    value: &Variant,
) -> Result<(), String> {
    // SAFETY: the caller guarantees `payload` is an allocation laid out by
    // `payload_layout(kind, len)` with `index < len`, so the offset stays
    // in-bounds and `ptr` is aligned for `kind`'s element type.
    let ptr = unsafe { payload.add(payload_offset(kind, index)) };
    if kind == SafeArrayElementKind::Variant {
        // SAFETY: kind == Variant, so the in-bounds slot computed above has
        // `Variant` size/alignment; the slot is written exactly once into the
        // fresh zeroed buffer, so `write` correctly skips dropping old contents.
        unsafe { ptr.cast::<Variant>().write(value.try_clone()?) };
        return Ok(());
    }
    match kind {
        SafeArrayElementKind::Variant => unreachable!("handled above"),
        // SAFETY: kind == I1 sized/aligned the slot for i8; plain-data write,
        // exactly once into the zeroed buffer.
        SafeArrayElementKind::I1 => unsafe {
            ptr.cast::<i8>()
                .write(i8::try_from(variant_i64(value)?).map_err(|_| {
                    format!("value {value:?} does not fit VT_I1 SAFEARRAY element")
                })?);
        },
        // SAFETY: kind == Ui1 sized/aligned the slot for u8; plain-data write,
        // exactly once into the zeroed buffer.
        SafeArrayElementKind::Ui1 => unsafe {
            ptr.cast::<u8>().write(
                u8::try_from(variant_i64(value)?).map_err(|_| {
                    format!("value {value:?} does not fit VT_UI1 SAFEARRAY element")
                })?,
            );
        },
        // SAFETY: kind == I2 sized/aligned the slot for i16; plain-data write,
        // exactly once into the zeroed buffer.
        SafeArrayElementKind::I2 => unsafe {
            ptr.cast::<i16>()
                .write(i16::try_from(variant_i64(value)?).map_err(|_| {
                    format!("value {value:?} does not fit VT_I2 SAFEARRAY element")
                })?);
        },
        // SAFETY: kind == Ui2 sized/aligned the slot for u16; plain-data write,
        // exactly once into the zeroed buffer.
        SafeArrayElementKind::Ui2 => unsafe {
            ptr.cast::<u16>().write(
                u16::try_from(variant_i64(value)?).map_err(|_| {
                    format!("value {value:?} does not fit VT_UI2 SAFEARRAY element")
                })?,
            );
        },
        // SAFETY: kind == I4/Int sized/aligned the slot for i32; plain-data
        // write, exactly once into the zeroed buffer.
        SafeArrayElementKind::I4 | SafeArrayElementKind::Int => unsafe {
            ptr.cast::<i32>()
                .write(i32::try_from(variant_i64(value)?).map_err(|_| {
                    format!("value {value:?} does not fit VT_I4 SAFEARRAY element")
                })?);
        },
        // SAFETY: kind == Ui4/UInt sized/aligned the slot for u32; plain-data
        // write, exactly once into the zeroed buffer.
        SafeArrayElementKind::Ui4 | SafeArrayElementKind::UInt => unsafe {
            ptr.cast::<u32>().write(
                u32::try_from(variant_i64(value)?).map_err(|_| {
                    format!("value {value:?} does not fit VT_UI4 SAFEARRAY element")
                })?,
            );
        },
        // SAFETY: kind == I8 sized/aligned the slot for i64; plain-data write,
        // exactly once into the zeroed buffer.
        SafeArrayElementKind::I8 => unsafe {
            ptr.cast::<i64>().write(variant_i64(value)?);
        },
        // SAFETY: kind == Ui8 sized/aligned the slot for u64; plain-data write,
        // exactly once into the zeroed buffer.
        SafeArrayElementKind::Ui8 => unsafe {
            ptr.cast::<u64>().write(
                u64::try_from(variant_i64(value)?).map_err(|_| {
                    format!("value {value:?} does not fit VT_UI8 SAFEARRAY element")
                })?,
            );
        },
        // SAFETY: kind == R4 sized/aligned the slot for f32; plain-data write,
        // exactly once into the zeroed buffer.
        SafeArrayElementKind::R4 => unsafe {
            ptr.cast::<f32>().write(variant_f64(value)? as f32);
        },
        // SAFETY: kind == R8 sized/aligned the slot for f64; plain-data write,
        // exactly once into the zeroed buffer.
        SafeArrayElementKind::R8 => unsafe {
            ptr.cast::<f64>().write(variant_f64(value)?);
        },
        // SAFETY: kind == Currency sized/aligned the slot for the scaled-i64
        // currency carrier; plain-data write, exactly once into the zeroed buffer.
        SafeArrayElementKind::Currency => unsafe {
            ptr.cast::<i64>().write(
                value
                    .as_currency_scaled_i64()
                    .ok_or_else(|| format!("expected Currency SAFEARRAY element, got {value:?}"))?,
            );
        },
        // SAFETY: kind == Date sized/aligned the slot for the f64 date serial;
        // plain-data write, exactly once into the zeroed buffer.
        SafeArrayElementKind::Date => unsafe {
            ptr.cast::<f64>().write(
                value
                    .as_date_f64()
                    .or_else(|| value.as_f64())
                    .ok_or_else(|| format!("expected Date SAFEARRAY element, got {value:?}"))?,
            );
        },
        // SAFETY: kind == Bool sized/aligned the slot for the VARIANT_BOOL i16;
        // plain-data write, exactly once into the zeroed buffer.
        SafeArrayElementKind::Bool => unsafe {
            ptr.cast::<i16>().write(
                if value
                    .as_bool()
                    .ok_or_else(|| format!("expected Bool SAFEARRAY element, got {value:?}"))?
                {
                    -1
                } else {
                    0
                },
            );
        },
        // SAFETY: kind == BStr sized/aligned the slot for a pointer; ownership
        // of the freshly allocated raw BSTR transfers into the slot, to be
        // freed exactly once by `drop_element`.
        SafeArrayElementKind::BStr => unsafe {
            let text = value
                .try_as_bstr()?
                .ok_or_else(|| format!("expected String SAFEARRAY element, got {value:?}"))?;
            ptr.cast::<*mut u16>().write(text.into_raw_bstr());
        },
        // SAFETY: kind == Dispatch/Unknown sized/aligned the slot for a pointer; the
        // element's retained reference (or null for Nothing) transfers into the slot, so the
        // array owns exactly one reference, released by `drop_element`.
        SafeArrayElementKind::Dispatch | SafeArrayElementKind::Unknown => unsafe {
            ptr.cast::<*mut RawRuntimeIUnknown>()
                .write(object_element(value)?);
        },
        // SAFETY: kind == Decimal sized/aligned the slot for the plain-data
        // RawDecimalArrayElement; written exactly once into the zeroed buffer.
        SafeArrayElementKind::Decimal => unsafe {
            ptr.cast::<RawDecimalArrayElement>()
                .write(RawDecimalArrayElement::from_decimal96(
                    value.as_decimal96().ok_or_else(|| {
                        format!("expected Decimal SAFEARRAY element, got {value:?}")
                    })?,
                ));
        },
    }
    Ok(())
}

/// # Safety
/// `payload` must point at this array's element buffer holding more than `index`
/// initialized elements of `kind`, so the `index`-th slot is live and writable.
unsafe fn replace_element_variant(
    kind: SafeArrayElementKind,
    payload: *mut u8,
    index: usize,
    value: &Variant,
) -> Result<(), String> {
    // SAFETY: `payload_offset` is `element_size(kind) * index`, and the contract
    // guarantees `index` is in range, so this lands on the `index`-th element slot.
    let ptr = unsafe { payload.add(payload_offset(kind, index)) };
    match kind {
        SafeArrayElementKind::Variant => {
            // SAFETY: kind == Variant, so the slot holds an initialized `Variant`;
            // `replace` writes the new clone and yields the old value to drop.
            let old = unsafe { ptr.cast::<Variant>().replace(value.try_clone()?) };
            drop(old);
        }
        // SAFETY: kind == I1, so the slot is a live, aligned `i8` scalar.
        SafeArrayElementKind::I1 => unsafe {
            *ptr.cast::<i8>() = i8::try_from(variant_i64(value)?)
                .map_err(|_| format!("value {value:?} does not fit VT_I1 SAFEARRAY element"))?;
        },
        // SAFETY: kind == Ui1, so the slot is a live, aligned `u8` scalar.
        SafeArrayElementKind::Ui1 => unsafe {
            *ptr.cast::<u8>() = u8::try_from(variant_i64(value)?)
                .map_err(|_| format!("value {value:?} does not fit VT_UI1 SAFEARRAY element"))?;
        },
        // SAFETY: kind == I2, so the slot is a live, aligned `i16` scalar.
        SafeArrayElementKind::I2 => unsafe {
            *ptr.cast::<i16>() = i16::try_from(variant_i64(value)?)
                .map_err(|_| format!("value {value:?} does not fit VT_I2 SAFEARRAY element"))?;
        },
        // SAFETY: kind == Ui2, so the slot is a live, aligned `u16` scalar.
        SafeArrayElementKind::Ui2 => unsafe {
            *ptr.cast::<u16>() = u16::try_from(variant_i64(value)?)
                .map_err(|_| format!("value {value:?} does not fit VT_UI2 SAFEARRAY element"))?;
        },
        // SAFETY: kind == I4/Int, so the slot is a live, aligned `i32` scalar.
        SafeArrayElementKind::I4 | SafeArrayElementKind::Int => unsafe {
            *ptr.cast::<i32>() = i32::try_from(variant_i64(value)?)
                .map_err(|_| format!("value {value:?} does not fit VT_I4 SAFEARRAY element"))?;
        },
        // SAFETY: kind == Ui4/UInt, so the slot is a live, aligned `u32` scalar.
        SafeArrayElementKind::Ui4 | SafeArrayElementKind::UInt => unsafe {
            *ptr.cast::<u32>() = u32::try_from(variant_i64(value)?)
                .map_err(|_| format!("value {value:?} does not fit VT_UI4 SAFEARRAY element"))?;
        },
        // SAFETY: kind == I8, so the slot is a live, aligned `i64` scalar.
        SafeArrayElementKind::I8 => unsafe {
            *ptr.cast::<i64>() = variant_i64(value)?;
        },
        // SAFETY: kind == Ui8, so the slot is a live, aligned `u64` scalar.
        SafeArrayElementKind::Ui8 => unsafe {
            *ptr.cast::<u64>() = u64::try_from(variant_i64(value)?)
                .map_err(|_| format!("value {value:?} does not fit VT_UI8 SAFEARRAY element"))?;
        },
        // SAFETY: kind == R4, so the slot is a live, aligned `f32` scalar.
        SafeArrayElementKind::R4 => unsafe {
            *ptr.cast::<f32>() = variant_f64(value)? as f32;
        },
        // SAFETY: kind == R8, so the slot is a live, aligned `f64` scalar.
        SafeArrayElementKind::R8 => unsafe {
            *ptr.cast::<f64>() = variant_f64(value)?;
        },
        // SAFETY: kind == Currency, stored as a scaled `i64`; slot is an aligned `i64`.
        SafeArrayElementKind::Currency => unsafe {
            *ptr.cast::<i64>() = value
                .as_currency_scaled_i64()
                .ok_or_else(|| format!("expected Currency SAFEARRAY element, got {value:?}"))?;
        },
        // SAFETY: kind == Date, stored as an `f64` serial; slot is an aligned `f64`.
        SafeArrayElementKind::Date => unsafe {
            *ptr.cast::<f64>() = value
                .as_date_f64()
                .or_else(|| value.as_f64())
                .ok_or_else(|| format!("expected Date SAFEARRAY element, got {value:?}"))?;
        },
        // SAFETY: kind == Bool, stored as a VBA `i16` (-1/0); slot is an aligned `i16`.
        SafeArrayElementKind::Bool => unsafe {
            *ptr.cast::<i16>() = if value
                .as_bool()
                .ok_or_else(|| format!("expected Bool SAFEARRAY element, got {value:?}"))?
            {
                -1
            } else {
                0
            };
        },
        // SAFETY: kind == BStr, so the slot owns a BSTR carrier pointer; replace it
        // with the freshly allocated BSTR and free the previous one.
        SafeArrayElementKind::BStr => unsafe {
            let text = value
                .try_as_bstr()?
                .ok_or_else(|| format!("expected String SAFEARRAY element, got {value:?}"))?;
            let old = ptr.cast::<*mut u16>().replace(text.into_raw_bstr());
            free_raw_bstr(old);
        },
        // SAFETY: kind == Dispatch/Unknown, so the slot owns an interface pointer (or null);
        // store the new owned reference and release the previous one.
        SafeArrayElementKind::Dispatch | SafeArrayElementKind::Unknown => unsafe {
            let old = ptr
                .cast::<*mut RawRuntimeIUnknown>()
                .replace(object_element(value)?);
            drop(ObjectRef::from_raw_iunknown_owned(old));
        },
        // SAFETY: kind == Decimal, so the slot is a live, aligned `RawDecimalArrayElement`.
        SafeArrayElementKind::Decimal => unsafe {
            *ptr.cast::<RawDecimalArrayElement>() = RawDecimalArrayElement::from_decimal96(
                value
                    .as_decimal96()
                    .ok_or_else(|| format!("expected Decimal SAFEARRAY element, got {value:?}"))?,
            );
        },
    }
    Ok(())
}

unsafe fn drop_element(kind: SafeArrayElementKind, payload: *mut u8, index: usize) {
    // SAFETY: the caller guarantees `payload` holds more than `index`
    // initialized (or zeroed) elements of `kind` and visits each index at most once, so
    // the offset is in-bounds and the slot still owns its resource.
    let ptr = unsafe { payload.add(payload_offset(kind, index)) };
    match kind {
        // SAFETY: kind == Variant, so the slot holds an initialized `Variant` (a zeroed slot
        // is `Empty`), dropped exactly once.
        SafeArrayElementKind::Variant => unsafe { core::ptr::drop_in_place(ptr.cast::<Variant>()) },
        // SAFETY: kind == BStr, so the slot holds the array-owned raw BSTR pointer (or null);
        // ownership transfers to `free_raw_bstr`, which frees it exactly once.
        SafeArrayElementKind::BStr => unsafe { free_raw_bstr(*ptr.cast::<*mut u16>()) },
        // SAFETY: kind == Dispatch/Unknown, so the slot holds the array's single retained
        // reference (or null); `from_raw_iunknown_owned` adopts and releases exactly that.
        SafeArrayElementKind::Dispatch | SafeArrayElementKind::Unknown => unsafe {
            drop(ObjectRef::from_raw_iunknown_owned(
                *ptr.cast::<*mut RawRuntimeIUnknown>(),
            ));
        },
        _ => {}
    }
}

/// Hidden bytes for a new descriptor: the element vt, or the interface IID.
fn hidden_prefix(element: &ElementType) -> [u8; HIDDEN_SIZE] {
    let mut hidden = [0u8; HIDDEN_SIZE];
    if let ElementType::Intrinsic(kind) = element {
        match kind {
            SafeArrayElementKind::Dispatch => hidden = IID_IDISPATCH,
            SafeArrayElementKind::Unknown => hidden = IID_IUNKNOWN,
            _ => hidden[HIDDEN_SIZE - 4..]
                .copy_from_slice(&u32::from(kind_vartype(*kind)).to_le_bytes()),
        }
    }
    hidden
}

fn kind_vartype(kind: SafeArrayElementKind) -> u16 {
    match kind {
        SafeArrayElementKind::Variant => VT_VARIANT_VALUE,
        SafeArrayElementKind::I1 => VT_I1_VALUE,
        SafeArrayElementKind::Ui1 => VT_UI1_VALUE,
        SafeArrayElementKind::I2 => VT_I2_VALUE,
        SafeArrayElementKind::Ui2 => VT_UI2_VALUE,
        SafeArrayElementKind::I4 => VT_I4_VALUE,
        SafeArrayElementKind::Ui4 => VT_UI4_VALUE,
        SafeArrayElementKind::I8 => VT_I8_VALUE,
        SafeArrayElementKind::Ui8 => VT_UI8_VALUE,
        SafeArrayElementKind::Int => VT_INT_VALUE,
        SafeArrayElementKind::UInt => VT_UINT_VALUE,
        SafeArrayElementKind::R4 => VT_R4_VALUE,
        SafeArrayElementKind::R8 => VT_R8_VALUE,
        SafeArrayElementKind::Currency => VT_CY_VALUE,
        SafeArrayElementKind::Date => VT_DATE_VALUE,
        SafeArrayElementKind::Bool => VT_BOOL_VALUE,
        SafeArrayElementKind::BStr => VT_BSTR_VALUE,
        SafeArrayElementKind::Dispatch => VT_DISPATCH_VALUE,
        SafeArrayElementKind::Unknown => VT_UNKNOWN_VALUE,
        SafeArrayElementKind::Decimal => VT_DECIMAL_VALUE,
    }
}

impl SafeArray {
    pub fn supports_intrinsic_element_vartype(element_vt: u16) -> bool {
        SafeArrayElementKind::from_vartype(element_vt).is_some()
    }

    fn supports_zeroed_scalar_payload(kind: SafeArrayElementKind) -> bool {
        !kind.owns_resources()
    }

    /// A descriptor of `bounds` with a zeroed payload when `with_data` (and the array has
    /// elements), else with null `pvData`. Zeroed elements are valid for every element type
    /// (Empty, a null BSTR, Nothing, a default record), so the returned array may be dropped
    /// at any point while its elements are filled in.
    fn allocate(
        bounds: &[SafeArrayBound],
        element: ElementType,
        with_data: bool,
    ) -> Result<Self, String> {
        let c_dims = u16::try_from(bounds.len()).map_err(|_| {
            format!(
                "SAFEARRAY rank {} exceeds the {MAX_SAFEARRAY_RANK}-dimension limit",
                bounds.len()
            )
        })?;
        if c_dims == 0 {
            return Err("SAFEARRAY must have at least one dimension".to_string());
        }
        let len = bounds_total_len(bounds)?;
        let cb_elements = u32::try_from(element.element_size())
            .map_err(|_| "SAFEARRAY cbElements exceeds u32 capacity".to_string())?;
        len.checked_mul(element.element_size())
            .ok_or_else(|| "SAFEARRAY payload layout overflow".to_string())?;
        let (f_features, vt) = match &element {
            ElementType::Intrinsic(kind) => (kind.feature_flags(), Some(kind_vartype(*kind))),
            ElementType::Record(_) => (0, None),
        };
        crate::vba_record::owning_boundary("safearray-header-adoption")?;
        let header = sys::alloc_descriptor(c_dims, vt)
            .and_then(NonNull::new)
            .ok_or_else(|| "failed to allocate SAFEARRAY header".to_string())?;
        crate::live_counters::safearray_allocated();
        // SAFETY: `header` is a fresh descriptor with `c_dims` bounds and its hidden bytes;
        // every write below stays inside it.
        unsafe {
            let raw = header.as_ptr();
            hidden_bytes(raw)
                .copy_from_nonoverlapping(hidden_prefix(&element).as_ptr(), HIDDEN_SIZE);
            (*raw).f_features = f_features;
            (*raw).cb_elements = cb_elements;
            (*raw).c_locks = 0;
            (*raw).pv_data = core::ptr::null_mut();
            // `rgsabound` lists the last dimension first, as oleaut32 and VBA store it.
            let dst = core::ptr::addr_of_mut!((*raw).rgsabound).cast::<SafeArrayBound>();
            for (slot, bound) in bounds.iter().rev().enumerate() {
                dst.add(slot).write(*bound);
            }
        }
        let array = Self {
            header,
            record: match element {
                ElementType::Record(layout) => Some(layout),
                ElementType::Intrinsic(_) => None,
            },
        };
        if with_data && len != 0 {
            crate::vba_record::owning_boundary(if array.record.is_some() {
                "safearray-record-payload-allocation"
            } else {
                "safearray-intrinsic-payload-allocation"
            })?;
            // SAFETY: the descriptor is live, its bounds and cbElements are set, pvData null.
            if !unsafe { sys::alloc_data(header.as_ptr()) } {
                return Err("failed to allocate SAFEARRAY payload".to_string());
            }
            // SAFETY: `alloc_data` set pvData to `len * cbElements` bytes; zero them, as the
            // payload's initial state (oleaut32 does not promise it).
            unsafe {
                core::ptr::write_bytes(
                    (*header.as_ptr()).pv_data.cast::<u8>(),
                    0,
                    len * element_size_of(&array),
                );
            }
            #[cfg(test)]
            note_payload_allocation();
        }
        Ok(array)
    }

    fn data(&self) -> *mut u8 {
        // SAFETY: `self.header` is this value's live descriptor.
        unsafe { (*self.header.as_ptr()).pv_data.cast::<u8>() }
    }

    fn from_bounds_and_variants(
        bounds: Vec<SafeArrayBound>,
        element_vt: u16,
        values: Option<Vec<Variant>>,
    ) -> Result<Self, String> {
        let kind = SafeArrayElementKind::from_vartype(element_vt).ok_or_else(|| {
            format!("unsupported intrinsic SAFEARRAY element vartype 0x{element_vt:04X}")
        })?;
        let expected_len = bounds_total_len(&bounds)?;
        if let Some(values) = &values
            && values.len() != expected_len
        {
            return Err(format!(
                "SAFEARRAY payload length {} does not match shape length {}",
                values.len(),
                expected_len
            ));
        }
        let array = Self::allocate(&bounds, ElementType::Intrinsic(kind), values.is_some())?;
        if let Some(values) = values {
            let data = array.data();
            for (index, value) in values.iter().enumerate() {
                crate::vba_record::owning_boundary("safearray-intrinsic-element-clone")?;
                // SAFETY: the payload has `values.len()` zeroed slots of `kind`; each index is
                // written once over its zeroed (resource-free) state.
                unsafe { encode_element_variant(kind, data, index, value)? };
            }
        }
        Ok(array)
    }

    pub fn from_vba_records_nd(
        bounds: Vec<SafeArrayBound>,
        layout: Arc<VbaRecordLayout>,
        values: Vec<VbaRecord>,
    ) -> Result<Self, String> {
        let expected_len = bounds_total_len(&bounds)?;
        if values.len() != expected_len {
            return Err(format!(
                "SAFEARRAY payload length {} does not match shape length {}",
                values.len(),
                expected_len
            ));
        }
        let array = Self::allocate(&bounds, ElementType::Record(Arc::clone(&layout)), true)?;
        let data = array.data();
        for (index, value) in values.iter().enumerate() {
            if value.layout().as_ref() != layout.as_ref() {
                return Err("SAFEARRAY record element layout mismatch".to_string());
            }
            crate::vba_record::owning_boundary("safearray-record-element-clone")?;
            // SAFETY: the zeroed `index`-th same-layout record slot owns nothing yet;
            // `clone_into_raw` leaves it untouched on failure.
            unsafe { value.clone_into_raw(data.add(record_payload_offset(&layout, index)))? };
        }
        Ok(array)
    }

    /// A UDT array of default records (zeroed memory is VBA's default record).
    pub fn from_default_records_nd(
        bounds: Vec<SafeArrayBound>,
        layout: Arc<VbaRecordLayout>,
    ) -> Result<Self, String> {
        Self::allocate(&bounds, ElementType::Record(layout), true)
    }

    pub fn vector(len: usize) -> Self {
        Self::from_bounds_and_variants(
            default_bounds_for_len(len).expect("vector bounds should fit SAFEARRAY capacity"),
            VT_VARIANT_VALUE,
            None,
        )
        .expect("shape-only SAFEARRAY allocation should succeed")
    }

    pub fn from_shape(bounds: Vec<SafeArrayBound>) -> Result<Self, String> {
        Self::from_bounds_and_variants(bounds, VT_VARIANT_VALUE, None)
    }

    pub fn from_shape_typed(bounds: Vec<SafeArrayBound>, element_vt: u16) -> Result<Self, String> {
        Self::from_bounds_and_variants(bounds, element_vt, None)
    }

    pub fn from_zeroed_typed_scalars_nd(
        bounds: Vec<SafeArrayBound>,
        element_vt: u16,
    ) -> Result<Self, String> {
        let kind = SafeArrayElementKind::from_vartype(element_vt).ok_or_else(|| {
            format!("unsupported intrinsic SAFEARRAY element vartype 0x{element_vt:04X}")
        })?;
        if !Self::supports_zeroed_scalar_payload(kind) {
            return Err(format!(
                "SAFEARRAY element vartype 0x{element_vt:04X} cannot be defaulted by zeroed scalar payload"
            ));
        }
        Self::allocate(&bounds, ElementType::Intrinsic(kind), true)
    }

    pub fn from_variants(values: Vec<Variant>) -> Self {
        let len = values.len();
        Self::from_bounds_and_variants(
            default_bounds_for_len(len).expect("value bounds should fit SAFEARRAY capacity"),
            VT_VARIANT_VALUE,
            Some(values),
        )
        .expect("SAFEARRAY payload allocation should succeed for supported canonical variants")
    }

    pub fn from_variants_nd(bounds: Vec<SafeArrayBound>, values: Vec<Variant>) -> Self {
        Self::from_bounds_and_variants(bounds, VT_VARIANT_VALUE, Some(values)).expect(
            "SAFEARRAY nd payload allocation should succeed for supported canonical variants",
        )
    }

    pub fn from_typed_variants(element_vt: u16, values: Vec<Variant>) -> Result<Self, String> {
        let len = values.len();
        Self::from_bounds_and_variants(default_bounds_for_len(len)?, element_vt, Some(values))
    }

    pub fn from_typed_variants_nd(
        bounds: Vec<SafeArrayBound>,
        element_vt: u16,
        values: Vec<Variant>,
    ) -> Result<Self, String> {
        Self::from_bounds_and_variants(bounds, element_vt, Some(values))
    }

    pub fn from_shape_and_variants(
        bounds: Vec<SafeArrayBound>,
        values: Vec<Variant>,
    ) -> Result<Self, String> {
        Self::from_bounds_and_variants(bounds, VT_VARIANT_VALUE, Some(values))
    }

    pub fn dimensions(&self) -> u16 {
        // SAFETY: `self.header` is this value's live descriptor.
        unsafe { (*self.header.as_ptr()).c_dims }
    }

    /// The element `VARTYPE`: `VT_RECORD` for a VBA UDT array, else as the descriptor gives it.
    pub fn element_vartype(&self) -> u16 {
        if self.record.is_some() {
            return VT_RECORD_VALUE;
        }
        // SAFETY: `self.header` is this value's live descriptor.
        unsafe { descriptor_vartype(self.header.as_ptr()) }.unwrap_or(0)
    }

    pub fn feature_flags(&self) -> u16 {
        // SAFETY: `self.header` is this value's live descriptor.
        unsafe { (*self.header.as_ptr()).f_features }
    }

    /// Whether this is a fixed-size array (`Dim a(1 To 3)`, `FADF_FIXEDSIZE`), whose `Erase`
    /// resets elements rather than freeing the storage. The bit lives in the descriptor's
    /// `fFeatures`, so it travels with [`Clone`] and across the COM boundary.
    pub fn is_fixed_size(&self) -> bool {
        self.feature_flags() & FADF_FIXEDSIZE_VALUE != 0
    }

    /// Mark the array fixed-size or dynamic in place: VBA sets `FADF_STATIC` and
    /// `FADF_FIXEDSIZE` together on a fixed-size array and neither on a dynamic one.
    pub fn set_fixed_size(&mut self, fixed: bool) {
        // SAFETY: `&mut self` gives exclusive access to this value's live descriptor.
        unsafe {
            let header = self.header.as_ptr();
            if fixed {
                (*header).f_features |= FADF_FIXED_ARRAY;
            } else {
                (*header).f_features &= !FADF_FIXED_ARRAY;
            }
        }
    }

    /// Consume `self`, marking it fixed-size or dynamic — a builder-style convenience over
    /// [`Self::set_fixed_size`].
    #[must_use]
    pub fn with_fixed_size(mut self, fixed: bool) -> Self {
        self.set_fixed_size(fixed);
        self
    }

    fn element_kind(&self) -> Option<SafeArrayElementKind> {
        if self.record.is_some() {
            return None;
        }
        SafeArrayElementKind::from_vartype(self.element_vartype())
    }

    fn intrinsic_kind(&self) -> Result<SafeArrayElementKind, String> {
        self.element_kind().ok_or_else(|| {
            format!(
                "unsupported SAFEARRAY element type (fFeatures 0x{:04X}, vt 0x{:04X})",
                self.feature_flags(),
                self.element_vartype()
            )
        })
    }

    /// The record layout of a VBA UDT array.
    pub fn vba_record_layout(&self) -> Option<Arc<VbaRecordLayout>> {
        self.record.clone()
    }

    /// Bounds in declared order (dimension 1 first); `rgsabound` stores them reversed.
    fn raw_bounds(&self) -> Vec<SafeArrayBound> {
        // SAFETY: `self.header` is this value's live descriptor.
        let mut bounds = unsafe { raw_bounds_slice(self.header.as_ptr()) }.to_vec();
        bounds.reverse();
        bounds
    }

    pub fn len(&self) -> usize {
        // SAFETY: `self.header` is this value's live descriptor.
        bounds_total_len(unsafe { raw_bounds_slice(self.header.as_ptr()) }).unwrap_or(0)
    }

    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    pub fn effective_len(&self) -> usize {
        self.len()
    }

    pub fn bounds(&self) -> Option<Vec<SafeArrayBound>> {
        let bounds = self.raw_bounds();
        if bounds.is_empty() {
            return None;
        }
        Some(bounds)
    }

    fn bounds_for_shape(&self) -> Vec<SafeArrayBound> {
        self.bounds()
            .unwrap_or_else(|| default_bounds_for_len(self.len()).unwrap_or_default())
    }

    pub fn variant_elements(&self) -> Option<Vec<Variant>> {
        self.try_variant_elements()
            .expect("SAFEARRAY payload should decode into canonical Variants")
    }

    /// VBA `For Each` walks the elements in storage order (first dimension fastest).
    pub fn enumeration_elements(&self) -> Option<Vec<Variant>> {
        self.variant_elements()
    }

    pub(crate) fn try_variant_elements(&self) -> Result<Option<Vec<Variant>>, String> {
        let data = self.data();
        if data.is_null() {
            return Ok(None);
        }
        let len = self.len();
        let mut values = Vec::with_capacity(len);
        for index in 0..len {
            values.push(self.element_at(data, index)?);
        }
        Ok(Some(values))
    }

    /// Element `index` of the non-null payload `data`, cloned out as a Variant.
    fn element_at(&self, data: *mut u8, index: usize) -> Result<Variant, String> {
        if let Some(layout) = &self.record {
            // SAFETY: `data` is the record payload of `self.len()` records strided by
            // `layout.size()`, and callers pass `index < self.len()`.
            let ptr = unsafe { data.add(record_payload_offset(layout, index)) };
            // SAFETY: `ptr` references the live `index`-th record; `clone_from_raw` copies it.
            return Ok(Variant::from_vba_record(unsafe {
                VbaRecord::clone_from_raw(ptr, Arc::clone(layout))
            }?));
        }
        // SAFETY: `data` is the payload of `self.len()` elements of this kind, and callers
        // pass `index < self.len()`.
        unsafe { decode_element_variant(self.intrinsic_kind()?, data, index) }
    }

    pub fn variant_element(&self, index: usize) -> Result<Variant, String> {
        let len = self.len();
        if index >= len {
            return Err(format!(
                "SAFEARRAY index {index} out of range for length {len}"
            ));
        }
        let data = self.data();
        if data.is_null() {
            return Err("SAFEARRAY has no materialized element payload".to_string());
        }
        self.element_at(data, index)
    }

    pub fn replace_variant_elements(&self, values: Vec<Variant>) -> Result<Self, String> {
        let bounds = self.bounds_for_shape();
        let mut array = if let Some(layout) = &self.record {
            let records = values
                .iter()
                .map(|value| {
                    value
                        .vba_record_ref()
                        .ok_or_else(|| "expected VBA record SAFEARRAY element".to_string())
                        .and_then(VbaRecord::try_clone)
                })
                .collect::<Result<Vec<_>, _>>()?;
            Self::from_vba_records_nd(bounds, Arc::clone(layout), records)?
        } else {
            Self::from_bounds_and_variants(bounds, self.element_vartype(), Some(values))?
        };
        array.copy_descriptor_traits_from(self);
        Ok(array)
    }

    /// Give `self` the fixed-size bits and hidden IID of `source` (a copy keeps them).
    fn copy_descriptor_traits_from(&mut self, source: &Self) {
        self.set_fixed_size(source.is_fixed_size());
        if source.feature_flags() & FADF_HAVEIID_VALUE != 0
            && self.feature_flags() & FADF_HAVEIID_VALUE != 0
        {
            // SAFETY: both descriptors are live and have their hidden bytes.
            unsafe {
                hidden_bytes(self.header.as_ptr())
                    .copy_from_nonoverlapping(hidden_bytes(source.header.as_ptr()), HIDDEN_SIZE);
            }
        }
    }

    pub fn set_variant_element(&mut self, index: usize, value: &Variant) -> Result<(), String> {
        let len = self.len();
        if index >= len {
            return Err(format!(
                "SAFEARRAY index {index} out of range for length {len}"
            ));
        }
        let data = self.data();
        if data.is_null() {
            return Err("SAFEARRAY has no materialized element payload".to_string());
        }
        if let Some(layout) = &self.record {
            let Some(source) = value.vba_record_ref() else {
                return Err(format!(
                    "expected VBA record SAFEARRAY element, got {:?}",
                    value.vtype()
                ));
            };
            if source.layout().as_ref() != layout.as_ref() {
                return Err("VBA record SAFEARRAY element layout mismatch".to_string());
            }
            let mut replacement = source.try_clone()?;
            // SAFETY: `index < self.len()` (checked above) and `data` is the record payload
            // strided by `layout.size()`, so this offset is the in-bounds slot.
            let ptr = unsafe { data.add(record_payload_offset(layout, index)) };
            // SAFETY: replacement is a complete same-layout owner and `ptr` is the live
            // `index`-th record. The byte swap cannot fail; replacement then drops the prior
            // element after the array already owns the new payload.
            unsafe { replacement.swap_with_raw(ptr) };
            return Ok(());
        }
        // SAFETY: `data` is the non-null payload of `self.len()` elements of this kind, and
        // `index < self.len()`, so the slot is writable.
        unsafe { replace_element_variant(self.intrinsic_kind()?, data, index, value) }
    }

    pub fn i32_element(&self, index: usize) -> Result<Option<i32>, String> {
        let Some(kind @ (SafeArrayElementKind::I4 | SafeArrayElementKind::Int)) =
            self.element_kind()
        else {
            return Ok(None);
        };
        let len = self.len();
        if index >= len {
            return Err(format!(
                "SAFEARRAY index {index} out of range for length {len}"
            ));
        }
        let data = self.data();
        if data.is_null() {
            return Err("SAFEARRAY has no materialized element payload".to_string());
        }
        // SAFETY: kind/len checks above prove this is an in-bounds, aligned i32 slot.
        Ok(Some(unsafe {
            *data.add(payload_offset(kind, index)).cast::<i32>()
        }))
    }

    pub fn set_i32_element(&mut self, index: usize, value: i32) -> Result<bool, String> {
        let Some(kind @ (SafeArrayElementKind::I4 | SafeArrayElementKind::Int)) =
            self.element_kind()
        else {
            return Ok(false);
        };
        let len = self.len();
        if index >= len {
            return Err(format!(
                "SAFEARRAY index {index} out of range for length {len}"
            ));
        }
        let data = self.data();
        if data.is_null() {
            return Err("SAFEARRAY has no materialized element payload".to_string());
        }
        // SAFETY: kind/len checks above and exclusive `&mut self` make this a unique, aligned
        // live i32 slot.
        unsafe { *data.add(payload_offset(kind, index)).cast::<i32>() = value };
        Ok(true)
    }

    pub fn raw_safearray_ptr(&self) -> *mut core::ffi::c_void {
        self.header.as_ptr().cast()
    }

    /// Transfer the descriptor to an owner that will release it with
    /// [`Self::from_raw_safearray_owned`] (passing the same record layout) on the owning
    /// thread. A UDT array's layout is not in the descriptor; the owner keeps it.
    pub fn into_raw_safearray_ptr(self) -> *mut core::ffi::c_void {
        let owner = core::mem::ManuallyDrop::new(self);
        // SAFETY: `owner` is never dropped, so moving its layout out once is sound.
        drop(unsafe { core::ptr::read(&owner.record) });
        owner.header.as_ptr().cast()
    }

    pub fn clone_raw_safearray_ptr(&self) -> *mut core::ffi::c_void {
        self.clone().into_raw_safearray_ptr()
    }

    /// A non-owning view of a raw descriptor (with the record layout its owner keeps).
    ///
    /// # Safety
    /// `raw` must remain a live descriptor for the complete returned view lifetime, with
    /// shared or exclusive access appropriate to its use, and `record` must be its elements'
    /// layout when it is a VBA UDT array.
    pub unsafe fn borrow_raw(
        raw: *mut core::ffi::c_void,
        record: Option<Arc<VbaRecordLayout>>,
    ) -> Option<BorrowedSafeArray> {
        let header = NonNull::new(raw.cast::<RawSafeArray>())?;
        Some(BorrowedSafeArray(core::mem::ManuallyDrop::new(Self {
            header,
            record,
        })))
    }

    /// # Safety
    /// As [`Self::borrow_raw`].
    unsafe fn borrow_raw_result(
        raw: *mut core::ffi::c_void,
        record: Option<Arc<VbaRecordLayout>>,
    ) -> Result<BorrowedSafeArray, String> {
        // SAFETY: forwarded from the caller.
        unsafe { Self::borrow_raw(raw, record) }
            .ok_or_else(|| "SAFEARRAY raw pointer is null".to_string())
    }

    /// Takes ownership of a raw SAFEARRAY descriptor: one this runtime handed out with
    /// [`Self::into_raw_safearray_ptr`], or one oleaut32 allocated (a callee's
    /// `SafeArrayCreate`, which Automation lets the caller own).
    ///
    /// # Safety
    /// `raw` must be null or a live descriptor the caller owns, and `record` its elements'
    /// layout when it is a VBA UDT array. Returns `None` (leaving `raw` alone) for null or for
    /// an element type this runtime cannot hold.
    pub unsafe fn from_raw_safearray_owned(
        raw: *mut core::ffi::c_void,
        record: Option<Arc<VbaRecordLayout>>,
    ) -> Option<Self> {
        let header = NonNull::new(raw.cast::<RawSafeArray>())?;
        let array = Self { header, record };
        if array.record.is_none() && array.element_kind().is_none() {
            let _ = core::mem::ManuallyDrop::new(array);
            return None;
        }
        Some(array)
    }

    /// Clones a raw SAFEARRAY descriptor.
    ///
    /// # Safety
    /// `raw` must be null or a live descriptor (with `record` its UDT layout, if any). The
    /// clone receives independent descriptor/payload storage; ownership of `raw` is unchanged.
    pub unsafe fn clone_from_raw_safearray(
        raw: *mut core::ffi::c_void,
        record: Option<Arc<VbaRecordLayout>>,
    ) -> Option<Self> {
        // SAFETY: forwards this public helper's raw descriptor contract.
        unsafe { Self::try_clone_from_raw_safearray(raw, record) }
            .expect("cloning a live canonical SAFEARRAY should succeed")
    }

    /// # Safety
    /// As [`Self::clone_from_raw_safearray`].
    pub(crate) unsafe fn try_clone_from_raw_safearray(
        raw: *mut core::ffi::c_void,
        record: Option<Arc<VbaRecordLayout>>,
    ) -> Result<Option<Self>, String> {
        // SAFETY: forwards this function's raw descriptor contract; the view cannot claim
        // source ownership during error or unwind.
        let Some(borrowed) = (unsafe { Self::borrow_raw(raw, record) }) else {
            return Ok(None);
        };
        crate::vba_record::owning_boundary("borrowed-safearray-clone")?;
        Ok(Some(borrowed.try_clone()?))
    }

    /// Mutates one element of a raw SAFEARRAY descriptor.
    ///
    /// # Safety
    /// `raw` must point at a live descriptor (with `record` its UDT layout, if any), and the
    /// caller must have exclusive access to it for the duration of this call.
    pub unsafe fn set_raw_safearray_variant_element(
        raw: *mut core::ffi::c_void,
        record: Option<Arc<VbaRecordLayout>>,
        index: usize,
        value: &Variant,
    ) -> Result<(), String> {
        // SAFETY: forwards this function's raw descriptor contract.
        let mut borrowed = unsafe { Self::borrow_raw_result(raw, record) }?;
        crate::vba_record::owning_boundary("borrowed-safearray-variant-set")?;
        borrowed.set_variant_element(index, value)
    }

    /// Reads one element of a raw SAFEARRAY descriptor **without cloning the array** — O(1)
    /// in array length, the read primitive that keeps `arr(i)` constant-time.
    ///
    /// # Safety
    /// `raw` must point at a live descriptor (with `record` its UDT layout, if any), and the
    /// caller must have shared access to it for the duration of this call.
    pub unsafe fn raw_safearray_variant_element(
        raw: *mut core::ffi::c_void,
        record: Option<Arc<VbaRecordLayout>>,
        index: usize,
    ) -> Result<Variant, String> {
        // SAFETY: forwards this function's raw descriptor contract.
        let borrowed = unsafe { Self::borrow_raw_result(raw, record) }?;
        crate::vba_record::owning_boundary("borrowed-safearray-variant-read")?;
        borrowed.variant_element(index)
    }

    /// Reads an I4/Int element from a raw SAFEARRAY descriptor.
    ///
    /// # Safety
    /// `raw` must point at a live descriptor, and the caller must have shared access to it for
    /// the duration of this call. Ownership of the descriptor remains with the caller.
    pub unsafe fn raw_safearray_i32_element(
        raw: *mut core::ffi::c_void,
        index: usize,
    ) -> Result<Option<i32>, String> {
        // SAFETY: forwards this function's raw descriptor contract.
        let borrowed = unsafe { Self::borrow_raw_result(raw, None) }?;
        crate::vba_record::owning_boundary("borrowed-safearray-i32-read")?;
        borrowed.i32_element(index)
    }

    /// Mutates an I4/Int element of a raw SAFEARRAY descriptor.
    ///
    /// # Safety
    /// `raw` must point at a live descriptor, and the caller must have exclusive access to it
    /// for the duration of this call. Ownership of the descriptor remains with the caller.
    pub unsafe fn set_raw_safearray_i32_element(
        raw: *mut core::ffi::c_void,
        index: usize,
        value: i32,
    ) -> Result<bool, String> {
        // SAFETY: forwards this function's raw descriptor contract.
        let mut borrowed = unsafe { Self::borrow_raw_result(raw, None) }?;
        crate::vba_record::owning_boundary("borrowed-safearray-i32-set")?;
        borrowed.set_i32_element(index, value)
    }

    /// Reads the bounds and element count of a raw SAFEARRAY descriptor **without cloning the
    /// element payload** — O(rank).
    ///
    /// # Safety
    /// `raw` must be null or point at a live descriptor, with shared access for this call.
    pub unsafe fn raw_safearray_bounds_len(
        raw: *mut core::ffi::c_void,
    ) -> Option<(Vec<SafeArrayBound>, usize)> {
        // SAFETY: forwarded from the caller.
        let borrowed = unsafe { Self::borrow_raw(raw, None) }?;
        crate::vba_record::owning_boundary("borrowed-safearray-bounds").ok()?;
        let bounds = borrowed.raw_bounds();
        let len = bounds_total_len(&bounds).ok()?;
        if bounds.is_empty() {
            None
        } else {
            Some((bounds, len))
        }
    }

    pub(crate) fn try_clone(&self) -> Result<Self, String> {
        let bounds = self.raw_bounds();
        let data = self.data();
        let element = match &self.record {
            Some(layout) => ElementType::Record(Arc::clone(layout)),
            None => ElementType::Intrinsic(self.intrinsic_kind()?),
        };
        let mut copy = Self::allocate(&bounds, element, !data.is_null())?;
        if !data.is_null() {
            let target = copy.data();
            for index in 0..self.len() {
                match &self.record {
                    Some(layout) => {
                        crate::vba_record::owning_boundary("safearray-record-element-clone")?;
                        let offset = record_payload_offset(layout, index);
                        // SAFETY: both payloads hold `self.len()` records of `layout`; the
                        // source record is copied into the zeroed target slot.
                        unsafe {
                            VbaRecord::clone_from_raw(data.add(offset), Arc::clone(layout))?
                                .clone_into_raw(target.add(offset))?
                        };
                    }
                    None => {
                        crate::vba_record::owning_boundary("safearray-intrinsic-element-clone")?;
                        let kind = self.intrinsic_kind()?;
                        // SAFETY: both payloads hold `self.len()` elements of `kind`; the
                        // zeroed target slot is written once.
                        unsafe {
                            let value = decode_element_variant(kind, data, index)?;
                            encode_element_variant(kind, target, index, &value)?;
                        }
                    }
                }
            }
        }
        copy.copy_descriptor_traits_from(self);
        Ok(copy)
    }
}

/// The element size of an allocated array.
fn element_size_of(array: &SafeArray) -> usize {
    // SAFETY: `array.header` is a live descriptor.
    unsafe { (*array.header.as_ptr()).cb_elements as usize }
}

/// A [`SafeArray`] view over a descriptor owned elsewhere: dropping it releases only its
/// handle on the record layout, never the descriptor.
pub struct BorrowedSafeArray(core::mem::ManuallyDrop<SafeArray>);

impl core::ops::Deref for BorrowedSafeArray {
    type Target = SafeArray;
    fn deref(&self) -> &SafeArray {
        &self.0
    }
}

impl core::ops::DerefMut for BorrowedSafeArray {
    fn deref_mut(&mut self) -> &mut SafeArray {
        &mut self.0
    }
}

impl Drop for BorrowedSafeArray {
    fn drop(&mut self) {
        drop(self.0.record.take());
    }
}

impl Clone for SafeArray {
    fn clone(&self) -> Self {
        self.try_clone()
            .expect("cloning canonical SAFEARRAY should succeed")
    }
}

impl Drop for SafeArray {
    fn drop(&mut self) {
        let len = self.len();
        let data = self.data();
        if !data.is_null() {
            if let Some(layout) = &self.record {
                for index in 0..len {
                    // SAFETY: each `index < len` names a live (or zeroed) record of `layout`,
                    // dropped exactly once.
                    unsafe {
                        VbaRecord::drop_raw(data.add(record_payload_offset(layout, index)), layout)
                    };
                }
            } else if let Some(kind) = self.element_kind()
                && kind.owns_resources()
            {
                for index in 0..len {
                    // SAFETY: each `index < len` names a live (or zeroed) element of `kind`,
                    // dropped exactly once.
                    unsafe { drop_element(kind, data, index) };
                }
                // SAFETY: the payload has `len * cbElements` bytes. Zeroed, oleaut32's own
                // element clear in `SafeArrayDestroy` finds nothing left to release.
                unsafe { core::ptr::write_bytes(data, 0, len * element_size_of(self)) };
            }
            #[cfg(test)]
            note_payload_free();
        }
        // SAFETY: this value owns the live descriptor. Clearing `FADF_STATIC` lets the
        // allocator free the payload; nothing touches the descriptor after `destroy`.
        unsafe {
            (*self.header.as_ptr()).f_features &= !FADF_STATIC_VALUE;
            sys::destroy(self.header.as_ptr());
        }
        crate::live_counters::safearray_freed();
    }
}

impl core::fmt::Debug for SafeArray {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        f.debug_struct("SafeArray")
            .field("dimensions", &self.dimensions())
            .field("len", &self.len())
            .field(
                "element_vartype",
                &format_args!("{:#06X}", self.element_vartype()),
            )
            .field("bounds", &self.bounds())
            .field("elements", &self.try_variant_elements().ok().flatten())
            .finish()
    }
}

impl PartialEq for SafeArray {
    fn eq(&self, other: &Self) -> bool {
        self.dimensions() == other.dimensions()
            && self.len() == other.len()
            && self.element_vartype() == other.element_vartype()
            && self.bounds() == other.bounds()
            && self.variant_elements() == other.variant_elements()
    }
}

impl Eq for SafeArray {}

pub fn is_array_tag(value: i32) -> bool {
    (ARRAY_TAG_BASE..=ARRAY_TAG_LIMIT).contains(&value)
}

pub fn array_len_from_tag(value: i32) -> Option<usize> {
    if !is_array_tag(value) {
        return None;
    }
    let count = value.checked_sub(ARRAY_TAG_BASE)?;
    usize::try_from(count).ok()
}

pub fn safe_array_from_tag(value: i32) -> Option<SafeArray> {
    array_len_from_tag(value).map(SafeArray::vector)
}

pub fn array_tag_from_safe_array(array: &SafeArray) -> Option<i32> {
    if array.dimensions() == 0 {
        return None;
    }
    let len_i32 = i32::try_from(array.effective_len()).ok()?;
    ARRAY_TAG_BASE
        .checked_add(len_i32)
        .filter(|v| *v <= ARRAY_TAG_LIMIT)
}

pub fn marshal_dispatch_argument(value: i32) -> i32 {
    let Some(array) = safe_array_from_tag(value) else {
        return value;
    };
    match i32::try_from(array.len()) {
        Ok(len) => DISPATCH_ARRAY_PAYLOAD_BASE.saturating_add(len),
        Err(_) => DISPATCH_ARRAY_PAYLOAD_BASE,
    }
}

#[cfg(test)]
// Test-support code exercising the documented production raw-pointer paths.
#[allow(clippy::undocumented_unsafe_blocks)]
mod tests {
    use super::{
        ARRAY_TAG_BASE, FADF_BSTR_VALUE, FADF_DISPATCH_VALUE, FADF_HAVEVARTYPE_VALUE,
        FADF_UNKNOWN_VALUE, FADF_VARIANT_VALUE, RawSafeArray, SafeArray, SafeArrayBound,
        VT_BOOL_VALUE, VT_BSTR_VALUE, VT_CY_VALUE, VT_DATE_VALUE, VT_DECIMAL_VALUE,
        VT_DISPATCH_VALUE, VT_I2_VALUE, VT_I4_VALUE, VT_RECORD_VALUE, VT_UNKNOWN_VALUE,
        VT_VARIANT_VALUE, advance_subscripts, array_len_from_tag, array_tag_from_safe_array,
        element_index, element_subscripts, marshal_dispatch_argument, payload_allocation_events,
        preserved_element_index, raw_bounds_slice, safe_array_from_tag,
    };
    use crate::live_counters::thread_live_handle_counts;
    use crate::vba_record::{
        OwningFailureMode, clear_owning_boundary_failure, inject_owning_boundary_failure,
        record_buffer_event_counts, take_owning_boundary_trace,
    };
    use crate::{
        Decimal96, ObjectRef, Variant, VbaRecord, VbaRecordFieldKind, VbaRecordFieldSpec,
        VbaRecordLayout, bstr::BStr,
    };
    use std::{
        panic::{AssertUnwindSafe, catch_unwind},
        sync::Arc,
    };

    fn offset_of<T, U>(field: fn(*const T) -> *const U) -> usize {
        let value = core::mem::MaybeUninit::<T>::uninit();
        let base = value.as_ptr();
        field(base) as usize - base as usize
    }

    fn transactional_record_layouts() -> (Arc<VbaRecordLayout>, Arc<VbaRecordLayout>) {
        let inner = Arc::new(
            VbaRecordLayout::new(vec![
                VbaRecordFieldSpec::named("InnerText", VbaRecordFieldKind::String),
                VbaRecordFieldSpec::named("InnerValue", VbaRecordFieldKind::Variant),
            ])
            .expect("transactional inner layout"),
        );
        let outer = Arc::new(
            VbaRecordLayout::new(vec![
                VbaRecordFieldSpec::named("Text", VbaRecordFieldKind::String),
                VbaRecordFieldSpec::named("Value", VbaRecordFieldKind::Variant),
                VbaRecordFieldSpec::named("Nested", VbaRecordFieldKind::Record(Arc::clone(&inner))),
            ])
            .expect("transactional record SAFEARRAY layout"),
        );
        (inner, outer)
    }

    fn transactional_record(
        inner_layout: &Arc<VbaRecordLayout>,
        outer_layout: &Arc<VbaRecordLayout>,
        label: &str,
        identity: i32,
    ) -> VbaRecord {
        let mut nested =
            VbaRecord::new_default(Arc::clone(inner_layout)).expect("nested record value");
        nested
            .write_field_variant(0, &Variant::from_string(format!("{label} nested")))
            .expect("nested String");
        nested
            .write_field_variant(
                1,
                &Variant::from_object_ref(ObjectRef::from_compat_identity(identity + 100)),
            )
            .expect("nested Variant");

        let mut record =
            VbaRecord::new_default(Arc::clone(outer_layout)).expect("outer record value");
        record
            .write_field_variant(0, &Variant::from_string(format!("{label} text")))
            .expect("outer String");
        record
            .write_field_variant(
                1,
                &Variant::from_object_ref(ObjectRef::from_compat_identity(identity)),
            )
            .expect("outer Variant");
        record
            .write_field_variant(2, &Variant::from_vba_record(nested))
            .expect("outer nested record");
        record
    }

    fn assert_transactional_record(value: &Variant, label: &str, identity: i32) {
        let record = value.vba_record_ref().expect("VBA record Variant");
        assert_eq!(
            record
                .read_field_variant(0)
                .expect("outer String read")
                .as_bstr()
                .expect("outer String")
                .as_str(),
            format!("{label} text")
        );
        assert_eq!(
            record
                .read_field_variant(1)
                .expect("outer Variant read")
                .as_object_ref()
                .expect("outer Object")
                .compat_identity(),
            identity
        );
        let nested = record.read_field_variant(2).expect("nested record read");
        let nested = nested.vba_record_ref().expect("nested VBA record");
        assert_eq!(
            nested
                .read_field_variant(0)
                .expect("nested String read")
                .as_bstr()
                .expect("nested String")
                .as_str(),
            format!("{label} nested")
        );
        assert_eq!(
            nested
                .read_field_variant(1)
                .expect("nested Variant read")
                .as_object_ref()
                .expect("nested Object")
                .compat_identity(),
            identity + 100
        );
    }

    fn transactional_record_array(
        inner_layout: &Arc<VbaRecordLayout>,
        outer_layout: &Arc<VbaRecordLayout>,
        label: &str,
        identity: i32,
    ) -> SafeArray {
        SafeArray::from_vba_records_nd(
            vec![SafeArrayBound { count: 1, lower: 7 }],
            Arc::clone(outer_layout),
            vec![transactional_record(
                inner_layout,
                outer_layout,
                label,
                identity,
            )],
        )
        .expect("record SAFEARRAY")
    }

    fn record_array_element_bytes(array: &SafeArray, layout: &VbaRecordLayout) -> Vec<u8> {
        // SAFETY: this helper is restricted to the one-element record SAFEARRAY
        // constructed above; its payload is non-null and exactly `layout.size()`.
        let data = unsafe { (*array.header.as_ptr()).pv_data.cast::<u8>() };
        assert!(!data.is_null());
        // SAFETY: the descriptor owns one initialized record of this exact layout.
        unsafe { core::slice::from_raw_parts(data, layout.size()) }.to_vec()
    }

    fn descriptor_fingerprint(
        array: &SafeArray,
    ) -> (usize, u16, u16, u32, u32, usize, Vec<SafeArrayBound>) {
        // SAFETY: `array.0` is its live descriptor for the duration of this borrow.
        let header = unsafe { &*array.header.as_ptr() };
        (
            array.header.as_ptr() as usize,
            header.c_dims,
            header.f_features,
            header.cb_elements,
            header.c_locks,
            header.pv_data as usize,
            array.raw_bounds(),
        )
    }

    fn assert_borrowed_bstr_array_source(
        array: &SafeArray,
        descriptor: &(usize, u16, u16, u32, u32, usize, Vec<SafeArrayBound>),
        expected: &[u8],
    ) {
        assert_eq!(&descriptor_fingerprint(array), descriptor);
        assert_eq!(
            array
                .variant_element(0)
                .expect("borrowed BSTR source remains readable")
                .string_bytes(),
            Some(expected.to_vec())
        );
        take_owning_boundary_trace();
    }

    fn assert_payload_event_balance(before: (usize, usize), label: &str) {
        let after = payload_allocation_events();
        assert_eq!(
            after.0 - before.0,
            after.1 - before.1,
            "{label} must free every raw payload allocation"
        );
    }

    fn assert_payload_constructor_sweep<T>(
        label: &str,
        make_values: impl Fn() -> T,
        construct: impl Fn(T) -> Result<SafeArray, String>,
        element_boundary: &'static str,
        expected_elements: usize,
    ) {
        let handles_before = thread_live_handle_counts();
        let payloads_before = payload_allocation_events();
        let values = make_values();
        take_owning_boundary_trace();
        let successful = construct(values).expect("uninjected SAFEARRAY payload construction");
        let success_trace = take_owning_boundary_trace();
        assert_eq!(successful.len(), expected_elements);
        assert_eq!(success_trace.first(), Some(&"safearray-header-adoption"));
        let element_boundaries = success_trace
            .iter()
            .filter(|boundary| **boundary == element_boundary)
            .count();
        assert!(
            element_boundaries >= expected_elements,
            "{label} trace must include every top-level element-clone boundary; \
             nested carrier clones may add more ({element_boundaries} observed)"
        );
        drop(successful);
        assert_eq!(
            thread_live_handle_counts(),
            handles_before,
            "uninjected {label} construction must balance every owned handle"
        );
        assert_payload_event_balance(payloads_before, label);

        for mode in [OwningFailureMode::Error, OwningFailureMode::Panic] {
            for nth in 0..success_trace.len() {
                let handles_before = thread_live_handle_counts();
                let payloads_before = payload_allocation_events();
                let values = make_values();
                inject_owning_boundary_failure(nth, mode);
                let outcome = catch_unwind(AssertUnwindSafe(|| construct(values)));
                clear_owning_boundary_failure();
                let failure_trace = take_owning_boundary_trace();
                assert_eq!(
                    failure_trace.as_slice(),
                    &success_trace[..=nth],
                    "{label} {mode:?} sweep missed boundary {nth}"
                );
                match (mode, outcome) {
                    (OwningFailureMode::Error, Ok(Err(error))) => assert!(
                        error.contains("injected owning clone/allocation failure"),
                        "unexpected {label} constructor error: {error}"
                    ),
                    (OwningFailureMode::Error, Err(_)) => {
                        panic!("fallible {label} constructor unexpectedly unwound")
                    }
                    (OwningFailureMode::Error, Ok(Ok(_))) => {
                        panic!("injected {label} constructor error unexpectedly succeeded")
                    }
                    (OwningFailureMode::Panic, Err(_)) => {}
                    (OwningFailureMode::Panic, Ok(result)) => {
                        panic!("injected {label} constructor panic did not unwind: {result:?}")
                    }
                }
                assert_eq!(
                    thread_live_handle_counts(),
                    handles_before,
                    "{label} {mode:?} boundary {nth} leaked an owned handle"
                );
                assert_payload_event_balance(
                    payloads_before,
                    &format!("{label} {mode:?} boundary {nth}"),
                );
            }
        }
        take_owning_boundary_trace();
    }

    #[test]
    fn safearray_descriptor_layout_matches_com_header_shape() {
        let pointer_sized_offsets = core::mem::size_of::<usize>() == 8;
        assert_eq!(core::mem::size_of::<SafeArrayBound>(), 8);
        assert_eq!(core::mem::align_of::<SafeArrayBound>(), 4);
        assert_eq!(
            core::mem::size_of::<RawSafeArray>(),
            if pointer_sized_offsets { 32 } else { 24 }
        );
        assert_eq!(
            core::mem::align_of::<RawSafeArray>(),
            core::mem::align_of::<*mut core::ffi::c_void>()
        );
        assert_eq!(
            offset_of::<RawSafeArray, _>(|base| unsafe { core::ptr::addr_of!((*base).c_dims) }),
            0
        );
        assert_eq!(
            offset_of::<RawSafeArray, _>(|base| unsafe { core::ptr::addr_of!((*base).f_features) }),
            2
        );
        assert_eq!(
            offset_of::<RawSafeArray, _>(|base| unsafe {
                core::ptr::addr_of!((*base).cb_elements)
            }),
            4
        );
        assert_eq!(
            offset_of::<RawSafeArray, _>(|base| unsafe { core::ptr::addr_of!((*base).c_locks) }),
            8
        );
        assert_eq!(
            offset_of::<RawSafeArray, _>(|base| unsafe { core::ptr::addr_of!((*base).pv_data) }),
            if pointer_sized_offsets { 16 } else { 12 }
        );
        assert_eq!(
            offset_of::<RawSafeArray, _>(|base| unsafe { core::ptr::addr_of!((*base).rgsabound) }),
            if pointer_sized_offsets { 24 } else { 16 }
        );
    }

    #[test]
    fn storage_order_is_first_dimension_fastest_with_reversed_rgsabound() {
        // Excel x64 (`arr6`): `f(0 To 1, 0 To 2)` stores `f(1,0)` at element 1 and `f(0,1)`
        // at element 2, with `rgsabound = {3,0},{2,0}`.
        let bounds = vec![
            SafeArrayBound { lower: 0, count: 2 },
            SafeArrayBound { lower: 0, count: 3 },
        ];
        assert_eq!(element_index(&bounds, &[1, 0]), Some(1));
        assert_eq!(element_index(&bounds, &[0, 1]), Some(2));
        assert_eq!(element_index(&bounds, &[1, 2]), Some(5));
        assert_eq!(element_index(&bounds, &[2, 0]), None);
        assert_eq!(element_index(&bounds, &[0]), None);
        assert_eq!(element_subscripts(&bounds, 3), Some(vec![1, 1]));
        assert_eq!(element_subscripts(&bounds, 6), None);
        let mut subscripts = [0, 0];
        let mut walked = vec![subscripts];
        while advance_subscripts(&mut subscripts, &bounds) {
            walked.push(subscripts);
        }
        assert_eq!(walked, [[0, 0], [1, 0], [0, 1], [1, 1], [0, 2], [1, 2]]);
        assert_eq!(subscripts, [0, 0]);
        let grown = vec![bounds[0], SafeArrayBound { lower: 0, count: 5 }];
        assert!((0..6).all(|i| preserved_element_index(i, &bounds, &grown) == Some(i)));
        let shrunk = vec![bounds[0], SafeArrayBound { lower: 0, count: 1 }];
        assert_eq!(preserved_element_index(1, &bounds, &shrunk), Some(1));
        assert_eq!(preserved_element_index(2, &bounds, &shrunk), None);

        let array =
            SafeArray::from_variants_nd(bounds.clone(), (0..6).map(Variant::from_i32).collect());
        assert_eq!(array.bounds(), Some(bounds.clone()));
        // SAFETY: the live descriptor has two bounds.
        let stored = unsafe { raw_bounds_slice(array.header.as_ptr()) };
        assert_eq!(stored, [bounds[1], bounds[0]]);
    }

    #[test]
    fn safe_array_tag_roundtrip_for_vector_shape() {
        let tag = ARRAY_TAG_BASE + 3;
        let array = safe_array_from_tag(tag).expect("array tag should decode");
        assert_eq!(array.len(), 3);
        assert_eq!(array.dimensions(), 1);
        assert_eq!(array.element_vartype(), VT_VARIANT_VALUE);
        assert_eq!(array_tag_from_safe_array(&array), Some(tag));
    }

    #[test]
    fn marshal_dispatch_argument_distinguishes_array_tags() {
        assert_eq!(marshal_dispatch_argument(9), 9);
        assert_eq!(marshal_dispatch_argument(ARRAY_TAG_BASE + 4), 20_004);
        assert_eq!(array_len_from_tag(ARRAY_TAG_BASE + 2), Some(2));
    }

    #[test]
    fn safe_array_from_values_preserves_owned_payload_shape() {
        let array = SafeArray::from_variants(vec![Variant::from_i32(4), Variant::from_i32(9)]);
        assert_eq!(array.dimensions(), 1);
        assert_eq!(array.len(), 2);
        assert_eq!(array.effective_len(), 2);
        assert_eq!(array.element_vartype(), VT_VARIANT_VALUE);
        assert_eq!(
            array.variant_elements(),
            Some(vec![Variant::from_i32(4), Variant::from_i32(9)])
        );
        assert_eq!(array_tag_from_safe_array(&array), Some(ARRAY_TAG_BASE + 2));
    }

    #[test]
    fn safe_array_variant_api_preserves_canonical_payload_shape() {
        let array = SafeArray::from_variants(vec![
            Variant::from_i32(4),
            Variant::from_string(BStr::from("A")),
        ]);
        let elements = array
            .variant_elements()
            .expect("variant SAFEARRAY should expose variants");
        assert_eq!(elements[0], Variant::from_i32(4));
        assert_eq!(elements[1], Variant::from_string(BStr::from("A")));
        let replaced = array
            .replace_variant_elements(vec![
                Variant::from_i32(9),
                Variant::from_string(BStr::from("B")),
            ])
            .expect("replace variant elements");
        assert_eq!(
            replaced.variant_elements(),
            Some(vec![
                Variant::from_i32(9),
                Variant::from_string(BStr::from("B"))
            ])
        );
    }

    #[test]
    fn safearray_descriptor_advertises_variant_vartype_metadata() {
        let array = SafeArray::from_variants(vec![Variant::from_i32(4)]);
        assert_eq!(array.element_vartype(), VT_VARIANT_VALUE);
        assert_eq!(
            array.feature_flags(),
            FADF_HAVEVARTYPE_VALUE | FADF_VARIANT_VALUE
        );
    }

    #[test]
    fn raw_safearray_adoption_reads_the_element_type_from_the_descriptor() {
        let array = SafeArray::from_variants(vec![Variant::from_i32(4)]);
        let raw = array.clone_raw_safearray_ptr();
        // A descriptor that names no element type (a UDT array without its layout) is
        // refused and left with its owner.
        let header = raw.cast::<RawSafeArray>();
        let features = unsafe { (*header).f_features };
        unsafe { (*header).f_features = 0 };
        assert!(unsafe { SafeArray::from_raw_safearray_owned(raw, None) }.is_none());
        unsafe { (*header).f_features = features };
        let adopted = unsafe { SafeArray::from_raw_safearray_owned(raw, None) }
            .expect("a Variant SAFEARRAY is adopted");
        assert_eq!(
            adopted.variant_elements().expect("elements"),
            vec![Variant::from_i32(4)]
        );
    }

    /// `fFeatures`, `cbElements` and the hidden bytes as Excel x64 shows them (2026-11-14,
    /// `temp/variant-layout-probe/arr6b`): Long `&H80` with vt 3 at `psa-4`, String `&H180`,
    /// Variant `&H880`, Object `&H440` with IID_IDispatch at `psa-16`, a UDT `0` with no vt,
    /// and a fixed-size array adds `FADF_STATIC | FADF_FIXEDSIZE` (`&H92`, `&H892`, `&H12`).
    #[test]
    fn descriptors_carry_the_features_and_hidden_bytes_vba_writes() {
        fn hidden(array: &SafeArray) -> [u8; 16] {
            let mut out = [0u8; 16];
            unsafe {
                out.copy_from_slice(core::slice::from_raw_parts(
                    super::hidden_bytes(array.header.as_ptr()),
                    16,
                ))
            };
            out
        }
        fn cb(array: &SafeArray) -> u32 {
            unsafe { (*array.header.as_ptr()).cb_elements }
        }
        let pointer = core::mem::size_of::<usize>() as u32;
        let long = SafeArray::from_typed_variants(VT_I4_VALUE, vec![Variant::from_i32(1)]).unwrap();
        assert_eq!((long.feature_flags(), cb(&long)), (0x80, 4));
        assert_eq!(hidden(&long)[12..], 3u32.to_le_bytes());
        let fixed = long.clone().with_fixed_size(true);
        assert_eq!(fixed.feature_flags(), 0x92);
        assert_eq!(fixed.with_fixed_size(false).feature_flags(), 0x80);

        let text =
            SafeArray::from_typed_variants(VT_BSTR_VALUE, vec![Variant::from_string("a")]).unwrap();
        assert_eq!((text.feature_flags(), cb(&text)), (0x180, pointer));
        assert_eq!(hidden(&text)[12..], 8u32.to_le_bytes());

        let variants = SafeArray::from_variants(vec![Variant::empty()]).with_fixed_size(true);
        assert_eq!(
            (variants.feature_flags(), cb(&variants) as usize),
            (0x892, core::mem::size_of::<Variant>())
        );

        let objects =
            SafeArray::from_typed_variants(VT_DISPATCH_VALUE, vec![Variant::nothing()]).unwrap();
        assert_eq!((objects.feature_flags(), cb(&objects)), (0x440, pointer));
        assert_eq!(
            hidden(&objects),
            [0, 4, 2, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46]
        );
        assert_eq!(objects.element_vartype(), VT_DISPATCH_VALUE);
        assert_eq!(objects.variant_element(0).unwrap(), Variant::nothing());
        assert_eq!(objects.clone().feature_flags(), 0x440);

        let (_, layout) = transactional_record_layouts();
        let records = SafeArray::from_default_records_nd(
            vec![SafeArrayBound { count: 2, lower: 0 }],
            Arc::clone(&layout),
        )
        .unwrap();
        assert_eq!(
            (records.feature_flags(), cb(&records) as usize),
            (0, layout.size())
        );
        assert_eq!(hidden(&records), [0; 16]);
        assert_eq!(records.element_vartype(), VT_RECORD_VALUE);
        assert_eq!(records.with_fixed_size(true).feature_flags(), 0x12);
    }

    #[test]
    fn fixed_size_flag_surfaces_in_feature_flags_and_survives_clone() {
        use super::FADF_FIXEDSIZE_VALUE;

        let dynamic = SafeArray::from_variants(vec![Variant::from_i32(7)]);
        assert!(!dynamic.is_fixed_size());
        assert_eq!(dynamic.feature_flags() & FADF_FIXEDSIZE_VALUE, 0);

        let fixed = SafeArray::from_variants(vec![Variant::from_i32(7)]).with_fixed_size(true);
        assert!(fixed.is_fixed_size());
        assert_eq!(
            fixed.feature_flags() & FADF_FIXEDSIZE_VALUE,
            FADF_FIXEDSIZE_VALUE
        );
        // The non-fixed FADF metadata is untouched alongside the fixed bits.
        assert_eq!(
            fixed.feature_flags(),
            dynamic.feature_flags() | FADF_FIXEDSIZE_VALUE | super::FADF_STATIC_VALUE
        );

        // The descriptor flag travels with the value copy (Variant assignment,
        // ByVal/ByRef copies all clone the SafeArray).
        let cloned = fixed.clone();
        assert!(cloned.is_fixed_size());
        assert_eq!(cloned.variant_elements(), Some(vec![Variant::from_i32(7)]));

        // Clearing the bit is observable, and a cleared array clones as dynamic.
        let recleared = cloned.with_fixed_size(false);
        assert!(!recleared.is_fixed_size());
        assert!(!recleared.clone().is_fixed_size());
    }

    #[test]
    fn fixed_size_flag_travels_through_raw_descriptor_roundtrip() {
        let fixed = SafeArray::from_typed_variants(VT_I2_VALUE, vec![Variant::from_i16(3)])
            .expect("typed array")
            .with_fixed_size(true);
        let raw = fixed.clone_raw_safearray_ptr();
        let adopted = unsafe { SafeArray::from_raw_safearray_owned(raw, None) }
            .expect("OxVba-owned descriptor should adopt");
        assert!(adopted.is_fixed_size());
    }

    #[test]
    fn safe_array_from_values_nd_preserves_multi_dimensional_shape() {
        let bounds = vec![
            SafeArrayBound { lower: 1, count: 3 },
            SafeArrayBound { lower: 1, count: 2 },
        ];
        let values = vec![
            Variant::from_i32(1),
            Variant::from_i32(2),
            Variant::from_i32(3),
            Variant::from_i32(4),
            Variant::from_i32(5),
            Variant::from_i32(6),
        ];
        let array = SafeArray::from_variants_nd(bounds.clone(), values.clone());
        assert_eq!(array.dimensions(), 2);
        assert_eq!(array.len(), 6);
        assert_eq!(array.effective_len(), 6);
        assert_eq!(array.bounds().as_ref(), Some(&bounds));
        assert_eq!(array.variant_elements().as_ref(), Some(&values));
    }

    fn unit_bounds(rank: usize) -> Vec<SafeArrayBound> {
        vec![SafeArrayBound { lower: 0, count: 1 }; rank]
    }

    #[test]
    fn safe_array_dimension_layout_keeps_exact_rank_at_boundaries() {
        for rank in [1usize, 2, 60, 61, 127, 128, 254, 255, 256, 257, 1000, 65535] {
            let before = thread_live_handle_counts();
            let array =
                SafeArray::from_shape_and_variants(unit_bounds(rank), vec![Variant::from_i32(7)])
                    .unwrap_or_else(|e| panic!("rank {rank}: {e}"));
            assert_eq!(array.dimensions() as usize, rank);
            assert_eq!(array.bounds().map(|b| b.len()), Some(rank));
            let clone = array.clone();
            assert_eq!(clone.dimensions() as usize, rank);
            assert_eq!(clone, array);
            drop(clone);
            drop(array);
            assert_eq!(thread_live_handle_counts(), before, "rank {rank} drifted");
        }
    }

    #[test]
    fn safe_array_dimension_layout_rejects_oversized_rank_before_allocation() {
        let before = thread_live_handle_counts();
        let payloads = payload_allocation_events();
        for rank in [65536usize, 65537, 70000] {
            assert!(
                SafeArray::from_shape_and_variants(unit_bounds(rank), vec![Variant::from_i32(1)])
                    .is_err(),
                "rank {rank}"
            );
            assert!(
                SafeArray::from_shape(unit_bounds(rank)).is_err(),
                "rank {rank}"
            );
        }
        assert_eq!(thread_live_handle_counts(), before);
        assert_eq!(payload_allocation_events(), payloads);
    }

    #[test]
    fn safe_array_vector_exposes_descriptor_bounds() {
        let array = SafeArray::vector(5);
        assert_eq!(array.dimensions(), 1);
        assert_eq!(
            array.bounds(),
            Some(vec![SafeArrayBound { lower: 0, count: 5 }])
        );
        assert_eq!(array.variant_elements(), None);
    }

    #[test]
    fn safe_array_zeroed_typed_scalar_payload_is_materialized_and_writable() {
        let mut array = SafeArray::from_zeroed_typed_scalars_nd(
            vec![SafeArrayBound {
                lower: -1,
                count: 3,
            }],
            VT_I4_VALUE,
        )
        .expect("zeroed Long payload");
        assert_eq!(
            array.variant_elements(),
            Some(vec![
                Variant::from_i32(0),
                Variant::from_i32(0),
                Variant::from_i32(0),
            ])
        );
        array
            .set_variant_element(1, &Variant::from_i32(42))
            .expect("write zeroed payload");
        assert_eq!(array.variant_element(1), Ok(Variant::from_i32(42)));
    }

    #[test]
    fn safe_array_replace_elements_preserves_shape() {
        let shape = SafeArray::from_shape(vec![
            SafeArrayBound { lower: 1, count: 2 },
            SafeArrayBound { lower: 4, count: 2 },
        ])
        .expect("shape");
        let replaced = shape
            .replace_variant_elements(vec![
                Variant::from_i32(1),
                Variant::from_i32(2),
                Variant::from_i32(3),
                Variant::from_i32(4),
            ])
            .expect("replace");
        assert_eq!(replaced.bounds(), shape.bounds());
        assert_eq!(replaced.variant_elements().expect("elements").len(), 4);
    }

    #[test]
    fn safe_array_set_variant_element_mutates_payload_without_rebuilding_descriptor() {
        let mut array = SafeArray::from_typed_variants(
            VT_I2_VALUE,
            vec![Variant::from_i16(4), Variant::from_i16(9)],
        )
        .expect("typed array");
        let raw = array.raw_safearray_ptr();

        array
            .set_variant_element(1, &Variant::from_i16(42))
            .expect("set element");

        assert_eq!(array.raw_safearray_ptr(), raw);
        assert_eq!(
            array.variant_elements(),
            Some(vec![Variant::from_i16(4), Variant::from_i16(42)])
        );
    }

    #[test]
    fn safe_array_set_variant_element_replaces_owned_bstr_payload() {
        let mut array =
            SafeArray::from_typed_variants(VT_BSTR_VALUE, vec![Variant::from_string("before")])
                .expect("typed bstr array");
        let raw = array.raw_safearray_ptr();

        array
            .set_variant_element(0, &Variant::from_string("after"))
            .expect("set bstr element");

        assert_eq!(array.raw_safearray_ptr(), raw);
        assert_eq!(
            array.variant_elements(),
            Some(vec![Variant::from_string("after")])
        );
    }

    #[test]
    fn safe_array_set_variant_element_replaces_variant_payload() {
        let mut array = SafeArray::from_variants(vec![Variant::from_i32(1)]);
        let raw = array.raw_safearray_ptr();

        array
            .set_variant_element(0, &Variant::from_string("variant"))
            .expect("set variant element");

        assert_eq!(array.raw_safearray_ptr(), raw);
        assert_eq!(
            array.variant_elements(),
            Some(vec![Variant::from_string("variant")])
        );
    }

    #[test]
    fn native_record_safearray_stores_contiguous_record_payloads() {
        let layout = Arc::new(
            VbaRecordLayout::new(vec![VbaRecordFieldSpec::named(
                "Value",
                VbaRecordFieldKind::Long,
            )])
            .expect("layout"),
        );
        let make_record = |value: i32| {
            let mut record = VbaRecord::new_default(layout.clone()).expect("record");
            let field = record.field_handle(0).expect("record field handle");
            unsafe {
                record
                    .field_mut_ptr(&field)
                    .expect("record field pointer")
                    .cast::<i32>()
                    .write(value);
            }
            record
        };
        let mut array = SafeArray::from_vba_records_nd(
            vec![SafeArrayBound { count: 2, lower: 0 }],
            layout.clone(),
            vec![make_record(10), make_record(20)],
        )
        .expect("record array");

        // A VBA UDT array names no element type in its descriptor (Excel: fFeatures 0).
        assert_eq!(array.element_vartype(), VT_RECORD_VALUE);
        assert_eq!(array.feature_flags(), 0);
        assert!(array.vba_record_layout().is_some());
        assert_eq!(
            array
                .variant_element(1)
                .expect("element")
                .as_vba_record()
                .expect("record")
                .read_field_variant(0)
                .expect("field")
                .as_i32(),
            Some(20)
        );

        array
            .set_variant_element(1, &Variant::from_vba_record(make_record(99)))
            .expect("replace record element");
        assert_eq!(
            array.variant_elements().expect("elements")[1]
                .as_vba_record()
                .expect("record")
                .read_field_variant(0)
                .expect("field")
                .as_i32(),
            Some(99)
        );
    }

    #[test]
    fn safe_array_record_transactional_write_preserves_element_and_descriptor_until_commit() {
        let before_buffers = record_buffer_event_counts();
        {
            let (inner_layout, outer_layout) = transactional_record_layouts();
            let replacement = Variant::from_vba_record(transactional_record(
                &inner_layout,
                &outer_layout,
                "new",
                72,
            ));

            let mut successful =
                transactional_record_array(&inner_layout, &outer_layout, "old", 71);
            let descriptor_before = descriptor_fingerprint(&successful);
            take_owning_boundary_trace();
            successful
                .set_variant_element(0, &replacement)
                .expect("uninjected record SAFEARRAY replacement");
            let success_trace = take_owning_boundary_trace();
            assert!(
                !success_trace.is_empty(),
                "record SAFEARRAY replacement must expose owning boundaries"
            );
            assert_eq!(descriptor_fingerprint(&successful), descriptor_before);
            assert_transactional_record(
                &successful.variant_element(0).expect("new array element"),
                "new",
                72,
            );
            drop(successful);

            for mode in [OwningFailureMode::Error, OwningFailureMode::Panic] {
                for nth in 0..success_trace.len() {
                    let mut target =
                        transactional_record_array(&inner_layout, &outer_layout, "old", 71);
                    let descriptor_before = descriptor_fingerprint(&target);
                    let element_before = record_array_element_bytes(&target, &outer_layout);
                    inject_owning_boundary_failure(nth, mode);
                    let outcome = catch_unwind(AssertUnwindSafe(|| {
                        target.set_variant_element(0, &replacement)
                    }));
                    clear_owning_boundary_failure();
                    let failure_trace = take_owning_boundary_trace();
                    assert_eq!(
                        failure_trace.as_slice(),
                        &success_trace[..=nth],
                        "injected {mode:?} failure missed SAFEARRAY boundary {nth}"
                    );
                    match (mode, outcome) {
                        (OwningFailureMode::Error, Ok(Err(error))) => assert!(
                            error.contains("injected owning clone/allocation failure"),
                            "unexpected record SAFEARRAY error: {error}"
                        ),
                        (OwningFailureMode::Error, Err(_)) => {
                            panic!("injected record SAFEARRAY error unexpectedly unwound")
                        }
                        (OwningFailureMode::Panic, Err(_)) => {}
                        (OwningFailureMode::Error, Ok(Ok(()))) => {
                            panic!("injected record SAFEARRAY error unexpectedly committed")
                        }
                        (OwningFailureMode::Panic, Ok(result)) => {
                            panic!("injected record SAFEARRAY panic did not unwind: {result:?}")
                        }
                    }
                    assert_eq!(descriptor_fingerprint(&target), descriptor_before);
                    assert_eq!(
                        record_array_element_bytes(&target, &outer_layout),
                        element_before,
                        "record SAFEARRAY element bytes changed after {mode:?} boundary {nth}"
                    );
                    assert_transactional_record(
                        &target
                            .variant_element(0)
                            .expect("old array element after failure"),
                        "old",
                        71,
                    );
                }
            }
        }
        clear_owning_boundary_failure();
        take_owning_boundary_trace();
        let after_buffers = record_buffer_event_counts();
        assert_eq!(
            after_buffers.0 - before_buffers.0,
            after_buffers.1 - before_buffers.1,
            "record SAFEARRAY success, error, and unwind paths must balance record owners"
        );
    }

    #[test]
    fn borrowed_carrier_unwind_safety_safearray_bstr_clone_read_and_set_helpers() {
        let handles_before = thread_live_handle_counts();
        {
            let original_bytes = vec![0x41, 0x00, 0x00, 0x42, 0x7f];
            let replacement_bytes = vec![0x52, 0x00, 0x00, 0x53, 0x7e];
            let mut source = SafeArray::from_typed_variants(
                VT_BSTR_VALUE,
                vec![Variant::from_bstr_bytes(&original_bytes)],
            )
            .expect("typed BSTR SAFEARRAY");
            let descriptor = descriptor_fingerprint(&source);
            let raw = source.raw_safearray_ptr();

            for mode in [OwningFailureMode::Error, OwningFailureMode::Panic] {
                inject_owning_boundary_failure(0, mode);
                let outcome = catch_unwind(AssertUnwindSafe(|| source.variant_element(0)));
                clear_owning_boundary_failure();
                assert_eq!(take_owning_boundary_trace(), vec!["borrowed-bstr-clone"]);
                match (mode, outcome) {
                    (OwningFailureMode::Error, Ok(Err(error))) => assert!(
                        error.contains("injected owning clone/allocation failure"),
                        "unexpected typed BSTR read error: {error}"
                    ),
                    (OwningFailureMode::Panic, Err(_)) => {}
                    (OwningFailureMode::Error, Err(_)) => {
                        panic!("fallible typed BSTR read unexpectedly unwound")
                    }
                    (OwningFailureMode::Error, Ok(Ok(_))) => {
                        panic!("injected typed BSTR read unexpectedly succeeded")
                    }
                    (OwningFailureMode::Panic, Ok(result)) => {
                        panic!("injected typed BSTR read panic did not unwind: {result:?}")
                    }
                }
                assert_borrowed_bstr_array_source(&source, &descriptor, &original_bytes);

                // Failure zero happens after the ManuallyDrop raw-descriptor view
                // exists; failure one reaches the nested destructor-free BSTR view.
                for nth in 0..=1 {
                    inject_owning_boundary_failure(nth, mode);
                    let outcome = catch_unwind(AssertUnwindSafe(|| {
                        // SAFETY: `raw` is the live descriptor owned by `source`.
                        unsafe { SafeArray::raw_safearray_variant_element(raw, None, 0) }
                    }));
                    clear_owning_boundary_failure();
                    let expected_trace = if nth == 0 {
                        vec!["borrowed-safearray-variant-read"]
                    } else {
                        vec!["borrowed-safearray-variant-read", "borrowed-bstr-clone"]
                    };
                    assert_eq!(take_owning_boundary_trace(), expected_trace);
                    match (mode, outcome) {
                        (OwningFailureMode::Error, Ok(Err(error))) => assert!(
                            error.contains("injected owning clone/allocation failure"),
                            "unexpected raw typed BSTR read error: {error}"
                        ),
                        (OwningFailureMode::Panic, Err(_)) => {}
                        (OwningFailureMode::Error, Err(_)) => {
                            panic!("fallible raw typed BSTR read unexpectedly unwound")
                        }
                        (OwningFailureMode::Error, Ok(Ok(_))) => {
                            panic!("injected raw typed BSTR read unexpectedly succeeded")
                        }
                        (OwningFailureMode::Panic, Ok(result)) => {
                            panic!("injected raw typed BSTR read panic did not unwind: {result:?}")
                        }
                    }
                    assert_borrowed_bstr_array_source(&source, &descriptor, &original_bytes);
                }

                // A clone allocates the descriptor and payload, then copies each element.
                let clone_trace = [
                    "borrowed-safearray-clone",
                    "safearray-header-adoption",
                    "safearray-intrinsic-payload-allocation",
                    "safearray-intrinsic-element-clone",
                    "borrowed-bstr-clone",
                ];
                for nth in 0..clone_trace.len() {
                    inject_owning_boundary_failure(nth, mode);
                    let outcome = catch_unwind(AssertUnwindSafe(|| {
                        // SAFETY: `raw` remains a live shared descriptor for this clone.
                        unsafe { SafeArray::try_clone_from_raw_safearray(raw, None) }
                    }));
                    clear_owning_boundary_failure();
                    let expected_trace = clone_trace[..=nth].to_vec();
                    assert_eq!(take_owning_boundary_trace(), expected_trace);
                    match (mode, outcome) {
                        (OwningFailureMode::Error, Ok(Err(error))) => assert!(
                            error.contains("injected owning clone/allocation failure"),
                            "unexpected raw SAFEARRAY clone error: {error}"
                        ),
                        (OwningFailureMode::Panic, Err(_)) => {}
                        (OwningFailureMode::Error, Err(_)) => {
                            panic!("fallible raw SAFEARRAY clone unexpectedly unwound")
                        }
                        (OwningFailureMode::Error, Ok(Ok(_))) => {
                            panic!("injected raw SAFEARRAY clone unexpectedly succeeded")
                        }
                        (OwningFailureMode::Panic, Ok(result)) => {
                            panic!("injected raw SAFEARRAY clone panic did not unwind: {result:?}")
                        }
                    }
                    assert_borrowed_bstr_array_source(&source, &descriptor, &original_bytes);
                }

                let replacement = Variant::from_bstr_bytes(&replacement_bytes);
                for nth in 0..=1 {
                    inject_owning_boundary_failure(nth, mode);
                    let outcome = catch_unwind(AssertUnwindSafe(|| {
                        // SAFETY: `raw` is the live, exclusively borrowed descriptor
                        // owned by `source` throughout this mutation attempt.
                        unsafe {
                            SafeArray::set_raw_safearray_variant_element(raw, None, 0, &replacement)
                        }
                    }));
                    clear_owning_boundary_failure();
                    let expected_trace = if nth == 0 {
                        vec!["borrowed-safearray-variant-set"]
                    } else {
                        vec!["borrowed-safearray-variant-set", "borrowed-bstr-clone"]
                    };
                    assert_eq!(take_owning_boundary_trace(), expected_trace);
                    match (mode, outcome) {
                        (OwningFailureMode::Error, Ok(Err(error))) => assert!(
                            error.contains("injected owning clone/allocation failure"),
                            "unexpected raw SAFEARRAY set error: {error}"
                        ),
                        (OwningFailureMode::Panic, Err(_)) => {}
                        (OwningFailureMode::Error, Err(_)) => {
                            panic!("fallible raw SAFEARRAY set unexpectedly unwound")
                        }
                        (OwningFailureMode::Error, Ok(Ok(()))) => {
                            panic!("injected raw SAFEARRAY set unexpectedly committed")
                        }
                        (OwningFailureMode::Panic, Ok(result)) => {
                            panic!("injected raw SAFEARRAY set panic did not unwind: {result:?}")
                        }
                    }
                    assert_borrowed_bstr_array_source(&source, &descriptor, &original_bytes);
                    assert_eq!(replacement.string_bytes(), Some(replacement_bytes.clone()));
                    take_owning_boundary_trace();
                }
            }

            source
                .set_variant_element(0, &Variant::from_bstr_bytes(&replacement_bytes))
                .expect("source remains mutable after injected failures");
            assert_eq!(
                source
                    .variant_element(0)
                    .expect("replacement read")
                    .string_bytes(),
                Some(replacement_bytes)
            );
            take_owning_boundary_trace();
        }
        clear_owning_boundary_failure();
        take_owning_boundary_trace();
        assert_eq!(
            thread_live_handle_counts(),
            handles_before,
            "SAFEARRAY BSTR clone/read/set failures and final drops must balance"
        );
    }

    #[test]
    fn borrowed_carrier_unwind_safety_safearray_scalar_and_composite_raw_helpers() {
        let handles_before = thread_live_handle_counts();
        {
            let mut scalars = SafeArray::from_typed_variants(
                VT_I4_VALUE,
                vec![Variant::from_i32(17), Variant::from_i32(29)],
            )
            .expect("typed I4 SAFEARRAY");
            let scalar_descriptor = descriptor_fingerprint(&scalars);
            let scalar_raw = scalars.raw_safearray_ptr();

            for mode in [OwningFailureMode::Error, OwningFailureMode::Panic] {
                inject_owning_boundary_failure(0, mode);
                let read = catch_unwind(AssertUnwindSafe(|| {
                    // SAFETY: the descriptor remains live and shared for this read.
                    unsafe { SafeArray::raw_safearray_i32_element(scalar_raw, 1) }
                }));
                clear_owning_boundary_failure();
                assert_eq!(
                    take_owning_boundary_trace(),
                    vec!["borrowed-safearray-i32-read"]
                );
                match (mode, read) {
                    (OwningFailureMode::Error, Ok(Err(error))) => assert!(
                        error.contains("injected owning clone/allocation failure"),
                        "unexpected raw I4 read error: {error}"
                    ),
                    (OwningFailureMode::Panic, Err(_)) => {}
                    (OwningFailureMode::Error, Err(_)) => {
                        panic!("fallible raw I4 read unexpectedly unwound")
                    }
                    (OwningFailureMode::Error, Ok(Ok(_))) => {
                        panic!("injected raw I4 read unexpectedly succeeded")
                    }
                    (OwningFailureMode::Panic, Ok(result)) => {
                        panic!("injected raw I4 read panic did not unwind: {result:?}")
                    }
                }
                assert_eq!(descriptor_fingerprint(&scalars), scalar_descriptor);
                assert_eq!(scalars.i32_element(1), Ok(Some(29)));

                inject_owning_boundary_failure(0, mode);
                let set = catch_unwind(AssertUnwindSafe(|| {
                    // SAFETY: the descriptor is exclusively borrowed through
                    // `scalars` for this attempted mutation.
                    unsafe { SafeArray::set_raw_safearray_i32_element(scalar_raw, 1, 99) }
                }));
                clear_owning_boundary_failure();
                assert_eq!(
                    take_owning_boundary_trace(),
                    vec!["borrowed-safearray-i32-set"]
                );
                match (mode, set) {
                    (OwningFailureMode::Error, Ok(Err(error))) => assert!(
                        error.contains("injected owning clone/allocation failure"),
                        "unexpected raw I4 set error: {error}"
                    ),
                    (OwningFailureMode::Panic, Err(_)) => {}
                    (OwningFailureMode::Error, Err(_)) => {
                        panic!("fallible raw I4 set unexpectedly unwound")
                    }
                    (OwningFailureMode::Error, Ok(Ok(_))) => {
                        panic!("injected raw I4 set unexpectedly committed")
                    }
                    (OwningFailureMode::Panic, Ok(result)) => {
                        panic!("injected raw I4 set panic did not unwind: {result:?}")
                    }
                }
                assert_eq!(descriptor_fingerprint(&scalars), scalar_descriptor);
                assert_eq!(scalars.i32_element(1), Ok(Some(29)));

                inject_owning_boundary_failure(0, mode);
                let bounds = catch_unwind(AssertUnwindSafe(|| {
                    // SAFETY: the descriptor remains live and shared for this query.
                    unsafe { SafeArray::raw_safearray_bounds_len(scalar_raw) }
                }));
                clear_owning_boundary_failure();
                assert_eq!(
                    take_owning_boundary_trace(),
                    vec!["borrowed-safearray-bounds"]
                );
                match (mode, bounds) {
                    (OwningFailureMode::Error, Ok(None)) => {}
                    (OwningFailureMode::Panic, Err(_)) => {}
                    (OwningFailureMode::Error, Err(_)) => {
                        panic!("fallible raw bounds query unexpectedly unwound")
                    }
                    (OwningFailureMode::Error, Ok(Some(result))) => {
                        panic!("injected raw bounds error returned a value: {result:?}")
                    }
                    (OwningFailureMode::Panic, Ok(result)) => {
                        panic!("injected raw bounds panic did not unwind: {result:?}")
                    }
                }
                assert_eq!(descriptor_fingerprint(&scalars), scalar_descriptor);
                assert_eq!(scalars.i32_element(1), Ok(Some(29)));
            }
            scalars
                .set_i32_element(1, 31)
                .expect("post-failure scalar mutation");
            assert_eq!(scalars.i32_element(1), Ok(Some(31)));

            let (inner_layout, outer_layout) = transactional_record_layouts();
            let composite = SafeArray::from_variants(vec![
                Variant::from_string("composite text"),
                Variant::from_object_ref(ObjectRef::from_compat_identity(631)),
                Variant::from_safearray(SafeArray::from_variants(vec![Variant::from_string(
                    "nested array text",
                )])),
                Variant::from_vba_record(transactional_record(
                    &inner_layout,
                    &outer_layout,
                    "composite record",
                    632,
                )),
            ]);
            let composite_descriptor = descriptor_fingerprint(&composite);
            let composite_raw = composite.raw_safearray_ptr();
            for mode in [OwningFailureMode::Error, OwningFailureMode::Panic] {
                inject_owning_boundary_failure(0, mode);
                let outcome = catch_unwind(AssertUnwindSafe(|| {
                    // SAFETY: `composite` owns this live descriptor throughout.
                    unsafe { SafeArray::try_clone_from_raw_safearray(composite_raw, None) }
                }));
                clear_owning_boundary_failure();
                assert_eq!(
                    take_owning_boundary_trace(),
                    vec!["borrowed-safearray-clone"]
                );
                match (mode, outcome) {
                    (OwningFailureMode::Error, Ok(Err(error))) => assert!(
                        error.contains("injected owning clone/allocation failure"),
                        "unexpected composite raw-clone error: {error}"
                    ),
                    (OwningFailureMode::Panic, Err(_)) => {}
                    (OwningFailureMode::Error, Err(_)) => {
                        panic!("fallible composite raw clone unexpectedly unwound")
                    }
                    (OwningFailureMode::Error, Ok(Ok(_))) => {
                        panic!("injected composite raw clone unexpectedly succeeded")
                    }
                    (OwningFailureMode::Panic, Ok(result)) => {
                        panic!("injected composite raw-clone panic did not unwind: {result:?}")
                    }
                }
                assert_eq!(descriptor_fingerprint(&composite), composite_descriptor);
                let values = composite
                    .variant_elements()
                    .expect("composite source remains readable");
                assert_eq!(
                    values[0].as_bstr().expect("String value").as_str(),
                    "composite text"
                );
                assert_eq!(
                    values[1]
                        .as_object_ref()
                        .expect("Object value")
                        .compat_identity(),
                    631
                );
                assert_eq!(
                    values[2]
                        .safearray_element(0)
                        .expect("nested SAFEARRAY")
                        .expect("nested element")
                        .as_bstr()
                        .expect("nested String")
                        .as_str(),
                    "nested array text"
                );
                assert_transactional_record(&values[3], "composite record", 632);
                take_owning_boundary_trace();
            }
        }
        clear_owning_boundary_failure();
        take_owning_boundary_trace();
        assert_eq!(
            thread_live_handle_counts(),
            handles_before,
            "scalar/composite raw-helper failures and final drops must balance every carrier"
        );
    }

    #[test]
    fn safe_array_payload_transactional_construction_cleans_prefix_and_header_failures() {
        let (inner_layout, outer_layout) = transactional_record_layouts();

        assert_payload_constructor_sweep(
            "record payload",
            || {
                vec![
                    transactional_record(&inner_layout, &outer_layout, "first", 81),
                    transactional_record(&inner_layout, &outer_layout, "second", 82),
                ]
            },
            |values| {
                SafeArray::from_vba_records_nd(
                    vec![SafeArrayBound {
                        count: 2,
                        lower: -1,
                    }],
                    Arc::clone(&outer_layout),
                    values,
                )
            },
            "safearray-record-element-clone",
            2,
        );

        assert_payload_constructor_sweep(
            "intrinsic Variant payload",
            || {
                vec![
                    Variant::from_string("prefix BSTR"),
                    Variant::from_object_ref(ObjectRef::from_compat_identity(91)),
                    Variant::from_safearray(SafeArray::from_variants(vec![
                        Variant::from_string("nested SAFEARRAY BSTR"),
                        Variant::from_object_ref(ObjectRef::from_compat_identity(92)),
                    ])),
                    Variant::from_vba_record(transactional_record(
                        &inner_layout,
                        &outer_layout,
                        "record Variant",
                        93,
                    )),
                ]
            },
            |values| {
                SafeArray::from_shape_and_variants(
                    vec![SafeArrayBound { count: 4, lower: 3 }],
                    values,
                )
            },
            "safearray-intrinsic-element-clone",
            4,
        );

        clear_owning_boundary_failure();
        take_owning_boundary_trace();
    }

    #[test]
    fn safe_array_variant_element_reads_one_typed_payload_slot() {
        let array = SafeArray::from_typed_variants(
            VT_I2_VALUE,
            vec![
                Variant::from_i16(4),
                Variant::from_i16(9),
                Variant::from_i16(16),
            ],
        )
        .expect("typed array");

        assert_eq!(
            array.variant_element(1).expect("element"),
            Variant::from_i16(9)
        );
        assert_eq!(
            array.variant_element(3).expect_err("out of range"),
            "SAFEARRAY index 3 out of range for length 3"
        );
    }

    #[test]
    fn safe_array_variant_element_reads_owned_bstr_payload_slot() {
        let array = SafeArray::from_typed_variants(
            VT_BSTR_VALUE,
            vec![
                Variant::from_string("first"),
                Variant::from_string("second"),
            ],
        )
        .expect("typed bstr array");

        assert_eq!(
            array.variant_element(1).expect("element"),
            Variant::from_string("second")
        );
    }

    #[test]
    fn typed_i2_safearray_preserves_intrinsic_element_vartype() {
        let array = SafeArray::from_typed_variants(
            VT_I2_VALUE,
            vec![Variant::from_i16(4), Variant::from_i16(9)],
        )
        .expect("typed array");
        assert_eq!(array.element_vartype(), VT_I2_VALUE);
        assert_eq!(array.feature_flags(), FADF_HAVEVARTYPE_VALUE);
        assert_eq!(
            array.variant_elements(),
            Some(vec![Variant::from_i16(4), Variant::from_i16(9)])
        );
    }

    #[test]
    fn safearray_descriptor_advertises_special_typed_element_metadata() {
        let bstr = SafeArray::from_typed_variants(
            VT_BSTR_VALUE,
            vec![Variant::from_string(BStr::from("Alpha"))],
        )
        .expect("typed bstr array");
        assert_eq!(
            bstr.feature_flags(),
            FADF_HAVEVARTYPE_VALUE | FADF_BSTR_VALUE
        );

        let dispatch = SafeArray::from_typed_variants(
            VT_DISPATCH_VALUE,
            vec![Variant::from_object_ref(ObjectRef::from_compat_identity(
                41,
            ))],
        )
        .expect("typed dispatch array");
        assert_eq!(
            dispatch.feature_flags(),
            super::FADF_HAVEIID_VALUE | FADF_DISPATCH_VALUE
        );

        let unknown = SafeArray::from_typed_variants(
            VT_UNKNOWN_VALUE,
            vec![Variant::from_object_ref(ObjectRef::from_compat_identity(
                77,
            ))],
        )
        .expect("typed unknown array");
        assert_eq!(
            unknown.feature_flags(),
            super::FADF_HAVEIID_VALUE | FADF_UNKNOWN_VALUE
        );
    }

    #[test]
    fn typed_safearray_variant_elements_preserve_intrinsic_carriers_before_projection() {
        let array = SafeArray::from_typed_variants(
            VT_I2_VALUE,
            vec![Variant::from_i16(4), Variant::from_i16(9)],
        )
        .expect("typed array");

        assert_eq!(
            array.variant_elements().expect("variant elements"),
            vec![Variant::from_i16(4), Variant::from_i16(9)]
        );
    }

    #[test]
    fn typed_safearray_variant_construction_encodes_intrinsic_carriers_directly() {
        let array = SafeArray::from_typed_variants(
            VT_I2_VALUE,
            vec![Variant::from_i16(4), Variant::from_i16(9)],
        )
        .expect("typed array from Variants");

        assert_eq!(array.element_vartype(), VT_I2_VALUE);
        assert_eq!(
            array.variant_elements().expect("variant elements"),
            vec![Variant::from_i16(4), Variant::from_i16(9)]
        );
    }

    #[test]
    fn typed_exact_safearray_carriers_preserve_intrinsic_payloads() {
        let currency = SafeArray::from_typed_variants(
            VT_CY_VALUE,
            vec![Variant::from_currency_scaled_i64(-42_500)],
        )
        .expect("typed currency array");
        assert_eq!(
            currency.variant_elements().expect("currency elements")[0].as_currency_scaled_i64(),
            Some(-42_500)
        );

        let decimal_value = Decimal96::from_parts(123_450, 0, 0, 3, true);
        let decimal = SafeArray::from_typed_variants(
            VT_DECIMAL_VALUE,
            vec![Variant::from_decimal96(decimal_value)],
        )
        .expect("typed decimal array");
        assert_eq!(
            decimal.variant_elements().expect("decimal elements")[0].as_decimal96(),
            Some(decimal_value)
        );

        let date =
            SafeArray::from_typed_variants(VT_DATE_VALUE, vec![Variant::from_date_f64(46_081.25)])
                .expect("typed date array");
        assert_eq!(
            date.variant_elements().expect("date elements")[0].as_date_f64(),
            Some(46_081.25)
        );

        let boolean = SafeArray::from_typed_variants(VT_BOOL_VALUE, vec![Variant::from_bool(true)])
            .expect("typed bool array");
        assert_eq!(
            boolean.variant_elements().expect("bool elements")[0].as_bool(),
            Some(true)
        );
    }

    #[test]
    fn typed_bstr_safearray_roundtrips_strings_without_variant_normalization() {
        let array = SafeArray::from_typed_variants(
            VT_BSTR_VALUE,
            vec![
                Variant::from_string(BStr::from("Alpha")),
                Variant::from_string(BStr::from("Beta")),
            ],
        )
        .expect("typed bstr array");
        assert_eq!(array.element_vartype(), VT_BSTR_VALUE);
        assert_eq!(
            array.variant_elements(),
            Some(vec![
                Variant::from_string(BStr::from("Alpha")),
                Variant::from_string(BStr::from("Beta")),
            ])
        );
    }

    #[test]
    fn typed_dispatch_safearray_preserves_intrinsic_object_payloads() {
        let object = ObjectRef::from_compat_identity(41);
        let array = SafeArray::from_typed_variants(
            VT_DISPATCH_VALUE,
            vec![Variant::from_object_ref(object.clone())],
        )
        .expect("typed dispatch array");
        assert_eq!(array.element_vartype(), VT_DISPATCH_VALUE);
        let elements = array.variant_elements().expect("typed dispatch elements");
        assert_eq!(elements.len(), 1);
        assert_eq!(elements[0], Variant::from_object_ref(object));
    }

    #[test]
    fn typed_unknown_safearray_preserves_intrinsic_object_payloads() {
        let object = ObjectRef::from_compat_identity(77);
        let array = SafeArray::from_typed_variants(
            VT_UNKNOWN_VALUE,
            vec![Variant::from_object_ref(object.clone())],
        )
        .expect("typed unknown array");
        assert_eq!(array.element_vartype(), VT_UNKNOWN_VALUE);
        let elements = array.variant_elements().expect("typed unknown elements");
        assert_eq!(elements.len(), 1);
        assert_eq!(elements[0], Variant::from_object_ref(object));
    }
}

#[cfg(test)]
mod proptests {
    use super::{ARRAY_TAG_BASE, ARRAY_TAG_LIMIT, array_tag_from_safe_array, safe_array_from_tag};
    use proptest::prelude::*;

    proptest! {
        #[test]
        fn prop_safe_array_tag_roundtrip(len in 0..=1_000_000usize) {
            let tag = ARRAY_TAG_BASE + len as i32;
            prop_assert!((ARRAY_TAG_BASE..=ARRAY_TAG_LIMIT).contains(&tag));

            let array = safe_array_from_tag(tag)
                .expect("tag in valid range should decode to SafeArray");
            prop_assert_eq!(array.len(), len);
            prop_assert_eq!(array.dimensions(), 1);

            let recovered_tag = array_tag_from_safe_array(&array)
                .expect("decoded SafeArray should encode back to a tag");
            prop_assert_eq!(recovered_tag, tag);
        }
    }
}

/// The storage index of `subscripts` (absolute, declared order) in an array with `bounds`
/// (declared order). Elements are stored first dimension fastest, as VBA and oleaut32 store
/// them: in `a(0 To 1, 0 To 2)`, `a(1, 0)` is element 1 and `a(0, 1)` element 2. `None` when
/// the rank differs or a subscript is outside its dimension (VBA error 9).
pub fn element_index(bounds: &[SafeArrayBound], subscripts: &[i64]) -> Option<usize> {
    if bounds.len() != subscripts.len() {
        return None;
    }
    let mut index = 0usize;
    let mut stride = 1usize;
    for (bound, &subscript) in bounds.iter().zip(subscripts) {
        let offset = subscript.checked_sub(i64::from(bound.lower))?;
        if offset < 0 || offset >= i64::from(bound.count) {
            return None;
        }
        index = index.checked_add(usize::try_from(offset).ok()?.checked_mul(stride)?)?;
        stride = stride.checked_mul(bound.count as usize)?;
    }
    Some(index)
}

/// The absolute subscripts (declared order) of storage element `index`; the inverse of
/// [`element_index`]. `None` when `index` is past the last element.
pub fn element_subscripts(bounds: &[SafeArrayBound], mut index: usize) -> Option<Vec<i64>> {
    let mut subscripts = Vec::with_capacity(bounds.len());
    for bound in bounds {
        let count = bound.count as usize;
        if count == 0 {
            return None;
        }
        subscripts.push(i64::from(bound.lower) + (index % count) as i64);
        index /= count;
    }
    (index == 0).then_some(subscripts)
}

/// Step absolute subscripts (declared order) to the next element in storage order, first
/// dimension fastest; the order `SafeArrayGetElement` tuples must follow to pair with storage
/// indices. Returns `false`, with every subscript back at its lower bound, after the last
/// element.
pub fn advance_subscripts(subscripts: &mut [i32], bounds: &[SafeArrayBound]) -> bool {
    for (subscript, bound) in subscripts.iter_mut().zip(bounds) {
        let next = i64::from(*subscript) + 1;
        if next < i64::from(bound.lower) + i64::from(bound.count)
            && let Ok(next) = i32::try_from(next)
        {
            *subscript = next;
            return true;
        }
        *subscript = bound.lower;
    }
    false
}

/// Where element `old_index` of an array with `old_bounds` lands after `ReDim Preserve` to
/// `new_bounds` (same rank), keeping its subscripts; `None` when the new bounds drop it.
/// Since VBA only lets `ReDim Preserve` change the last dimension's upper bound, the
/// surviving elements keep their storage index.
pub fn preserved_element_index(
    old_index: usize,
    old_bounds: &[SafeArrayBound],
    new_bounds: &[SafeArrayBound],
) -> Option<usize> {
    element_index(new_bounds, &element_subscripts(old_bounds, old_index)?)
}
