//! Native Automation carriers. These helpers perform individual value
//! operations; no source, bytecode, procedure token or CFG is executed here.
use calcvba_runtime::bstr::BStr;
use std::{ffi::c_void, ptr};
use windows_sys::Win32::{
    Foundation::{SysAllocStringByteLen, SysFreeString, SysStringByteLen, SysStringLen},
    Globalization::{CP_ACP, MultiByteToWideChar, WideCharToMultiByte},
    System::{Com::SAFEARRAY, Ole::*, Variant::*},
};

fn status(hr: i32) -> Result<(), i32> {
    if hr >= 0 {
        return Ok(());
    }
    Err(match hr as u32 {
        0x80020005 => 13,
        0x8002000a => 6,
        0x8002000b => 9,
        0x80020012 => 11,
        0x8007000e => 7,
        0x80070057 => 5,
        _ => hr,
    })
}

fn boundary(work: impl FnOnce() -> Result<(), i32>) -> i32 {
    // SAFETY: GetLastError reads only the calling thread's Win32 state.
    let saved = unsafe { windows_sys::Win32::Foundation::GetLastError() };
    let result = match std::panic::catch_unwind(std::panic::AssertUnwindSafe(work)) {
        Ok(Ok(())) => 0,
        Ok(Err(code)) => code,
        Err(_) => 7,
    };
    // SAFETY: restore the saved calling-thread error after runtime housekeeping.
    unsafe {
        windows_sys::Win32::Foundation::SetLastError(saved);
    }
    result
}

struct OwnedVariant(VARIANT);
impl OwnedVariant {
    fn empty() -> Self {
        // SAFETY: all-zero is VT_EMPTY, including the reserved union tail.
        Self(unsafe { std::mem::zeroed() })
    }
}
impl Drop for OwnedVariant {
    fn drop(&mut self) {
        // SAFETY: this wrapper owns the initialized Automation value.
        unsafe {
            VariantClear(&mut self.0);
        }
    }
}

fn size(kind: u32) -> usize {
    match kind {
        2 | 11 => 2,
        3 | 4 | 10 => 4,
        17 => 1,
        12 => std::mem::size_of::<VARIANT>(),
        _ => 8,
    }
}

// No ownership is taken: the generated caller retains the source cell.
unsafe fn borrowed(src: *const u8, kind: u32) -> VARIANT {
    // SAFETY: emitted code supplies the exact live native carrier for kind.
    unsafe {
        if kind == 12 {
            return src.cast::<VARIANT>().read();
        }
        let mut value: VARIANT = std::mem::zeroed();
        value.Anonymous.Anonymous.vt = kind as u16;
        ptr::copy_nonoverlapping(
            src,
            (&mut value as *mut VARIANT).cast::<u8>().add(8),
            size(kind),
        );
        value
    }
}

unsafe fn retain(pointer: *mut c_void) {
    if !pointer.is_null() {
        // SAFETY: Object/Interface parameters are live IUnknown-shaped pointers.
        unsafe {
            let table = *pointer.cast::<*const *const c_void>();
            let call: unsafe extern "system" fn(*mut c_void) -> u32 =
                std::mem::transmute(*table.add(1));
            call(pointer);
        }
    }
}
unsafe fn release(pointer: *mut c_void) {
    if !pointer.is_null() {
        // SAFETY: the native cell owns one live COM reference.
        unsafe {
            let table = *pointer.cast::<*const *const c_void>();
            let call: unsafe extern "system" fn(*mut c_void) -> u32 =
                std::mem::transmute(*table.add(2));
            call(pointer);
        }
    }
}

unsafe fn clear(cell: *mut u8, kind: u32) -> Result<(), i32> {
    // SAFETY: cell is initialized native storage of this exact kind.
    unsafe {
        match kind {
            8 => SysFreeString(cell.cast::<*mut u16>().read()),
            9 | 13 => release(cell.cast::<*mut c_void>().read()),
            12 => status(VariantClear(cell.cast()))?,
            k if k & 0x2000 != 0 => {
                let array = cell.cast::<*mut SAFEARRAY>().read();
                if !array.is_null() {
                    status(SafeArrayDestroy(array))?;
                }
            }
            _ => {}
        }
        ptr::write_bytes(cell, 0, size(kind));
    }
    Ok(())
}

unsafe fn copy(dst: *mut u8, src: *const u8, to: u32, from: u32) -> Result<(), i32> {
    // SAFETY: caller supplies live exact typed cells, permitting self-aliasing.
    unsafe {
        if to == from {
            match to {
                8 => {
                    let value =
                        BStr::clone_from_raw_bstr(src.cast::<*mut u16>().read()).map_err(|_| 7)?;
                    clear(dst, to)?;
                    dst.cast::<*mut u16>().write(value.into_raw_bstr());
                    return Ok(());
                }
                9 | 13 => {
                    let value = src.cast::<*mut c_void>().read();
                    retain(value);
                    clear(dst, to)?;
                    dst.cast::<*mut c_void>().write(value);
                    return Ok(());
                }
                k if k & 0x2000 != 0 => {
                    let mut value = ptr::null_mut();
                    let source = src.cast::<*mut SAFEARRAY>().read();
                    if !source.is_null() {
                        status(SafeArrayCopy(source, &mut value))?;
                    }
                    if let Err(code) = clear(dst, to) {
                        if !value.is_null() {
                            SafeArrayDestroy(value);
                        }
                        return Err(code);
                    }
                    dst.cast::<*mut SAFEARRAY>().write(value);
                    return Ok(());
                }
                12 => {}
                _ => {
                    ptr::copy(src, dst, size(to));
                    return Ok(());
                }
            }
        }
        let input = borrowed(src, from);
        if from == 12 && input.Anonymous.Anonymous.vt == 1 && to != 12 {
            return Err(94);
        }
        let mut result = OwnedVariant::empty();
        if to == 12 {
            status(VariantCopyInd(&mut result.0, &input))?;
        } else {
            status(VariantChangeType(
                &mut result.0,
                &input,
                if to == 8 { 2 } else { 0 },
                to as u16,
            ))?;
        }
        clear(dst, to)?;
        if to == 12 {
            dst.cast::<VARIANT>().write(result.0);
        } else {
            ptr::copy_nonoverlapping(
                (&result.0 as *const VARIANT).cast::<u8>().add(8),
                dst,
                size(to),
            );
        }
        // Ownership of any BSTR/object/array payload now belongs to dst.
        result.0.Anonymous.Anonymous.vt = 0;
    }
    Ok(())
}

/// # Safety
/// Exact initialized native cells for the given Automation types; aliases are
/// allowed. Strings are BSTR, arrays SAFEARRAY*, and Variant is Windows VARIANT.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_value_copy(dst: *mut u8, src: *const u8, to: u32, from: u32) -> i32 {
    // SAFETY: forwards the exact live-cell contract documented on this export.
    boundary(|| unsafe { copy(dst, src, to, from) })
}
/// # Safety
/// cell owns one initialized native value of kind and is writable.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_value_clear(cell: *mut u8, kind: u32) -> i32 {
    // SAFETY: the caller owns this initialized, writable native cell.
    boundary(|| unsafe { clear(cell, kind) })
}

/// # Safety
/// out is a writable VARIANT_BOOL; src borrows an initialized cell of kind.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_value_truthy(out: *mut i16, src: *const u8, kind: u32) -> i32 {
    boundary(|| {
        // SAFETY: exact retained native cells, with conversion in private storage.
        unsafe {
            let value = borrowed(src, kind);
            let truth = if value.Anonymous.Anonymous.vt <= 1 {
                false
            } else {
                let mut number = OwnedVariant::empty();
                status(VariantChangeType(&mut number.0, &value, 0, 5))?;
                number.0.Anonymous.Anonymous.Anonymous.dblVal != 0.0
            };
            out.write(if truth { -1 } else { 0 });
        }
        Ok(())
    })
}

/// # Safety
/// src is a nullable terminated LPWSTR (encoding 1) or system-codepage LPSTR
/// (encoding 2); out owns an initialized BSTR cell. The input remains borrowed.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_string_import(src: *const u8, encoding: u32, out: *mut u8) -> i32 {
    boundary(|| {
        // SAFETY: the native caller supplies a terminated input or null.
        let units = unsafe {
            if src.is_null() {
                Vec::new()
            } else if encoding == 1 {
                let mut len = 0;
                while src.cast::<u16>().add(len).read() != 0 {
                    len += 1;
                }
                std::slice::from_raw_parts(src.cast::<u16>(), len).to_vec()
            } else if encoding == 2 {
                let count = MultiByteToWideChar(CP_ACP, 0, src, -1, ptr::null_mut(), 0);
                if count <= 0 {
                    return Err(5);
                }
                let mut units = vec![0; count as usize];
                if MultiByteToWideChar(CP_ACP, 0, src, -1, units.as_mut_ptr(), count) != count {
                    return Err(5);
                }
                units.pop();
                units
            } else {
                return Err(5);
            }
        };
        let value = BStr::from_utf16_units(&units).map_err(|_| 7)?;
        // SAFETY: out is a caller-owned initialized BSTR cell.
        unsafe {
            clear(out, 8)?;
            out.cast::<*mut u16>().write(value.into_raw_bstr());
        }
        Ok(())
    })
}

/// # Safety
/// out/a/b are initialized BSTR cells, with arbitrary aliases.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_string_concat(out: *mut u8, a: *const u8, b: *const u8) -> i32 {
    boundary(|| {
        // SAFETY: borrowed BSTRs stay owned by their input cells until copy-out.
        let units = unsafe {
            let a = a.cast::<*mut u16>().read();
            let b = b.cast::<*mut u16>().read();
            let mut units = Vec::new();
            if !a.is_null() {
                units.extend_from_slice(std::slice::from_raw_parts(a, SysStringLen(a) as usize));
            }
            if !b.is_null() {
                units.extend_from_slice(std::slice::from_raw_parts(b, SysStringLen(b) as usize));
            }
            units
        };
        let result = BStr::from_utf16_units(&units).map_err(|_| 7)?;
        // SAFETY: allocate before releasing out, including when inputs alias it.
        unsafe {
            clear(out, 8)?;
            out.cast::<*mut u16>().write(result.into_raw_bstr());
        }
        Ok(())
    })
}

/// # Safety
/// input is a borrowed initialized BSTR cell and out is a writable Long cell.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_string_len(input: *const *mut u16, out: *mut i32) -> i32 {
    boundary(|| {
        // SAFETY: SysStringLen accepts null and reads the BSTR length prefix.
        let len = unsafe { SysStringLen(input.read()) };
        let len = i32::try_from(len).map_err(|_| 6)?;
        // SAFETY: the generated caller supplied a writable Long output cell.
        unsafe {
            out.write(len);
        }
        Ok(())
    })
}

type Binary = unsafe extern "system" fn(*const VARIANT, *const VARIANT, *mut VARIANT) -> i32;
unsafe fn binary(
    out: *mut u8,
    a: *const u8,
    b: *const u8,
    to: u32,
    ak: u32,
    bk: u32,
    op: Binary,
) -> Result<(), i32> {
    // SAFETY: generated operands retain both initialized exact carriers.
    unsafe {
        let a = borrowed(a, ak);
        let b = borrowed(b, bk);
        let mut result = OwnedVariant::empty();
        status(op(&a, &b, &mut result.0))?;
        copy(out, (&result.0 as *const VARIANT).cast(), to, 12)
    }
}
macro_rules! binary_helper {
    ($name:ident, $operation:ident) => {
        /// # Safety
        /// Live initialized native cells of to/ak/bk; arbitrary aliases allowed.
        #[unsafe(no_mangle)]
        pub unsafe extern "C" fn $name(
            out: *mut u8,
            a: *const u8,
            b: *const u8,
            to: u32,
            ak: u32,
            bk: u32,
        ) -> i32 {
            // SAFETY: forwards the retained typed input/output carrier contract.
            boundary(|| unsafe { binary(out, a, b, to, ak, bk, $operation) })
        }
    };
}
binary_helper!(cvb_value_add, VarAdd);
binary_helper!(cvb_value_sub, VarSub);
binary_helper!(cvb_value_mul, VarMul);
binary_helper!(cvb_value_div, VarDiv);
binary_helper!(cvb_value_concat, VarCat);
binary_helper!(cvb_value_and, VarAnd);
binary_helper!(cvb_value_or, VarOr);
binary_helper!(cvb_value_xor, VarXor);
binary_helper!(cvb_value_eqv, VarEqv);
binary_helper!(cvb_value_imp, VarImp);

/// # Safety
/// Live initialized typed input cells and a writable Boolean output cell.
#[unsafe(no_mangle)]
#[allow(clippy::too_many_arguments)] // Explicit typed operand and result carriers.
pub unsafe extern "C" fn cvb_value_compare(
    out: *mut u8,
    a: *const u8,
    b: *const u8,
    ak: u32,
    bk: u32,
    op: u32,
    text: u32,
    to: u32,
) -> i32 {
    boundary(|| {
        // SAFETY: caller supplies exact native cells; comparisons borrow them.
        let order = unsafe {
            if ak == 8 && bk == 8 {
                let a = a.cast::<*mut u16>().read();
                let b = b.cast::<*mut u16>().read();
                if text != 0 {
                    let result = VarBstrCmp(a, b, 0x400, 1);
                    status(result)?;
                    result as u32
                } else {
                    let read = |p: *mut u16| {
                        if p.is_null() {
                            Vec::new()
                        } else {
                            std::slice::from_raw_parts(p, SysStringLen(p) as usize)
                                .iter()
                                .map(|&u| {
                                    if text != 0 && (65..=90).contains(&u) {
                                        u + 32
                                    } else {
                                        u
                                    }
                                })
                                .collect::<Vec<_>>()
                        }
                    };
                    match read(a).cmp(&read(b)) {
                        std::cmp::Ordering::Less => 0,
                        std::cmp::Ordering::Equal => 1,
                        std::cmp::Ordering::Greater => 2,
                    }
                }
            } else {
                let a = borrowed(a, ak);
                let b = borrowed(b, bk);
                let result = VarCmp(&a, &b, 0x400, u32::from(text != 0));
                if result == 3 {
                    if to != 12 {
                        return Err(94);
                    }
                    let mut null = OwnedVariant::empty();
                    null.0.Anonymous.Anonymous.vt = 1;
                    return copy(out, (&null.0 as *const VARIANT).cast(), to, 12);
                }
                status(result as i32)?;
                result as u32
            }
        };
        let value = match op {
            0 => order == 1,
            1 => order != 1,
            2 => order == 0,
            3 => order != 2,
            4 => order == 2,
            5 => order != 0,
            _ => return Err(5),
        };
        let value: i16 = if value { -1 } else { 0 };
        // SAFETY: the local Boolean is live and out is the declared result cell.
        unsafe { copy(out, (&value as *const i16).cast(), to, 11) }
    })
}

macro_rules! unary_helper {
    ($name:ident, $operation:ident) => {
        /// # Safety
        /// Live initialized native input/output cells of from/to.
        #[unsafe(no_mangle)]
        pub unsafe extern "C" fn $name(out: *mut u8, src: *const u8, to: u32, from: u32) -> i32 {
            // SAFETY: source/output are the retained exact native typed cells.
            boundary(|| unsafe {
                let value = borrowed(src, from);
                let mut result = OwnedVariant::empty();
                status($operation(&value, &mut result.0))?;
                copy(out, (&result.0 as *const VARIANT).cast(), to, 12)
            })
        }
    };
}
unary_helper!(cvb_value_neg, VarNeg);
unary_helper!(cvb_value_not, VarNot);

// Record descriptor: size:u32, owning_count:u32, then (offset:u32, kind:u32).
// Nested records/fixed arrays are flattened at compile time. Plain bytes retain
// their C/VBA offsets; only owning fields need cloning and destruction.
unsafe fn record_shape<'a>(descriptor: *const u32) -> (usize, &'a [u32]) {
    // SAFETY: compiler-generated immutable, aligned, bounded descriptor.
    unsafe {
        (
            descriptor.read() as usize,
            std::slice::from_raw_parts(descriptor.add(2), descriptor.add(1).read() as usize * 2),
        )
    }
}
unsafe fn record_clear(cell: *mut u8, descriptor: *const u32) -> Result<(), i32> {
    // SAFETY: matching initialized record storage and compiler descriptor.
    unsafe {
        let (size, fields) = record_shape(descriptor);
        for field in fields.chunks_exact(2) {
            clear(cell.add(field[0] as usize), field[1])?;
        }
        ptr::write_bytes(cell, 0, size);
    }
    Ok(())
}
/// # Safety
/// cell is a writable initialized record with the compiler-generated descriptor.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_record_clear(cell: *mut u8, descriptor: *const u32) -> i32 {
    // SAFETY: caller supplies initialized storage and its immutable descriptor.
    boundary(|| unsafe { record_clear(cell, descriptor) })
}
/// # Safety
/// dst/src are matching initialized records; self-aliasing is supported.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_record_copy(
    dst: *mut u8,
    src: *const u8,
    descriptor: *const u32,
) -> i32 {
    boundary(|| {
        // SAFETY: clone every owner before replacing any destination bytes.
        unsafe {
            let (size, fields) = record_shape(descriptor);
            let mut words = vec![0u64; size.div_ceil(8)];
            let staging = words.as_mut_ptr().cast::<u8>();
            ptr::copy_nonoverlapping(src, staging, size);
            for field in fields.chunks_exact(2) {
                ptr::write_bytes(staging.add(field[0] as usize), 0, self::size(field[1]));
            }
            for field in fields.chunks_exact(2) {
                let offset = field[0] as usize;
                if let Err(error) = copy(staging.add(offset), src.add(offset), field[1], field[1]) {
                    record_clear(staging, descriptor)?;
                    return Err(error);
                }
            }
            if let Err(error) = record_clear(dst, descriptor) {
                record_clear(staging, descriptor)?;
                return Err(error);
            }
            ptr::copy_nonoverlapping(staging, dst, size);
        }
        Ok(())
    })
}

/// # Safety
/// array is a live SAFEARRAY* cell, indices spans rank Longs, and out owns an
/// initialized element cell of the exact array element kind.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_array_get(
    array: *const *mut SAFEARRAY,
    indices: *const i32,
    rank: u32,
    kind: u32,
    out: *mut u8,
    to: u32,
) -> i32 {
    boundary(|| {
        // SAFETY: inputs retain a live matching descriptor and index vector.
        unsafe {
            let array = array.read();
            if array.is_null() {
                return Err(9);
            }
            if SafeArrayGetDim(array) != rank {
                return Err(9);
            }
            let mut actual = 0;
            status(SafeArrayGetVartype(array, &mut actual))?;
            // VBA Declare converts String() to byte BSTRs in an otherwise
            // ordinary SAFEARRAY. Its descriptor reports VT_EMPTY, FADF_BSTR.
            let ansi = kind == 8 && actual == 0 && (*array).fFeatures & 0x100 != 0;
            if u32::from(actual) != kind && !ansi {
                return Err(13);
            }
            let mut words = [0u64; 3];
            let staging = words.as_mut_ptr().cast::<u8>();
            status(SafeArrayGetElement(array, indices, staging.cast()))?;
            if ansi {
                let original = staging.cast::<*mut u16>().read();
                let bytes = SysStringByteLen(original);
                let wide = ansi_to_bstr(original.cast(), bytes as i32);
                SysFreeString(original);
                staging.cast::<*mut u16>().write(wide?);
            }
            let result = copy(out, staging, to, kind);
            clear(staging, kind)?;
            result
        }
    })
}
/// # Safety
/// Exact live array, index vector and borrowed element cell, as for array_get.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_array_set(
    array: *const *mut SAFEARRAY,
    indices: *const i32,
    rank: u32,
    kind: u32,
    value: *const u8,
    from: u32,
) -> i32 {
    boundary(|| {
        // SAFETY: Automation checks bounds and retains owning element payloads.
        unsafe {
            let array = array.read();
            if array.is_null() {
                return Err(9);
            }
            if SafeArrayGetDim(array) != rank {
                return Err(9);
            }
            let mut actual = 0;
            status(SafeArrayGetVartype(array, &mut actual))?;
            let ansi = kind == 8 && actual == 0 && (*array).fFeatures & 0x100 != 0;
            if u32::from(actual) != kind && !ansi {
                return Err(13);
            }
            let mut words = [0u64; 3];
            let cell = words.as_mut_ptr().cast::<u8>();
            copy(cell, value, kind, from)?;
            if ansi {
                let original = cell.cast::<*mut u16>().read();
                let bytes = bstr_to_ansi(original);
                SysFreeString(original);
                cell.cast::<*mut u16>().write(bytes?);
            }
            let value = if matches!(kind, 8 | 9 | 13) {
                cell.cast::<*const u8>().read()
            } else {
                cell
            };
            let result = status(SafeArrayPutElement(array, indices, value.cast()));
            clear(cell, kind)?;
            result
        }
    })
}

// Lengths are explicit so embedded NULs survive VBA's ANSI Declare conversion.
unsafe fn ansi_to_bstr(src: *const u8, len: i32) -> Result<*mut u16, i32> {
    if len == 0 {
        return Ok(ptr::null_mut());
    }
    // SAFETY: the retained byte BSTR spans len initialized bytes.
    unsafe {
        let count = MultiByteToWideChar(CP_ACP, 0, src, len, ptr::null_mut(), 0);
        if count == 0 {
            return Err(5);
        }
        let mut units = vec![0u16; count as usize];
        if MultiByteToWideChar(CP_ACP, 0, src, len, units.as_mut_ptr(), count) != count {
            return Err(5);
        }
        Ok(BStr::from_utf16_units(&units)
            .map_err(|_| 7)?
            .into_raw_bstr())
    }
}
pub(super) unsafe fn bstr_to_ansi(src: *mut u16) -> Result<*mut u16, i32> {
    // SAFETY: input is an owning Unicode BSTR retained by the caller.
    unsafe {
        let len = SysStringLen(src) as i32;
        if len == 0 {
            return Ok(ptr::null_mut());
        }
        let count = WideCharToMultiByte(
            CP_ACP,
            0,
            src,
            len,
            ptr::null_mut(),
            0,
            ptr::null(),
            ptr::null_mut(),
        );
        if count == 0 {
            return Err(5);
        }
        let out = SysAllocStringByteLen(ptr::null(), count as u32).cast_mut();
        if out.is_null() {
            return Err(7);
        }
        if WideCharToMultiByte(
            CP_ACP,
            0,
            src,
            len,
            out.cast(),
            count,
            ptr::null(),
            ptr::null_mut(),
        ) != count
        {
            SysFreeString(out);
            return Err(5);
        }
        Ok(out)
    }
}
/// # Safety
/// array is a live SAFEARRAY* cell; out is a writable Long cell.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_array_bound(
    array: *const *mut SAFEARRAY,
    dimension: u32,
    upper: u32,
    out: *mut i32,
) -> i32 {
    boundary(|| {
        // SAFETY: OS validates dimension against the retained descriptor.
        unsafe {
            let array = array.read();
            if array.is_null() {
                return Err(9);
            }
            let hr = if upper != 0 {
                SafeArrayGetUBound(array, dimension, out)
            } else {
                SafeArrayGetLBound(array, dimension, out)
            };
            status(hr)
        }
    })
}

/// # Safety
/// cell owns a nullable SAFEARRAY*; pairs spans rank (lower,upper) Long pairs.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_array_redim(
    cell: *mut *mut SAFEARRAY,
    kind: u32,
    rank: u32,
    pairs: *const i32,
    preserve: u32,
    fixed: u32,
) -> i32 {
    boundary(|| {
        if rank == 0 || rank > 60 || kind == 36 {
            return Err(5);
        }
        // SAFETY: generated bounds span rank pairs; the input array is retained.
        unsafe {
            let pairs = std::slice::from_raw_parts(pairs, rank as usize * 2);
            let mut bounds = Vec::new();
            for pair in pairs.chunks_exact(2) {
                let count = i64::from(pair[1]) - i64::from(pair[0]) + 1;
                let count = u32::try_from(count).map_err(|_| 9)?;
                bounds.push(windows_sys::Win32::System::Com::SAFEARRAYBOUND {
                    cElements: count,
                    lLbound: pair[0],
                });
            }
            let old = cell.read();
            if !old.is_null() && ((*old).fFeatures & 0x10 != 0 || (*old).cLocks != 0) {
                return Err(10);
            }
            if preserve != 0 && !old.is_null() {
                if SafeArrayGetDim(old) != rank {
                    return Err(9);
                }
                for (i, pair) in pairs.chunks_exact(2).enumerate() {
                    let mut lower = 0;
                    let mut upper = 0;
                    status(SafeArrayGetLBound(old, i as u32 + 1, &mut lower))?;
                    status(SafeArrayGetUBound(old, i as u32 + 1, &mut upper))?;
                    if lower != pair[0] || (i + 1 != rank as usize && upper != pair[1]) {
                        return Err(9);
                    }
                }
                return status(SafeArrayRedim(old, bounds.last().expect("nonzero rank")));
            }
            let value = SafeArrayCreate(kind as u16, rank, bounds.as_ptr());
            if value.is_null() {
                return Err(7);
            }
            if let Err(error) = clear(cell.cast(), 0x2000 | kind) {
                SafeArrayDestroy(value);
                return Err(error);
            }
            if fixed != 0 {
                (*value).fFeatures |= 0x10;
            }
            cell.write(value);
        }
        Ok(())
    })
}
