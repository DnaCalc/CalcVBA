//! Individual native Declare calls. Fixed signatures are prepared once per
//! compiled call site; accepted calls and callbacks remain on the owning VBA thread.
use calcvba_automation::abi::VARIANT_SIZE;
use calcvba_com::windows_ffi_bridge::{
    self as ffi, FfiReturnType, OwnedLibrary, PreparedNativeCall,
};
use calcvba_runtime::bstr::BStr;
use calcvba_runtime::{pointer_helpers::AutomationRecord, vba_record::VbaRecordLayout};
use std::sync::Arc;
use std::{
    ffi::{CStr, c_void},
    ptr,
};
use windows_sys::Win32::Foundation::{GetLastError, SetLastError, SysFreeString, SysStringByteLen};

const BYREF: u32 = 0x10000;
const WRITE_STRING: u32 = 0x20000;

struct AnsiReference {
    address: *mut usize,
    destination: *mut u8,
}
impl Drop for AnsiReference {
    fn drop(&mut self) {
        // SAFETY: this cell owns the current ANSI byte BSTR. A conforming callee
        // frees its old BSTR before replacing the pointer through ByRef String.
        unsafe {
            SysFreeString((*self.address as *mut u16).cast_const());
        }
    }
}

struct Import {
    _owner: OwnedLibrary,
    call: PreparedNativeCall,
    parameters: Vec<u32>,
    records: Vec<Option<Arc<VbaRecordLayout>>>,
    result_record: Option<Arc<VbaRecordLayout>>,
    result_kind: u32,
}
fn scalar(kind: u32) -> Result<FfiReturnType, i32> {
    Ok(match kind {
        0 => FfiReturnType::Void,
        17 => FfiReturnType::Byte,
        2 => FfiReturnType::Integer,
        3 => FfiReturnType::Long,
        11 => FfiReturnType::Boolean,
        20 | 6 => FfiReturnType::LongLong,
        8 => FfiReturnType::LongPtr,
        4 => FfiReturnType::Single,
        5 | 7 => FfiReturnType::Double,
        _ => return Err(13),
    })
}
impl Import {
    unsafe fn prepare(metadata: *const u32) -> Result<Self, i32> {
        // SAFETY: all strings, counts and flags are retained compiler data.
        unsafe {
            let count = metadata.add(1).read() as usize;
            if count > 32 {
                return Err(5);
            }
            let library = CStr::from_ptr(
                metadata
                    .cast::<u8>()
                    .add(metadata.add(2).read() as usize)
                    .cast(),
            )
            .to_str()
            .map_err(|_| 53)?;
            let symbol = CStr::from_ptr(
                metadata
                    .cast::<u8>()
                    .add(metadata.add(3).read() as usize)
                    .cast(),
            )
            .to_str()
            .map_err(|_| 453)?;
            let owner = OwnedLibrary::load(library).map_err(|_| 53)?;
            let address = ffi::get_proc_address(owner.handle(), symbol).map_err(|_| 453)?;
            unsafe fn record(
                metadata: *const u32,
                offset: u32,
            ) -> Result<Option<Arc<VbaRecordLayout>>, i32> {
                if offset == 0 {
                    return Ok(None);
                }
                // SAFETY: these compiler-produced offsets address retained type metadata.
                unsafe {
                    let words = metadata.cast::<u8>().add(offset as usize).cast::<u32>();
                    let count = words.read() as usize;
                    if count > 1048576 {
                        return Err(13);
                    }
                    VbaRecordLayout::from_boundary_schema(std::slice::from_raw_parts(
                        words.add(1),
                        count,
                    ))
                    .map(Some)
                    .map_err(|_| 13)
                }
            }
            let parameters = (0..count)
                .map(|i| metadata.add(5 + i * 2).read())
                .collect::<Vec<_>>();
            let records = (0..count)
                .map(|i| record(metadata, metadata.add(6 + i * 2).read()))
                .collect::<Result<Vec<_>, _>>()?;
            let result_record = record(metadata, metadata.add(4).read())?;
            let result_kind = metadata.read();
            let types = parameters
                .iter()
                .map(|flags| {
                    if flags & BYREF != 0 || flags & 0xffff == 8 {
                        Ok(FfiReturnType::LongPtr)
                    } else {
                        scalar(flags & 0xffff)
                    }
                })
                .collect::<Result<Vec<_>, _>>()?;
            let call = if let Some(layout) = &result_record {
                PreparedNativeCall::aggregate_return(
                    address,
                    &types,
                    layout.declare_shape().map_err(|_| 13)?.0,
                )
            } else if result_kind == 12 {
                PreparedNativeCall::aggregate_return(address, &types, VARIANT_SIZE)
            } else {
                PreparedNativeCall::new(address, &types, scalar(result_kind)?)
            };
            Ok(Self {
                _owner: owner,
                call,
                parameters,
                records,
                result_record,
                result_kind,
            })
        }
    }
}

unsafe fn invoke(
    state: *mut usize,
    metadata: *const u32,
    cells: *const *mut u8,
    out: *mut u8,
    last_error: &mut u32,
) -> Result<(), i32> {
    // SAFETY: state is private to this image and its owning thread. Each call
    // keeps its own argument values, so synchronous reentry cannot overwrite a
    // suspended invocation's cells or ANSI conversion buffers.
    unsafe {
        if state.read() == 0 {
            let import = Box::new(Import::prepare(metadata)?);
            state.write(Box::into_raw(import) as usize);
        }
        let import = &*(state.read() as *const Import);
        let mut addresses = [0usize; 32];
        let mut indirects = [0usize; 32];
        let mut arguments = [ptr::null_mut::<c_void>(); 32];
        let mut strings = Vec::new();
        let mut references = Vec::new();
        let mut records = Vec::new();
        for (i, flags) in import.parameters.iter().copied().enumerate() {
            let cell = cells.add(i).read();
            if let Some(layout) = &import.records[i] {
                let mut record =
                    AutomationRecord::from_automation(cell, layout.clone()).map_err(|_| 13)?;
                addresses[i] = record.as_mut_ptr() as usize;
                arguments[i] = (&mut addresses[i] as *mut usize).cast();
                records.push((cell, record));
                continue;
            }
            if flags & 0xffff == 8 {
                let source = cell.cast::<*mut u16>().read();
                let ansi = super::values::bstr_to_ansi(source)?;
                if flags & BYREF != 0 {
                    addresses[i] = ansi as usize;
                    references.push(AnsiReference {
                        address: &mut addresses[i],
                        destination: cell,
                    });
                    indirects[i] = (&mut addresses[i] as *mut usize) as usize;
                    arguments[i] = (&mut indirects[i] as *mut usize).cast();
                    continue;
                }
                let len = SysStringByteLen(ansi) as usize;
                let mut buffer = if len == 0 {
                    Vec::new()
                } else {
                    std::slice::from_raw_parts(ansi.cast::<u8>(), len).to_vec()
                };
                SysFreeString(ansi);
                buffer.push(0);
                let original = buffer.clone();
                addresses[i] = buffer.as_mut_ptr() as usize;
                strings.push((cell, buffer, original, flags & WRITE_STRING != 0));
                arguments[i] = (&mut addresses[i] as *mut usize).cast();
            } else if flags & BYREF != 0 {
                addresses[i] = cell as usize;
                arguments[i] = (&mut addresses[i] as *mut usize).cast();
            } else {
                arguments[i] = cell.cast();
            }
        }
        // Restore the input Win32 error after first-call loading and argument
        // preparation; capture the callee's value before any cleanup runs.
        let mut record_result = import
            .result_record
            .as_ref()
            .map(|layout| AutomationRecord::empty_declared(layout.clone()))
            .transpose()
            .map_err(|_| 13)?;
        let mut variant_result: windows_sys::Win32::System::Variant::VARIANT = std::mem::zeroed();
        let result_pointer = if let Some(record) = &mut record_result {
            record.as_mut_ptr().cast()
        } else if import.result_kind == 12 {
            (&mut variant_result as *mut windows_sys::Win32::System::Variant::VARIANT).cast()
        } else {
            out.cast()
        };
        SetLastError(*last_error);
        *last_error = import
            .call
            .invoke_cells(&mut arguments[..import.parameters.len()], result_pointer);
        let record_status = if let Some(record) = &record_result {
            record.copy_back_automation(out).map_err(|_| 13)
        } else {
            Ok(())
        };
        let variant_status = if import.result_kind == 12 {
            let result = super::values::copy(
                out,
                (&variant_result as *const windows_sys::Win32::System::Variant::VARIANT).cast(),
                12,
                12,
            );
            windows_sys::Win32::System::Variant::VariantClear(&mut variant_result);
            result
        } else {
            Ok(())
        };
        record_status?;
        variant_status?;
        for (cell, record) in records {
            record.copy_back_automation(cell).map_err(|_| 13)?;
        }
        for reference in references {
            let source = (*reference.address as *mut u16).cast_const();
            let value =
                super::values::ansi_to_bstr(source.cast(), SysStringByteLen(source) as i32)?;
            SysFreeString(reference.destination.cast::<*mut u16>().read());
            reference.destination.cast::<*mut u16>().write(value);
        }
        for (cell, buffer, original, writable) in strings {
            if writable && buffer != original {
                let value =
                    BStr::from_utf16_units(&ffi::utf16_from_ansi(&buffer[..buffer.len() - 1]))
                        .map_err(|_| 7)?
                        .into_raw_bstr();
                SysFreeString(cell.cast::<*mut u16>().read());
                cell.cast::<*mut u16>().write(value);
            }
        }
        if metadata.read() == 8 {
            let ansi = out.cast::<*mut u16>().read();
            let unicode = super::values::ansi_to_bstr(ansi.cast(), SysStringByteLen(ansi) as i32);
            SysFreeString(ansi);
            out.cast::<*mut u16>().write(unicode?);
        }
        Ok(())
    }
}

/// # Safety
/// The compiler supplies the retained descriptor, zeroed two-word state,
/// invocation-local argument cells and initialized private result storage.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_declare(
    state: *mut usize,
    metadata: *const u32,
    cells: *const *mut u8,
    out: *mut u8,
) -> i32 {
    // SAFETY: GetLastError reads only calling-thread state.
    let mut error = unsafe { GetLastError() };
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        // SAFETY: forwarded compiler-generated invocation contract above.
        unsafe { invoke(state, metadata, cells, out, &mut error) }
    }));
    let status = match result {
        Ok(Ok(())) => 0,
        Ok(Err(code)) => code,
        Err(_) => 7,
    };
    // SAFETY: all temporary cleanup is done before restoring the callee's error.
    unsafe {
        SetLastError(error);
    }
    status
}

/// # Safety
/// state is this image's private import entry, on its owning thread at shutdown.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_declare_clear(state: *mut usize) -> i32 {
    // SAFETY: release only this entry's recorded LoadLibrary reference and CIF.
    unsafe {
        let error = GetLastError();
        let owner = std::mem::replace(&mut *state, 0);
        state.add(1).write(0);
        if owner != 0 {
            drop(Box::from_raw(owner as *mut Import));
        }
        SetLastError(error);
    }
    0
}
