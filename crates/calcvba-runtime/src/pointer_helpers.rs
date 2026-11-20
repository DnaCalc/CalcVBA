//! Automation-layout projections of VBA records and record arrays for `Declare` calls: the
//! Unicode layout and the ANSI one a `Declare` passes. Records and UDT arrays (`FADF_RECORD`
//! `SAFEARRAY`s with registration-free `IRecordInfo`) project on every platform (the oleaut32
//! calls go through `calcvba_automation::oleaut`, calling real oleaut32 on Windows and the
//! binary-identical portable implementation elsewhere). `VarPtr`, `StrPtr` and `ObjPtr` give
//! real addresses (`crate::address`).

mod record_pin {
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    enum Representation {
        Canonical,
        Unicode,
        Ansi,
    }
    use crate::safe_array::{SafeArray, SafeArrayBound};
    use crate::{
        Variant as Value,
        bstr::BStr,
        vba_record::{VbaRecord, VbaRecordFieldKind as Field, VbaRecordLayout},
    };
    use calcvba_automation::oleaut::{
        safe_array_create_ex as SafeArrayCreateEx, safe_array_get_dim as SafeArrayGetDim,
        safe_array_get_elemsize as SafeArrayGetElemsize,
        safe_array_get_lbound as SafeArrayGetLBound, safe_array_get_ubound as SafeArrayGetUBound,
        safe_array_ptr_of_index as SafeArrayPtrOfIndex,
        safe_array_put_element as SafeArrayPutElement,
        sys_alloc_string_byte_len as SysAllocStringByteLen, sys_free_string as SysFreeString,
        sys_string_byte_len as SysStringByteLen,
    };
    use std::{ffi::c_void, ptr, sync::Arc};
    #[cfg(windows)]
    use windows_sys::Win32::Globalization::{CP_ACP, MultiByteToWideChar, WideCharToMultiByte};

    /// An Automation `VARIANT` cell (the runtime cell has its layout).
    #[cfg(windows)]
    #[allow(clippy::upper_case_acronyms)]
    pub(super) type VARIANT = windows_sys::Win32::System::Variant::VARIANT;
    #[cfg(not(windows))]
    #[allow(clippy::upper_case_acronyms)]
    pub(super) type VARIANT = crate::VariantCore;

    /// `VariantCopyInd` over the platform cell.
    unsafe fn variant_copy_ind(dst: *mut VARIANT, source: *const VARIANT) -> i32 {
        // SAFETY: forwarded from the caller.
        unsafe { calcvba_automation::oleaut::variant_copy_ind(dst.cast(), source.cast()) }
    }

    /// `VariantClear` over the platform cell.
    unsafe fn variant_clear(cell: *mut VARIANT) -> i32 {
        // SAFETY: forwarded from the caller.
        unsafe { calcvba_automation::oleaut::variant_clear(cell.cast()) }
    }

    /// The `vt` of a cell.
    fn vt_of(cell: &VARIANT) -> u16 {
        // SAFETY: every VARIANT starts with its two-byte vt.
        unsafe { (&raw const *cell).cast::<u16>().read() }
    }

    /// Sets the `vt` of a cell (its payload stays as it is).
    fn set_vt(cell: &mut VARIANT, vt: u16) {
        // SAFETY: as above; writing the vt alone leaves a valid cell for a POD payload.
        unsafe { (&raw mut *cell).cast::<u16>().write(vt) }
    }

    fn checked(hr: i32) -> Result<(), String> {
        if hr < 0 {
            Err(format!("Automation carrier operation failed: {hr:#010x}"))
        } else {
            Ok(())
        }
    }

    pub(super) struct NativeValue(pub(super) VARIANT);
    impl NativeValue {
        fn empty() -> Self {
            // SAFETY: all-zero VARIANT is VT_EMPTY.
            Self(unsafe { std::mem::zeroed() })
        }
        /// A copy of a runtime cell as an Automation `VARIANT` (`VariantCopyInd`, references
        /// read through): the cell is the platform's own layout, so oleaut32 copies it, and the
        /// copy is oleaut32's to free (`VariantClear` in `Drop`, or the record's owner). The
        /// runtime's live counters see neither, as they see no other foreign allocation.
        /// Objects and records are refused, as before: a record field cannot carry their binding.
        pub(super) fn from_value(value: &Value) -> Result<Self, String> {
            if matches!(value.vtype(), crate::VarType::ProcRef) {
                return Err("procedure references cannot be marshaled as VARIANT values".into());
            }
            let mut copy = Self::empty();
            // SAFETY: `value` is a live runtime cell laid out as a VARIANT; `copy` is empty.
            checked(unsafe {
                variant_copy_ind(&mut copy.0, (&raw const *value).cast::<VARIANT>())
            })?;
            let vt = vt_of(&copy.0);
            if matches!(vt & 0x0fff, 9 | 13 | 36) {
                // `copy` is cleared on return.
                return Err(
                    "Record pointer Variant payload requires host object/record binding".into(),
                );
            }
            Ok(copy)
        }
    }
    impl Drop for NativeValue {
        fn drop(&mut self) {
            // SAFETY: this guard owns every native payload in its VARIANT.
            unsafe {
                variant_clear(&mut self.0);
            }
        }
    }

    /// A runtime copy of a borrowed `VARIANT` (`VariantCopyInd`), read in place so its pointers
    /// keep their provenance. Objects and records are refused, as `from_value` refuses them.
    pub(super) unsafe fn value_from_native(input: *const VARIANT) -> Result<Value, String> {
        // SAFETY: the caller's live VARIANT, laid out as the runtime cell.
        let copy = unsafe { Value::from_trusted_cell(&input.cast::<crate::VariantCore>().read()) }?;
        let copy = if copy.vtype() == crate::VarType::ByRef {
            copy.try_clone()?
        } else {
            copy
        };
        if matches!(
            copy.vtype(),
            crate::VarType::Object | crate::VarType::Record
        ) {
            return Err(
                "Record pointer Variant copyback requires host object/record binding".into(),
            );
        }
        if copy.vtype() == crate::VarType::ArrayVariant
            && matches!(copy.core().raw_vt() & 0x0fff, 9 | 13 | 36)
        {
            return Err(
                "Record pointer array copyback requires host object/record metadata".into(),
            );
        }
        Ok(copy)
    }

    fn shape(kind: &Field, representation: Representation) -> Result<(usize, usize), String> {
        match representation {
            Representation::Canonical | Representation::Unicode => kind.storage_shape(),
            Representation::Ansi => kind.declare_storage_shape(),
        }
    }
    fn offset(
        layout: &VbaRecordLayout,
        index: usize,
        representation: Representation,
    ) -> Result<usize, String> {
        match representation {
            Representation::Canonical | Representation::Unicode => {
                Ok(layout.fields()[index].offset())
            }
            Representation::Ansi => layout.declare_field_offset(index),
        }
    }
    #[cfg(not(windows))]
    fn ansi_bytes(units: &[u16]) -> Result<Vec<u8>, String> {
        Ok(crate::ansi::ansi_encode_units(units))
    }
    #[cfg(not(windows))]
    fn wide_units(bytes: &[u8]) -> Result<Vec<u16>, String> {
        Ok(crate::ansi::ansi_decode_units(bytes))
    }
    #[cfg(windows)]
    fn ansi_bytes(units: &[u16]) -> Result<Vec<u8>, String> {
        if units.is_empty() {
            return Ok(Vec::new());
        }
        let len = i32::try_from(units.len()).map_err(|_| "String extent overflow")?;
        // SAFETY: both API calls use the exact retained source and output extents.
        unsafe {
            let count = WideCharToMultiByte(
                CP_ACP,
                0,
                units.as_ptr(),
                len,
                ptr::null_mut(),
                0,
                ptr::null(),
                ptr::null_mut(),
            );
            if count == 0 {
                return Err("ANSI conversion failed".into());
            }
            let mut result = vec![0; count as usize];
            if WideCharToMultiByte(
                CP_ACP,
                0,
                units.as_ptr(),
                len,
                result.as_mut_ptr(),
                count,
                ptr::null(),
                ptr::null_mut(),
            ) != count
            {
                return Err("ANSI conversion failed".into());
            }
            Ok(result)
        }
    }
    #[cfg(windows)]
    fn wide_units(bytes: &[u8]) -> Result<Vec<u16>, String> {
        if bytes.is_empty() {
            return Ok(Vec::new());
        }
        let len = i32::try_from(bytes.len()).map_err(|_| "String extent overflow")?;
        // SAFETY: the bounded byte BSTR/field may include embedded NULs.
        unsafe {
            let count = MultiByteToWideChar(CP_ACP, 0, bytes.as_ptr(), len, ptr::null_mut(), 0);
            if count == 0 {
                return Err("ANSI decoding failed".into());
            }
            let mut result = vec![0; count as usize];
            if MultiByteToWideChar(CP_ACP, 0, bytes.as_ptr(), len, result.as_mut_ptr(), count)
                != count
            {
                return Err("ANSI decoding failed".into());
            }
            Ok(result)
        }
    }
    unsafe fn project_field(
        kind: &Field,
        src: *const u8,
        dst: *mut u8,
        from: Representation,
        to: Representation,
    ) -> Result<(), String> {
        // SAFETY: callers provide distinct initialized fields of the exact
        // representations; prepare each replacement before releasing its owner.
        unsafe {
            match kind {
                Field::Variant => {
                    if to == Representation::Canonical {
                        *dst.cast::<Value>() = value_from_native(src.cast())?;
                    } else {
                        let mut value = if from == Representation::Canonical {
                            NativeValue::from_value(&*src.cast::<Value>())?
                        } else {
                            let mut value = NativeValue::empty();
                            checked(variant_copy_ind(&mut value.0, src.cast()))?;
                            value
                        };
                        checked(variant_clear(dst.cast()))?;
                        dst.cast::<VARIANT>().write(value.0);
                        set_vt(&mut value.0, 0);
                    }
                }
                Field::String => {
                    let raw = src.cast::<*mut u16>().read();
                    let units = if from == Representation::Ansi {
                        let len = SysStringByteLen(raw) as usize;
                        wide_units(if len == 0 {
                            &[]
                        } else {
                            std::slice::from_raw_parts(raw.cast(), len)
                        })?
                    } else {
                        BStr::clone_from_raw_bstr(raw)?.to_utf16_units()
                    };
                    let value = if to == Representation::Canonical {
                        BStr::from_utf16_units(&units)?.into_raw_bstr()
                    } else {
                        let bytes;
                        let (pointer, len) = if to == Representation::Ansi {
                            bytes = ansi_bytes(&units)?;
                            (bytes.as_ptr(), bytes.len())
                        } else {
                            (units.as_ptr().cast(), units.len() * 2)
                        };
                        let value = SysAllocStringByteLen(
                            pointer.cast(),
                            u32::try_from(len).map_err(|_| "BSTR extent overflow")?,
                        );
                        if len != 0 && value.is_null() {
                            return Err("Native record BSTR allocation failed".into());
                        }
                        value
                    };
                    let old = dst.cast::<*mut u16>().replace(value);
                    if to == Representation::Canonical {
                        if !old.is_null() {
                            drop(BStr::from_raw_bstr(old));
                        }
                    } else {
                        SysFreeString(old);
                    }
                }
                Field::FixedString { len } => {
                    let units = if from == Representation::Ansi {
                        wide_units(std::slice::from_raw_parts(src, *len))?
                    } else {
                        (0..*len)
                            .map(|i| src.add(i * 2).cast::<u16>().read_unaligned())
                            .collect()
                    };
                    if to == Representation::Ansi {
                        let bytes = ansi_bytes(&units)?;
                        for i in 0..*len {
                            dst.add(i).write(bytes.get(i).copied().unwrap_or(32));
                        }
                    } else {
                        for i in 0..*len {
                            dst.add(i * 2)
                                .cast::<u16>()
                                .write_unaligned(units.get(i).copied().unwrap_or(32));
                        }
                    }
                }
                Field::Record(layout) => project_record(layout, src, dst, from, to)?,
                Field::FixedArray { element, bounds } => {
                    let count = bounds
                        .iter()
                        .try_fold(1usize, |n, b| n.checked_mul(b.count as usize))
                        .ok_or("Record array extent overflow")?;
                    let (source, align) = shape(element, from)?;
                    let source = source.next_multiple_of(align);
                    let (target, align) = shape(element, to)?;
                    let target = target.next_multiple_of(align);
                    // Every representation stores the first dimension fastest, as VBA does.
                    for i in 0..count {
                        project_field(element, src.add(i * source), dst.add(i * target), from, to)?;
                    }
                }
                // The projection borrows object and array pointers: it holds no reference of
                // its own and never releases them. Direct record passing (dc-yqq0.7) removes it.
                Field::Object | Field::DynamicArray { .. } => {
                    ptr::copy_nonoverlapping(src, dst, core::mem::size_of::<*const ()>())
                }
                _ => ptr::copy_nonoverlapping(src, dst, kind.storage_shape()?.0),
            }
        }
        Ok(())
    }
    unsafe fn project_record(
        layout: &VbaRecordLayout,
        src: *const u8,
        dst: *mut u8,
        from: Representation,
        to: Representation,
    ) -> Result<(), String> {
        for (index, field) in layout.fields().iter().enumerate() {
            // SAFETY: recursive shape checks bound each source/destination field.
            unsafe {
                project_field(
                    field.kind(),
                    src.add(offset(layout, index, from)?),
                    dst.add(offset(layout, index, to)?),
                    from,
                    to,
                )?;
            }
        }
        Ok(())
    }
    fn owners(
        kind: &Field,
        base: usize,
        representation: Representation,
        result: &mut Vec<(usize, u32)>,
    ) -> Result<(), String> {
        match kind {
            Field::String => result.push((base, 8)),
            Field::Variant => result.push((base, 12)),
            Field::Record(layout) => {
                for (i, field) in layout.fields().iter().enumerate() {
                    owners(
                        field.kind(),
                        base + offset(layout, i, representation)?,
                        representation,
                        result,
                    )?;
                }
            }
            Field::FixedArray { element, bounds } => {
                let count = bounds
                    .iter()
                    .try_fold(1usize, |n, b| n.checked_mul(b.count as usize))
                    .ok_or("Record array extent overflow")?;
                let (width, align) = shape(element, representation)?;
                for i in 0..count {
                    owners(
                        element,
                        base + i * width.next_multiple_of(align),
                        representation,
                        result,
                    )?;
                }
            }
            _ => {}
        }
        Ok(())
    }

    #[derive(Debug)]
    pub struct AutomationRecord {
        pub(super) words: Vec<u64>,
        layout: Arc<VbaRecordLayout>,
        owners: Vec<(usize, u32)>,
        representation: Representation,
    }
    impl AutomationRecord {
        fn empty(
            layout: Arc<VbaRecordLayout>,
            representation: Representation,
        ) -> Result<Self, String> {
            let size = if representation == Representation::Ansi {
                layout.declare_shape()?.0
            } else {
                layout.size()
            };
            let mut fields = Vec::new();
            for (i, field) in layout.fields().iter().enumerate() {
                owners(
                    field.kind(),
                    offset(&layout, i, representation)?,
                    representation,
                    &mut fields,
                )?;
            }
            Ok(Self {
                words: vec![0; size.div_ceil(8)],
                layout,
                owners: fields,
                representation,
            })
        }
        pub fn from_record(record: &VbaRecord) -> Result<Self, String> {
            let mut result = Self::empty(record.layout().clone(), Representation::Unicode)?;
            // SAFETY: zeroed Automation storage and a retained canonical source.
            unsafe {
                project_record(
                    &result.layout,
                    record.data_ptr(),
                    result.words.as_mut_ptr().cast(),
                    Representation::Canonical,
                    Representation::Unicode,
                )?;
            }
            Ok(result)
        }
        pub fn from_declared_record(record: &VbaRecord) -> Result<Self, String> {
            let mut result = Self::empty(record.layout().clone(), Representation::Ansi)?;
            // SAFETY: zeroed ANSI record storage and a retained canonical source.
            unsafe {
                project_record(
                    &result.layout,
                    record.data_ptr(),
                    result.words.as_mut_ptr().cast(),
                    Representation::Canonical,
                    Representation::Ansi,
                )?;
            }
            Ok(result)
        }
        /// # Safety
        /// source is initialized Unicode Automation record storage of layout.
        pub unsafe fn from_automation(
            source: *const u8,
            layout: Arc<VbaRecordLayout>,
        ) -> Result<Self, String> {
            let mut result = Self::empty(layout, Representation::Ansi)?;
            // SAFETY: prepare an independent ANSI owner before the native call.
            unsafe {
                project_record(
                    &result.layout,
                    source,
                    result.words.as_mut_ptr().cast(),
                    Representation::Unicode,
                    Representation::Ansi,
                )?;
            }
            Ok(result)
        }
        pub fn empty_declared(layout: Arc<VbaRecordLayout>) -> Result<Self, String> {
            Self::empty(layout, Representation::Ansi)
        }
        pub fn as_mut_ptr(&mut self) -> *mut u8 {
            self.words.as_mut_ptr().cast()
        }
        pub fn read(&self) -> Result<Value, String> {
            let mut record = VbaRecord::new_default(self.layout.clone())?;
            // SAFETY: source owners are retained until canonical copyback ends.
            unsafe {
                project_record(
                    &self.layout,
                    self.words.as_ptr().cast(),
                    record.data_mut_ptr(),
                    self.representation,
                    Representation::Canonical,
                )?;
            }
            Ok(Value::from_vba_record(record))
        }
        /// # Safety
        /// destination is initialized same-layout Unicode Automation storage.
        pub unsafe fn copy_back_automation(&self, destination: *mut u8) -> Result<(), String> {
            // SAFETY: caller retains the destination through conversion/copyback.
            unsafe {
                project_record(
                    &self.layout,
                    self.words.as_ptr().cast(),
                    destination,
                    self.representation,
                    Representation::Unicode,
                )
            }
        }
    }
    impl Drop for AutomationRecord {
        fn drop(&mut self) {
            for (offset, kind) in &self.owners {
                // SAFETY: each owner is initialized at its exact native offset.
                unsafe {
                    let cell = self.words.as_mut_ptr().cast::<u8>().add(*offset);
                    if *kind == 8 {
                        SysFreeString(cell.cast::<*mut u16>().read());
                    } else {
                        variant_clear(cell.cast());
                    }
                }
            }
        }
    }
    /// Address of the `parray` pointer slot at `PAYLOAD_OFFSET` (8) inside a platform `VARIANT`.
    fn cell_parray_slot(cell: &mut VARIANT) -> *mut *mut c_void {
        // SAFETY: every platform VARIANT stores its pointer union (`parray`) at `PAYLOAD_OFFSET`.
        unsafe {
            (&raw mut *cell)
                .cast::<u8>()
                .add(calcvba_automation::abi::PAYLOAD_OFFSET)
                .cast::<*mut c_void>()
        }
    }

    /// Reads the `parray` pointer from a platform `VARIANT`.
    fn read_cell_parray(cell: &VARIANT) -> *mut c_void {
        // SAFETY: every platform VARIANT stores its pointer union (`parray`) at `PAYLOAD_OFFSET`.
        unsafe {
            (&raw const *cell)
                .cast::<u8>()
                .add(calcvba_automation::abi::PAYLOAD_OFFSET)
                .cast::<*mut c_void>()
                .read()
        }
    }

    pub struct AutomationRecordArray {
        cell: Box<VARIANT>,
        layout: Arc<VbaRecordLayout>,
    }
    impl AutomationRecordArray {
        /// The native form of a UDT array value: `array` is `None` while unallocated.
        pub fn from_array(
            layout: Arc<VbaRecordLayout>,
            array: Option<&SafeArray>,
        ) -> Result<Self, String> {
            let mut result = Self {
                cell: Box::new(NativeValue::empty().0),
                layout,
            };
            // SAFETY: the private VARIANT exclusively owns the new descriptor.
            unsafe {
                set_vt(&mut result.cell, 0x2024);
                if let Some(array) = array
                    && let Some(bounds) = array.bounds()
                {
                    let mut fields = Vec::new();
                    for (i, field) in result.layout.fields().iter().enumerate() {
                        owners(
                            field.kind(),
                            result.layout.fields()[i].offset(),
                            Representation::Unicode,
                            &mut fields,
                        )?;
                    }
                    let mut descriptor = vec![result.layout.size() as u32, fields.len() as u32];
                    for (offset, kind) in fields {
                        descriptor.extend([offset as u32, kind]);
                    }
                    let mut hash = 0xcbf29ce484222325u64;
                    for word in &descriptor {
                        for byte in word.to_le_bytes() {
                            hash = (hash ^ u64::from(byte)).wrapping_mul(0x100000001b3);
                        }
                    }
                    descriptor.extend([hash as u32, (hash >> 32) as u32, 0x45564143, 0x5245434f]);
                    let info = crate::automation_record::create_record_info(descriptor.as_ptr());
                    let raw = SafeArrayCreateEx(36, bounds.len() as u32, bounds.as_ptr(), info);
                    crate::automation_record::release_record_info(info);
                    if raw.is_null() {
                        return Err("Native record array allocation failed".into());
                    }
                    cell_parray_slot(&mut result.cell).write(raw);
                    let values = array.try_variant_elements()?.unwrap_or_default();
                    let mut indices: Vec<_> = bounds.iter().map(|b| b.lower).collect();
                    for (i, value) in values.iter().enumerate() {
                        let record = value
                            .as_vba_record()
                            .ok_or("Record array element missing")?;
                        let mut native = AutomationRecord::from_record(&record)?;
                        checked(SafeArrayPutElement(
                            raw,
                            indices.as_ptr(),
                            native.as_mut_ptr().cast(),
                        ))?;
                        if i + 1 != values.len() {
                            super::advance_safearray_indices(&mut indices, &bounds)?;
                        }
                    }
                }
            }
            Ok(result)
        }
        pub fn as_cell_ptr(&mut self) -> *mut c_void {
            cell_parray_slot(&mut self.cell).cast()
        }
        pub fn read(&self) -> Result<Value, String> {
            // SAFETY: the callee returned a nullable live same-layout SAFEARRAY.
            unsafe {
                let array = read_cell_parray(&self.cell);
                if array.is_null() {
                    return Ok(Value::unallocated_record_array(self.layout.clone()));
                }
                if SafeArrayGetElemsize(array) as usize != self.layout.size() {
                    return Err("Record array native element size changed".into());
                }
                let rank = SafeArrayGetDim(array);
                let mut count = 1usize;
                let mut bounds = Vec::new();
                for dimension in 1..=rank {
                    let mut lower = 0;
                    let mut upper = 0;
                    checked(SafeArrayGetLBound(array, dimension, &mut lower))?;
                    checked(SafeArrayGetUBound(array, dimension, &mut upper))?;
                    let length = u32::try_from(i64::from(upper) - i64::from(lower) + 1)
                        .map_err(|_| "Invalid record array bounds")?;
                    count = count
                        .checked_mul(length as usize)
                        .ok_or("Record array extent overflow")?;
                    bounds.push(SafeArrayBound {
                        lower,
                        count: length,
                    });
                }
                let mut indices: Vec<_> = bounds.iter().map(|b| b.lower).collect();
                let mut records = Vec::with_capacity(count);
                for i in 0..count {
                    let mut pointer = ptr::null_mut();
                    checked(SafeArrayPtrOfIndex(array, indices.as_ptr(), &mut pointer))?;
                    let mut record = VbaRecord::new_default(self.layout.clone())?;
                    project_record(
                        &self.layout,
                        pointer.cast(),
                        record.data_mut_ptr(),
                        Representation::Unicode,
                        Representation::Canonical,
                    )?;
                    records.push(record);
                    if i + 1 != count {
                        super::advance_safearray_indices(&mut indices, &bounds)?;
                    }
                }
                Ok(Value::from_safearray(SafeArray::from_vba_records_nd(
                    bounds,
                    self.layout.clone(),
                    records,
                )?))
            }
        }
    }
    impl Drop for AutomationRecordArray {
        fn drop(&mut self) {
            // SAFETY: consume this cell's one retained array/record-info owner.
            unsafe {
                variant_clear(&mut *self.cell);
            }
        }
    }
}
pub use record_pin::{AutomationRecord, AutomationRecordArray};

fn advance_safearray_indices(
    indices: &mut [i32],
    bounds: &[crate::safe_array::SafeArrayBound],
) -> Result<(), String> {
    if crate::safe_array::advance_subscripts(indices, bounds) {
        return Ok(());
    }
    if let Some(dimension) = bounds
        .iter()
        .position(|b| i64::from(b.lower) + i64::from(b.count) - 1 > i64::from(i32::MAX))
    {
        return Err(format!(
            "SAFEARRAY bound at dimension {dimension} exceeds the Windows LONG index range"
        ));
    }
    Err("SAFEARRAY multidimensional index progression exhausted early".to_string())
}

#[cfg(test)]
mod tests {
    use super::{AutomationRecord, AutomationRecordArray};
    use crate::{
        Variant,
        bstr::BStr,
        current_thread_live_handle_counts,
        safe_array::{SafeArray, SafeArrayBound},
        vba_record::{VbaRecord, VbaRecordFieldKind, VbaRecordFieldSpec, VbaRecordLayout},
    };
    use std::sync::Arc;

    /// The Automation projection of a record's Variant fields copies them through oleaut32 and
    /// frees them there: the runtime's BSTR and SAFEARRAY counters stay balanced.
    #[test]
    fn variant_fields_projected_for_a_declare_balance_the_live_counters() {
        let layout = Arc::new(
            VbaRecordLayout::new(vec![
                VbaRecordFieldSpec::named("Text", VbaRecordFieldKind::Variant),
                VbaRecordFieldSpec::named("Items", VbaRecordFieldKind::Variant),
            ])
            .expect("record layout"),
        );
        let mut record = VbaRecord::new_default(layout).expect("record");
        record
            .write_field_variant(0, &Variant::from_string(BStr::from("payload")))
            .expect("String field");
        let items = SafeArray::from_variants(vec![
            Variant::from_string(BStr::from("a")),
            Variant::from_string(BStr::from("b")),
        ]);
        record
            .write_field_variant(1, &Variant::from_safearray(items))
            .expect("array field");
        let before = current_thread_live_handle_counts();
        for _ in 0..3 {
            let unicode = AutomationRecord::from_record(&record).expect("Unicode projection");
            let back = unicode.read().expect("read back");
            assert_eq!(
                back.read_record_field_variant(0)
                    .expect("field")
                    .as_bstr()
                    .map(|text| text.to_string()),
                Some("payload".into())
            );
            drop(back);
            drop(unicode);
            drop(AutomationRecord::from_declared_record(&record).expect("ANSI projection"));
        }
        assert_eq!(
            before.balance_to(current_thread_live_handle_counts()),
            Default::default()
        );
    }

    #[test]
    fn record_array_projected_for_a_declare_handles_nd_bounds_strings_variants_and_balances_counters()
     {
        let layout = Arc::new(
            VbaRecordLayout::new(vec![
                VbaRecordFieldSpec::named("Id", VbaRecordFieldKind::Long),
                VbaRecordFieldSpec::named("Tag", VbaRecordFieldKind::FixedString { len: 4 }),
                VbaRecordFieldSpec::named("Name", VbaRecordFieldKind::String),
                VbaRecordFieldSpec::named("Payload", VbaRecordFieldKind::Variant),
            ])
            .expect("record layout"),
        );
        let bounds = vec![
            SafeArrayBound { count: 2, lower: 1 },
            SafeArrayBound {
                count: 2,
                lower: 10,
            },
        ];
        let mut records = Vec::new();
        for n in 0..4i32 {
            let mut rec = VbaRecord::new_default(Arc::clone(&layout)).expect("record");
            rec.write_field_variant(0, &Variant::from_i32(100 + n))
                .expect("Id");
            rec.write_field_variant(1, &Variant::from_string(format!("T{n:03}")))
                .expect("Tag");
            rec.write_field_variant(2, &Variant::from_string(format!("item-{n}")))
                .expect("Name");
            rec.write_field_variant(3, &Variant::from_string(format!("v-{n}")))
                .expect("Payload");
            records.push(rec);
        }
        let array = SafeArray::from_vba_records_nd(bounds.clone(), Arc::clone(&layout), records)
            .expect("record safearray");
        let before = current_thread_live_handle_counts();
        {
            // Unallocated record array round-trips as unallocated.
            let empty =
                AutomationRecordArray::from_array(Arc::clone(&layout), None).expect("empty array");
            let read_empty = empty.read().expect("read empty");
            assert_eq!(read_empty.record_array_layout(), Some(Arc::clone(&layout)));
            assert!(read_empty.try_as_safearray().unwrap().is_none());
            drop(read_empty);
            drop(empty);

            // Allocated 2D record array: verify native descriptor, mutate element (2, 11) in place, read back.
            let mut projected =
                AutomationRecordArray::from_array(Arc::clone(&layout), Some(&array))
                    .expect("projected record array");
            // SAFETY: `as_cell_ptr` points to the live `SAFEARRAY*` slot owned by `projected`.
            unsafe {
                let psa = *projected.as_cell_ptr().cast::<*mut core::ffi::c_void>();
                assert!(!psa.is_null());
                assert_eq!(calcvba_automation::oleaut::safe_array_get_dim(psa), 2);
                assert_eq!(
                    calcvba_automation::oleaut::safe_array_get_elemsize(psa) as usize,
                    layout.size()
                );
                let mut elem = core::ptr::null_mut();
                let idx = [2i32, 11i32]; // flat index 3
                assert_eq!(
                    calcvba_automation::oleaut::safe_array_ptr_of_index(
                        psa,
                        idx.as_ptr(),
                        &mut elem
                    ),
                    0
                );
                assert_eq!(elem.cast::<i32>().read(), 103);
                elem.cast::<i32>().write(777);
            }
            let back = projected.read().expect("read projected array");
            let back_sa = back
                .try_as_safearray()
                .expect("safearray")
                .expect("allocated");
            assert_eq!(back_sa.bounds(), Some(bounds));
            let elems = back_sa.variant_elements().expect("elements");
            assert_eq!(elems.len(), 4);
            let rec3 = elems[3].as_vba_record().expect("record 3");
            assert_eq!(rec3.read_field_variant(0).unwrap().as_i32(), Some(777));
            assert_eq!(
                rec3.read_field_variant(1)
                    .unwrap()
                    .as_bstr()
                    .unwrap()
                    .as_str(),
                "T003"
            );
            assert_eq!(
                rec3.read_field_variant(2)
                    .unwrap()
                    .as_bstr()
                    .unwrap()
                    .as_str(),
                "item-3"
            );
            assert_eq!(
                rec3.read_field_variant(3)
                    .unwrap()
                    .as_bstr()
                    .unwrap()
                    .as_str(),
                "v-3"
            );
            drop(elems);
            drop(back_sa);
            drop(back);
            drop(projected);
        }
        assert_eq!(
            before.balance_to(current_thread_live_handle_counts()),
            Default::default()
        );
    }
}
