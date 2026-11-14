//! Automation-layout projections of VBA records and record arrays for `Declare` calls
//! (Windows). `VarPtr`, `StrPtr` and `ObjPtr` give real addresses (`crate::address`).

#[cfg(windows)]
mod native_record_pin {
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    enum Representation {
        Canonical,
        Unicode,
        Ansi,
    }
    use crate::{
        Variant as Value,
        bstr::BStr,
        safe_array::SafeArray,
        vba_record::{VbaRecord, VbaRecordFieldKind as Field, VbaRecordLayout},
    };
    use std::{ptr, sync::Arc};
    use windows_sys::Win32::{
        Foundation::{SysAllocStringByteLen, SysFreeString, SysStringByteLen},
        Globalization::{CP_ACP, MultiByteToWideChar, WideCharToMultiByte},
        System::{Com::SAFEARRAYBOUND, Ole::*, Variant::*},
    };

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
        pub(super) fn from_value(value: &Value) -> Result<Self, String> {
            let mut result = Self::empty();
            // SAFETY: result owns fresh VT_EMPTY storage; each branch initializes
            // a complete matching payload before exposing the type tag.
            unsafe {
                match value.vtype() {
                    crate::VarType::ProcRef => {
                        return Err(
                            "procedure references cannot be marshaled as VARIANT values".into()
                        );
                    }
                    crate::VarType::Error => {
                        let code = value.as_error_code().ok_or("Error payload missing")?;
                        result.0.Anonymous.Anonymous.vt = 10;
                        result.0.Anonymous.Anonymous.Anonymous.scode =
                            if (0..=65535).contains(&code) {
                                (0x800a_0000u32 | code as u32) as i32
                            } else {
                                code
                            };
                    }
                    crate::VarType::String => {
                        let text = value.try_as_bstr()?.ok_or("String payload missing")?;
                        let bytes = text
                            .to_utf16_units()
                            .len()
                            .checked_mul(2)
                            .ok_or("BSTR extent overflow")?;
                        let units = text.to_utf16_units();
                        let raw = SysAllocStringByteLen(
                            units.as_ptr().cast(),
                            u32::try_from(bytes).map_err(|_| "BSTR too large")?,
                        )
                        .cast_mut();
                        if bytes != 0 && raw.is_null() {
                            return Err("BSTR allocation failed".into());
                        }
                        result.0.Anonymous.Anonymous.vt = 8;
                        result.0.Anonymous.Anonymous.Anonymous.bstrVal = raw;
                    }
                    crate::VarType::ArrayVariant => {
                        let array = value.try_as_safearray()?.ok_or("Array payload missing")?;
                        let kind = array.element_vartype();
                        if matches!(kind, 9 | 13 | 36) {
                            return Err(
                                "Record pointer Variant array requires host object/record metadata"
                                    .into(),
                            );
                        }
                        result.0.Anonymous.Anonymous.vt = 0x2000 | kind;
                        if let Some(bounds) = array.bounds() {
                            let native_bounds: Vec<_> = bounds
                                .iter()
                                .map(|b| SAFEARRAYBOUND {
                                    cElements: b.count,
                                    lLbound: b.lower,
                                })
                                .collect();
                            let raw = SafeArrayCreate(
                                kind,
                                native_bounds.len() as u32,
                                native_bounds.as_ptr(),
                            );
                            if raw.is_null() {
                                return Err("SAFEARRAY allocation failed".into());
                            }
                            result.0.Anonymous.Anonymous.Anonymous.parray = raw;
                            let values = array.try_variant_elements()?.unwrap_or_default();
                            let mut indices: Vec<_> = bounds.iter().map(|b| b.lower).collect();
                            for (i, value) in values.iter().enumerate() {
                                let native = Self::from_value(value)?;
                                let data = if kind == 12 {
                                    (&native.0 as *const VARIANT).cast()
                                } else if kind == 8 {
                                    native.0.Anonymous.Anonymous.Anonymous.bstrVal.cast()
                                } else {
                                    (&native.0 as *const VARIANT).cast::<u8>().add(8).cast()
                                };
                                checked(SafeArrayPutElement(raw, indices.as_ptr(), data))?;
                                if i + 1 != values.len() {
                                    super::advance_windows_safearray_indices(
                                        &mut indices,
                                        &bounds,
                                    )?;
                                }
                            }
                        }
                    }
                    crate::VarType::Object | crate::VarType::Record => {
                        return Err(
                            "Record pointer Variant payload requires host object/record binding"
                                .into(),
                        );
                    }
                    _ => {
                        let bytes = value.to_wire_bytes();
                        ptr::copy_nonoverlapping(
                            bytes.as_ptr(),
                            (&mut result.0 as *mut VARIANT).cast(),
                            bytes.len(),
                        );
                    }
                }
            }
            Ok(result)
        }
    }
    impl Drop for NativeValue {
        fn drop(&mut self) {
            // SAFETY: this guard owns every native payload in its VARIANT.
            unsafe {
                VariantClear(&mut self.0);
            }
        }
    }

    pub(super) unsafe fn value_from_native(input: *const VARIANT) -> Result<Value, String> {
        let mut value = NativeValue::empty();
        // SAFETY: copy/dereference the borrowed live VARIANT into a private owner.
        unsafe {
            checked(VariantCopyInd(&mut value.0, input))?;
            let kind = value.0.Anonymous.Anonymous.vt;
            if kind == 10 {
                let code = value.0.Anonymous.Anonymous.Anonymous.scode;
                return Ok(Value::from_error_code(
                    if code as u32 & 0xffff_0000 == 0x800a_0000 {
                        code & 0xffff
                    } else {
                        code
                    },
                ));
            }
            if kind == 8 {
                return Ok(Value::from_string(BStr::clone_from_raw_bstr(
                    value.0.Anonymous.Anonymous.Anonymous.bstrVal.cast_mut(),
                )?));
            }
            if kind & 0x2000 != 0 {
                let array = value.0.Anonymous.Anonymous.Anonymous.parray;
                if array.is_null() {
                    return Err("Variant record field contains an unallocated SAFEARRAY".into());
                }
                let element = kind & 0xfff;
                if matches!(element, 9 | 13 | 36) {
                    return Err(
                        "Record pointer array copyback requires host object/record metadata".into(),
                    );
                }
                let rank = SafeArrayGetDim(array);
                let mut bounds = Vec::new();
                let mut count = 1usize;
                for dimension in 1..=rank {
                    let mut lower = 0;
                    let mut upper = 0;
                    checked(SafeArrayGetLBound(array, dimension, &mut lower))?;
                    checked(SafeArrayGetUBound(array, dimension, &mut upper))?;
                    let length = u32::try_from(i64::from(upper) - i64::from(lower) + 1)
                        .map_err(|_| "Invalid SAFEARRAY bounds")?;
                    count = count
                        .checked_mul(length as usize)
                        .ok_or("SAFEARRAY extent overflow")?;
                    bounds.push(crate::safe_array::SafeArrayBound {
                        lower,
                        count: length,
                    });
                }
                let mut indices: Vec<_> = bounds.iter().map(|b| b.lower).collect();
                let mut values = Vec::with_capacity(count);
                for index in 0..count {
                    let mut item = NativeValue::empty();
                    let data = if element == 12 {
                        (&mut item.0 as *mut VARIANT).cast()
                    } else {
                        item.0.Anonymous.Anonymous.vt = element;
                        (&mut item.0 as *mut VARIANT).cast::<u8>().add(8).cast()
                    };
                    checked(SafeArrayGetElement(array, indices.as_ptr(), data))?;
                    values.push(value_from_native(&item.0)?);
                    if index + 1 != count {
                        super::advance_windows_safearray_indices(&mut indices, &bounds)?;
                    }
                }
                return Ok(Value::from_safearray(SafeArray::from_typed_variants_nd(
                    bounds, element, values,
                )?));
            }
            if matches!(kind, 9 | 13 | 36) {
                return Err(
                    "Record pointer Variant copyback requires host object/record binding".into(),
                );
            }
            let mut bytes = [0u8; crate::variant::WIRE_SIZE];
            ptr::copy_nonoverlapping(
                (&value.0 as *const VARIANT).cast(),
                bytes.as_mut_ptr(),
                bytes.len(),
            );
            Value::from_wire_bytes(bytes)
        }
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
                            checked(VariantCopyInd(&mut value.0, src.cast()))?;
                            value
                        };
                        checked(VariantClear(dst.cast()))?;
                        dst.cast::<VARIANT>().write(value.0);
                        value.0.Anonymous.Anonymous.vt = 0;
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
                        )
                        .cast_mut();
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
                        VariantClear(cell.cast());
                    }
                }
            }
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
                result.cell.Anonymous.Anonymous.vt = 0x2024;
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
                    let native_bounds: Vec<_> = bounds
                        .iter()
                        .map(|b| SAFEARRAYBOUND {
                            cElements: b.count,
                            lLbound: b.lower,
                        })
                        .collect();
                    let raw = SafeArrayCreateEx(
                        36,
                        native_bounds.len() as u32,
                        native_bounds.as_ptr(),
                        info,
                    );
                    crate::automation_record::release_record_info(info);
                    if raw.is_null() {
                        return Err("Native record array allocation failed".into());
                    }
                    result.cell.Anonymous.Anonymous.Anonymous.parray = raw;
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
                            super::advance_windows_safearray_indices(&mut indices, &bounds)?;
                        }
                    }
                }
            }
            Ok(result)
        }
        pub fn as_cell_ptr(&mut self) -> *mut std::ffi::c_void {
            // SAFETY: the boxed VARIANT's array pointer cell has a stable address.
            unsafe {
                (&mut self.cell.Anonymous.Anonymous.Anonymous.parray
                    as *mut *mut windows_sys::Win32::System::Com::SAFEARRAY)
                    .cast()
            }
        }
        pub fn read(&self) -> Result<Value, String> {
            // SAFETY: the callee returned a nullable live same-layout SAFEARRAY.
            unsafe {
                let array = self.cell.Anonymous.Anonymous.Anonymous.parray;
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
                    bounds.push(crate::safe_array::SafeArrayBound {
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
                        super::advance_windows_safearray_indices(&mut indices, &bounds)?;
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
                VariantClear(&mut *self.cell);
            }
        }
    }
}
#[cfg(windows)]
pub use native_record_pin::{AutomationRecord, AutomationRecordArray};
#[cfg(target_os = "windows")]
fn advance_windows_safearray_indices(
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
