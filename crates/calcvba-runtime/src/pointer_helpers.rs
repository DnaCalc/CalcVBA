use std::{
    collections::HashMap,
    ffi::c_void,
    sync::{Mutex, OnceLock},
};

use crate::{Variant, bstr::BStr};

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
                            16,
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
            let mut bytes = [0u8; 16];
            ptr::copy_nonoverlapping((&value.0 as *const VARIANT).cast(), bytes.as_mut_ptr(), 16);
            Value::from_wire_bytes(bytes)
        }
    }

    fn shape(kind: &Field, representation: Representation) -> Result<(usize, usize), String> {
        match representation {
            Representation::Canonical => kind.storage_shape(),
            Representation::Unicode => kind.automation_storage_shape(),
            Representation::Ansi => kind.declare_storage_shape(),
        }
    }
    fn offset(
        layout: &VbaRecordLayout,
        index: usize,
        representation: Representation,
    ) -> Result<usize, String> {
        match representation {
            Representation::Canonical => Ok(layout.fields()[index].offset()),
            Representation::Unicode => layout.automation_field_offset(index),
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
                    for i in 0..count {
                        // Canonical arrays enumerate the last dimension fastest;
                        // Office inline arrays store the first dimension fastest.
                        let mut remainder = i;
                        let mut native = 0;
                        for bound in bounds.iter().rev() {
                            native =
                                native * bound.count as usize + remainder % bound.count as usize;
                            remainder /= bound.count as usize;
                        }
                        let source_index = if from == Representation::Canonical {
                            i
                        } else {
                            native
                        };
                        let target_index = if to == Representation::Canonical {
                            i
                        } else {
                            native
                        };
                        project_field(
                            element,
                            src.add(source_index * source),
                            dst.add(target_index * target),
                            from,
                            to,
                        )?;
                    }
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
            let (size, _) = if representation == Representation::Ansi {
                layout.declare_shape()?
            } else {
                layout.automation_shape()?
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
        pub fn from_array(array: &SafeArray) -> Result<Self, String> {
            let layout = array
                .vba_record_layout()
                .ok_or("Record array layout missing")?;
            let mut result = Self {
                cell: Box::new(NativeValue::empty().0),
                layout,
            };
            // SAFETY: the private VARIANT exclusively owns the new descriptor.
            unsafe {
                result.cell.Anonymous.Anonymous.vt = 0x2024;
                if let Some(bounds) = array.bounds() {
                    let mut fields = Vec::new();
                    for (i, field) in result.layout.fields().iter().enumerate() {
                        owners(
                            field.kind(),
                            result.layout.automation_field_offset(i)?,
                            Representation::Unicode,
                            &mut fields,
                        )?;
                    }
                    let mut descriptor = vec![
                        result.layout.automation_shape()?.0 as u32,
                        fields.len() as u32,
                    ];
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
                    return Ok(Value::from_safearray(SafeArray::from_vba_records_nd(
                        Vec::new(),
                        self.layout.clone(),
                        Vec::new(),
                    )?));
                }
                if SafeArrayGetElemsize(array) as usize != self.layout.automation_shape()?.0 {
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
#[cfg(windows)]
type OwnedNativeRecord = AutomationRecord;

#[cfg(target_os = "windows")]
use windows_sys::{
    Win32::Foundation::SysFreeString,
    Win32::System::Variant::{
        VARIANT, VT_BOOL, VT_BSTR, VT_CY, VT_DATE, VT_EMPTY, VT_ERROR, VT_I4, VT_I8, VT_NULL,
        VT_UI1, VT_UNKNOWN, VariantClear,
    },
    core::BSTR,
};

#[cfg(target_os = "windows")]
const VT_R4: u16 = 4;
#[cfg(target_os = "windows")]
const VT_R8: u16 = 5;
#[cfg(target_os = "windows")]
const VT_I1: u16 = 16;
#[cfg(target_os = "windows")]
const VT_UI2: u16 = 18;
#[cfg(target_os = "windows")]
const VT_UI4: u16 = 19;
#[cfg(target_os = "windows")]
const VT_UI8: u16 = 21;
#[cfg(target_os = "windows")]
const VT_UINT: u16 = 23;

#[cfg(target_os = "windows")]
#[derive(Debug)]
// Pointer helpers expose real Windows BSTR cells by cloning the canonical
// runtime BSTR payload.
struct OwnedBstr(Option<BStr>);

#[cfg(target_os = "windows")]
impl OwnedBstr {
    fn from_bstr(text: &BStr) -> Result<Self, String> {
        if text.raw_bstr().is_null() {
            return Ok(Self(None));
        }
        Ok(Self(Some(BStr::from_bytes(text.payload_bytes())?)))
    }

    fn from_text(text: &str) -> Result<Self, String> {
        Ok(Self(Some(BStr::from(text))))
    }

    fn as_ptr(&self) -> *mut c_void {
        self.raw_bstr().cast_mut().cast()
    }

    fn raw_bstr(&self) -> BSTR {
        self.0
            .as_ref()
            .map(BStr::raw_bstr)
            .unwrap_or(std::ptr::null_mut())
    }

    /// Transfer this canonical BSTR owner into a fresh Windows VARIANT and
    /// attach the independent accounting token before returning.
    ///
    /// # Safety
    /// `variant` must point to a writable, zero/VT_EMPTY-initialized VARIANT
    /// cell not currently owning a payload. `accounting` must belong to that
    /// same cell and contain no prior BSTR transfer token.
    unsafe fn transfer_into_variant(
        mut self,
        variant: *mut VARIANT,
        accounting: &mut Option<TrackedBstrAccountingToken>,
    ) -> Result<(), String> {
        if accounting.is_some() {
            return Err("Windows VARIANT already carries a tracked BSTR transfer".to_string());
        }
        let Some(value) = self.0.take() else {
            // SAFETY: guaranteed by this method's fresh writable-cell contract.
            unsafe {
                (*variant).Anonymous.Anonymous.Anonymous.bstrVal = std::ptr::null_mut();
                (*variant).Anonymous.Anonymous.vt = VT_BSTR;
            }
            return Ok(());
        };

        // No fallible operation follows this ownership transfer. The pointer
        // is installed in the destination before the accounting token is
        // attached; both writes target the caller's valid fresh cell.
        let raw = value.into_raw_bstr();
        // SAFETY: guaranteed by this method's fresh writable-cell contract.
        unsafe {
            (*variant).Anonymous.Anonymous.Anonymous.bstrVal = raw;
            (*variant).Anonymous.Anonymous.vt = VT_BSTR;
        }
        *accounting = Some(TrackedBstrAccountingToken);
        Ok(())
    }
}

/// Counter ownership for one canonical BSTR allocation transferred into a
/// Windows VARIANT. Windows owns and frees the native payload after transfer;
/// this token owns only the matching OxVba live-counter debit.
#[cfg(target_os = "windows")]
#[derive(Debug)]
struct TrackedBstrAccountingToken;

#[cfg(target_os = "windows")]
impl Drop for TrackedBstrAccountingToken {
    fn drop(&mut self) {
        crate::live_counters::bstr_freed();
    }
}

#[cfg(target_os = "windows")]
#[derive(Debug)]
// `VarPtr(String)` exposes a pointer to a BSTR cell. The helper therefore owns
// the cell itself and whichever BSTR pointer a native call leaves in that cell.
struct OwnedBstrCell {
    cell: Box<BSTR>,
    original: Option<BStr>,
}

#[cfg(target_os = "windows")]
impl OwnedBstrCell {
    fn from_bstr(text: &BStr) -> Result<Self, String> {
        let mut owner = OwnedBstr::from_bstr(text)?;
        let cell = Box::new(owner.raw_bstr());
        Ok(Self {
            cell,
            original: owner.0.take(),
        })
    }

    fn as_ptr(&mut self) -> *mut c_void {
        (&mut *self.cell as *mut BSTR).cast()
    }
}

#[cfg(target_os = "windows")]
impl Drop for OwnedBstrCell {
    fn drop(&mut self) {
        let current = std::mem::replace(&mut *self.cell, std::ptr::null_mut());
        let Some(original) = self.original.take() else {
            if !current.is_null() {
                // SAFETY: A non-null cell with no original OxVba owner can only
                // contain a native replacement whose ownership transferred to
                // this cell. The consuming call has completed before pin drop.
                unsafe { SysFreeString(current) };
            }
            return;
        };
        let original_raw = original.raw_bstr();
        if current == original_raw {
            // The cell still contains its original canonical owner. Dropping
            // through `BStr` performs both the OS free and live-counter debit.
            drop(original);
            return;
        }

        // The native LPBSTR write consumed/freed the original before replacing
        // the cell. Reconcile that original tracked allocation without freeing
        // it twice. The replacement was allocated by native code, so it must be
        // freed below without debiting the OxVba BSTR counter.
        crate::live_counters::bstr_freed();
        // Native code already consumed/freed this allocation. Taking its raw
        // address only disarms the stale Rust owner; the dangling value must
        // never be dereferenced or freed again.
        let _consumed_raw = original.into_raw_bstr();
        if !current.is_null() {
            // SAFETY: `current` is the native replacement now solely owned by
            // this cell, and the consuming native call has completed.
            unsafe { SysFreeString(current) };
        }
    }
}

#[cfg(target_os = "windows")]
// SAFETY: The `Box<BSTR>` cell and the BSTR allocation it points to are owned
// exclusively by this wrapper; BSTRs are thread-agnostic OLE-heap allocations
// and `SysFreeString` is callable from any thread, so transferring the sole
// owner across threads (required by the global `Mutex<PointerRegistry>`) is
// sound.
unsafe impl Send for OwnedBstrCell {}

#[cfg(target_os = "windows")]
// `VarPtr(Variant)` materializes a Windows-observable VARIANT cell from the
// canonical semantic Variant carrier; the raw VARIANT is still a boundary
// projection rather than the canonical runtime container itself.
struct OwnedVariant {
    cell: VARIANT,
    tracked_bstr: Option<TrackedBstrAccountingToken>,
    original_object: Option<(Variant, usize)>,
}

#[cfg(target_os = "windows")]
impl std::fmt::Debug for OwnedVariant {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("OwnedVariant")
            .field("tracked_bstr", &self.tracked_bstr.is_some())
            .finish_non_exhaustive()
    }
}

#[cfg(target_os = "windows")]
impl OwnedVariant {
    fn from_variant(value: &Variant) -> Result<Self, String> {
        // SAFETY: The all-zero bit pattern is a valid VARIANT — vt == VT_EMPTY
        // with no owned payload, the same state `VariantInit` produces; every
        // union field (integers, raw pointers) admits the zero bit pattern.
        let mut owner = Self {
            // SAFETY: the all-zero bit pattern is a valid VT_EMPTY VARIANT.
            cell: unsafe { std::mem::zeroed() },
            tracked_bstr: None,
            original_object: None,
        };
        // SAFETY: `&mut owner.cell` points to a writable VARIANT local that was
        // just zeroed to VT_EMPTY, so the callee overwrites no pre-owned
        // payload; it stores only payloads matching the vt it sets, which
        // `OwnedVariant`'s Drop later releases via VariantClear.
        unsafe {
            set_windows_variant_from_variant(&mut owner.cell, value, &mut owner.tracked_bstr)?
        };
        if value.vtype() == crate::VarType::Object {
            // SAFETY: the just-projected object cell owns this IUnknown pointer.
            let pointer = unsafe { owner.cell.Anonymous.Anonymous.Anonymous.punkVal } as usize;
            owner.original_object = Some((value.clone(), pointer));
        }
        Ok(owner)
    }

    fn as_ptr(&mut self) -> *mut c_void {
        (&mut self.cell as *mut VARIANT).cast()
    }
}

#[cfg(target_os = "windows")]
// SAFETY: ASSUMPTION — the VARIANT struct itself is plain data, and the BSTR /
// SAFEARRAY payloads `set_windows_variant_from_variant` writes are fresh,
// exclusively owned, thread-agnostic OLE-heap allocations; the punkVal payload
// is one retained `ObjectRef` reference whose runtime refcount is atomic
// (object_ref.rs `ref_count: AtomicU32`), so retain/release of runtime objects
// is thread-safe. What must additionally hold for foreign (CreateObject) COM
// payloads, which may be apartment-affine: the pin is created and dropped
// (VariantClear → Release) on the same VM thread within one statement, per
// docs/spec/OXVBA_POINTER_HELPERS_CONTRACT_V1.md — Send is demanded only
// because the registry lives behind a process-global Mutex, not because
// entries actually migrate threads.
unsafe impl Send for OwnedVariant {}

#[cfg(target_os = "windows")]
fn advance_windows_safearray_indices(
    indices: &mut [i32],
    bounds: &[crate::safe_array::SafeArrayBound],
) -> Result<(), String> {
    let mut carry = true;
    for (dimension, bound) in bounds.iter().enumerate().rev() {
        if !carry {
            break;
        }
        let relative = i64::from(indices[dimension]) - i64::from(bound.lower);
        let next_relative = relative + 1;
        if next_relative >= i64::from(bound.count) {
            indices[dimension] = bound.lower;
        } else {
            let next = i64::from(bound.lower) + next_relative;
            indices[dimension] = i32::try_from(next).map_err(|_| {
                format!(
                    "SAFEARRAY bound at dimension {dimension} exceeds the Windows LONG index range"
                )
            })?;
            carry = false;
        }
    }
    if carry {
        return Err("SAFEARRAY multidimensional index progression exhausted early".to_string());
    }
    Ok(())
}

#[cfg(target_os = "windows")]
fn retained_iunknown_pointer(object: &crate::ObjectRef) -> *mut c_void {
    object.query_iunknown().into_raw_iunknown().cast::<c_void>()
}

#[cfg(target_os = "windows")]
#[allow(unsafe_op_in_unsafe_fn)]
unsafe fn set_windows_variant_array_arg(
    variant: *mut VARIANT,
    array: &crate::safe_array::SafeArray,
) -> Result<(), String> {
    let source = Variant::from_safearray(array.clone());
    let mut native = native_record_pin::NativeValue::from_value(&source)?;
    // SAFETY: the caller supplies a fresh VT_EMPTY cell. Transfer the complete
    // typed descriptor, retaining both element type and all original bounds.
    unsafe {
        *variant = native.0;
        native.0.Anonymous.Anonymous.vt = 0;
    }
    Ok(())
}
#[cfg(target_os = "windows")]
#[allow(unsafe_op_in_unsafe_fn)]
unsafe fn set_windows_variant_from_variant(
    variant: *mut VARIANT,
    value: &Variant,
    tracked_bstr: &mut Option<TrackedBstrAccountingToken>,
) -> Result<(), String> {
    // This projection is only valid for the fresh VT_EMPTY cell and matching
    // empty accounting slot constructed by OwnedVariant::from_variant.
    match value.vtype() {
        crate::VarType::Empty => {
            (*variant).Anonymous.Anonymous.vt = VT_EMPTY;
        }
        crate::VarType::Null => {
            (*variant).Anonymous.Anonymous.vt = VT_NULL;
        }
        crate::VarType::Error => {
            (*variant).Anonymous.Anonymous.vt = VT_ERROR;
            let code = value.as_error_code().expect("error payload");
            (*variant).Anonymous.Anonymous.Anonymous.scode = if (0..=65535).contains(&code) {
                (0x800a_0000u32 | code as u32) as i32
            } else {
                code
            };
        }
        crate::VarType::Integer => {
            (*variant).Anonymous.Anonymous.vt = 2;
            (*variant).Anonymous.Anonymous.Anonymous.iVal =
                value.as_i16().expect("integer payload");
        }
        crate::VarType::Long => {
            (*variant).Anonymous.Anonymous.vt = VT_I4;
            (*variant).Anonymous.Anonymous.Anonymous.lVal = value.as_i32().expect("long payload");
        }
        crate::VarType::SignedByte => {
            (*variant).Anonymous.Anonymous.vt = VT_I1;
            (*variant).Anonymous.Anonymous.Anonymous.cVal = value.as_i8().expect("i1 payload");
        }
        crate::VarType::Byte => {
            (*variant).Anonymous.Anonymous.vt = VT_UI1;
            (*variant).Anonymous.Anonymous.Anonymous.bVal = value.as_u8().expect("byte payload");
        }
        crate::VarType::UnsignedInteger => {
            (*variant).Anonymous.Anonymous.vt = VT_UI2;
            (*variant).Anonymous.Anonymous.Anonymous.uiVal = value.as_u16().expect("ui2 payload");
        }
        crate::VarType::UnsignedLong => {
            (*variant).Anonymous.Anonymous.vt = VT_UI4;
            (*variant).Anonymous.Anonymous.Anonymous.ulVal = value.as_u32().expect("ui4 payload");
        }
        crate::VarType::UnsignedInt => {
            (*variant).Anonymous.Anonymous.vt = VT_UINT;
            (*variant).Anonymous.Anonymous.Anonymous.uintVal =
                value.as_u32().expect("uint payload");
        }
        crate::VarType::LongLong => {
            (*variant).Anonymous.Anonymous.vt = VT_I8;
            (*variant).Anonymous.Anonymous.Anonymous.llVal = value.as_i64().expect("i64 payload");
        }
        crate::VarType::UnsignedLongLong => {
            (*variant).Anonymous.Anonymous.vt = VT_UI8;
            (*variant).Anonymous.Anonymous.Anonymous.ullVal = value.as_u64().expect("ui8 payload");
        }
        crate::VarType::Boolean => {
            (*variant).Anonymous.Anonymous.vt = VT_BOOL;
            (*variant).Anonymous.Anonymous.Anonymous.boolVal =
                if value.as_bool().expect("bool payload") {
                    -1
                } else {
                    0
                };
        }
        crate::VarType::Single => {
            (*variant).Anonymous.Anonymous.vt = VT_R4;
            (*variant).Anonymous.Anonymous.Anonymous.fltVal =
                value.as_f32().expect("single payload");
        }
        crate::VarType::Double => {
            (*variant).Anonymous.Anonymous.vt = VT_R8;
            (*variant).Anonymous.Anonymous.Anonymous.dblVal =
                value.as_f64().expect("double payload");
        }
        crate::VarType::Date => {
            (*variant).Anonymous.Anonymous.vt = VT_DATE;
            (*variant).Anonymous.Anonymous.Anonymous.dblVal =
                value.as_date_f64().expect("date payload");
        }
        crate::VarType::Currency => {
            (*variant).Anonymous.Anonymous.vt = VT_CY;
            (*variant).Anonymous.Anonymous.Anonymous.cyVal.int64 =
                value.as_currency_scaled_i64().expect("currency payload");
        }
        crate::VarType::String => {
            let text = value
                .try_as_bstr()?
                .ok_or_else(|| "canonical string Variant lost owned BSTR payload".to_string())?;
            // SAFETY: this function's caller supplies the fresh writable
            // VARIANT cell and the matching empty accounting slot.
            unsafe { OwnedBstr::from_bstr(&text)?.transfer_into_variant(variant, tracked_bstr)? };
        }
        crate::VarType::Decimal => {
            let bytes = value.to_wire_bytes();
            std::ptr::copy_nonoverlapping(bytes.as_ptr(), variant.cast::<u8>(), bytes.len());
        }
        crate::VarType::Object => {
            let Some(object) = value.as_object_ref() else {
                (*variant).Anonymous.Anonymous.vt = VT_UNKNOWN;
                (*variant).Anonymous.Anonymous.Anonymous.punkVal = std::ptr::null_mut();
                return Ok(());
            };
            (*variant).Anonymous.Anonymous.vt = VT_UNKNOWN;
            (*variant).Anonymous.Anonymous.Anonymous.punkVal = if object.raw() == 0 {
                std::ptr::null_mut()
            } else {
                retained_iunknown_pointer(&object)
            };
        }
        crate::VarType::ArrayVariant => {
            let Some(array) = value.try_as_safearray()? else {
                return Err(
                    "VarPtr over Variant containing null SAFEARRAY payload is not yet supported"
                        .to_string(),
                );
            };
            set_windows_variant_array_arg(variant, &array)?;
        }
        crate::VarType::Record => {
            return Err(
                "COM records cannot be marshaled by the generic pointer helper".to_string(),
            );
        }
        crate::VarType::ProcRef => {
            return Err("procedure references cannot be marshaled as VARIANT values".to_string());
        }
    }
    Ok(())
}

#[cfg(target_os = "windows")]
impl Drop for OwnedVariant {
    fn drop(&mut self) {
        // SAFETY: `self.cell` was fully initialized by `from_variant` (zeroed to
        // VT_EMPTY, then populated by `set_windows_variant_from_variant`). The
        // VARIANT owns its BSTR/SAFEARRAY/IUnknown payload until VariantClear,
        // and this cell is that payload's sole owner (the registry holds the
        // entry until `free_pins` removes it, after the consuming native call
        // per docs/spec/OXVBA_POINTER_HELPERS_CONTRACT_V1.md), so this releases
        // the current payload exactly once; VariantClear on VT_EMPTY is a no-op.
        // The transfer token is independent of the cell's current pointer/type:
        // native code may already have cleared the original and installed null
        // or a replacement. Settle the original tracked allocation only after
        // VariantClear has released whichever valid native value remains.
        unsafe {
            let _ = VariantClear(&mut self.cell);
        }
        drop(self.tracked_bstr.take());
    }
}

#[derive(Debug)]
enum PointerEntry {
    #[cfg(windows)]
    Record(OwnedNativeRecord),
    #[cfg(target_os = "windows")]
    Bstr(OwnedBstr),
    #[cfg(target_os = "windows")]
    BstrCell(OwnedBstrCell),
    #[cfg(target_os = "windows")]
    // Heap ownership keeps the VARIANT cell address stable while the enum and
    // HashMap buckets move. PointerRegistry::insert derives its public pin ID
    // from this allocation only after the Box exists.
    VariantCell(Box<OwnedVariant>),
    #[cfg(not(target_os = "windows"))]
    Utf16(Box<[u16]>),
    Bytes(Box<[u8]>),
    I8(Box<i8>),
    U8(Box<u8>),
    I16(Box<i16>),
    U16(Box<u16>),
    I32(Box<i32>),
    U32(Box<u32>),
    I64(Box<i64>),
    U64(Box<u64>),
    F32(Box<f32>),
    F64(Box<f64>),
    Bool(Box<i16>),
    ObjectIdentity(Box<i64>),
}

impl PointerEntry {
    fn byte_len(&self) -> usize {
        match self {
            #[cfg(windows)]
            Self::Record(value) => value.words.len() * 8,
            #[cfg(windows)]
            Self::Bstr(value) => value.0.as_ref().map_or(0, |b| b.payload_bytes().len() + 2),
            #[cfg(windows)]
            Self::BstrCell(_) => std::mem::size_of::<BSTR>(),
            #[cfg(windows)]
            Self::VariantCell(_) => std::mem::size_of::<VARIANT>(),
            #[cfg(not(windows))]
            Self::Utf16(value) => value.len() * 2,
            Self::Bytes(value) => value.len(),
            Self::I8(_) | Self::U8(_) => 1,
            Self::I16(_) | Self::U16(_) | Self::Bool(_) => 2,
            Self::I32(_) | Self::U32(_) | Self::F32(_) => 4,
            Self::I64(_) | Self::U64(_) | Self::F64(_) | Self::ObjectIdentity(_) => 8,
        }
    }
    fn as_ptr(&mut self) -> *mut c_void {
        match self {
            #[cfg(windows)]
            Self::Record(value) => value.words.as_mut_ptr().cast(),
            #[cfg(target_os = "windows")]
            Self::Bstr(value) => value.as_ptr(),
            #[cfg(target_os = "windows")]
            Self::BstrCell(value) => value.as_ptr(),
            #[cfg(target_os = "windows")]
            Self::VariantCell(value) => value.as_ptr(),
            #[cfg(not(target_os = "windows"))]
            Self::Utf16(value) => value.as_mut_ptr().cast(),
            Self::Bytes(value) => value.as_mut_ptr().cast(),
            Self::I8(value) => (&mut **value as *mut i8).cast(),
            Self::U8(value) => (&mut **value as *mut u8).cast(),
            Self::I16(value) => (&mut **value as *mut i16).cast(),
            Self::U16(value) => (&mut **value as *mut u16).cast(),
            Self::I32(value) => (&mut **value as *mut i32).cast(),
            Self::U32(value) => (&mut **value as *mut u32).cast(),
            Self::I64(value) => (&mut **value as *mut i64).cast(),
            Self::U64(value) => (&mut **value as *mut u64).cast(),
            Self::F32(value) => (&mut **value as *mut f32).cast(),
            Self::F64(value) => (&mut **value as *mut f64).cast(),
            Self::Bool(value) => (&mut **value as *mut i16).cast(),
            Self::ObjectIdentity(value) => (&mut **value as *mut i64).cast(),
        }
    }
}

#[derive(Debug, Default)]
struct PointerRegistry {
    entries: HashMap<usize, PointerEntry>,
}

impl PointerRegistry {
    fn base(&self, pointer: i64) -> Option<usize> {
        let pointer = pointer as usize;
        if self.entries.contains_key(&pointer) {
            return Some(pointer);
        }
        self.entries.iter().find_map(|(&base, entry)| {
            pointer
                .checked_sub(base)
                .filter(|offset| *offset < entry.byte_len())
                .map(|_| base)
        })
    }
    fn entry(&self, pointer: i64) -> Option<&PointerEntry> {
        self.base(pointer).and_then(|base| self.entries.get(&base))
    }
    fn insert(&mut self, mut entry: PointerEntry) -> i64 {
        let addr = entry.as_ptr() as usize;
        self.entries.insert(addr, entry);
        addr as i64
    }

    fn read_back_string_payload_variant(&self, pointer: i64) -> Result<Variant, String> {
        if pointer == 0 {
            return Ok(Variant::from_string(BStr::empty()));
        }
        let Some(entry) = self.entry(pointer) else {
            return Err(format!(
                "pointer helper registry does not contain string payload pointer {pointer}"
            ));
        };
        match entry {
            #[cfg(target_os = "windows")]
            PointerEntry::Bstr(value) => {
                let Some(value) = value.0.as_ref() else {
                    return Ok(Variant::from_string(BStr::empty()));
                };
                Ok(Variant::from_string(BStr::from_bytes(
                    value.payload_bytes(),
                )?))
            }
            #[cfg(target_os = "windows")]
            PointerEntry::BstrCell(value) => {
                if (*value.cell).is_null() {
                    return Ok(Variant::from_string(BStr::empty()));
                }
                // SAFETY: `*value.cell` was checked non-null just above. The
                // cell holds either our original allocation or the BSTR a
                // native call wrote through the VarPtr(String) cell, whose
                // ownership transferred to the cell with that write; per
                // docs/spec/OXVBA_POINTER_HELPERS_CONTRACT_V1.md the pin (and
                // thus the BSTR) stays live until `free_pins` runs after this
                // read-back, so SysStringByteLen reads a live allocation's 4-byte
                // byte-length prefix at ptr-4.
                let len = unsafe { windows_sys::Win32::Foundation::SysStringByteLen(*value.cell) }
                    as usize;
                // SAFETY: Same live, non-null BSTR as above; `len` is the
                // payload byte count from the length prefix, so the first
                // `len` bytes are within the allocation and readable. Embedded
                // NULs and odd byte lengths are legal because the length is
                // prefix-derived; the registry guard held by the caller keeps
                // the entry alive for the duration of this borrow.
                let bytes = unsafe { std::slice::from_raw_parts((*value.cell).cast::<u8>(), len) };
                Ok(Variant::from_string(BStr::from_bytes(bytes)?))
            }
            #[cfg(not(target_os = "windows"))]
            PointerEntry::Utf16(value) => {
                let end = value
                    .iter()
                    .position(|unit| *unit == 0)
                    .unwrap_or(value.len());
                Ok(Variant::from_string(BStr::from_utf16_units(&value[..end])?))
            }
            other => Err(format!(
                "pointer helper entry {other:?} cannot be read back as a string payload"
            )),
        }
    }

    fn read_back_byte_array_payload_variant(&self, pointer: i64) -> Result<Variant, String> {
        if pointer == 0 {
            return Ok(Variant::from_safearray(
                crate::safe_array::SafeArray::from_variants(Vec::new()),
            ));
        }
        let Some(entry) = self.entry(pointer) else {
            return Err(format!(
                "pointer helper registry does not contain byte-array payload pointer {pointer}"
            ));
        };
        match entry {
            PointerEntry::Bytes(bytes) => Ok(Variant::from_safearray(
                crate::safe_array::SafeArray::from_variants(
                    bytes
                        .iter()
                        .map(|byte| crate::Variant::from_u8(*byte))
                        .collect(),
                ),
            )),
            other => Err(format!(
                "pointer helper entry {other:?} cannot be read back as a byte-array payload"
            )),
        }
    }

    fn read_back_scalar_payload_variant(
        &self,
        pointer: i64,
        kind: ScalarPointerKind,
    ) -> Result<Variant, String> {
        let Some(entry) = self.entry(pointer) else {
            return Err(format!(
                "pointer helper registry does not contain scalar payload pointer {pointer}"
            ));
        };
        match (kind, entry) {
            (ScalarPointerKind::Boolean, PointerEntry::Bool(value)) => {
                Ok(Variant::from_bool(**value != 0))
            }
            (ScalarPointerKind::Byte, PointerEntry::U8(value)) => Ok(Variant::from_u8(**value)),
            (ScalarPointerKind::Integer, PointerEntry::I16(value)) => {
                Ok(Variant::from_i16(**value))
            }
            (ScalarPointerKind::Long, PointerEntry::I32(value)) => Ok(Variant::from_i32(**value)),
            (ScalarPointerKind::LongLong, PointerEntry::I64(value)) => {
                Ok(Variant::from_i64(**value))
            }
            (ScalarPointerKind::LongPtr, PointerEntry::I64(value)) => {
                Ok(Variant::from_i64(**value))
            }
            (ScalarPointerKind::Single, PointerEntry::F32(value)) => Ok(Variant::from_f32(**value)),
            (ScalarPointerKind::Double, PointerEntry::F64(value)) => Ok(Variant::from_f64(**value)),
            (ScalarPointerKind::Currency, PointerEntry::I64(value)) => {
                Ok(Variant::from_currency_scaled_i64(**value))
            }
            (ScalarPointerKind::Date, PointerEntry::F64(value)) => {
                Ok(Variant::from_date_f64(**value))
            }
            (_, other) => Err(format!(
                "pointer helper entry {other:?} cannot be read back as {kind:?}"
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ScalarPointerKind {
    Boolean,
    Byte,
    Integer,
    Long,
    LongLong,
    LongPtr,
    Single,
    Double,
    Currency,
    Date,
}

fn registry() -> &'static Mutex<PointerRegistry> {
    static REGISTRY: OnceLock<Mutex<PointerRegistry>> = OnceLock::new();
    REGISTRY.get_or_init(|| Mutex::new(PointerRegistry::default()))
}

pub fn register_utf16_string(text: &str) -> Result<i64, String> {
    #[cfg(target_os = "windows")]
    let entry = PointerEntry::Bstr(OwnedBstr::from_text(text)?);

    #[cfg(not(target_os = "windows"))]
    let entry = {
        let core = BStr::from(text).owned_core();
        PointerEntry::Utf16(core.payload_units_with_nul().to_vec().into_boxed_slice())
    };

    let mut guard = registry()
        .lock()
        .map_err(|_| "pointer helper registry lock poisoned".to_string())?;
    Ok(guard.insert(entry))
}

pub fn register_variant_pointer(value: &Variant) -> Result<i64, String> {
    let entry = match value.vtype() {
        crate::VarType::Empty | crate::VarType::Null => return Ok(0),
        crate::VarType::Integer => {
            PointerEntry::I16(Box::new(value.as_i16().expect("Integer Variant payload")))
        }
        crate::VarType::Long => {
            PointerEntry::I32(Box::new(value.as_i32().expect("Long Variant payload")))
        }
        crate::VarType::LongLong => {
            PointerEntry::I64(Box::new(value.as_i64().expect("LongLong Variant payload")))
        }
        crate::VarType::SignedByte => {
            PointerEntry::I8(Box::new(value.as_i8().expect("SignedByte Variant payload")))
        }
        crate::VarType::Byte => {
            PointerEntry::U8(Box::new(value.as_u8().expect("Byte Variant payload")))
        }
        crate::VarType::UnsignedInteger => PointerEntry::U16(Box::new(
            value.as_u16().expect("UnsignedInteger Variant payload"),
        )),
        crate::VarType::UnsignedLong | crate::VarType::UnsignedInt => PointerEntry::U32(Box::new(
            value.as_u32().expect("unsigned 32-bit Variant payload"),
        )),
        crate::VarType::UnsignedLongLong => PointerEntry::U64(Box::new(
            value.as_u64().expect("UnsignedLongLong Variant payload"),
        )),
        crate::VarType::Single => {
            PointerEntry::F32(Box::new(value.as_f32().expect("Single Variant payload")))
        }
        crate::VarType::Double => {
            PointerEntry::F64(Box::new(value.as_f64().expect("Double Variant payload")))
        }
        crate::VarType::Date => {
            PointerEntry::F64(Box::new(value.as_date_f64().expect("Date Variant payload")))
        }
        crate::VarType::Currency => PointerEntry::I64(Box::new(
            value
                .as_currency_scaled_i64()
                .expect("Currency Variant payload"),
        )),
        crate::VarType::Boolean => PointerEntry::Bool(Box::new(
            if value.as_bool().expect("Boolean Variant payload") {
                -1
            } else {
                0
            },
        )),
        crate::VarType::String => {
            let text = value
                .as_bstr()
                .ok_or_else(|| "String Variant lost BSTR payload".to_string())?;
            #[cfg(target_os = "windows")]
            {
                PointerEntry::Bstr(OwnedBstr::from_bstr(&text)?)
            }
            #[cfg(not(target_os = "windows"))]
            {
                let core = text.owned_core();
                let data: Box<[u16]> = core.payload_units_with_nul().to_vec().into_boxed_slice();
                PointerEntry::Utf16(data)
            }
        }
        crate::VarType::ArrayVariant => {
            let Some(array) = value.as_safearray() else {
                return Err("VarPtr over null SAFEARRAY payload is not yet supported".to_string());
            };
            byte_array_variant_pointer_entry(&array)?
        }
        crate::VarType::Object => {
            let Some(handle) = value.as_object_ref() else {
                return Ok(0);
            };
            PointerEntry::ObjectIdentity(Box::new(handle.raw_iunknown() as usize as i64))
        }
        crate::VarType::Error => PointerEntry::I32(Box::new(
            value.as_error_code().expect("Error Variant payload"),
        )),
        crate::VarType::Decimal => {
            return Err("VarPtr/ObjPtr over Decimal is not yet supported".to_string());
        }
        crate::VarType::Record => {
            #[cfg(windows)]
            {
                PointerEntry::Record(OwnedNativeRecord::from_record(
                    &value
                        .as_vba_record()
                        .ok_or("VarPtr over COM records is not supported")?,
                )?)
            }
            #[cfg(not(windows))]
            {
                return Err("VarPtr over records is currently Windows-only".into());
            }
        }
        crate::VarType::ProcRef => {
            return Err("VarPtr/ObjPtr over procedure references is not supported".to_string());
        }
    };

    let mut guard = registry()
        .lock()
        .map_err(|_| "pointer helper registry lock poisoned".to_string())?;
    Ok(guard.insert(entry))
}

fn byte_array_variant_pointer_entry(
    array: &crate::safe_array::SafeArray,
) -> Result<PointerEntry, String> {
    let Some(elements) = array.variant_elements() else {
        return Err(
            "VarPtr over array shape without element payload is not yet supported".to_string(),
        );
    };
    let mut bytes = Vec::with_capacity(elements.len());
    for element in elements {
        match element.vtype() {
            crate::VarType::Empty | crate::VarType::Null => bytes.push(0),
            crate::VarType::Byte => bytes.push(element.as_u8().expect("Byte Variant payload")),
            crate::VarType::Boolean => {
                bytes.push(if element.as_bool().expect("Boolean Variant payload") {
                    1
                } else {
                    0
                })
            }
            crate::VarType::Integer => {
                let value = i32::from(element.as_i16().expect("Integer Variant payload"));
                if !(0..=255).contains(&value) {
                    return Err(format!(
                        "VarPtr over array payload currently requires byte-compatible elements, got Integer({value})"
                    ));
                }
                bytes.push(value as u8);
            }
            crate::VarType::Long => {
                let value = element.as_i32().expect("Long Variant payload");
                if !(0..=255).contains(&value) {
                    return Err(format!(
                        "VarPtr over array payload currently requires byte-compatible elements, got Long({value})"
                    ));
                }
                bytes.push(value as u8);
            }
            other => {
                return Err(format!(
                    "VarPtr over array payload currently requires byte-compatible Variant elements, got {other:?}"
                ));
            }
        }
    }
    Ok(PointerEntry::Bytes(bytes.into_boxed_slice()))
}

pub fn register_string_variant_pointer(value: &Variant) -> Result<i64, String> {
    #[cfg(target_os = "windows")]
    {
        let entry = match value.vtype() {
            crate::VarType::String => {
                let text = value
                    .as_bstr()
                    .ok_or_else(|| "String Variant lost BSTR payload".to_string())?;
                PointerEntry::BstrCell(OwnedBstrCell::from_bstr(&text)?)
            }
            crate::VarType::Empty | crate::VarType::Null => return Ok(0),
            _ => return Err("VarPtr over String requires a string variable".to_string()),
        };

        let mut guard = registry()
            .lock()
            .map_err(|_| "pointer helper registry lock poisoned".to_string())?;
        Ok(guard.insert(entry))
    }

    #[cfg(not(target_os = "windows"))]
    {
        register_variant_pointer(value)
    }
}

pub fn register_variant_var_variant_pointer(value: &Variant) -> Result<i64, String> {
    #[cfg(target_os = "windows")]
    {
        let entry = PointerEntry::VariantCell(Box::new(OwnedVariant::from_variant(value)?));
        let mut guard = registry()
            .lock()
            .map_err(|_| "pointer helper registry lock poisoned".to_string())?;
        Ok(guard.insert(entry))
    }

    #[cfg(not(target_os = "windows"))]
    {
        register_variant_pointer(value)
    }
}

pub fn register_object_variant_pointer(value: &Variant) -> Result<i64, String> {
    match value.vtype() {
        crate::VarType::Empty | crate::VarType::Null => Ok(0),
        crate::VarType::Object => match value.as_object_ref() {
            Some(handle) if handle.raw() == 0 => Ok(0),
            Some(handle) => Ok(handle.raw_iunknown() as usize as i64),
            None => Ok(0),
        },
        _ => Err("ObjPtr requires an object reference".to_string()),
    }
}

pub fn register_array_payload_pointer(array: &crate::safe_array::SafeArray) -> Result<i64, String> {
    let entry = byte_array_variant_pointer_entry(array)?;
    let mut guard = registry()
        .lock()
        .map_err(|_| "pointer helper registry lock poisoned".to_string())?;
    Ok(guard.insert(entry))
}

pub fn lookup_pointer(pointer: i64) -> Option<*mut c_void> {
    if pointer == 0 {
        return Some(std::ptr::null_mut());
    }
    let mut guard = registry().lock().ok()?;
    let entry = guard.entries.get_mut(&(pointer as usize))?;
    Some(entry.as_ptr())
}

pub fn read_back_string_payload_variant(pointer: i64) -> Result<Variant, String> {
    let guard = registry()
        .lock()
        .map_err(|_| "pointer helper registry lock poisoned".to_string())?;
    guard.read_back_string_payload_variant(pointer)
}

pub fn read_back_byte_array_payload_variant(pointer: i64) -> Result<Variant, String> {
    let guard = registry()
        .lock()
        .map_err(|_| "pointer helper registry lock poisoned".to_string())?;
    guard.read_back_byte_array_payload_variant(pointer)
}

pub fn read_back_scalar_payload_variant(
    pointer: i64,
    kind: ScalarPointerKind,
) -> Result<Variant, String> {
    let guard = registry()
        .lock()
        .map_err(|_| "pointer helper registry lock poisoned".to_string())?;
    guard.read_back_scalar_payload_variant(pointer, kind)
}

pub fn read_back_record(pointer: i64) -> Result<Variant, String> {
    let guard = registry()
        .lock()
        .map_err(|_| "pointer helper registry lock poisoned".to_string())?;
    #[cfg(windows)]
    if let Some(PointerEntry::Record(record)) = guard.entry(pointer) {
        return record.read();
    }
    let _ = (&guard, pointer);
    Err("pointer helper registry does not contain a native UDT pin".into())
}

pub fn read_back_variant(pointer: i64) -> Result<Variant, String> {
    let guard = registry()
        .lock()
        .map_err(|_| "pointer helper registry lock poisoned".to_string())?;
    #[cfg(windows)]
    if let Some(PointerEntry::VariantCell(value)) = guard.entry(pointer) {
        // SAFETY: this pin retains the initialized native VARIANT until copyback
        // completes. Same-object reads preserve the canonical host binding.
        unsafe {
            let kind = value.cell.Anonymous.Anonymous.vt;
            if matches!(kind, 9 | 13) {
                let pointer = value.cell.Anonymous.Anonymous.Anonymous.punkVal as usize;
                if pointer == 0 {
                    return Ok(Variant::nothing());
                }
                if let Some((original, address)) = &value.original_object
                    && pointer == *address
                {
                    return Ok(original.clone());
                }
                return Err("VarPtr Variant object replacement requires host binding".into());
            }
            return native_record_pin::value_from_native(&value.cell);
        }
    }
    let _ = (&guard, pointer);
    Err("pointer helper registry does not contain a native Variant pin".into())
}

pub fn register_byte_buffer(bytes: Vec<u8>) -> Result<i64, String> {
    let mut guard = registry()
        .lock()
        .map_err(|_| "pointer helper registry lock poisoned".to_string())?;
    Ok(guard.insert(PointerEntry::Bytes(bytes.into_boxed_slice())))
}

/// Release the pinned cells at the given addresses, dropping each entry (which
/// frees its native BSTR/VARIANT/buffer). Addresses not in the registry — `0`,
/// already-freed, or never-pinned integers — are ignored. Pins are scoped to the
/// VBA statement that creates them: the VM drains the addresses it pinned at each
/// statement boundary (after any pointer write-back has read them), matching VBA's
/// "the pointer is valid for the duration of the call" contract and keeping the
/// registry bounded for long-running/looping code instead of leaking every pin.
pub fn free_pins(addrs: &[i64]) {
    if addrs.is_empty() {
        return;
    }
    let Ok(mut guard) = registry().lock() else {
        return;
    };
    for addr in addrs {
        if *addr != 0
            && let Some(base) = guard.base(*addr)
        {
            guard.entries.remove(&base);
        }
    }
}

/// Number of pinned cells currently live in the registry. Diagnostic — used by
/// tests to assert that pointer-helper pins do not accumulate across a loop.
pub fn live_pin_count() -> usize {
    registry()
        .lock()
        .map(|guard| guard.entries.len())
        .unwrap_or(0)
}

#[cfg(test)]
// Test-support code exercising the documented production pointer-helper paths.
#[allow(clippy::undocumented_unsafe_blocks)]
mod tests {
    #[cfg(windows)]
    #[test]
    fn interior_variant_writeback_retains_payload_and_releases_its_pin() {
        let pin = super::register_variant_var_variant_pointer(&crate::Variant::empty()).unwrap();
        let native = super::lookup_pointer(pin)
            .unwrap()
            .cast::<windows_sys::Win32::System::Variant::VARIANT>();
        // SAFETY: this test owns the live empty pin and transfers a fresh BSTR
        // into its VARIANT. The registry releases it after readback.
        unsafe {
            (*native).Anonymous.Anonymous.vt = 8;
            (*native).Anonymous.Anonymous.Anonymous.bstrVal =
                windows_sys::Win32::Foundation::SysAllocStringLen([65u16, 0, 955].as_ptr(), 3);
        }
        let read = super::read_back_variant(pin + 8).unwrap();
        assert_eq!(read.as_bstr().unwrap().to_utf16_units(), [65, 0, 955]);
        super::free_pins(&[pin + 8]);
        assert!(super::lookup_pointer(pin).is_none());
        assert_eq!(read.as_bstr().unwrap().to_utf16_units(), [65, 0, 955]);
    }

    #[cfg(windows)]
    #[test]
    fn variant_pointer_tags_and_error_codes_match_automation() {
        use windows_sys::Win32::System::Variant::VARIANT;
        let integer =
            super::register_variant_var_variant_pointer(&crate::Variant::from_i16(7)).unwrap();
        let error =
            super::register_variant_var_variant_pointer(&crate::Variant::from_error_code(2042))
                .unwrap();
        // SAFETY: both pins retain these initialized cells through inspection.
        unsafe {
            let value = super::lookup_pointer(integer).unwrap().cast::<VARIANT>();
            assert_eq!((*value).Anonymous.Anonymous.vt, 2);
            let value = super::lookup_pointer(error).unwrap().cast::<VARIANT>();
            assert_eq!(
                (*value).Anonymous.Anonymous.Anonymous.scode as u32,
                0x800a07fa
            );
        }
        assert_eq!(
            super::read_back_variant(error).unwrap().as_error_code(),
            Some(2042)
        );
        super::free_pins(&[integer, error]);
    }

    use super::{
        ScalarPointerKind, free_pins, lookup_pointer, register_byte_buffer,
        register_object_variant_pointer, register_utf16_string, register_variant_pointer,
    };
    #[cfg(target_os = "windows")]
    use super::{register_string_variant_pointer, register_variant_var_variant_pointer};
    #[cfg(target_os = "windows")]
    use crate::live_counters::thread_live_handle_counts;
    #[cfg(target_os = "windows")]
    use crate::{Decimal96, VarType};
    use crate::{ObjectRef, Variant, bstr::BStr};
    #[cfg(target_os = "windows")]
    use windows_sys::{
        Win32::Foundation::{
            SysAllocStringByteLen, SysAllocStringLen, SysFreeString, SysStringLen,
        },
        Win32::System::Ole::{SafeArrayGetDim, SafeArrayGetElement},
        Win32::System::Variant::{
            VARIANT, VT_ARRAY, VT_BSTR, VT_DATE, VT_I4, VT_NULL, VT_UNKNOWN, VT_VARIANT,
            VariantClear,
        },
        core::BSTR,
    };

    #[test]
    fn utf16_pointer_helper_allocates_terminated_text() {
        let ptr = register_utf16_string("abc").expect("register string");
        assert_ne!(ptr, 0);
        let raw = lookup_pointer(ptr).expect("lookup pointer") as *const u16;
        assert!(!raw.is_null());
        let slice = unsafe { std::slice::from_raw_parts(raw, 4) };
        assert_eq!(slice, &[97, 98, 99, 0]);
    }

    #[cfg(target_os = "windows")]
    #[test]
    fn string_pointer_helper_materializes_real_bstr_payload() {
        let ptr = register_utf16_string("abc").expect("register string");
        assert_ne!(ptr, 0);
        let raw = lookup_pointer(ptr).expect("lookup pointer").cast::<u16>();
        assert!(!raw.is_null());
        let len = unsafe { SysStringLen(raw.cast()) };
        assert_eq!(len, 3);
    }

    #[test]
    fn free_pins_releases_only_the_named_addresses() {
        // `free_pins` removes exactly the listed pins (idempotent on absent
        // addresses) so the VM can scope a statement's pins without disturbing
        // any other live pin. Keyed by concrete addresses, so it is robust to
        // pins other tests leave in the shared global registry.
        let a = register_byte_buffer(vec![1, 2, 3]).expect("register a");
        let b = register_byte_buffer(vec![4, 5, 6]).expect("register b");
        assert!(lookup_pointer(a).is_some());
        assert!(lookup_pointer(b).is_some());

        free_pins(&[a, 0, a]); // includes a no-op `0` and a duplicate
        assert!(lookup_pointer(a).is_none(), "a should be released");
        assert!(lookup_pointer(b).is_some(), "b must be untouched");

        free_pins(&[b]);
        assert!(lookup_pointer(b).is_none(), "b should be released");
    }

    #[test]
    fn variant_pointer_handles_canonical_scalars_and_strings() {
        let string_ptr = register_variant_pointer(&Variant::from_string(BStr::from("xyz")))
            .expect("register string variant");
        assert_ne!(string_ptr, 0);
        let scalar_ptr =
            register_variant_pointer(&Variant::from_i64(42)).expect("register i64 variant");
        assert_ne!(scalar_ptr, 0);
    }

    #[test]
    fn variant_pointer_preserves_date_and_byte_array_payloads() {
        let date_ptr =
            register_variant_pointer(&Variant::from_date_f64(42.25)).expect("register date");
        let date_raw = lookup_pointer(date_ptr)
            .expect("lookup date pointer")
            .cast::<f64>();
        assert!(!date_raw.is_null());
        assert_eq!(unsafe { *date_raw }, 42.25);

        let array_ptr = register_variant_pointer(&Variant::from_safearray(
            crate::safe_array::SafeArray::from_variants(vec![
                Variant::from_u8(1),
                Variant::from_bool(false),
                Variant::from_i32(3),
            ]),
        ))
        .expect("register variant array");
        let array_raw = lookup_pointer(array_ptr)
            .expect("lookup array pointer")
            .cast::<u8>();
        assert!(!array_raw.is_null());
        let bytes = unsafe { std::slice::from_raw_parts(array_raw, 3) };
        assert_eq!(bytes, &[1, 0, 3]);
    }

    #[test]
    fn scalar_pointer_readback_uses_declared_storage_widths() {
        let long_ptr = register_variant_pointer(&Variant::from_i32(10)).expect("register long");
        let long_raw = lookup_pointer(long_ptr)
            .expect("lookup long pointer")
            .cast::<i32>();
        unsafe {
            *long_raw = 0x11223344;
        }
        assert_eq!(
            super::read_back_scalar_payload_variant(long_ptr, ScalarPointerKind::Long)
                .expect("read back long"),
            Variant::from_i32(0x11223344)
        );

        let int_ptr = register_variant_pointer(&Variant::from_i16(7)).expect("register integer");
        let int_raw = lookup_pointer(int_ptr)
            .expect("lookup integer pointer")
            .cast::<i16>();
        unsafe {
            *int_raw = -1234;
        }
        assert_eq!(
            super::read_back_scalar_payload_variant(int_ptr, ScalarPointerKind::Integer)
                .expect("read back integer"),
            Variant::from_i16(-1234)
        );

        let single_ptr =
            register_variant_pointer(&Variant::from_f32(1.5)).expect("register single");
        let single_raw = lookup_pointer(single_ptr)
            .expect("lookup single pointer")
            .cast::<f32>();
        unsafe {
            *single_raw = 2.25;
        }
        assert_eq!(
            super::read_back_scalar_payload_variant(single_ptr, ScalarPointerKind::Single)
                .expect("read back single"),
            Variant::from_f32(2.25)
        );
        free_pins(&[long_ptr, int_ptr, single_ptr]);
    }

    #[cfg(target_os = "windows")]
    #[test]
    fn string_variant_pointer_exposes_bstr_cell_not_payload() {
        let ptr = register_string_variant_pointer(&Variant::from_string(BStr::from("abc")))
            .expect("register string variant var");
        assert_ne!(ptr, 0);
        let raw = lookup_pointer(ptr)
            .expect("lookup string variant var")
            .cast::<usize>();
        assert!(!raw.is_null());
        let payload = unsafe { *raw as *const u16 };
        assert!(!payload.is_null());
        let len = unsafe { SysStringLen(payload.cast_mut()) };
        assert_eq!(len, 3);
    }

    #[cfg(target_os = "windows")]
    #[test]
    fn string_var_pointer_readback_tracks_updated_bstr_cell() {
        let ptr = register_string_variant_pointer(&Variant::from_string(BStr::from("abc")))
            .expect("register string var");
        let raw = lookup_pointer(ptr)
            .expect("lookup string var")
            .cast::<BSTR>();
        assert!(!raw.is_null());
        let old_payload = unsafe { *raw };
        if !old_payload.is_null() {
            unsafe { SysFreeString(old_payload) };
        }
        let replacement = BStr::from("alpha").owned_core();
        let replacement_len = u32::try_from(replacement.len_code_units()).expect("u32 length");
        unsafe {
            *raw = SysAllocStringLen(replacement.payload_ptr(), replacement_len);
        }
        let variant =
            super::read_back_string_payload_variant(ptr).expect("read back updated string var");
        assert_eq!(variant, Variant::from_string(BStr::from("alpha")));
    }

    #[cfg(target_os = "windows")]
    #[test]
    fn string_var_pointer_preserves_embedded_nul_bstr_payload() {
        let original = Variant::from_string(BStr::from("abc"));
        let ptr = register_string_variant_pointer(&original).expect("register string var");
        let raw = lookup_pointer(ptr)
            .expect("lookup string var")
            .cast::<BSTR>();
        assert!(!raw.is_null());
        let old_payload = unsafe { *raw };
        if !old_payload.is_null() {
            unsafe { SysFreeString(old_payload) };
        }
        let replacement = BStr::from("a\0bc").owned_core();
        let replacement_len = u32::try_from(replacement.len_code_units()).expect("u32 length");
        unsafe {
            *raw = SysAllocStringLen(replacement.payload_ptr(), replacement_len);
        }
        let payload = unsafe { *raw };
        let len = unsafe { SysStringLen(payload) };
        assert_eq!(len, 4);
        let variant =
            super::read_back_string_payload_variant(ptr).expect("read back embedded nul string");
        assert_eq!(variant, Variant::from_string(BStr::from("a\0bc")));
    }

    #[cfg(target_os = "windows")]
    #[test]
    fn string_var_pointer_readback_preserves_odd_byte_bstr_payload() {
        let original = Variant::from_string(BStr::from("abc"));
        let ptr = register_string_variant_pointer(&original).expect("register string var");
        let raw = lookup_pointer(ptr)
            .expect("lookup string var")
            .cast::<BSTR>();
        assert!(!raw.is_null());
        let old_payload = unsafe { *raw };
        if !old_payload.is_null() {
            unsafe { SysFreeString(old_payload) };
        }
        let bytes = [0x00, 0x43, 0x00];
        unsafe {
            *raw = SysAllocStringByteLen(bytes.as_ptr(), bytes.len() as u32);
        }
        let variant =
            super::read_back_string_payload_variant(ptr).expect("read back odd-byte string");
        assert_eq!(variant.string_byte_len(), Some(3));
        assert_eq!(variant.string_bytes(), Some(bytes.to_vec()));
        assert_eq!(variant.string_units(), Some(vec![0x4300]));
    }

    #[test]
    fn byte_buffer_pointer_readback_exposes_variant_array_payload() {
        let ptr = super::register_byte_buffer(vec![1, 2, 3]).expect("register byte buffer");
        let variant =
            super::read_back_byte_array_payload_variant(ptr).expect("read back byte buffer");
        let elements = variant
            .as_safearray()
            .expect("variant should carry SAFEARRAY")
            .variant_elements()
            .expect("SAFEARRAY should expose variant elements");
        assert_eq!(
            elements,
            vec![
                Variant::from_u8(1),
                Variant::from_u8(2),
                Variant::from_u8(3)
            ]
        );
        assert_eq!(
            Variant::from_safearray(crate::safe_array::SafeArray::from_variants(elements)),
            variant
        );
    }

    #[cfg(target_os = "windows")]
    #[test]
    fn windows_variant_pointer_bstr_balance() {
        let first_source = Variant::from_string(BStr::from("first\0source"));
        let second_source = Variant::from_string(BStr::from("second\0source"));
        let before = thread_live_handle_counts();

        let first = register_variant_var_variant_pointer(&first_source).expect("first Variant pin");
        let second =
            register_variant_var_variant_pointer(&second_source).expect("second Variant pin");
        assert_ne!(first, second, "simultaneous Variant pin IDs must differ");
        let first_cell = lookup_pointer(first)
            .expect("first Variant cell")
            .cast::<VARIANT>();
        let second_cell = lookup_pointer(second)
            .expect("second Variant cell")
            .cast::<VARIANT>();
        assert_ne!(first_cell, second_cell, "Variant cells must not alias");
        assert_eq!(
            before.balance_to(thread_live_handle_counts()).bstrs,
            2,
            "each live cell must own one accounting token"
        );

        // SAFETY: both pointers are distinct live registry-owned VARIANT cells.
        // Clear the first to a non-BSTR and replace the second with an
        // independently allocated native BSTR to exercise pointer-independent
        // accounting for both mutations.
        unsafe {
            assert!(VariantClear(first_cell) >= 0);
            (*first_cell).Anonymous.Anonymous.vt = VT_NULL;

            assert!(VariantClear(second_cell) >= 0);
            let replacement_units: Vec<u16> = "native\0replacement".encode_utf16().collect();
            let replacement = SysAllocStringLen(
                replacement_units.as_ptr(),
                u32::try_from(replacement_units.len()).expect("replacement length"),
            );
            assert!(!replacement.is_null());
            (*second_cell).Anonymous.Anonymous.Anonymous.bstrVal = replacement;
            (*second_cell).Anonymous.Anonymous.vt = VT_BSTR;
        }

        free_pins(&[first, first]);
        assert!(lookup_pointer(first).is_none());
        assert_eq!(
            lookup_pointer(second).map(|pointer| pointer.cast::<VARIANT>()),
            Some(second_cell),
            "freeing one pin must not invalidate the other"
        );
        assert_eq!(
            before.balance_to(thread_live_handle_counts()).bstrs,
            1,
            "the second accounting token must remain live"
        );

        free_pins(&[second, second]);
        free_pins(&[first, second]);
        let drift = before.balance_to(thread_live_handle_counts());
        assert!(
            drift.is_zero(),
            "all Variant pin carriers must balance: {drift:?}"
        );
    }

    #[cfg(target_os = "windows")]
    #[test]
    fn variant_var_pointer_materializes_variant_container() {
        let ptr = register_variant_var_variant_pointer(&Variant::from_string(BStr::from("abc")))
            .expect("register variant");
        assert_ne!(ptr, 0);
        let raw = lookup_pointer(ptr)
            .expect("lookup variant")
            .cast::<VARIANT>();
        assert!(!raw.is_null());
        let variant = unsafe { &*raw };
        assert_eq!(unsafe { variant.Anonymous.Anonymous.vt }, VT_BSTR);
        let payload = unsafe { variant.Anonymous.Anonymous.Anonymous.bstrVal };
        let len = unsafe { SysStringLen(payload) };
        assert_eq!(len, 3);

        let int_ptr = register_variant_var_variant_pointer(&Variant::from_i32(42))
            .expect("register i32 variant");
        let int_raw = lookup_pointer(int_ptr)
            .expect("lookup i32 variant")
            .cast::<VARIANT>();
        let int_variant = unsafe { &*int_raw };
        assert_eq!(unsafe { int_variant.Anonymous.Anonymous.vt }, VT_I4);
        assert_eq!(
            unsafe { int_variant.Anonymous.Anonymous.Anonymous.lVal },
            42
        );

        let date_ptr = register_variant_var_variant_pointer(&Variant::from_date_f64(42.25))
            .expect("register date variant");
        let date_raw = lookup_pointer(date_ptr)
            .expect("lookup date variant")
            .cast::<VARIANT>();
        let date_variant = unsafe { &*date_raw };
        assert_eq!(unsafe { date_variant.Anonymous.Anonymous.vt }, VT_DATE);
        assert_eq!(
            unsafe { date_variant.Anonymous.Anonymous.Anonymous.dblVal },
            42.25
        );

        let decimal_ptr = register_variant_var_variant_pointer(&Variant::from_decimal96(
            Decimal96::from_parts(123_450, 0, 0, 3, false),
        ))
        .expect("register decimal variant");
        let decimal_raw = lookup_pointer(decimal_ptr)
            .expect("lookup decimal variant")
            .cast::<u8>();
        let bytes = unsafe { std::slice::from_raw_parts(decimal_raw, 16) };
        let mut wire = [0u8; 16];
        wire.copy_from_slice(bytes);
        let decimal_variant = Variant::from_wire_bytes(wire).expect("decimal compat-slot wire");
        assert_eq!(decimal_variant.vtype(), VarType::Decimal);
        assert_eq!(
            decimal_variant.as_decimal96(),
            Some(Decimal96::from_parts(123_450, 0, 0, 3, false))
        );
        free_pins(&[ptr, int_ptr, date_ptr, decimal_ptr]);
    }

    #[cfg(target_os = "windows")]
    #[test]
    fn variant_var_pointer_accepts_canonical_variant_input() {
        let ptr = register_variant_var_variant_pointer(&Variant::from_string(BStr::from("abc")))
            .expect("register canonical variant");
        assert_ne!(ptr, 0);
        let raw = lookup_pointer(ptr)
            .expect("lookup canonical variant")
            .cast::<VARIANT>();
        assert!(!raw.is_null());
        let variant = unsafe { &*raw };
        assert_eq!(unsafe { variant.Anonymous.Anonymous.vt }, VT_BSTR);
        let payload = unsafe { variant.Anonymous.Anonymous.Anonymous.bstrVal };
        let len = unsafe { SysStringLen(payload) };
        assert_eq!(len, 3);
        free_pins(&[ptr]);
    }

    #[test]
    fn object_pointer_requires_object_like_value() {
        assert_eq!(
            register_object_variant_pointer(&Variant::from_object_ref(
                ObjectRef::from_compat_identity(0)
            ))
            .expect("nothing"),
            0
        );
        assert!(register_object_variant_pointer(&Variant::from_i32(5)).is_err());
    }

    #[cfg(target_os = "windows")]
    #[test]
    fn object_pointer_returns_raw_iunknown_address() {
        let object = ObjectRef::from_compat_identity(42);
        let pointer = register_object_variant_pointer(&Variant::from_object_ref(object.clone()))
            .expect("object pointer");
        assert_eq!(pointer, object.raw_iunknown() as usize as i64);
    }

    #[cfg(target_os = "windows")]
    #[test]
    fn variant_var_pointer_supports_object_payload() {
        let object = ObjectRef::from_compat_identity(42);
        let ptr = register_variant_var_variant_pointer(&Variant::from_object_ref(object.clone()))
            .expect("register object-valued variant");
        let raw = lookup_pointer(ptr)
            .expect("lookup variant var")
            .cast::<VARIANT>();
        assert!(!raw.is_null());
        let variant = unsafe { &*raw };
        assert_eq!(unsafe { variant.Anonymous.Anonymous.vt }, VT_UNKNOWN);
        assert_eq!(
            unsafe { variant.Anonymous.Anonymous.Anonymous.punkVal },
            object.raw_iunknown().cast()
        );
        free_pins(&[ptr]);
    }

    #[cfg(target_os = "windows")]
    #[test]
    fn variant_var_pointer_supports_array_payload() {
        let ptr = register_variant_var_variant_pointer(&Variant::from_safearray(
            crate::safe_array::SafeArray::from_variants(vec![
                Variant::from_i32(4),
                Variant::from_string(BStr::from("abc")),
            ]),
        ))
        .expect("register array-valued variant");
        let raw = lookup_pointer(ptr)
            .expect("lookup variant var")
            .cast::<VARIANT>();
        assert!(!raw.is_null());
        let variant = unsafe { &*raw };
        assert_eq!(
            unsafe { variant.Anonymous.Anonymous.vt },
            VT_ARRAY | VT_VARIANT
        );
        let psa = unsafe { variant.Anonymous.Anonymous.Anonymous.parray };
        assert!(!psa.is_null());
        assert_eq!(unsafe { SafeArrayGetDim(psa) }, 1);

        let mut first: VARIANT = unsafe { std::mem::zeroed() };
        let index = 0i32;
        let hr = unsafe {
            SafeArrayGetElement(
                psa.cast_const(),
                &index,
                (&mut first as *mut VARIANT).cast(),
            )
        };
        assert!(hr >= 0);
        assert_eq!(unsafe { first.Anonymous.Anonymous.vt }, VT_I4);
        assert_eq!(unsafe { first.Anonymous.Anonymous.Anonymous.lVal }, 4);
        unsafe {
            VariantClear(&mut first);
        }
        free_pins(&[ptr]);
    }
}
