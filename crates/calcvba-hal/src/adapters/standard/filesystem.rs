use crate::{
    error::{HalError, HalResult},
    model::CapabilityId,
    traits::FileSystemHal,
};
use calcvba_runtime::ansi::{ansi_decode, ansi_encode};
use calcvba_runtime::vba_record::VbaRecordFieldKind;
use calcvba_runtime::{VarType, Variant, bstr::BStr, vba_date::ymd_to_serial};
use std::collections::{BTreeMap, BTreeSet};
use std::fs::{self, OpenOptions};
use std::path::{Component, Path, PathBuf};
#[cfg(target_os = "windows")]
use std::{thread, time::Duration};

use super::{StandardHostServices, time::system_time_to_local_vba_serial};

// VBA file-attribute bits. These are deliberately identical to the Win32
// `FILE_ATTRIBUTE_*` values, which is what `GetAttr`/`SetAttr` expose.
const VB_READONLY: i32 = 0x0001;
const VB_HIDDEN: i32 = 0x0002;
const VB_SYSTEM: i32 = 0x0004;
const VB_DIRECTORY: i32 = 0x0010;
const VB_ARCHIVE: i32 = 0x0020;
/// The bits `SetAttr` is permitted to change (directory/volume are intrinsic).
const VB_SETTABLE_ATTRS: i32 = VB_READONLY | VB_HIDDEN | VB_SYSTEM | VB_ARCHIVE;

/// Attribute bits used when filtering `Dir` entries.
fn file_attribute_bits(meta: &fs::Metadata) -> i32 {
    #[cfg(target_os = "windows")]
    {
        use std::os::windows::fs::MetadataExt;
        (meta.file_attributes() as i32)
            & (VB_READONLY | VB_HIDDEN | VB_SYSTEM | VB_DIRECTORY | VB_ARCHIVE)
    }
    #[cfg(not(target_os = "windows"))]
    {
        // Only read-only and directory are observable through portable std.
        let mut bits = 0;
        if meta.permissions().readonly() {
            bits |= VB_READONLY;
        }
        if meta.is_dir() {
            bits |= VB_DIRECTORY;
        }
        bits
    }
}

/// `GetAttr` preserves flags absent from Rust's handle-based metadata result.
fn path_attribute_bits(path: &Path, mapped: bool) -> std::io::Result<i32> {
    #[cfg(windows)]
    {
        let _ = mapped;
        use std::os::windows::ffi::OsStrExt;
        use windows_sys::Win32::Storage::FileSystem::{
            GetFileAttributesW, INVALID_FILE_ATTRIBUTES,
        };
        let wide: Vec<u16> = path.as_os_str().encode_wide().chain(Some(0)).collect();
        // SAFETY: wide is a NUL-terminated path and remains live for the call.
        let attrs = unsafe { GetFileAttributesW(wide.as_ptr()) };
        if attrs == INVALID_FILE_ATTRIBUTES {
            Err(std::io::Error::last_os_error())
        } else {
            Ok((attrs & !0x80) as i32)
        }
    }
    #[cfg(not(windows))]
    {
        let meta = fs::metadata(path)?;
        let bits = file_attribute_bits(&meta);
        if mapped && let Some(windows_bits) = read_mapped_attributes(path)? {
            return Ok(bits | windows_bits);
        }
        Ok(if mapped && meta.is_file() {
            bits | VB_ARCHIVE
        } else {
            bits
        })
    }
}

/// Apply the settable attribute bits to `path`, preserving the bits `SetAttr`
/// does not manage (e.g. `vbDirectory`). `settable` is pre-validated to contain
/// only [`VB_SETTABLE_ATTRS`].
fn apply_file_attributes(path: &Path, settable: i32, mapped: bool) -> std::io::Result<()> {
    #[cfg(target_os = "windows")]
    {
        let _ = mapped;
        use std::os::windows::ffi::OsStrExt;
        use windows_sys::Win32::Storage::FileSystem::{
            FILE_ATTRIBUTE_NORMAL, GetFileAttributesW, INVALID_FILE_ATTRIBUTES, SetFileAttributesW,
        };

        let wide: Vec<u16> = path
            .as_os_str()
            .encode_wide()
            .chain(std::iter::once(0))
            .collect();
        // SAFETY: `wide` is a NUL-terminated UTF-16 buffer that outlives the call;
        // `GetFileAttributesW` only reads it and returns the attribute DWORD (or
        // INVALID_FILE_ATTRIBUTES on failure, which we surface as an OS error).
        let current = unsafe { GetFileAttributesW(wide.as_ptr()) };
        if current == INVALID_FILE_ATTRIBUTES {
            return Err(std::io::Error::last_os_error());
        }
        let preserved = current & !(VB_SETTABLE_ATTRS as u32);
        let mut new_attrs = preserved | (settable as u32);
        if new_attrs == 0 {
            new_attrs = FILE_ATTRIBUTE_NORMAL;
        }
        // SAFETY: same NUL-terminated UTF-16 buffer; `SetFileAttributesW` reads the
        // path and applies `new_attrs`, returning 0 (FALSE) on failure.
        let ok = unsafe { SetFileAttributesW(wide.as_ptr(), new_attrs) };
        if ok == 0 {
            return Err(std::io::Error::last_os_error());
        }
        Ok(())
    }
    #[cfg(not(target_os = "windows"))]
    {
        let meta = fs::metadata(path)?;
        let original = meta.permissions();
        let mut perms = original.clone();
        let readonly = settable & VB_READONLY != 0;
        if mapped {
            let temporary_write =
                original.readonly() && cfg!(any(target_os = "linux", target_os = "macos"));
            if temporary_write {
                // On Unix the inode must be writable to change a user xattr.
                make_owner_writable(&mut perms);
                fs::set_permissions(path, perms.clone())?;
            }
            let requested = settable & (VB_HIDDEN | VB_SYSTEM | VB_ARCHIVE);
            let default = if meta.is_file() { VB_ARCHIVE } else { 0 };
            if let Err(error) = write_mapped_attributes(path, requested)
                && (error.kind() != std::io::ErrorKind::Unsupported || requested != default)
            {
                if temporary_write {
                    fs::set_permissions(path, original)?;
                }
                return Err(error);
            }
            if !readonly && perms.readonly() {
                make_owner_writable(&mut perms);
            } else if readonly {
                perms.set_readonly(true);
            }
        } else {
            perms.set_readonly(readonly);
        }
        fs::set_permissions(path, perms)
    }
}

#[cfg(not(windows))]
fn make_owner_writable(permissions: &mut fs::Permissions) {
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        permissions.set_mode(permissions.mode() | 0o200);
    }
    #[cfg(not(unix))]
    permissions.set_readonly(false);
}

#[cfg(any(target_os = "linux", target_os = "macos"))]
fn read_mapped_attributes(path: &Path) -> std::io::Result<Option<i32>> {
    use std::{ffi::CString, os::unix::ffi::OsStrExt};
    let c_path = CString::new(path.as_os_str().as_bytes())?;
    let mut bitset = 0u8;
    // SAFETY: path/name are terminated strings and bitset is a live one-byte buffer.
    #[cfg(target_os = "linux")]
    let read = unsafe {
        libc::getxattr(
            c_path.as_ptr(),
            c"user.calcvba.attributes".as_ptr(),
            (&raw mut bitset).cast(),
            1,
        )
    };
    // SAFETY: path/name are terminated strings and bitset is a live one-byte buffer.
    #[cfg(target_os = "macos")]
    let read = unsafe {
        libc::getxattr(
            c_path.as_ptr(),
            c"org.dnacalc.calcvba.attributes".as_ptr(),
            (&raw mut bitset).cast(),
            1,
            0,
            0,
        )
    };
    #[cfg(target_os = "linux")]
    let missing_attribute = libc::ENODATA;
    #[cfg(target_os = "macos")]
    let missing_attribute = libc::ENOATTR;
    match read {
        1 => Ok(Some(
            i32::from(bitset) & (VB_HIDDEN | VB_SYSTEM | VB_ARCHIVE),
        )),
        -1 => {
            let error = std::io::Error::last_os_error();
            #[cfg(target_os = "linux")]
            if error.kind() == std::io::ErrorKind::Unsupported {
                return super::windows_attributes::read(path);
            }
            if error.raw_os_error() == Some(missing_attribute)
                || error.kind() == std::io::ErrorKind::Unsupported
            {
                Ok(None)
            } else {
                Err(error)
            }
        }
        _ => Err(std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            "invalid mapped file attributes",
        )),
    }
}

#[cfg(not(any(windows, target_os = "linux", target_os = "macos")))]
fn read_mapped_attributes(_path: &Path) -> std::io::Result<Option<i32>> {
    Ok(None)
}

#[cfg(any(target_os = "linux", target_os = "macos"))]
fn write_mapped_attributes(path: &Path, bitset: i32) -> std::io::Result<()> {
    use std::{ffi::CString, os::unix::ffi::OsStrExt};
    let c_path = CString::new(path.as_os_str().as_bytes())?;
    let bitset = bitset as u8;
    // SAFETY: path/name are terminated strings and bitset is a live one-byte buffer.
    #[cfg(target_os = "linux")]
    let written = unsafe {
        libc::setxattr(
            c_path.as_ptr(),
            c"user.calcvba.attributes".as_ptr(),
            (&raw const bitset).cast(),
            1,
            0,
        )
    };
    // SAFETY: path/name are terminated strings and bitset is a live one-byte buffer.
    #[cfg(target_os = "macos")]
    let written = unsafe {
        libc::setxattr(
            c_path.as_ptr(),
            c"org.dnacalc.calcvba.attributes".as_ptr(),
            (&raw const bitset).cast(),
            1,
            0,
            0,
        )
    };
    if written == 0 {
        Ok(())
    } else {
        let error = std::io::Error::last_os_error();
        #[cfg(target_os = "linux")]
        if error.kind() == std::io::ErrorKind::Unsupported {
            return super::windows_attributes::write(path, bitset);
        }
        Err(error)
    }
}

#[cfg(not(any(windows, target_os = "linux", target_os = "macos")))]
fn write_mapped_attributes(_path: &Path, _bitset: i32) -> std::io::Result<()> {
    Err(std::io::Error::from(std::io::ErrorKind::Unsupported))
}

/// Change the current drive (VBA `ChDrive`). On Windows this selects the drive's
/// remembered working directory; elsewhere there is no drive concept, so it is a
/// no-op.
fn change_drive(letter: char) -> std::io::Result<()> {
    #[cfg(target_os = "windows")]
    {
        // `set_current_dir("C:")` with no trailing separator selects drive C's
        // current directory, which is exactly ChDrive's effect.
        std::env::set_current_dir(format!("{letter}:"))
    }
    #[cfg(not(target_os = "windows"))]
    {
        let _ = letter;
        Ok(())
    }
}

/// The VBA run-time error number for a failed host I/O call: `not_found` for a missing
/// file or path (53 file, 76 path), 70 permission denied, 75 for an existing or
/// non-empty target, else `other`.
fn io_error_code(err: &std::io::Error, not_found: i32, other: i32) -> i32 {
    use std::io::ErrorKind;
    match err.kind() {
        ErrorKind::NotFound => not_found,
        ErrorKind::PermissionDenied => 70,
        ErrorKind::AlreadyExists | ErrorKind::DirectoryNotEmpty => 75,
        _ => other,
    }
}

/// The path-reading functions distinguish a missing file from a missing parent.
fn metadata_error_code(err: &std::io::Error, path: &Path, mapped: bool) -> i32 {
    if path.as_os_str().is_empty() {
        return 53;
    }
    #[cfg(windows)]
    match err.raw_os_error() {
        Some(3) => return 76,   // ERROR_PATH_NOT_FOUND
        Some(123) => return 52, // ERROR_INVALID_NAME (including wildcards)
        _ => {}
    }
    if mapped
        && matches!(
            err.kind(),
            std::io::ErrorKind::NotFound | std::io::ErrorKind::NotADirectory
        )
        && path.parent().is_some_and(|parent| !parent.is_dir())
    {
        return 76;
    }
    io_error_code(err, 53, 53)
}

fn directory_error_code(err: &std::io::Error) -> i32 {
    #[cfg(windows)]
    if matches!(err.raw_os_error(), Some(123 | 267)) {
        return 76; // invalid name / not a directory
    }
    if err.kind() == std::io::ErrorKind::PermissionDenied {
        75
    } else {
        io_error_code(err, 76, 75)
    }
}

#[derive(Debug, Default)]
pub(super) struct FileSystemState {
    pub(super) handles: BTreeMap<i32, FileHandleState>,
}

/// Fixed on-disk byte size for a Random/Binary record of `vtype` (None = variable
/// length, i.e. `String`, or unsupported).
fn record_size(vtype: VarType) -> Option<usize> {
    Some(match vtype {
        VarType::Byte => 1,
        VarType::Integer | VarType::Boolean => 2,
        VarType::Long | VarType::Single => 4,
        VarType::Double | VarType::Currency | VarType::Date | VarType::LongLong => 8,
        _ => return None,
    })
}

fn arr<const N: usize>(b: &[u8]) -> [u8; N] {
    let mut out = [0u8; N];
    let n = b.len().min(N);
    out[..n].copy_from_slice(&b[..n]);
    out
}

/// VBA `Open` mode codes (`Input`=0, `Output`=1 are the sequential `_` cases).
const MODE_APPEND: i32 = 2;
const MODE_BINARY: i32 = 3;
const MODE_RANDOM: i32 = 4;

/// Serialize a value to its VBA on-disk byte representation (little-endian for
/// numerics). A `String` is encoded as ANSI (system code page): a fixed-length
/// string (`fixed`) is raw with no prefix in any mode; a variable string carries
/// a 2-byte length prefix in Random mode and is raw in Binary mode.
fn serialize_record(value: &Variant, mode: i32, fixed: bool) -> Vec<u8> {
    match value.vtype() {
        VarType::Byte => vec![value.as_u8().unwrap_or(0)],
        VarType::Integer => value.as_i16().unwrap_or(0).to_le_bytes().to_vec(),
        VarType::Boolean => (if value.as_bool().unwrap_or(false) {
            -1i16
        } else {
            0
        })
        .to_le_bytes()
        .to_vec(),
        VarType::Long => value.as_i32().unwrap_or(0).to_le_bytes().to_vec(),
        VarType::Single => value.as_f32().unwrap_or(0.0).to_le_bytes().to_vec(),
        VarType::Double => value.as_f64().unwrap_or(0.0).to_le_bytes().to_vec(),
        VarType::Currency => value
            .as_currency_scaled_i64()
            .unwrap_or(0)
            .to_le_bytes()
            .to_vec(),
        VarType::Date => value.as_date_f64().unwrap_or(0.0).to_le_bytes().to_vec(),
        VarType::LongLong => value.as_i64().unwrap_or(0).to_le_bytes().to_vec(),
        VarType::String => {
            let s = value
                .as_bstr()
                .map(|b| b.as_str().to_string())
                .unwrap_or_default();
            let ansi = calcvba_com::ansi::ansi_encode(&s);
            if !fixed && mode == MODE_RANDOM {
                // Variable-length string in a Random record: 2-byte length prefix.
                let len = ansi.len().min(u16::MAX as usize);
                let mut out = (len as u16).to_le_bytes().to_vec();
                out.extend_from_slice(&ansi[..len]);
                out
            } else {
                // Fixed-length string (already padded to N), or Binary mode: raw ANSI.
                ansi
            }
        }
        _ => Vec::new(),
    }
}

/// Decode a fixed-size record of `vtype` from `bytes` (the caller has already
/// sliced to the record's bytes; for strings `bytes` is the ANSI payload).
fn deserialize_record(vtype: VarType, bytes: &[u8]) -> Variant {
    match vtype {
        VarType::Byte => Variant::from_u8(bytes.first().copied().unwrap_or(0)),
        VarType::Integer => Variant::from_i16(i16::from_le_bytes(arr(bytes))),
        VarType::Boolean => Variant::from_bool(i16::from_le_bytes(arr(bytes)) != 0),
        VarType::Long => Variant::from_i32(i32::from_le_bytes(arr(bytes))),
        VarType::Single => Variant::from_f32(f32::from_le_bytes(arr(bytes))),
        VarType::Double => Variant::from_f64(f64::from_le_bytes(arr(bytes))),
        VarType::Currency => Variant::from_currency_scaled_i64(i64::from_le_bytes(arr(bytes))),
        VarType::Date => Variant::from_date_f64(f64::from_le_bytes(arr(bytes))),
        VarType::LongLong => Variant::from_i64(i64::from_le_bytes(arr(bytes))),
        VarType::String => Variant::from_string(calcvba_com::ansi::ansi_decode(bytes)),
        _ => Variant::empty(),
    }
}

/// Fixed UDT file storage is packed field-by-field, without native alignment
/// padding. Fixed strings contain ANSI bytes rather than the native UTF-16.
fn serialize_udt_field(kind: &VbaRecordFieldKind, value: &Variant) -> Result<Vec<u8>, String> {
    use VbaRecordFieldKind as F;
    let size = kind.file_len()?;
    let mut bytes = Vec::with_capacity(size);
    match kind {
        F::Record(layout) => {
            for (index, field) in layout.fields().iter().enumerate() {
                bytes.extend(serialize_udt_field(
                    field.kind(),
                    &value.read_record_field_variant(index)?,
                )?);
            }
        }
        F::FixedArray { element, .. } => {
            let array = value.as_safearray().ok_or("expected fixed array field")?;
            for index in 0..array.len() {
                bytes.extend(serialize_udt_field(
                    element,
                    &array.variant_element(index)?,
                )?);
            }
        }
        F::FixedString { .. } => {
            bytes = serialize_record(value, MODE_BINARY, true);
            bytes.resize(size, 0);
            bytes.truncate(size);
        }
        _ => bytes = serialize_record(value, MODE_BINARY, false),
    }
    Ok(bytes)
}

fn deserialize_udt_field(
    kind: &VbaRecordFieldKind,
    shape: &Variant,
    bytes: &[u8],
) -> Result<Variant, String> {
    use VbaRecordFieldKind as F;
    match kind {
        F::Record(layout) => {
            let mut record = shape.clone();
            let mut offset = 0;
            for (index, field) in layout.fields().iter().enumerate() {
                let end = offset + field.kind().file_len()?;
                let value = deserialize_udt_field(
                    field.kind(),
                    &shape.read_record_field_variant(index)?,
                    &bytes[offset..end],
                )?;
                record.write_record_field_variant(index, &value)?;
                offset = end;
            }
            Ok(record)
        }
        F::FixedArray { element, .. } => {
            let mut array = shape.as_safearray().ok_or("expected fixed array field")?;
            let size = element.file_len()?;
            for index in 0..array.len() {
                let offset = index * size;
                let value = deserialize_udt_field(
                    element,
                    &array.variant_element(index)?,
                    &bytes[offset..offset + size],
                )?;
                array.set_variant_element(index, &value)?;
            }
            Ok(Variant::from_safearray(array))
        }
        _ => Ok(deserialize_record(shape.vtype(), bytes)),
    }
}

impl FileSystemState {
    pub(super) fn first_free_in(&self, start: i32, end: i32) -> Option<i32> {
        let in_use: BTreeSet<i32> = self.handles.keys().copied().collect();
        (start..=end).find(|candidate| !in_use.contains(candidate))
    }

    pub(super) fn is_handle_in_use(&self, handle: i32) -> bool {
        self.handles.contains_key(&handle)
    }
}

#[derive(Debug, Clone)]
pub(super) struct FileHandleState {
    pub(super) mode: i32,
    pub(super) position: i32,
    input_buffer_origin: i32,
    input_text_eof: bool,
    pub(super) len: i32,
    pub(super) host_path: Option<PathBuf>,
    pub(super) data: Vec<u8>,
    modified: bool,
    /// Output line width set by `Width #n` (`0` disables wrapping).
    pub(super) width: i32,
    /// Zero-based print column for sequential formatted output.
    pub(super) print_column: i32,
    /// Locked 1-based byte/record ranges (inclusive) for `Lock`/`Unlock`. A `Lock`
    /// overlapping an already-locked range raises VBA error 70.
    pub(super) locks: Vec<(i32, i32)>,
    /// Fixed record length from `Open … For Random … Len = N` (0 = unset; Random
    /// records then use the written value's natural size as the stride).
    pub(super) record_len: i32,
    /// The last `Get` could not read a whole record: `EOF` of a Binary or Random file.
    get_hit_eof: bool,
}

/// The 1-based position VBA `Seek(filenumber)` (FUNCTION form) reports: the next
/// byte to read/write (Binary/sequential) or the next record (Random). The
/// internal `position` is a 0-based byte cursor, so the byte form is `cursor + 1`;
/// Random divides by the fixed record length to report a 1-based record number.
fn seek_report(entry: &FileHandleState) -> i32 {
    if entry.mode == MODE_RANDOM && entry.record_len > 0 {
        entry.position / entry.record_len + 1
    } else {
        entry.position.saturating_add(1)
    }
}

/// The value VBA `Loc(filenumber)` reports: the number of the last record
/// read/written (Random); the 1-based offset of the last byte (Binary), which is
/// exactly the 0-based next-byte cursor. Input rounds up to 128-byte blocks,
/// with a minimum of one; Output/Append count complete blocks (Excel P_Loc).
fn loc_report(entry: &FileHandleState) -> i32 {
    match entry.mode {
        MODE_RANDOM if entry.record_len > 0 => entry.position / entry.record_len,
        MODE_BINARY | MODE_RANDOM => entry.position,
        0 => (entry.position.saturating_add(127) / 128).max(1),
        _ => entry.position / 128,
    }
}

/// The byte offset a sequential write (`Print #`/`Write #`) starts at. In
/// `Append` mode every write lands at the current end-of-file regardless of the
/// reported seek cursor (VBA keeps the two decoupled); the other modes write at
/// the 0-based cursor.
fn seq_write_offset(entry: &FileHandleState) -> usize {
    if entry.mode == MODE_APPEND {
        entry.len.max(0) as usize
    } else {
        entry.position.max(0) as usize
    }
}

fn print_column_after_text(start_column: i32, text: &str) -> i32 {
    let mut col = start_column.max(0) as usize;
    for ch in text.chars() {
        match ch {
            '\r' | '\n' => col = 0,
            _ => col = col.saturating_add(1),
        }
    }
    col.min(i32::MAX as usize) as i32
}

fn line_column_at(data: &[u8], cursor: usize) -> i32 {
    let upto = cursor.min(data.len());
    let mut col = 0usize;
    for ch in ansi_decode(&data[..upto]).chars() {
        match ch {
            '\r' | '\n' => col = 0,
            _ => col = col.saturating_add(1),
        }
    }
    if cursor > data.len() {
        col = col.saturating_add(cursor - data.len());
    }
    col.min(i32::MAX as usize) as i32
}

fn initial_print_column(mode: i32, data: &[u8], position: i32) -> i32 {
    // Append starts a new Print column counter even when the existing file
    // ends with an unterminated line. Writes still append at EOF.
    if mode == MODE_APPEND {
        0
    } else {
        line_column_at(data, position.max(0) as usize)
    }
}

pub(super) fn pseudo_file_len_from_path_token(path: i32) -> i32 {
    let magnitude = path.saturating_abs();
    1 + (magnitude % 4096)
}

pub(super) fn clamp_u64_to_i32(value: u64) -> i32 {
    value.min(i32::MAX as u64) as i32
}

fn parse_input_field(data: &[u8], mut cursor: usize) -> (String, usize) {
    let len = data.len();
    while cursor < len && matches!(data[cursor], b' ' | b'\t') {
        cursor += 1;
    }
    if cursor >= len {
        return (String::new(), cursor);
    }

    if data[cursor] == b'"' {
        // Input # ends the field at the next quote. A second quote starts another field.
        let start = cursor + 1;
        let end = data[start..]
            .iter()
            .position(|&byte| byte == b'"')
            .map_or(len, |offset| start + offset);
        (ansi_decode(&data[start..end]), (end + 1).min(len))
    } else {
        let start = cursor;
        while cursor < len && !matches!(data[cursor], b',' | b'\r' | b'\n') {
            cursor += 1;
        }
        (ansi_decode(&data[start..cursor]).trim().to_owned(), cursor)
    }
}

fn variant_from_input_field(field: &str) -> Variant {
    if field.eq_ignore_ascii_case("#TRUE#") {
        return Variant::from_bool(true);
    }
    if field.eq_ignore_ascii_case("#FALSE#") {
        return Variant::from_bool(false);
    }
    if field.eq_ignore_ascii_case("#NULL#") {
        return Variant::null();
    }
    if let Some(value) = parse_write_date_field(field) {
        return value;
    }
    if let Ok(value) = field.parse::<i32>() {
        return Variant::from_i32(value);
    }
    if let Ok(value) = field.parse::<f64>() {
        return Variant::from_f64(value);
    }
    Variant::from_string(field)
}

fn parse_write_date_field(field: &str) -> Option<Variant> {
    let inner = field.strip_prefix('#')?.strip_suffix('#')?;
    let (date_part, time_part) = inner.split_once(' ').unwrap_or((inner, ""));
    let mut date_bits = date_part.split('-');
    let y = date_bits.next()?.parse::<i64>().ok()?;
    let m = date_bits.next()?.parse::<i64>().ok()?;
    let d = date_bits.next()?.parse::<i64>().ok()?;
    if date_bits.next().is_some() || !(100..=9999).contains(&y) {
        return None;
    }
    let time = if time_part.is_empty() {
        0.0
    } else {
        let mut time_bits = time_part.split(':');
        let h = time_bits.next()?.parse::<i64>().ok()?;
        let mi = time_bits.next()?.parse::<i64>().ok()?;
        let s = time_bits.next()?.parse::<i64>().ok()?;
        if time_bits.next().is_some()
            || !(0..=23).contains(&h)
            || !(0..=59).contains(&mi)
            || !(0..=59).contains(&s)
        {
            return None;
        }
        (h * 3600 + mi * 60 + s) as f64 / 86_400.0
    };
    Some(Variant::from_date_f64(ymd_to_serial(y, m, d) + time))
}

fn advance_input_separator(data: &[u8], mut cursor: usize) -> usize {
    let len = data.len();
    while cursor < len && matches!(data[cursor], b' ' | b'\t') {
        cursor += 1;
    }
    if cursor < len && data[cursor] == b',' {
        cursor += 1;
    } else if cursor < len && data[cursor] == b'\r' {
        cursor += 1;
        if cursor < len && data[cursor] == b'\n' {
            cursor += 1;
        }
    } else if cursor < len && data[cursor] == b'\n' {
        cursor += 1;
    }
    cursor
}

#[cfg(target_os = "windows")]
fn remove_file_with_retry(host_path: &PathBuf) -> std::io::Result<()> {
    const MAX_ATTEMPTS: usize = 8;
    const RETRY_DELAY_MS: u64 = 25;

    let mut last_err = None;
    for attempt in 0..MAX_ATTEMPTS {
        match fs::remove_file(host_path) {
            Ok(()) => return Ok(()),
            Err(err) if err.kind() == std::io::ErrorKind::PermissionDenied => {
                last_err = Some(err);
                if attempt + 1 < MAX_ATTEMPTS {
                    thread::sleep(Duration::from_millis(RETRY_DELAY_MS));
                    continue;
                }
            }
            Err(err) => return Err(err),
        }
    }
    Err(last_err.unwrap_or_else(|| std::io::Error::other("remove_file retry loop exhausted")))
}

fn same_host_path(left: &Path, right: &Path) -> bool {
    match (fs::canonicalize(left), fs::canonicalize(right)) {
        (Ok(left), Ok(right)) => {
            #[cfg(windows)]
            {
                left.to_string_lossy()
                    .eq_ignore_ascii_case(&right.to_string_lossy())
            }
            #[cfg(not(windows))]
            {
                left == right
            }
        }
        _ => false,
    }
}

impl StandardHostServices {
    fn flush_host_file(&self, entry: &FileHandleState) -> HalResult<()> {
        let Some(path) = entry.host_path.as_ref().filter(|_| entry.mode != 0) else {
            return Ok(());
        };
        let flush = || -> std::io::Result<()> {
            fs::write(path, &entry.data)?;
            #[cfg(not(windows))]
            if entry.modified
                && self.uses_windows_paths()
                && let Some(bits) = read_mapped_attributes(path)?
                && bits & VB_ARCHIVE == 0
            {
                write_mapped_attributes(path, bits | VB_ARCHIVE)?;
            }
            Ok(())
        };
        flush().map_err(|error| {
            HalError::adapter_fault(
                self.profile,
                CapabilityId::FileSystemIo,
                "close",
                format!("failed to flush host path {}: {error}", path.display()),
            )
        })
    }

    fn require_file_parent(&self, path: &Path, operation: &'static str) -> HalResult<()> {
        let Some(parent) = path
            .parent()
            .filter(|parent| !parent.as_os_str().is_empty())
        else {
            return Ok(());
        };
        let (number, message) = match fs::metadata(parent) {
            Ok(meta) if meta.is_dir() => return Ok(()),
            Ok(_) => (76, "parent is not a directory".into()),
            Err(error) if error.kind() == std::io::ErrorKind::NotADirectory => {
                (76, error.to_string())
            }
            Err(error) => (directory_error_code(&error), error.to_string()),
        };
        Err(HalError::adapter_fault(
            self.profile,
            CapabilityId::FileSystemIo,
            operation,
            format!("{}: {message}", parent.display()),
        )
        .with_host_error_code(number))
    }

    fn kill_host_file(&self, path: &PathBuf) -> HalResult<()> {
        let capability = CapabilityId::FileSystemIo;
        let fault = |code| {
            HalError::adapter_fault(self.profile, capability, "kill", path.display().to_string())
                .with_host_error_code(code)
        };
        let attrs = path_attribute_bits(path, self.uses_windows_paths()).map_err(|_| fault(53))?;
        if attrs & (VB_HIDDEN | VB_SYSTEM | VB_DIRECTORY) != 0 {
            return Err(fault(53));
        }
        if attrs & VB_READONLY != 0 {
            return Err(fault(75));
        }
        if self
            .borrow_fs(capability, "kill")?
            .handles
            .values()
            .any(|entry| {
                entry
                    .host_path
                    .as_deref()
                    .is_some_and(|open| same_host_path(open, path))
            })
        {
            return Err(fault(55));
        }
        #[cfg(windows)]
        let result = remove_file_with_retry(path);
        #[cfg(not(windows))]
        let result = fs::remove_file(path);
        result.map_err(|err| {
            fault(if err.kind() == std::io::ErrorKind::NotFound {
                53
            } else {
                75
            })
        })
    }

    /// The 1-based inclusive lock range for `Lock`/`Unlock`: no args = whole file,
    /// one arg = a single record, `start To end` = the explicit range.
    fn lock_range(
        &self,
        start: &Variant,
        end: &Variant,
        capability: CapabilityId,
        op: &'static str,
    ) -> HalResult<(i32, i32)> {
        let start = if matches!(start.vtype(), VarType::Empty | VarType::Null) {
            None
        } else {
            Some(self.variant_to_i32(start, capability, op, "start")?)
        };
        let end = if matches!(end.vtype(), VarType::Empty | VarType::Null) {
            None
        } else {
            Some(self.variant_to_i32(end, capability, op, "end")?)
        };
        Ok(match (start, end) {
            (None, _) => (1, i32::MAX),
            (Some(s), None) => (s, s),
            (Some(s), Some(e)) => (s, e),
        })
    }
}

impl FileSystemHal for StandardHostServices {
    fn windows_path_for_host(&self, path: &Path) -> Option<String> {
        #[cfg(windows)]
        if self.native_fs_enabled() && path.is_absolute() {
            return path.to_str().map(str::to_owned);
        }
        #[cfg(not(windows))]
        if self.uses_windows_paths() {
            return StandardHostServices::windows_path_for_host(self, path);
        }
        None
    }

    fn open_variant(&self, path: Variant, mode: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "open"));
        }
        let mode_raw = self.variant_to_i32(&mode, capability, "open", "mode")?;
        let requested_handle = mode_raw >> 16;
        let mode = mode_raw & 0xFFFF;
        if mode != 0 && !self.policy.allow_filesystem_mutation {
            return Err(self.denied(capability, "open"));
        }
        if let Some(path_text) = path.as_bstr() {
            let mut state = self.borrow_fs(capability, "open")?;
            self.assert_fs_invariants(&state, "open-pre");
            let handle = if requested_handle > 0 && requested_handle <= 511 {
                if state.is_handle_in_use(requested_handle) {
                    return Err(HalError::adapter_fault(
                        self.profile,
                        capability,
                        "open",
                        format!("file handle #{requested_handle} is already in use"),
                    )
                    .with_host_error_code(55));
                }
                requested_handle
            } else {
                state.first_free_in(1, 511).ok_or_else(|| {
                    HalError::adapter_fault(
                        self.profile,
                        capability,
                        "open",
                        "no free file handles available in supported range",
                    )
                })?
            };
            let host_path = if self.native_fs_enabled() {
                let host_path =
                    self.map_path(&path_text.as_str(), capability, "open", mode != 0)?;
                self.require_file_parent(&host_path, "open")?;
                Some(host_path)
            } else {
                None
            };
            let (initial_data, initial_len, initial_position) = if let Some(host_path) =
                host_path.as_ref()
            {
                match mode {
                    0 => {
                        let data = fs::read(host_path).map_err(|err| {
                            HalError::adapter_fault(
                                self.profile,
                                capability,
                                "open",
                                format!("failed to read host path {}: {err}", host_path.display()),
                            )
                            .with_host_error_code(io_error_code(&err, 53, 53))
                        })?;
                        let len = clamp_u64_to_i32(data.len() as u64);
                        (data, len, 0)
                    }
                    1 => {
                        OpenOptions::new()
                            .read(true)
                            .write(true)
                            .create(true)
                            .truncate(true)
                            .open(host_path.as_path())
                            .map_err(|err| {
                                HalError::adapter_fault(
                                    self.profile,
                                    capability,
                                    "open",
                                    format!(
                                        "failed to create output host path {}: {err}",
                                        host_path.display()
                                    ),
                                )
                            })?;
                        (Vec::new(), 0, 0)
                    }
                    2 => {
                        OpenOptions::new()
                            .read(true)
                            .write(true)
                            .create(true)
                            .truncate(false)
                            .open(host_path.as_path())
                            .map_err(|err| {
                                HalError::adapter_fault(
                                    self.profile,
                                    capability,
                                    "open",
                                    format!(
                                        "failed to open append host path {}: {err}",
                                        host_path.display()
                                    ),
                                )
                            })?;
                        let data = fs::read(host_path).unwrap_or_default();
                        let len = clamp_u64_to_i32(data.len() as u64);
                        // VBA `Append` reports `Seek = 1` / `Loc = 0` on a fresh
                        // open even though the file is non-empty; the write cursor
                        // is decoupled and writes always land at EOF (handled in
                        // the sequential write paths). So the reported cursor is 0.
                        (data, len, 0)
                    }
                    _ => {
                        OpenOptions::new()
                            .read(true)
                            .write(true)
                            .create(true)
                            .truncate(false)
                            .open(host_path.as_path())
                            .map_err(|err| {
                                HalError::adapter_fault(
                                    self.profile,
                                    capability,
                                    "open",
                                    format!(
                                        "failed to open host path {}: {err}",
                                        host_path.display()
                                    ),
                                )
                            })?;
                        let data = fs::read(host_path).unwrap_or_default();
                        let len = clamp_u64_to_i32(data.len() as u64);
                        (data, len, 0)
                    }
                }
            } else if mode == 0 {
                (Vec::new(), i32::from(!path_text.is_empty()), 0)
            } else {
                (Vec::new(), 0, 0)
            };
            let print_column = initial_print_column(mode, &initial_data, initial_position);
            state.handles.insert(
                handle,
                FileHandleState {
                    mode,
                    position: initial_position,
                    input_buffer_origin: initial_position,
                    input_text_eof: false,
                    len: initial_len,
                    host_path,
                    data: initial_data,
                    modified: mode == 1,
                    width: 0,
                    print_column,
                    locks: Vec::new(),
                    record_len: 0,
                    get_hit_eof: false,
                },
            );
            self.assert_fs_invariants(&state, "open-post");
            return Ok(Variant::from_i32(handle));
        }
        let path = self.variant_to_i32(&path, capability, "open", "path")?;
        let mut state = self.borrow_fs(capability, "open")?;
        self.assert_fs_invariants(&state, "open-pre");
        let Some(handle) = state.first_free_in(1, 511) else {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "open",
                "no free file handles available in supported range",
            ));
        };
        let host_path = if self.native_fs_enabled() {
            let host_path = self.host_path_from_token(path);
            if let Some(parent) = host_path.parent() {
                fs::create_dir_all(parent).map_err(|err| {
                    HalError::adapter_fault(
                        self.profile,
                        capability,
                        "open",
                        format!("failed to create host fs directory: {err}"),
                    )
                })?;
            }
            Some(host_path)
        } else {
            None
        };
        let (initial_data, initial_len, initial_position) = if let Some(host_path) =
            host_path.as_ref()
        {
            match mode {
                0 => {
                    let data = fs::read(host_path).map_err(|err| {
                        HalError::adapter_fault(
                            self.profile,
                            capability,
                            "open",
                            format!("failed to read host path {}: {err}", host_path.display()),
                        )
                    })?;
                    let len = clamp_u64_to_i32(data.len() as u64);
                    (data, len, 0)
                }
                1 => {
                    OpenOptions::new()
                        .read(true)
                        .write(true)
                        .create(true)
                        .truncate(true)
                        .open(host_path)
                        .map_err(|err| {
                            HalError::adapter_fault(
                                self.profile,
                                capability,
                                "open",
                                format!(
                                    "failed to create output host path {}: {err}",
                                    host_path.display()
                                ),
                            )
                        })?;
                    (Vec::new(), 0, 0)
                }
                2 => {
                    OpenOptions::new()
                        .read(true)
                        .write(true)
                        .create(true)
                        .truncate(false)
                        .open(host_path)
                        .map_err(|err| {
                            HalError::adapter_fault(
                                self.profile,
                                capability,
                                "open",
                                format!(
                                    "failed to open append host path {}: {err}",
                                    host_path.display()
                                ),
                            )
                        })?;
                    let data = fs::read(host_path).unwrap_or_default();
                    let len = clamp_u64_to_i32(data.len() as u64);
                    (data, len, len)
                }
                _ => {
                    OpenOptions::new()
                        .read(true)
                        .write(true)
                        .create(true)
                        .truncate(false)
                        .open(host_path)
                        .map_err(|err| {
                            HalError::adapter_fault(
                                self.profile,
                                capability,
                                "open",
                                format!("failed to open host path {}: {err}", host_path.display()),
                            )
                        })?;
                    let data = fs::read(host_path).unwrap_or_default();
                    let len = clamp_u64_to_i32(data.len() as u64);
                    (data, len, 0)
                }
            }
        } else if mode == 0 {
            (Vec::new(), pseudo_file_len_from_path_token(path), 0)
        } else {
            (Vec::new(), 0, 0)
        };
        let print_column = initial_print_column(mode, &initial_data, initial_position);
        state.handles.insert(
            handle,
            FileHandleState {
                mode,
                position: initial_position,
                input_buffer_origin: initial_position,
                input_text_eof: false,
                len: initial_len,
                host_path,
                data: initial_data,
                modified: mode == 1,
                width: 0,
                print_column,
                locks: Vec::new(),
                record_len: 0,
                get_hit_eof: false,
            },
        );
        self.assert_fs_invariants(&state, "open-post");
        hal_contract_assert!(
            (1..=511).contains(&handle),
            "op=open returned out-of-range handle {}",
            handle
        );
        Ok(Variant::from_i32(handle))
    }

    fn open_with_record_len(
        &self,
        path: Variant,
        mode: Variant,
        record_len: Variant,
    ) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        let mode_raw = self.variant_to_i32(&mode, capability, "open", "mode")?;
        // `Open … As #n` takes a file number 1..=511 (error 52 otherwise); `Len` is at
        // most 32767 (overflow, 6).
        if !(1..=511).contains(&(mode_raw >> 16)) {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "open",
                format!("bad file name or number: #{}", mode_raw >> 16),
            )
            .with_host_error_code(52));
        }
        let len = self
            .variant_to_i32(&record_len, capability, "open", "len")
            .unwrap_or(0);
        if len > 32767 {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "open",
                format!("record length {len} is out of range"),
            )
            .with_host_error_code(6));
        }
        let handle = self.open_variant(path, mode)?;
        // Random files use `Len = N` as the record length, 128 when it is omitted.
        let len = match (mode_raw & 0xFFFF, len) {
            (MODE_RANDOM, len) if len <= 0 => 128,
            (_, len) => len,
        };
        if len > 0
            && let Ok(handle_id) = self.variant_to_i32(&handle, capability, "open", "handle")
        {
            let mut state = self.borrow_fs(capability, "open")?;
            if let Ok(entry) = self.fs_entry_mut(&mut state, handle_id, "open") {
                entry.record_len = len;
            }
        }
        Ok(handle)
    }

    fn close_variant(&self, handle: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "close"));
        }
        let handle = self.variant_to_i32(&handle, capability, "close", "handle")?;
        let mut state = self.borrow_fs(capability, "close")?;
        self.assert_fs_invariants(&state, "close-pre");
        if handle == 0 {
            let drained: Vec<FileHandleState> = state.handles.values().cloned().collect();
            let count = drained.len() as i32;
            state.handles.clear();
            for entry in drained {
                self.flush_host_file(&entry)?;
            }
            self.assert_fs_invariants(&state, "close-all-post");
            Ok(Variant::from_i32(count))
        } else if let Some(entry) = state.handles.remove(&handle) {
            self.flush_host_file(&entry)?;
            self.assert_fs_invariants(&state, "close-post");
            Ok(Variant::from_i32(1))
        } else {
            Err(HalError::adapter_fault(
                self.profile,
                capability,
                "close",
                format!("invalid file handle: {handle}"),
            )
            .with_host_error_code(52))
        }
    }

    fn kill_variant(&self, path: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "kill"));
        }
        if !self.policy.allow_filesystem_mutation {
            return Err(self.denied(capability, "kill"));
        }
        let Some(path_text) = path.as_bstr() else {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "kill",
                "path must be a string",
            ));
        };
        if path_text.as_str().contains('*') || path_text.as_str().contains('?') {
            if self.native_fs_enabled() {
                let wildcard_path = self.map_path(&path_text.as_str(), capability, "kill", true)?;
                let matched_paths =
                    expand_host_wildcard_paths(&wildcard_path, self.uses_windows_paths())
                        .map_err(|err| {
                            if self.uses_windows_paths()
                                && err.kind() == std::io::ErrorKind::AlreadyExists
                            {
                                return self.windows_path_error(capability, "kill", err);
                            }
                            HalError::adapter_fault(
                                self.profile,
                                capability,
                                "kill",
                                format!("failed to expand wildcard Kill path {path_text}: {err}"),
                            )
                        })?
                        .into_iter()
                        .filter(|path| path.is_file())
                        .filter(|path| {
                            path_attribute_bits(path, self.uses_windows_paths())
                                .is_ok_and(|attrs| attrs & (VB_HIDDEN | VB_SYSTEM) == 0)
                        })
                        .collect::<Vec<_>>();
                if matched_paths.is_empty() {
                    return Err(HalError::adapter_fault(
                        self.profile,
                        capability,
                        "kill",
                        format!("wildcard Kill path matched no files: {path_text}"),
                    )
                    .with_host_error_code(53));
                }
                for matched_path in matched_paths {
                    self.kill_host_file(&matched_path)?;
                }
            }
            return Ok(Variant::from_i32(0));
        }
        if self.native_fs_enabled() {
            let host_path = self.map_path(&path_text.as_str(), capability, "kill", true)?;
            self.kill_host_file(&host_path)?;
        }
        Ok(Variant::from_i32(0))
    }

    fn mkdir_variant(&self, path: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "mkdir"));
        }
        if !self.policy.allow_filesystem_mutation {
            return Err(self.denied(capability, "mkdir"));
        }
        let Some(path_text) = path.as_bstr() else {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "mkdir",
                "path must be a string",
            ));
        };
        if self.native_fs_enabled() {
            // VBA `MkDir` creates a single directory (error 75 if it already
            // exists / the parent is missing); mirror with a non-recursive create.
            fs::create_dir(self.map_path(&path_text.as_str(), capability, "mkdir", true)?)
                .map_err(|err| {
                    HalError::adapter_fault(
                        self.profile,
                        capability,
                        "mkdir",
                        format!("failed to create directory {path_text}: {err}"),
                    )
                    .with_host_error_code(directory_error_code(&err))
                })?;
        }
        Ok(Variant::from_i32(0))
    }

    fn rmdir_variant(&self, path: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "rmdir"));
        }
        if !self.policy.allow_filesystem_mutation {
            return Err(self.denied(capability, "rmdir"));
        }
        let Some(path_text) = path.as_bstr() else {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "rmdir",
                "path must be a string",
            ));
        };
        if self.native_fs_enabled() {
            // VBA `RmDir` removes an empty directory; `remove_dir` (non-recursive)
            // matches (it errors if the directory is non-empty or missing).
            let directory = self.map_path(&path_text.as_str(), capability, "rmdir", true)?;
            let fault = |reason: String, code| {
                HalError::adapter_fault(
                    self.profile,
                    capability,
                    "rmdir",
                    format!("failed to remove directory {path_text}: {reason}"),
                )
                .with_host_error_code(code)
            };
            let mapped = self.uses_windows_paths();
            if mapped {
                let current = self
                    .mapped_cur_dir(None)
                    .and_then(|path| self.map_path(&path, capability, "rmdir", false));
                if current.is_ok_and(|path| path == directory || same_host_path(&path, &directory))
                    || fs::metadata(&directory)
                        .is_ok_and(|meta| meta.is_dir() && meta.permissions().readonly())
                {
                    return Err(fault("directory is current or read-only".into(), 75));
                }
            }
            fs::remove_dir(directory).map_err(|err| {
                let code = if mapped && err.kind() == std::io::ErrorKind::NotADirectory {
                    76
                } else {
                    directory_error_code(&err)
                };
                fault(err.to_string(), code)
            })?;
        }
        Ok(Variant::from_i32(0))
    }

    fn seek_variant(&self, handle: Variant, position: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "seek"));
        }
        let handle = self.variant_to_i32(&handle, capability, "seek", "handle")?;
        // The `Seek(filenumber)` FUNCTION form omits the position — it READS the
        // current 1-based position (next byte, or next record in Random) without
        // moving the cursor.
        if matches!(position.vtype(), VarType::Empty | VarType::Null) {
            let mut state = self.borrow_fs(capability, "seek")?;
            let entry = self.fs_entry_mut(&mut state, handle, "seek")?;
            return Ok(Variant::from_i32(seek_report(entry)));
        }
        // The `Seek #n, pos` STATEMENT repositions. `pos` is 1-based: a byte offset
        // for Binary/sequential, a record number for Random (mapped to its byte
        // slot via the fixed record length). VBA requires `pos >= 1`, and a bare
        // Seek past EOF does NOT extend the file — only a subsequent write does.
        let position = self.variant_to_i32(&position, capability, "seek", "position")?;
        if position < 1 {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "seek",
                format!("bad record number for seek: {position}"),
            )
            .with_host_error_code(63));
        }
        let mut state = self.borrow_fs(capability, "seek")?;
        self.assert_fs_invariants(&state, "seek-pre");
        {
            let entry = self.fs_entry_mut(&mut state, handle, "seek")?;
            entry.position = if entry.mode == MODE_RANDOM && entry.record_len > 0 {
                (position - 1).saturating_mul(entry.record_len)
            } else {
                position - 1
            };
            entry.input_buffer_origin = entry.position;
            entry.input_text_eof = false;
            entry.get_hit_eof = false;
            let print_cursor = if entry.mode == MODE_APPEND {
                entry.len.max(0) as usize
            } else {
                entry.position.max(0) as usize
            };
            entry.print_column = line_column_at(&entry.data, print_cursor);
        }
        self.assert_fs_invariants(&state, "seek-post");
        // The Seek STATEMENT yields no value; the discarded result is `Empty`.
        Ok(Variant::empty())
    }

    fn file_attr_variant(&self, handle: Variant, kind: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "fileattr"));
        }
        let handle = self.variant_to_i32(&handle, capability, "fileattr", "handle")?;
        let kind = self.variant_to_i32(&kind, capability, "fileattr", "kind")?;
        let mut state = self.borrow_fs(capability, "fileattr")?;
        let entry = self.fs_entry_mut(&mut state, handle, "fileattr")?;
        if kind != 1 {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "fileattr",
                "only FileAttr(n, 1) is supported",
            )
            .with_host_error_code(5));
        }
        // VBA's open-mode codes: Input 1, Output 2, Random 4, Append 8, Binary 32.
        Ok(Variant::from_i32(match entry.mode {
            0 => 1,
            1 => 2,
            MODE_APPEND => 8,
            MODE_BINARY => 32,
            _ => 4,
        }))
    }

    fn eof_variant(&self, handle: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "eof"));
        }
        let handle = self.variant_to_i32(&handle, capability, "eof", "handle")?;
        let mut state = self.borrow_fs(capability, "eof")?;
        let entry = self.fs_entry_mut(&mut state, handle, "eof")?;
        // A Binary or Random file reports EOF only once a `Get` could not read a whole record.
        if matches!(entry.mode, MODE_BINARY | MODE_RANDOM) {
            return Ok(Variant::from_i32(i32::from(entry.get_hit_eof)));
        }
        let text_end = entry.mode == 0
            && (entry.input_text_eof
                || entry.data.get(entry.position as usize).copied() == Some(26));
        Ok(Variant::from_i32(
            if entry.position >= entry.len || text_end {
                1
            } else {
                0
            },
        ))
    }

    fn lof_variant(&self, handle: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "lof"));
        }
        let handle = self.variant_to_i32(&handle, capability, "lof", "handle")?;
        let mut state = self.borrow_fs(capability, "lof")?;
        let entry = self.fs_entry_mut(&mut state, handle, "lof")?;
        Ok(Variant::from_i32(entry.len))
    }

    fn free_file_variant(&self, range_selector: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "free_file"));
        }
        let range_selector =
            self.variant_to_i32(&range_selector, capability, "free_file", "range_selector")?;
        let (start, end) = if range_selector != 0 {
            (256, 511)
        } else {
            (1, 255)
        };
        let state = self.borrow_fs(capability, "free_file")?;
        self.assert_fs_invariants(&state, "free_file");
        let candidate = state.first_free_in(start, end).ok_or_else(|| {
            HalError::adapter_fault(
                self.profile,
                capability,
                "free_file",
                format!("no free file number in range {start}..={end}"),
            )
            .with_host_error_code(67)
        })?;
        hal_contract_assert!(
            (start..=end).contains(&candidate),
            "op=free_file returned {} outside range {}..={}",
            candidate,
            start,
            end
        );
        Ok(Variant::from_i32(candidate))
    }

    fn input_chars_variant(&self, handle: Variant, count: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "input_chars"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "input_chars", "handle")?;
        let count = self.variant_to_i32(&count, capability, "input_chars", "count")?;
        let error = |code| {
            HalError::adapter_fault(
                self.profile,
                capability,
                "input_chars",
                calcvba_runtime::default_error_message(code),
            )
            .with_host_error_code(code)
        };
        if count < 0 {
            return Err(error(5));
        }
        let mut state = self.borrow_fs(capability, "input_chars")?;
        let entry = self
            .fs_entry_mut(&mut state, handle_id, "input_chars")
            .map_err(|_| error(52))?;
        if !matches!(entry.mode, 0 | MODE_BINARY) {
            return Err(error(54));
        }
        if entry.mode == 0 && entry.input_text_eof && count != 0 {
            return Err(error(62));
        }
        let mut bytes = Vec::new();
        let mut cursor = entry.position as usize;
        for _ in 0..count {
            if entry.mode == 0 {
                while cursor < entry.data.len() && entry.data[cursor] == 0 {
                    cursor += 1;
                }
            }
            if cursor >= entry.data.len() {
                break;
            }
            let byte = entry.data[cursor];
            if entry.mode == 0 && byte == 26 {
                // Office's sequential text reader reports DOS EOF. Its Seek
                // result after this failed read includes the buffered cursor;
                // a later failed read retains that position until Seek resets it.
                let origin = entry.input_buffer_origin as usize;
                let block = origin + (cursor - origin) / 512 * 512;
                let end = block.saturating_add(512).min(entry.data.len());
                cursor = end.saturating_add(cursor - block).saturating_add(1);
                entry.input_text_eof = true;
                break;
            }
            bytes.push(byte);
            cursor += 1;
            if calcvba_runtime::ansi::is_dbcs_lead_byte(byte) && cursor < entry.data.len() {
                bytes.push(entry.data[cursor]);
                cursor += 1;
            }
        }
        entry.position = i32::try_from(cursor).map_err(|_| error(6))?;
        let text = ansi_decode(&bytes);
        if entry.mode == 0 && text.encode_utf16().count() < count as usize {
            return Err(error(62));
        }
        Ok(Variant::from_string(text))
    }

    fn read_bytes_variant(&self, handle: Variant, count: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "read_bytes"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "read_bytes", "handle")?;
        let count = self.variant_to_i32(&count, capability, "read_bytes", "count")?;
        let mut state = self.borrow_fs(capability, "read_bytes")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "read_bytes")?;
        let pos = entry.position as usize;
        let count = count.max(0) as usize;
        let available = entry.data.len().saturating_sub(pos);
        let actual = count.min(available);
        let bytes = entry.data[pos..pos + actual].to_vec();
        entry.position += actual as i32;
        Ok(Variant::from_string(ansi_decode(&bytes)))
    }

    fn write_bytes_variant(&self, handle: Variant, data: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "write_bytes"));
        }
        if !self.policy.allow_filesystem_mutation {
            return Err(self.denied(capability, "write_bytes"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "write_bytes", "handle")?;
        let text = format!("{}\r\n", calcvba_runtime::write_display_text(&data));
        let bytes = ansi_encode(&text);
        let mut state = self.borrow_fs(capability, "write_bytes")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "write_bytes")?;
        let pos = seq_write_offset(entry);
        let end = pos + bytes.len();
        if end > entry.data.len() {
            entry.data.resize(end, 0);
        }
        entry.data[pos..end].copy_from_slice(&bytes);
        entry.modified = true;
        entry.position = if entry.mode == MODE_APPEND {
            entry.position.saturating_add(bytes.len() as i32)
        } else {
            end as i32
        };
        entry.len = entry.data.len() as i32;
        entry.print_column = print_column_after_text(entry.print_column, &text);
        Ok(Variant::from_i32(bytes.len() as i32))
    }

    fn print_line_variant(&self, handle: Variant, data: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "print_line"));
        }
        if !self.policy.allow_filesystem_mutation {
            return Err(self.denied(capability, "print_line"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "print_line", "handle")?;
        // The caller (calcvba-lib `file_print`/`file_write`) assembles the complete VBA
        // record — per-field formatting, separators, print zones, and the `\r\n`
        // terminator (when not suppressed by a trailing `;`/`,`) — so this sink writes
        // the text verbatim at the current position.
        let text = data
            .as_bstr()
            .map(|t| t.as_str().to_string())
            .unwrap_or_default();
        let bytes = ansi_encode(&text);
        let mut state = self.borrow_fs(capability, "print_line")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "print_line")?;
        let pos = seq_write_offset(entry);
        let end = pos + bytes.len();
        if end > entry.data.len() {
            entry.data.resize(end, 0);
        }
        entry.data[pos..end].copy_from_slice(&bytes);
        entry.modified |= !bytes.is_empty();
        entry.position = if entry.mode == MODE_APPEND {
            entry.position.saturating_add(bytes.len() as i32)
        } else {
            end as i32
        };
        entry.len = entry.data.len() as i32;
        entry.print_column = print_column_after_text(entry.print_column, &text);
        Ok(Variant::from_i32(0))
    }

    fn print_column_variant(&self, handle: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "print_column"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "print_column", "handle")?;
        let mut state = self.borrow_fs(capability, "print_column")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "print_column")?;
        Ok(Variant::from_i32(entry.print_column))
    }

    fn print_width_variant(&self, handle: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "print_width"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "print_width", "handle")?;
        let mut state = self.borrow_fs(capability, "print_width")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "print_width")?;
        Ok(Variant::from_i32(entry.width))
    }

    fn input_fields_variant(&self, handle: Variant, count: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "input_fields"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "input_fields", "handle")?;
        let count = self.variant_to_i32(&count, capability, "input_fields", "count")?;
        let count = count.max(1) as usize;
        let mut state = self.borrow_fs(capability, "input_fields")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "input_fields")?;
        let pos = entry.position as usize;
        let mut fields = Vec::new();
        let mut cursor = pos;
        while fields.len() < count && cursor < entry.data.len() {
            let (field, next_cursor) = parse_input_field(&entry.data, cursor);
            fields.push(field);
            cursor = advance_input_separator(&entry.data, next_cursor);
        }
        entry.position = cursor as i32;
        if count == 1 {
            if let Some(field) = fields.first() {
                return Ok(variant_from_input_field(field));
            }
            return Ok(Variant::from_string(BStr::empty()));
        }
        let result = fields.join(",");
        Ok(Variant::from_string(result))
    }

    fn line_input_variant(&self, handle: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "line_input"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "line_input", "handle")?;
        let mut state = self.borrow_fs(capability, "line_input")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "line_input")?;
        let pos = entry.position as usize;
        if pos >= entry.data.len() {
            return Ok(Variant::from_string(BStr::empty()));
        }
        let remaining = &entry.data[pos..];
        let dos_eof = (entry.mode == 0)
            .then(|| remaining.iter().position(|&b| b == 26))
            .flatten();
        let remaining = &remaining[..dos_eof.unwrap_or(remaining.len())];
        let line_end = remaining
            .iter()
            .position(|&b| b == b'\n')
            .map(|i| {
                if i > 0 && remaining[i - 1] == b'\r' {
                    (i - 1, i + 1)
                } else {
                    (i, i + 1)
                }
            })
            .unwrap_or((remaining.len(), remaining.len()));
        let line = ansi_decode(&remaining[..line_end.0]);
        entry.position += line_end.1 as i32;
        if dos_eof == Some(line_end.1) {
            entry.input_text_eof = true;
        }
        Ok(Variant::from_string(line))
    }

    fn loc_variant(&self, handle: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "loc"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "loc", "handle")?;
        let mut state = self.borrow_fs(capability, "loc")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "loc")?;
        Ok(Variant::from_i32(loc_report(entry)))
    }

    fn put_record_variant(
        &self,
        handle: Variant,
        position: Variant,
        value: Variant,
        fixed: Variant,
    ) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "put"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "put", "handle")?;
        let explicit = if matches!(position.vtype(), VarType::Empty | VarType::Null) {
            None
        } else {
            Some(self.variant_to_i32(&position, capability, "put", "position")?)
        };
        // A fixed-length-string source writes raw (no length prefix) in any mode.
        let fixed = self
            .variant_to_i32(&fixed, capability, "put", "fixed")
            .unwrap_or(0)
            != 0;
        let mut state = self.borrow_fs(capability, "put")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "put")?;
        let mode = entry.mode;
        let bytes = if let Some(layout) = value.vba_record_layout() {
            serialize_udt_field(&VbaRecordFieldKind::Record(layout.clone()), &value).map_err(
                |message| {
                    HalError::adapter_fault(self.profile, capability, "put", message)
                        .with_host_error_code(13)
                },
            )?
        } else {
            serialize_record(&value, mode, fixed)
        };
        // Random mode positions by fixed-width record slot (`(rec-1)*stride`),
        // padding the slot; Binary mode positions by byte offset.
        let (offset, slot) = if mode == MODE_RANDOM {
            let stride = if entry.record_len > 0 {
                entry.record_len as usize
            } else {
                bytes.len()
            };
            if entry.record_len > 0 && bytes.len() > stride {
                return Err(HalError::adapter_fault(
                    self.profile,
                    capability,
                    "put",
                    "Bad record length: value exceeds the file's Len",
                )
                .with_host_error_code(59));
            }
            let off = match explicit {
                Some(rec) if rec >= 1 => ((rec - 1) as usize).saturating_mul(stride),
                _ => entry.position.max(0) as usize,
            };
            (off, stride)
        } else {
            let off = match explicit {
                Some(rec) if rec >= 1 => (rec - 1) as usize,
                _ => entry.position.max(0) as usize,
            };
            (off, bytes.len())
        };
        let write_end = offset + bytes.len();
        let slot_end = offset + slot;
        // The file grows only to the last byte written; the unwritten tail of a
        // final record slot is not padded.
        if write_end > entry.data.len() {
            entry.data.resize(write_end, 0);
        }
        entry.data[offset..write_end].copy_from_slice(&bytes);
        entry.modified |= !bytes.is_empty();
        entry.position = slot_end as i32;
        entry.len = entry.len.max(write_end as i32);
        Ok(Variant::empty())
    }

    fn get_record_variant(
        &self,
        handle: Variant,
        position: Variant,
        vtype: Variant,
        str_len: Variant,
    ) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "get"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "get", "handle")?;
        let explicit = if matches!(position.vtype(), VarType::Empty | VarType::Null) {
            None
        } else {
            Some(self.variant_to_i32(&position, capability, "get", "position")?)
        };
        let code = self.variant_to_i32(&vtype, capability, "get", "type")? as u16;
        // `VT_ARRAY | element`: fill an array shaped like the one passed in the length slot.
        if code & 0x2000 != 0 {
            return self.get_array_record(handle_id, explicit, code & 0x0FFF, &str_len);
        }
        let vt = VarType::from_u16(code).unwrap_or(VarType::Empty);
        if vt == VarType::Record {
            return self.get_udt_record(handle_id, explicit, &str_len);
        }
        let fixed_numeric = record_size(vt);
        // `str_len` carries the string length intent: > 0 = a fixed-length string of
        // that many bytes (no prefix, any mode); <= 0 = a variable string, where
        // `-str_len` is the target's current Len (used for a Binary read; Random
        // uses the on-disk 2-byte length prefix).
        let str_len = self
            .variant_to_i32(&str_len, capability, "get", "str_len")
            .unwrap_or(0);
        let mut state = self.borrow_fs(capability, "get")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "get")?;
        let mode = entry.mode;
        // Random mode strides by the fixed record length (or the value's natural
        // size: numeric size, or a fixed string's byte length); Binary by byte offset.
        let stride = if entry.record_len > 0 {
            entry.record_len as usize
        } else if let Some(size) = fixed_numeric {
            size
        } else if str_len > 0 {
            str_len as usize
        } else {
            0
        };
        let offset = if mode == MODE_RANDOM {
            match explicit {
                Some(rec) if rec >= 1 => ((rec - 1) as usize).saturating_mul(stride),
                _ => entry.position.max(0) as usize,
            }
        } else {
            match explicit {
                Some(rec) if rec >= 1 => (rec - 1) as usize,
                _ => entry.position.max(0) as usize,
            }
        };
        let data_len = entry.data.len();
        let at = |start: usize, n: usize| -> (usize, usize) {
            (start.min(data_len), (start + n).min(data_len))
        };
        let (value, advance) = match (fixed_numeric, vt) {
            (Some(size), _) => {
                let (s, e) = at(offset, size);
                let v = deserialize_record(vt, &entry.data[s..e]);
                (
                    v,
                    if mode == MODE_RANDOM && stride > 0 {
                        stride
                    } else {
                        size
                    },
                )
            }
            (None, VarType::String) if str_len > 0 => {
                // Fixed-length string: exactly `str_len` ANSI bytes, no prefix.
                let n = str_len as usize;
                let (s, e) = at(offset, n);
                let v = deserialize_record(VarType::String, &entry.data[s..e]);
                (
                    v,
                    if mode == MODE_RANDOM && stride > 0 {
                        stride
                    } else {
                        n
                    },
                )
            }
            (None, VarType::String) if mode == MODE_RANDOM => {
                // Variable-length string in a Random record: 2-byte length prefix.
                let len = if offset + 2 <= data_len {
                    u16::from_le_bytes([entry.data[offset], entry.data[offset + 1]]) as usize
                } else {
                    0
                };
                let (s, e) = at(offset + 2, len);
                let v = deserialize_record(VarType::String, &entry.data[s..e]);
                (v, if stride > 0 { stride } else { 2 + len })
            }
            (None, VarType::String) if mode == MODE_BINARY => {
                // Variable-length string in Binary: read the target's current Len
                // (`-str_len`) ANSI bytes, no prefix.
                let n = (-str_len).max(0) as usize;
                let (s, e) = at(offset, n);
                (deserialize_record(VarType::String, &entry.data[s..e]), n)
            }
            (None, _) => (Variant::empty(), 0),
        };
        // EOF of a Binary or Random file: this `Get` could not read a whole record.
        let needed = match (fixed_numeric, vt) {
            (Some(size), _) => size,
            (None, VarType::String) if str_len > 0 => str_len as usize,
            (None, VarType::String) if mode == MODE_RANDOM => 2,
            (None, VarType::String) if mode == MODE_BINARY => (-str_len).max(0) as usize,
            _ => 0,
        };
        entry.get_hit_eof = needed > 0 && offset.saturating_add(needed) > data_len;
        entry.position = (offset + advance) as i32;
        Ok(value)
    }

    fn width_variant(&self, handle: Variant, width: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "width"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "width", "handle")?;
        if width.vtype() == VarType::Null {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "width",
                "Invalid use of Null",
            )
            .with_host_error_code(94));
        }
        let width = self.variant_to_i32(&width, capability, "width", "width")?;
        if !(0..=255).contains(&width) {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "width",
                "Invalid procedure call or argument",
            )
            .with_host_error_code(5));
        }
        let mut state = self.borrow_fs(capability, "width")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "width")?;
        entry.width = width;
        Ok(Variant::empty())
    }

    fn cur_dir_variant(&self, drive: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "curdir"));
        }
        // A read of process state; the deterministic lane reports the empty
        // string rather than leaking the real working directory.
        if self.native_fs_enabled() {
            let letter = drive.as_bstr().and_then(|s| s.as_str().chars().next());
            if letter.is_some_and(|letter| !letter.is_ascii_alphabetic()) {
                return Err(HalError::adapter_fault(
                    self.profile,
                    capability,
                    "curdir",
                    "invalid drive letter",
                ));
            }
            if self.uses_windows_paths() {
                return self.mapped_cur_dir(letter).map(Variant::from_string);
            }
            #[cfg(windows)]
            let dir = match letter {
                Some(letter) => std::path::absolute(format!("{letter}:")),
                None => std::env::current_dir(),
            };
            #[cfg(not(windows))]
            let dir = std::env::current_dir();
            let dir = dir.map_err(|err| {
                HalError::adapter_fault(self.profile, capability, "curdir", err.to_string())
            })?;
            Ok(Variant::from_string(BStr::from(
                dir.to_string_lossy().as_ref(),
            )))
        } else {
            Ok(Variant::from_string(BStr::empty()))
        }
    }

    fn ch_dir_variant(&self, path: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "chdir"));
        }
        if !self.policy.allow_filesystem_mutation {
            return Err(self.denied(capability, "chdir"));
        }
        if self.native_fs_enabled() {
            if self.uses_windows_paths()
                && let Some(text) = path.as_bstr()
            {
                self.mapped_ch_dir(&text.as_str())?;
                return Ok(Variant::from_i32(0));
            }
            let dir = self.variant_to_path(&path, capability, "chdir", "path", false)?;
            std::env::set_current_dir(&dir).map_err(|err| {
                HalError::adapter_fault(
                    self.profile,
                    capability,
                    "chdir",
                    format!("failed to change directory to {}: {err}", dir.display()),
                )
                .with_host_error_code(io_error_code(&err, 76, 76))
            })?;
        }
        Ok(Variant::from_i32(0))
    }

    fn file_len_variant(&self, path: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "filelen"));
        }
        if self.native_fs_enabled() {
            let p = self.variant_to_path(&path, capability, "filelen", "path", false)?;
            let len = fs::metadata(&p).map_err(|err| {
                HalError::adapter_fault(
                    self.profile,
                    capability,
                    "filelen",
                    format!("failed to stat {}: {err}", p.display()),
                )
                .with_host_error_code(metadata_error_code(
                    &err,
                    &p,
                    self.uses_windows_paths(),
                ))
            })?;
            let bytes = if self.uses_windows_paths() && len.is_dir() {
                0
            } else {
                len.len()
            };
            Ok(Variant::from_i32(bytes.min(i32::MAX as u64) as i32))
        } else {
            Ok(Variant::from_i32(0))
        }
    }

    fn file_copy_variant(&self, source: Variant, dest: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "filecopy"));
        }
        if !self.policy.allow_filesystem_mutation {
            return Err(self.denied(capability, "filecopy"));
        }
        if self.native_fs_enabled() {
            let src = self.variant_to_path(&source, capability, "filecopy", "source", false)?;
            let dst = self.variant_to_path(&dest, capability, "filecopy", "destination", true)?;
            let fault = |code| {
                HalError::adapter_fault(self.profile, capability, "filecopy", "file copy failed")
                    .with_host_error_code(code)
            };
            if src.as_os_str().is_empty() || dst.as_os_str().is_empty() {
                return Err(fault(75));
            }
            if same_host_path(&src, &dst) {
                return Err(fault(70));
            }
            if self
                .borrow_fs(capability, "filecopy")?
                .handles
                .values()
                .any(|entry| {
                    entry.host_path.as_deref().is_some_and(|open| {
                        (entry.mode != 0 && same_host_path(open, &src))
                            || same_host_path(open, &dst)
                    })
                })
            {
                return Err(fault(55));
            }
            fs::copy(&src, &dst).map_err(|err| {
                HalError::adapter_fault(
                    self.profile,
                    capability,
                    "filecopy",
                    format!(
                        "failed to copy {} to {}: {err}",
                        src.display(),
                        dst.display()
                    ),
                )
                .with_host_error_code(
                    if matches!(
                        err.kind(),
                        std::io::ErrorKind::PermissionDenied
                            | std::io::ErrorKind::InvalidInput
                            | std::io::ErrorKind::IsADirectory
                    ) {
                        75
                    } else {
                        let mapped = self.uses_windows_paths();
                        let missing = if mapped && src.exists() { &dst } else { &src };
                        metadata_error_code(&err, missing, mapped)
                    },
                )
            })?;
            // VBA creates ordinary destination attributes rather than copying
            // the source's read-only/hidden/system flags.
            apply_file_attributes(&dst, VB_ARCHIVE, self.uses_windows_paths())
                .map_err(|_| fault(75))?;
        }
        Ok(Variant::from_i32(0))
    }

    fn get_attr_variant(&self, path: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "getattr"));
        }
        if self.native_fs_enabled() {
            let p = self.variant_to_path(&path, capability, "getattr", "path", false)?;
            let bits = path_attribute_bits(&p, self.uses_windows_paths()).map_err(|err| {
                HalError::adapter_fault(
                    self.profile,
                    capability,
                    "getattr",
                    format!("failed to stat {}: {err}", p.display()),
                )
                .with_host_error_code(metadata_error_code(
                    &err,
                    &p,
                    self.uses_windows_paths(),
                ))
            })?;
            // Excel returns a Long, including filesystem-specific high bits.
            Ok(Variant::from_i32(bits))
        } else {
            Ok(Variant::from_i32(0))
        }
    }

    fn set_attr_variant(&self, path: Variant, attributes: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "setattr"));
        }
        if !self.policy.allow_filesystem_mutation {
            return Err(self.denied(capability, "setattr"));
        }
        let requested = self.variant_to_i32(&attributes, capability, "setattr", "attributes")?;
        // `SetAttr` manages only these bits; `vbDirectory`/`vbVolume` are intrinsic
        // to the entry and cannot be assigned, so a request to set them is an error
        // (real VBA raises 5, "Invalid procedure call") rather than a silent no-op.
        if requested & !VB_SETTABLE_ATTRS != 0 {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "setattr",
                format!(
                    "attribute value {requested} sets bits SetAttr cannot change \
                     (only vbReadOnly|vbHidden|vbSystem|vbArchive)"
                ),
            ));
        }
        if self.native_fs_enabled() {
            let p = self.variant_to_path(&path, capability, "setattr", "path", true)?;
            apply_file_attributes(&p, requested, self.uses_windows_paths()).map_err(|err| {
                HalError::adapter_fault(
                    self.profile,
                    capability,
                    "setattr",
                    format!("failed to set attributes on {}: {err}", p.display()),
                )
                .with_host_error_code(
                    if cfg!(windows) && err.raw_os_error() == Some(123) {
                        52
                    } else {
                        io_error_code(&err, 53, 5)
                    },
                )
            })?;
        }
        Ok(Variant::from_i32(0))
    }

    fn ch_drive_variant(&self, drive: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "chdrive"));
        }
        if !self.policy.allow_filesystem_mutation {
            return Err(self.denied(capability, "chdrive"));
        }
        let Some(text) = drive.as_bstr() else {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "chdrive",
                "drive must be a string",
            ));
        };
        // VBA `ChDrive` looks only at the first character; "" is a no-op.
        if let Some(letter) = text.as_str().chars().next()
            && self.native_fs_enabled()
        {
            if self.uses_windows_paths() {
                self.mapped_ch_drive(letter)?;
                return Ok(Variant::from_i32(0));
            }
            change_drive(letter).map_err(|err| {
                let fault = HalError::adapter_fault(
                    self.profile,
                    capability,
                    "chdrive",
                    format!("failed to change drive to {letter:?}: {err}"),
                );
                if letter.is_ascii_alphabetic() && matches!(err.raw_os_error(), Some(3 | 15 | 21)) {
                    fault.with_host_error_code(68)
                } else {
                    fault
                }
            })?;
        }
        Ok(Variant::from_i32(0))
    }

    fn file_date_time_variant(&self, path: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "filedatetime"));
        }
        if self.native_fs_enabled() {
            let p = self.variant_to_path(&path, capability, "filedatetime", "path", false)?;
            // A bare directory name is accepted, but its trailing separator is
            // not (Excel P_FileDateTime.14).
            if path
                .as_bstr()
                .is_some_and(|text| text.as_str().ends_with(['\\', '/']))
                || p.as_os_str().to_string_lossy().ends_with(['\\', '/'])
            {
                return Err(HalError::adapter_fault(
                    self.profile,
                    capability,
                    "filedatetime",
                    "file not found",
                )
                .with_host_error_code(53));
            }
            let meta = fs::metadata(&p).map_err(|err| {
                HalError::adapter_fault(
                    self.profile,
                    capability,
                    "filedatetime",
                    format!("failed to stat {}: {err}", p.display()),
                )
                .with_host_error_code(metadata_error_code(
                    &err,
                    &p,
                    self.uses_windows_paths(),
                ))
            })?;
            let modified = meta.modified().map_err(|err| {
                HalError::adapter_fault(
                    self.profile,
                    capability,
                    "filedatetime",
                    format!("modification time unavailable for {}: {err}", p.display()),
                )
            })?;
            // VBA `FileDateTime` reports the local wall-clock timestamp shown by
            // the host filesystem, matching the native `Now`/`Date`/`Time` facet.
            Ok(Variant::from_date_f64(system_time_to_local_vba_serial(
                modified,
            )))
        } else {
            // Deterministic lane: no real files; a stable zero serial (1899-12-30).
            Ok(Variant::from_date_f64(0.0))
        }
    }

    fn name_variant(&self, old_path: Variant, new_path: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "name"));
        }
        if !self.policy.allow_filesystem_mutation {
            return Err(self.denied(capability, "name"));
        }
        // Real rename when host filesystem access is enabled; otherwise the
        // deterministic lane has no host file to move (a documented no-op).
        if self.native_fs_enabled() {
            let old = self.variant_to_path(&old_path, capability, "name", "old", true)?;
            let mut new = self.variant_to_path(&new_path, capability, "name", "new", true)?;
            if self.uses_windows_paths() && fs::symlink_metadata(&new).is_ok() {
                // The mapped lookup resolves both spellings to the same entry.
                // Keep the requested basename for the actual case-only rename.
                let spelling = new_path.as_bstr().and_then(|text| {
                    text.as_str()
                        .trim_end_matches(['/', '\\'])
                        .rsplit(['/', '\\', ':'])
                        .next()
                        .map(str::to_owned)
                });
                let case_only = spelling.filter(|requested| {
                    old == new
                        && new
                            .file_name()
                            .and_then(|name| name.to_str())
                            .is_some_and(|existing| {
                                existing != requested
                                    && existing.to_lowercase() == requested.to_lowercase()
                            })
                });
                let Some(case_only) = case_only else {
                    return Err(HalError::adapter_fault(
                        self.profile,
                        capability,
                        "name",
                        format!("destination {} already exists", new.display()),
                    )
                    .with_host_error_code(58));
                };
                new.set_file_name(case_only);
            }
            fs::rename(&old, &new).map_err(|err| {
                let fault = HalError::adapter_fault(
                    self.profile,
                    capability,
                    "name",
                    format!(
                        "failed to rename {} -> {}: {err}",
                        old.display(),
                        new.display()
                    ),
                );
                // Excel VBA's Name reports File not found for a missing file or parent.
                if matches!(
                    err.kind(),
                    std::io::ErrorKind::NotFound | std::io::ErrorKind::NotADirectory
                ) {
                    fault.with_host_error_code(53)
                } else {
                    fault
                }
            })?;
        }
        Ok(Variant::empty())
    }

    fn lock_variant(&self, handle: Variant, start: Variant, end: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "lock"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "lock", "handle")?;
        let range = self.lock_range(&start, &end, capability, "lock")?;
        let mut state = self.borrow_fs(capability, "lock")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "lock")?;
        // Re-locking an already-locked range is VBA error 70 (permission denied).
        if entry
            .locks
            .iter()
            .any(|&(s, e)| range.0 <= e && s <= range.1)
        {
            return Err(HalError::adapter_fault(
                self.profile,
                capability,
                "lock",
                "Permission denied: range already locked",
            ));
        }
        entry.locks.push(range);
        Ok(Variant::empty())
    }

    fn unlock_variant(&self, handle: Variant, start: Variant, end: Variant) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        if !self.supports(capability) {
            return Err(self.unsupported(capability, "unlock"));
        }
        let handle_id = self.variant_to_i32(&handle, capability, "unlock", "handle")?;
        let range = self.lock_range(&start, &end, capability, "unlock")?;
        let mut state = self.borrow_fs(capability, "unlock")?;
        let entry = self.fs_entry_mut(&mut state, handle_id, "unlock")?;
        entry.locks.retain(|&r| r != range);
        Ok(Variant::empty())
    }
}

impl StandardHostServices {
    fn get_udt_record(
        &self,
        handle: i32,
        explicit: Option<i32>,
        shape: &Variant,
    ) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        let fault = |message| {
            HalError::adapter_fault(self.profile, capability, "get", message)
                .with_host_error_code(13)
        };
        let layout = shape
            .vba_record_layout()
            .ok_or_else(|| fault("expected record layout".to_string()))?;
        let kind = VbaRecordFieldKind::Record(layout.clone());
        let size = kind.file_len().map_err(fault)?;
        let mut state = self.borrow_fs(capability, "get")?;
        let entry = self.fs_entry_mut(&mut state, handle, "get")?;
        let stride = if entry.mode == MODE_RANDOM && entry.record_len > 0 {
            entry.record_len as usize
        } else {
            size
        };
        let offset = match explicit {
            Some(rec) if rec >= 1 => ((rec - 1) as usize)
                .saturating_mul(if entry.mode == MODE_RANDOM { stride } else { 1 }),
            _ => entry.position.max(0) as usize,
        };
        let start = offset.min(entry.data.len());
        let end = offset.saturating_add(size).min(entry.data.len());
        let mut bytes = entry.data[start..end].to_vec();
        let short = bytes.len() < size;
        bytes.resize(size, 0);
        let value = deserialize_udt_field(&kind, shape, &bytes).map_err(fault)?;
        entry.get_hit_eof = short;
        entry.position = offset.saturating_add(stride) as i32;
        Ok(value)
    }

    /// `Get #n, [rec], array`: read one element per slot of `shape` (the target array) into
    /// a new array of the same bounds. A short read zero-fills and marks the file at EOF.
    fn get_array_record(
        &self,
        handle: i32,
        explicit: Option<i32>,
        element_code: u16,
        shape: &Variant,
    ) -> HalResult<Variant> {
        let capability = CapabilityId::FileSystemIo;
        let fault = |message: &str, code: i32| {
            HalError::adapter_fault(self.profile, capability, "get", message.to_string())
                .with_host_error_code(code)
        };
        let Some(shape) = shape.as_safearray() else {
            return Err(fault("Get target is not an array", 13));
        };
        let element = VarType::from_u16(element_code).unwrap_or(VarType::Empty);
        let Some(size) = record_size(element) else {
            return Err(fault("Get into an array of this element type", 13));
        };
        let count = shape.len();
        let total = size.saturating_mul(count);
        let mut state = self.borrow_fs(capability, "get")?;
        let entry = self.fs_entry_mut(&mut state, handle, "get")?;
        let stride = if entry.record_len > 0 {
            entry.record_len as usize
        } else {
            total
        };
        let offset = match explicit {
            Some(rec) if rec >= 1 => {
                if entry.mode == MODE_RANDOM {
                    ((rec - 1) as usize).saturating_mul(stride)
                } else {
                    (rec - 1) as usize
                }
            }
            _ => entry.position.max(0) as usize,
        };
        let start = offset.min(entry.data.len());
        let end = offset.saturating_add(total).min(entry.data.len());
        let mut bytes = entry.data[start..end].to_vec();
        entry.get_hit_eof = bytes.len() < total;
        bytes.resize(total, 0);
        let values: Vec<Variant> = (0..count)
            .map(|index| deserialize_record(element, &bytes[index * size..(index + 1) * size]))
            .collect();
        entry.position = offset.saturating_add(if entry.mode == MODE_RANDOM && stride > 0 {
            stride
        } else {
            total
        }) as i32;
        let array = match shape.bounds() {
            Some(bounds) => calcvba_runtime::safe_array::SafeArray::from_typed_variants_nd(
                bounds,
                element_code,
                values,
            ),
            None => {
                calcvba_runtime::safe_array::SafeArray::from_typed_variants(element_code, values)
            }
        }
        .map_err(|message: String| fault(&message, 5))?;
        Ok(Variant::from_safearray(array))
    }
}

/// The entry names `Dir(path, attributes)` lists, in name order. A path ending in a
/// separator lists that folder's contents; a bare folder name matches nothing unless
/// `vbDirectory` is requested. An entry is listed when its hidden/system/directory bits
/// are all among the requested `attributes` (so `vbNormal` lists plain files only), and
/// `.` and `..` join a wildcard listing that asks for directories.
pub(super) fn dir_listing(
    raw: &str,
    target: &Path,
    attrs: i32,
    case_insensitive: bool,
) -> std::io::Result<Vec<String>> {
    let pattern = if raw.ends_with(['\\', '/']) {
        target.join("*")
    } else {
        target.to_path_buf()
    };
    let wildcard = path_contains_wildcards(&pattern);
    let candidates = if wildcard {
        match expand_host_wildcard_paths(&pattern, case_insensitive) {
            Ok(paths) => paths,
            Err(err) if err.kind() == std::io::ErrorKind::NotFound => Vec::new(),
            Err(err) => return Err(err),
        }
    } else if pattern.exists() {
        vec![pattern.clone()]
    } else {
        Vec::new()
    };
    let filtered = VB_HIDDEN | VB_SYSTEM | VB_DIRECTORY;
    let wanted = attrs & filtered;
    let name_of = |path: &Path| {
        #[cfg(windows)]
        if !wildcard {
            let parent = path
                .parent()
                .filter(|p| !p.as_os_str().is_empty())
                .unwrap_or_else(|| Path::new("."));
            if let Some(name) = path.file_name()
                && let Ok(entries) = fs::read_dir(parent)
                && let Some(entry) = entries.filter_map(Result::ok).find(|entry| {
                    entry
                        .file_name()
                        .to_string_lossy()
                        .eq_ignore_ascii_case(&name.to_string_lossy())
                })
            {
                return entry.file_name().to_string_lossy().to_string();
            }
        }
        path.file_name()
            .map(|value| value.to_string_lossy().to_string())
            .unwrap_or_else(|| path.display().to_string())
    };
    let mut names: Vec<String> = candidates
        .iter()
        .filter(|path| {
            let attributes = if case_insensitive {
                path_attribute_bits(path, true)
            } else {
                fs::metadata(path).map(|meta| file_attribute_bits(&meta))
            };
            attributes
                .map(|bits| bits & filtered & !wanted == 0)
                .unwrap_or(false)
        })
        .map(|path| name_of(path))
        .collect();
    if wildcard
        && attrs & VB_DIRECTORY != 0
        && let Some(last) = pattern.file_name()
    {
        let last = last.to_string_lossy();
        for dots in [".", ".."] {
            if wildcard_match(&last, dots, case_insensitive) {
                names.push(dots.to_string());
            }
        }
    }
    names.sort_by_key(|name| name.to_lowercase());
    Ok(names)
}

pub(super) fn path_contains_wildcards(path: &Path) -> bool {
    path.components().any(|component| match component {
        Component::Normal(part) => {
            let text = part.to_string_lossy();
            text.contains('*') || text.contains('?')
        }
        _ => false,
    })
}

pub(super) fn expand_host_wildcard_paths(
    path: &Path,
    case_insensitive: bool,
) -> std::io::Result<Vec<PathBuf>> {
    let components = path.components().collect::<Vec<_>>();
    let mut roots = vec![PathBuf::new()];
    let mut index = 0;

    while index < components.len() {
        match components[index] {
            Component::Prefix(prefix) => {
                for root in &mut roots {
                    root.push(prefix.as_os_str());
                }
                index += 1;
            }
            Component::RootDir => {
                for root in &mut roots {
                    root.push(Component::RootDir.as_os_str());
                }
                index += 1;
            }
            Component::CurDir => {
                index += 1;
            }
            Component::ParentDir => {
                for root in &mut roots {
                    root.push("..");
                }
                index += 1;
            }
            Component::Normal(_) => break,
        }
    }

    let mut matches = Vec::new();
    expand_host_wildcard_components(&roots, &components[index..], &mut matches, case_insensitive)?;
    matches.sort_by_key(|path| wildcard_casefold(&path.to_string_lossy()));
    matches.dedup();
    Ok(matches)
}

fn expand_host_wildcard_components(
    bases: &[PathBuf],
    components: &[Component<'_>],
    matches: &mut Vec<PathBuf>,
    case_insensitive: bool,
) -> std::io::Result<()> {
    if components.is_empty() {
        matches.extend(bases.iter().cloned());
        return Ok(());
    }

    let component = components[0];
    match component {
        Component::CurDir => {
            expand_host_wildcard_components(bases, &components[1..], matches, case_insensitive)
        }
        Component::ParentDir => {
            let next = bases
                .iter()
                .map(|base| {
                    let mut path = base.clone();
                    path.push("..");
                    path
                })
                .collect::<Vec<_>>();
            expand_host_wildcard_components(&next, &components[1..], matches, case_insensitive)
        }
        Component::Normal(part) => {
            let text = part.to_string_lossy();
            let has_wildcards = text.contains('*') || text.contains('?');
            if !has_wildcards {
                let next = bases
                    .iter()
                    .map(|base| {
                        let mut path = base.clone();
                        path.push(part);
                        path
                    })
                    .collect::<Vec<_>>();
                return expand_host_wildcard_components(
                    &next,
                    &components[1..],
                    matches,
                    case_insensitive,
                );
            }

            let mut next = Vec::new();
            for base in bases {
                let dir = if base.as_os_str().is_empty() {
                    PathBuf::from(".")
                } else {
                    base.clone()
                };
                let mut names = BTreeMap::new();
                for entry in fs::read_dir(&dir)? {
                    let entry = entry?;
                    let entry_path = entry.path();
                    let entry_name = entry.file_name();
                    if !wildcard_match(&text, &entry_name.to_string_lossy(), case_insensitive) {
                        continue;
                    }
                    if case_insensitive
                        && let Some(first) = names.insert(
                            entry_name.to_string_lossy().to_lowercase(),
                            entry_name.clone(),
                        )
                    {
                        return Err(std::io::Error::new(
                            std::io::ErrorKind::AlreadyExists,
                            format!("both {first:?} and {entry_name:?} match {text:?}"),
                        ));
                    }
                    if components.len() > 1 && !entry.file_type()?.is_dir() {
                        continue;
                    }
                    next.push(entry_path);
                }
            }
            expand_host_wildcard_components(&next, &components[1..], matches, case_insensitive)
        }
        Component::Prefix(_) | Component::RootDir => {
            unreachable!("prefix/root handled before recursion")
        }
    }
}

fn wildcard_match(pattern: &str, candidate: &str, case_insensitive: bool) -> bool {
    let fold = |value: &str| {
        if case_insensitive {
            value.to_lowercase()
        } else {
            wildcard_casefold(value)
        }
    };
    let pattern: Vec<char> = fold(pattern).chars().collect();
    let candidate: Vec<char> = fold(candidate).chars().collect();
    let mut dp = vec![vec![false; candidate.len() + 1]; pattern.len() + 1];
    dp[0][0] = true;

    for i in 0..pattern.len() {
        if pattern[i] == '*' {
            dp[i + 1][0] = dp[i][0];
        }
        for j in 0..candidate.len() {
            dp[i + 1][j + 1] = match pattern[i] {
                '*' => dp[i][j + 1] || dp[i + 1][j] || dp[i][j],
                '?' => dp[i][j],
                ch => dp[i][j] && ch == candidate[j],
            };
        }
    }

    dp[pattern.len()][candidate.len()]
}

fn wildcard_casefold(text: &str) -> String {
    #[cfg(target_os = "windows")]
    {
        text.to_ascii_lowercase()
    }
    #[cfg(not(target_os = "windows"))]
    {
        text.to_string()
    }
}

#[cfg(test)]
mod record_io_tests {
    use super::{MODE_BINARY, MODE_RANDOM, deserialize_record, record_size, serialize_record};
    use calcvba_runtime::{VarType, Variant};

    #[test]
    fn udt_file_bytes_omit_native_padding_and_encode_fixed_strings_as_ansi() {
        use calcvba_runtime::vba_record::{
            VbaRecord, VbaRecordFieldKind as F, VbaRecordFieldSpec, VbaRecordLayout,
        };
        let layout = std::sync::Arc::new(
            VbaRecordLayout::new(vec![
                VbaRecordFieldSpec::anonymous(F::Long),
                VbaRecordFieldSpec::anonymous(F::FixedString { len: 6 }),
            ])
            .unwrap(),
        );
        let mut record = VbaRecord::new_default(layout.clone()).unwrap();
        record
            .write_field_variant(0, &Variant::from_i32(29))
            .unwrap();
        record
            .write_field_variant(1, &Variant::from_string("abc"))
            .unwrap();
        let source = Variant::from_vba_record(record);
        let kind = F::Record(layout);
        let bytes = super::serialize_udt_field(&kind, &source).unwrap();
        assert_eq!(bytes, [29, 0, 0, 0, b'a', b'b', b'c', b' ', b' ', b' ']);
        let mut target = source.clone();
        target
            .write_record_field_variant(0, &Variant::from_i32(99))
            .unwrap();
        target
            .write_record_field_variant(1, &Variant::from_string("xxxxxx"))
            .unwrap();
        let result = super::deserialize_udt_field(&kind, &target, &bytes).unwrap();
        assert_eq!(
            result.read_record_field_variant(0).unwrap().as_i32(),
            Some(29)
        );
        assert_eq!(
            result
                .read_record_field_variant(1)
                .unwrap()
                .as_bstr()
                .unwrap()
                .as_str(),
            "abc   "
        );
    }

    #[test]
    fn long_serializes_little_endian_and_round_trips() {
        let bytes = serialize_record(&Variant::from_i32(0x1234_5678), MODE_RANDOM, false);
        assert_eq!(bytes, vec![0x78, 0x56, 0x34, 0x12]);
        assert_eq!(
            deserialize_record(VarType::Long, &bytes).as_i32(),
            Some(0x1234_5678)
        );
    }

    #[test]
    fn fixed_record_sizes_match_vba() {
        assert_eq!(record_size(VarType::Byte), Some(1));
        assert_eq!(record_size(VarType::Integer), Some(2));
        assert_eq!(record_size(VarType::Long), Some(4));
        assert_eq!(record_size(VarType::Double), Some(8));
        assert_eq!(record_size(VarType::String), None);
    }

    #[test]
    fn double_and_boolean_round_trip() {
        let d = serialize_record(&Variant::from_f64(3.5), MODE_RANDOM, false);
        assert_eq!(d.len(), 8);
        assert_eq!(deserialize_record(VarType::Double, &d).as_f64(), Some(3.5));
        let b = serialize_record(&Variant::from_bool(true), MODE_RANDOM, false);
        assert_eq!(b, vec![0xFF, 0xFF]);
        assert_eq!(
            deserialize_record(VarType::Boolean, &b).as_bool(),
            Some(true)
        );
    }

    #[test]
    fn string_framing_varies_by_mode_and_fixedness() {
        // Variable string, Random: 2-byte ANSI length prefix + ANSI bytes.
        let r = serialize_record(&Variant::from_string("Hi"), MODE_RANDOM, false);
        assert_eq!(r, vec![0x02, 0x00, b'H', b'i']);
        assert_eq!(
            deserialize_record(VarType::String, &r[2..])
                .as_bstr()
                .unwrap()
                .as_str(),
            "Hi"
        );
        // Variable string, Binary: raw ANSI, no length prefix.
        assert_eq!(
            serialize_record(&Variant::from_string("Hi"), MODE_BINARY, false),
            vec![b'H', b'i']
        );
        // Fixed-length string (already padded to N): raw ANSI, no prefix, even in Random.
        let f = serialize_record(&Variant::from_string("ab  "), MODE_RANDOM, true);
        assert_eq!(f, vec![b'a', b'b', b' ', b' ']);
    }
}

/// File-system behaviour measured in Excel (dc-ez3rd, probe module SpecStmt): VBA error
/// numbers, default record length, `FileAttr`, Binary `EOF` and `Dir` attributes, run on
/// the real file system under a temp directory.
#[cfg(test)]
mod vba_file_tests {
    use crate::{
        adapters::standard::StandardHostServices,
        model::{HostPolicy, native_host_profile},
        traits::{FileSystemHal, ProcessEnvHal},
    };
    use calcvba_runtime::Variant;
    use std::path::PathBuf;

    struct Scratch(PathBuf);

    impl Scratch {
        fn new(tag: &str) -> Self {
            let dir = std::env::temp_dir().join(format!("cvp-hal-{tag}-{}", std::process::id()));
            let _ = std::fs::remove_dir_all(&dir);
            std::fs::create_dir_all(&dir).unwrap();
            Scratch(dir)
        }
        fn path(&self, name: &str) -> Variant {
            Variant::from_string(self.0.join(name).to_string_lossy().as_ref())
        }
    }

    impl Drop for Scratch {
        fn drop(&mut self) {
            let _ = std::fs::remove_dir_all(&self.0);
        }
    }

    fn host() -> StandardHostServices {
        StandardHostServices::new(native_host_profile(), HostPolicy::native())
    }

    fn code<T: std::fmt::Debug>(result: crate::error::HalResult<T>) -> Option<i32> {
        result.expect_err("expected a fault").host_error_code
    }

    /// `Open path For <mode> As #handle [Len = len]` (mode codes: Input 0, Output 1, Binary 3,
    /// Random 4).
    fn open(
        host: &StandardHostServices,
        path: Variant,
        mode: i32,
        handle: i32,
        len: i32,
    ) -> crate::error::HalResult<Variant> {
        host.open_with_record_len(
            path,
            Variant::from_i32((handle << 16) | mode),
            Variant::from_i32(len),
        )
    }

    fn int(n: i32) -> Variant {
        Variant::from_i32(n)
    }

    #[test]
    fn append_starts_a_new_print_column_and_null_width_raises_94() {
        let dir = Scratch::new("append-width");
        let host = host();
        std::fs::write(dir.0.join("a.txt"), b"ab").unwrap();
        open(&host, dir.path("a.txt"), super::MODE_APPEND, 1, 0).unwrap();
        assert_eq!(host.print_column_variant(int(1)).unwrap(), int(0));
        host.width_variant(int(1), int(3)).unwrap();
        assert_eq!(code(host.width_variant(int(1), Variant::null())), Some(94));
        assert_eq!(host.print_width_variant(int(1)).unwrap(), int(3));
        host.close_variant(int(1)).unwrap();
    }

    #[test]
    fn file_faults_carry_vba_error_numbers() {
        let dir = Scratch::new("codes");
        let host = host();
        assert_eq!(code(open(&host, dir.path("none.txt"), 0, 1, 0)), Some(53));
        open(&host, dir.path("a.txt"), 1, 1, 0).unwrap();
        assert_eq!(code(open(&host, dir.path("b.txt"), 1, 1, 0)), Some(55));
        assert_eq!(code(open(&host, dir.path("b.txt"), 1, 0, 0)), Some(52));
        assert_eq!(code(open(&host, dir.path("b.txt"), 1, 600, 0)), Some(52));
        assert_eq!(code(host.eof_variant(int(9))), Some(52));
        assert_eq!(code(host.close_variant(int(9))), Some(52));
        host.close_variant(int(0)).unwrap();

        open(&host, dir.path("a.txt"), 3, 1, 0).unwrap();
        assert_eq!(code(host.seek_variant(int(1), int(0))), Some(63));
        host.close_variant(int(1)).unwrap();
        open(&host, dir.path("r.txt"), 4, 1, 2).unwrap();
        let put = host.put_record_variant(int(1), int(1), Variant::from_string("abcdef"), int(0));
        assert_eq!(code(put), Some(59));
        host.close_variant(int(1)).unwrap();
        assert_eq!(code(open(&host, dir.path("r.txt"), 4, 1, 40000)), Some(6));

        assert_eq!(code(host.kill_variant(dir.path("none.txt"))), Some(53));
        host.mkdir_variant(dir.path("d")).unwrap();
        assert_eq!(code(host.mkdir_variant(dir.path("d"))), Some(75));
        open(&host, dir.path("d/x.txt"), 1, 1, 0).unwrap();
        host.close_variant(int(1)).unwrap();
        assert_eq!(code(host.rmdir_variant(dir.path("d"))), Some(75));
        assert_eq!(code(host.ch_dir_variant(dir.path("nodir"))), Some(76));
        assert_eq!(code(host.file_len_variant(dir.path("none.txt"))), Some(53));
    }

    #[test]
    fn random_file_defaults_to_128_byte_records() {
        let dir = Scratch::new("random");
        let host = host();
        open(&host, dir.path("r.txt"), 4, 1, 0).unwrap();
        // A string is a 2-byte length plus the data, so record 2 ends at 128 + 3.
        host.put_record_variant(int(1), int(2), Variant::from_string("x"), int(0))
            .unwrap();
        assert_eq!(host.lof_variant(int(1)).unwrap().as_i32(), Some(131));
    }

    #[test]
    fn file_attr_reports_the_open_mode() {
        let dir = Scratch::new("attr");
        let host = host();
        for (mode, want) in [(0, 1), (1, 2), (2, 8), (3, 32), (4, 4)] {
            let name = format!("f{mode}.txt");
            std::fs::write(dir.0.join(&name), b"x").unwrap();
            open(&host, dir.path(&name), mode, 2, 0).unwrap();
            assert_eq!(
                host.file_attr_variant(int(2), int(1)).unwrap().as_i32(),
                Some(want),
                "mode {mode}"
            );
            host.close_variant(int(2)).unwrap();
        }
    }

    #[test]
    fn binary_eof_is_set_by_a_short_get() {
        let dir = Scratch::new("eof");
        let host = host();
        std::fs::write(dir.0.join("b.bin"), [1u8]).unwrap();
        open(&host, dir.path("b.bin"), 3, 2, 0).unwrap();
        let byte = int(17); // Byte
        host.get_record_variant(int(2), Variant::empty(), byte.clone(), int(0))
            .unwrap();
        // The only byte was read whole, so EOF is still False.
        assert_eq!(host.eof_variant(int(2)).unwrap().as_i32(), Some(0));
        host.get_record_variant(int(2), Variant::empty(), byte, int(0))
            .unwrap();
        assert_eq!(host.eof_variant(int(2)).unwrap().as_i32(), Some(1));
    }

    fn dir_all(host: &StandardHostServices, pattern: Variant, attrs: i32) -> Vec<String> {
        let mut names = Vec::new();
        let mut next = host.dir_variant(pattern, int(attrs)).unwrap();
        while let Some(name) = next.as_bstr().filter(|n| !n.is_empty()) {
            names.push(name.as_str().to_string());
            next = host.dir_variant(Variant::empty(), int(0)).unwrap();
        }
        names
    }

    #[test]
    fn dir_filters_by_attribute_and_lists_a_folder_only_with_a_trailing_separator() {
        let dir = Scratch::new("dir");
        let host = host();
        std::fs::create_dir(dir.0.join("sub")).unwrap();
        std::fs::write(dir.0.join("a.txt"), b"").unwrap();
        let star =
            |suffix: &str| Variant::from_string(format!("{}/{suffix}", dir.0.display()).as_str());
        assert_eq!(dir_all(&host, star("*"), 0), ["a.txt"]);
        assert_eq!(dir_all(&host, star("*"), 16), [".", "..", "a.txt", "sub"]);
        // A bare folder name finds nothing; with a trailing separator it lists the contents.
        let folder = dir.path("");
        let bare = Variant::from_string(dir.0.to_string_lossy().as_ref());
        assert_eq!(dir_all(&host, bare, 0), Vec::<String>::new());
        assert_eq!(dir_all(&host, folder, 0), ["a.txt"]);
        // A bare `Dir()` after the listing ended has no pattern to continue (error 5).
        assert_eq!(code(host.dir_variant(Variant::empty(), int(0))), Some(5));
    }
}
