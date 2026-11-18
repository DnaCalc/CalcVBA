//! Pure (host-free) base-library bodies: strings, math, date/time arithmetic,
//! conversion, random, financial, information, and a first-cut Collection.
//!
//! The `Format` engine lives in [`crate::format`]. Keyed `Collection` access is
//! implemented (see `calcvba-runtime`/`calcvba-eval` `collection`); the remaining
//! `FIDELITY:` marker is `StrConv`'s CJK/encoding modes.

use crate::{
    LibContext, LibError, LibResult, alloc_count, as_f64, as_i32, as_i64, as_str, as_usize, need,
    opt, vbool, vf64, vi32, vstr, vunit,
};
use calcvba_runtime::{
    Decimal96, Variant, default_error_message,
    safe_array::{SafeArray, VT_BSTR_VALUE, VT_UI1_VALUE, VT_VARIANT_VALUE},
    variant::VarType,
};
// The VBA serial ↔ civil calendar math is canonical in `calcvba_runtime::vba_date`; re-export it
// at `crate::pure::*` so this module's date functions (and `format.rs`) keep their call sites.
pub(crate) use calcvba_runtime::{
    civil_from_days, days_from_civil, days_in_month, serial_to_hms, serial_to_ymd, ymd_to_serial,
};

const VBA_DATE_MIN_SERIAL: f64 = -657_434.0; // 0100-01-01
const VBA_DATE_MAX_EXCLUSIVE_SERIAL: f64 = 2_958_466.0; // 10000-01-01

// ── Strings ──────────────────────────────────────────────────────────────────

pub fn len(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    if let Some(len) = value.vba_record_len() {
        return Ok(vi32(len as i32));
    }
    // Count UTF-16 CODE UNITS (VBA's string length): read them directly off a String
    // Variant's BSTR so a lone/paired surrogate half counts as one unit each (a UTF-8
    // round-trip via `as_str` would fold a lone surrogate to U+FFFD). Non-strings coerce.
    let units = match value.string_units() {
        Some(units) => units.len(),
        None => as_str(value)?.encode_utf16().count(),
    };
    Ok(vi32(units as i32))
}

/// Optional trailing compare-mode argument: 0 = binary (default), 1 = text. The
/// front-end supplies that mode as a trailing argument here, resolved from the
/// source's `Option Compare` (the binder injects `1` for `Option Compare Text`
/// when the call omits an explicit compare).
fn text_compare(args: &[Variant], index: usize) -> LibResult<bool> {
    match arg_present(args, index) {
        Some(v) => Ok(as_i32(v)? == 1),
        None => Ok(false),
    }
}

/// True for the `Missing` optional-argument sentinel the VM passes for an omitted
/// optional (`CoreArg::Omitted` → vbError `DISP_E_PARAMNOTFOUND`).
fn missing_sentinel(v: &Variant) -> bool {
    const MISSING_ARG: i32 = 0x8002_0004u32 as i32;
    v.vtype() == VarType::Error && v.as_error_code() == Some(MISSING_ARG)
}

/// An optional argument that is actually present — not past the end, not `Empty`,
/// and not the `Missing` sentinel. This lets the binder pad an omitted
/// *intermediate* optional (so a synthesized trailing `compare` lands at the right
/// index) without the lib mistaking the pad for a real value: an omitted optional
/// in such a padded slot arrives as `Empty`/Missing and must read as absent.
fn arg_present(args: &[Variant], index: usize) -> Option<&Variant> {
    opt(args, index).filter(|v| !matches!(v.vtype(), VarType::Empty) && !missing_sentinel(v))
}

fn norm_compare(s: String, text: bool) -> String {
    if !text {
        return s;
    }
    // Case-fold each character on its own (Unicode lower case), so byte offsets found in the
    // folded text stay valid in the original unless a fold changes the encoded length.
    s.chars()
        .map(|c| {
            let mut lower = c.to_lowercase();
            match (lower.next(), lower.next()) {
                (Some(l), None) if l.len_utf8() == c.len_utf8() => l,
                _ => c,
            }
        })
        .collect()
}

/// The operand's UTF-16 code units for comparison: a String reads its BSTR units VERBATIM
/// (lone surrogate halves preserved); a non-string coerces through its VBA string form
/// (which carries no surrogates). Mirrors the VM operator path (`arith::cmp_order`) so
/// `StrComp` distinguishes distinct lone surrogate halves instead of folding both to U+FFFD
/// via a lossy `as_str` round-trip — VBA strings are binary UTF-16 code-unit sequences.
fn compare_units(value: &Variant) -> LibResult<Vec<u16>> {
    match value.string_units() {
        Some(units) => Ok(units),
        None => Ok(as_str(value)?.encode_utf16().collect()),
    }
}

fn string_units(value: &Variant) -> LibResult<Vec<u16>> {
    match value.string_units() {
        Some(units) => Ok(units),
        None => Ok(as_str(value)?.encode_utf16().collect()),
    }
}

fn string_bytes(value: &Variant) -> LibResult<Vec<u8>> {
    match value.string_bytes() {
        Some(bytes) => Ok(bytes),
        None => Ok(as_str(value)?
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect()),
    }
}

pub fn left(args: &[Variant]) -> LibResult<Variant> {
    let units = string_units(need(args, 0)?)?;
    let n = as_usize(need(args, 1)?)?;
    Ok(Variant::from_utf16_units(&units[..n.min(units.len())]))
}

pub fn left_b(args: &[Variant]) -> LibResult<Variant> {
    let bytes = string_bytes(need(args, 0)?)?;
    let n = as_usize(need(args, 1)?)?;
    Ok(Variant::from_bstr_bytes(&bytes[..n.min(bytes.len())]))
}

/// `AscB(s)` — the first BYTE of the string's UTF-16LE storage, as a `Byte`.
pub fn asc_b(args: &[Variant]) -> LibResult<Variant> {
    let bytes = string_bytes(need(args, 0)?)?;
    let first = bytes
        .first()
        .copied()
        .ok_or_else(|| LibError::invalid_call("AscB of empty string"))?;
    Ok(Variant::from_u8(first))
}

/// `ChrB(code)` — a one-BYTE string (`LenB` 1, `Len` 0); `code` is `0..=255`.
pub fn chr_b(args: &[Variant]) -> LibResult<Variant> {
    let code = as_i32(need(args, 0)?)?;
    let byte = u8::try_from(code).map_err(|_| LibError::invalid_call("invalid character code"))?;
    Ok(Variant::from_bstr_bytes(&[byte]))
}

/// `MidB(s, start, [length])` — `start` and `length` count BYTES.
pub fn mid_b(args: &[Variant]) -> LibResult<Variant> {
    let bytes = string_bytes(need(args, 0)?)?;
    let start = one_based_mid_start(need(args, 1)?)?;
    let from = (start - 1).min(bytes.len());
    let take = match arg_present(args, 2) {
        Some(v) => as_usize(v)?,
        None => bytes.len() - from,
    };
    let to = from.saturating_add(take).min(bytes.len());
    Ok(Variant::from_bstr_bytes(&bytes[from..to]))
}

/// `InStrB([start], s1, s2, [compare])` — the 1-based BYTE position of `s2` in `s1`.
pub fn instr_b(args: &[Variant]) -> LibResult<Variant> {
    let has_start = args.len() >= 3;
    let base = usize::from(has_start);
    let start = if has_start {
        as_i32(need(args, 0)?)?
    } else {
        1
    };
    if start < 1 {
        return Err(LibError::invalid_call("InStrB start must be >= 1"));
    }
    let text = text_compare(args, base + 2)?;
    let bytes = |index: usize| -> LibResult<Vec<u8>> {
        let value = need(args, index)?;
        if text {
            Ok(norm_compare(as_str(value)?, true)
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect())
        } else {
            string_bytes(value)
        }
    };
    let hay = bytes(base)?;
    let needle = bytes(base + 1)?;
    let from = start as usize - 1;
    if from > hay.len() || needle.len() > hay.len() - from {
        return Ok(vi32(0));
    }
    if needle.is_empty() {
        return Ok(vi32(0));
    }
    let found = hay[from..].windows(needle.len()).position(|w| w == needle);
    Ok(vi32(found.map_or(0, |at| (from + at + 1) as i32)))
}

pub fn right(args: &[Variant]) -> LibResult<Variant> {
    let units = string_units(need(args, 0)?)?;
    let n = as_usize(need(args, 1)?)?.min(units.len());
    Ok(Variant::from_utf16_units(&units[units.len() - n..]))
}

pub fn right_b(args: &[Variant]) -> LibResult<Variant> {
    let bytes = string_bytes(need(args, 0)?)?;
    let n = as_usize(need(args, 1)?)?.min(bytes.len());
    Ok(Variant::from_bstr_bytes(&bytes[bytes.len() - n..]))
}

pub fn mid(args: &[Variant]) -> LibResult<Variant> {
    let units = string_units(need(args, 0)?)?;
    let start = one_based_mid_start(need(args, 1)?)?;
    let from = (start - 1).min(units.len());
    let take = match opt(args, 2) {
        Some(v) => as_usize(v)?,
        None => units.len() - from,
    };
    let to = from.saturating_add(take).min(units.len());
    Ok(Variant::from_utf16_units(&units[from..to]))
}

pub fn mid_stmt(args: &[Variant]) -> LibResult<Variant> {
    let mut units = string_units(need(args, 0)?)?;
    let start = one_based_mid_start(need(args, 1)?)? - 1;
    // A start past the end of the target string is an error, not a no-op.
    if start >= units.len() {
        return Err(LibError::invalid_call(
            "Mid start is past the end of the string",
        ));
    }
    // args = [target, start, value] or [target, start, count, value]
    let (count, value) = if args.len() >= 4 {
        (
            Some(as_usize(need(args, 2)?)?),
            string_units(need(args, 3)?)?,
        )
    } else {
        (None, string_units(need(args, 2)?)?)
    };
    let limit = count.unwrap_or(value.len()).min(value.len());
    for (i, unit) in value.into_iter().take(limit).enumerate() {
        if let Some(index) = start.checked_add(i)
            && index < units.len()
        {
            units[index] = unit;
        }
    }
    Ok(Variant::from_utf16_units(&units))
}

/// `MidB(s, start, [count]) = value` — as [`mid_stmt`], with `start` and `count` in bytes.
pub fn mid_b_stmt(args: &[Variant]) -> LibResult<Variant> {
    let mut bytes = string_bytes(need(args, 0)?)?;
    let start = one_based_mid_start(need(args, 1)?)? - 1;
    if start >= bytes.len() {
        return Err(LibError::invalid_call(
            "MidB start is past the end of the string",
        ));
    }
    let (count, value) = if args.len() >= 4 {
        (
            Some(as_usize(need(args, 2)?)?),
            string_bytes(need(args, 3)?)?,
        )
    } else {
        (None, string_bytes(need(args, 2)?)?)
    };
    let limit = count.unwrap_or(value.len()).min(value.len());
    for (i, byte) in value.into_iter().take(limit).enumerate() {
        if let Some(slot) = bytes.get_mut(start + i) {
            *slot = byte;
        }
    }
    Ok(Variant::from_bstr_bytes(&bytes))
}

pub fn lset_stmt(args: &[Variant]) -> LibResult<Variant> {
    align_stmt(args, false)
}

pub fn rset_stmt(args: &[Variant]) -> LibResult<Variant> {
    align_stmt(args, true)
}

/// `LSet s = value` / `RSet s = value` for string targets. The first argument is the
/// target's current value; its UTF-16 length is the alignment width for variable
/// strings, while fixed-length strings naturally carry their declared width.
fn align_stmt(args: &[Variant], right_align: bool) -> LibResult<Variant> {
    let width = string_units(need(args, 0)?)?.len();
    let mut value = string_units(need(args, 1)?)?;
    if value.len() > width {
        value.truncate(width);
    }
    let pad = width.saturating_sub(value.len());
    let mut out = Vec::with_capacity(width);
    if right_align {
        out.extend(std::iter::repeat_n(0x20, pad));
        out.extend(value);
    } else {
        out.extend(value);
        out.extend(std::iter::repeat_n(0x20, pad));
    }
    Ok(Variant::from_utf16_units(&out))
}

fn one_based_mid_start(value: &Variant) -> LibResult<usize> {
    let start = as_i64(value)?;
    if start < 1 {
        return Err(LibError::invalid_call(
            "Mid start must be greater than or equal to 1",
        ));
    }
    usize::try_from(start).map_err(|_| LibError::overflow("Mid start overflow"))
}

/// `InStr([start], string1, string2, [compare])` and
/// `InStrRev(string1, string2, [compare])`.
///
/// For `InStr`, an optional leading numeric-typed `start` (1-based) is detected
/// by argument type — the legacy 2-operand form passes a string-typed
/// `string1`, so this is unambiguous. The trailing compare mode (0 = binary,
/// 1 = text → ASCII case-insensitive) is supplied by the lowering from the
/// source's `Option Compare`. (`InStrRev`'s own `start` argument awaits the
/// canonical arg layout.)
pub fn instr(args: &[Variant], rev: bool) -> LibResult<Variant> {
    if rev {
        return instr_rev(args);
    }
    // `InStr([start], string1, string2, [compare])`. Per VBA `start` is detected by ARITY,
    // not by type: 2 args = `(string1, string2)`; 3–4 args = `(start, string1, string2,
    // [compare])` — `compare` is legal only when `start` is present. (The old code treated a
    // numeric first arg as `start`, mis-binding e.g. `InStr(numericVar, s)`.)
    let has_start = args.len() >= 3;
    let base = usize::from(has_start);
    let start = if has_start {
        as_i32(need(args, 0)?)?
    } else {
        1
    };
    if start < 1 {
        return Err(LibError::invalid_call("InStr start must be >= 1"));
    }
    let text = text_compare(args, base + 2)?;
    let hay = norm_compare(as_str(need(args, base)?)?, text);
    let needle = norm_compare(as_str(need(args, base + 1)?)?, text);
    // Nothing is found in an empty string, whatever the search text.
    if hay.is_empty() {
        return Ok(vi32(0));
    }

    let byte_start = utf16_index_to_byte(&hay, start as usize - 1);
    if byte_start > hay.len() {
        return Ok(vi32(0));
    }
    let pos = hay[byte_start..].find(&needle).map(|b| byte_start + b);
    Ok(vi32(match pos {
        Some(byte) => hay[..byte].encode_utf16().count() as i32 + 1,
        None => 0,
    }))
}

/// `InStrRev(stringcheck, stringmatch, [start = -1], [compare])` — the position (1-based, from
/// the START of the string) of the LAST occurrence of `stringmatch` whose match lies at or
/// before `start`; `-1` searches from the end. NOTE the arg layout differs from `InStr`:
/// `start` is the 3rd arg and `compare` the 4th (the old code ignored `start` and misread
/// arg 2 as the compare mode).
fn instr_rev(args: &[Variant]) -> LibResult<Variant> {
    let text = text_compare(args, 3)?;
    let hay = norm_compare(as_str(need(args, 0)?)?, text);
    let needle = norm_compare(as_str(need(args, 1)?)?, text);
    let hay_units = hay.encode_utf16().count();
    let start = match arg_present(args, 2) {
        Some(v) => as_i32(v)?,
        None => -1,
    };
    if start == 0 || start < -1 {
        return Err(LibError::invalid_call("InStrRev start must be -1 or >= 1"));
    }
    if start != -1 && start as usize > hay_units {
        return Ok(vi32(0));
    }
    // Search the first `start` code units (the match must end at or before `start`).
    let region_units = if start == -1 {
        hay_units
    } else {
        (start as usize).min(hay_units)
    };
    let region_end = utf16_index_to_byte(&hay, region_units);
    let region = &hay[..region_end];
    Ok(vi32(match region.rfind(&needle) {
        Some(byte) => hay[..byte].encode_utf16().count() as i32 + 1,
        None => 0,
    }))
}

/// Byte offset of the char at a given UTF-16 code-unit index (VBA string index).
fn utf16_index_to_byte(s: &str, utf16_index: usize) -> usize {
    let mut units = 0;
    for (byte, ch) in s.char_indices() {
        if units >= utf16_index {
            return byte;
        }
        units += ch.len_utf16();
    }
    s.len()
}

pub fn lcase(args: &[Variant]) -> LibResult<Variant> {
    Ok(vstr(as_str(need(args, 0)?)?.to_lowercase()))
}
pub fn ucase(args: &[Variant]) -> LibResult<Variant> {
    Ok(vstr(as_str(need(args, 0)?)?.to_uppercase()))
}

pub fn split(args: &[Variant]) -> LibResult<Variant> {
    let s = as_str(need(args, 0)?)?;
    let delim = match arg_present(args, 1) {
        Some(v) => as_str(v)?,
        None => " ".to_string(),
    };
    let limit = match arg_present(args, 2) {
        Some(v) => as_i32(v)?,
        None => -1,
    };
    if limit < -1 {
        return Err(LibError::invalid_call("Split limit must be -1 or greater"));
    }
    let text = text_compare(args, 3)?;
    let parts: Vec<Variant> = if s.is_empty() || limit == 0 {
        // An empty string, or a zero limit, gives an empty (`LBound 0, UBound -1`) array.
        Vec::new()
    } else if delim.is_empty() {
        // An empty delimiter never splits — the whole string is the single element.
        vec![vstr(s)]
    } else {
        split_with_limit(&s, &delim, limit, text)
    };
    string_array(parts)
}

/// A `String()` array (the element type `Split` and `Filter` return, not `Variant()`).
fn string_array(parts: Vec<Variant>) -> LibResult<Variant> {
    SafeArray::from_typed_variants(VT_BSTR_VALUE, parts)
        .map(Variant::from_safearray)
        .map_err(LibError::type_mismatch)
}

/// Split `s` on `delim` into at most `limit` parts (negative = unlimited), matching the
/// delimiter case-insensitively when `text`. Substrings are taken from the ORIGINAL string
/// (original case preserved); only the delimiter *positions* are found case-insensitively.
fn split_with_limit(s: &str, delim: &str, limit: i32, text: bool) -> Vec<Variant> {
    // For binary compare `hay`/`needle` equal the originals, so byte offsets are exact; for
    // text compare they are the lowercased forms (byte-length-preserving for ASCII and most
    // text — the same approximation `InStr`/`norm_compare` already make).
    let hay = norm_compare(s.to_string(), text);
    let needle = norm_compare(delim.to_string(), text);
    let mut parts: Vec<Variant> = Vec::new();
    let mut cursor = 0usize;
    loop {
        // Once `limit - 1` parts are emitted, the final element holds the whole remainder.
        if limit >= 0 && parts.len() as i32 == limit - 1 {
            break;
        }
        match hay[cursor..].find(&needle) {
            Some(rel) => {
                let pos = cursor + rel;
                parts.push(vstr(s[cursor..pos].to_string()));
                cursor = pos + needle.len();
            }
            None => break,
        }
    }
    parts.push(vstr(s[cursor..].to_string()));
    parts
}

/// `Filter(source, match, [include=True], [compare=binary])` — the subset of a
/// string array whose elements contain (or, when `include` is False, omit) the
/// `match` substring, returned as a fresh 0-based array.
pub fn filter(args: &[Variant]) -> LibResult<Variant> {
    let array = need(args, 0)?
        .as_safearray()
        .ok_or_else(|| LibError::type_mismatch("Filter expects a source array"))?;
    let needle = as_str(need(args, 1)?)?;
    let include = match arg_present(args, 2) {
        Some(v) => as_bool_lenient(v),
        None => true,
    };
    let text = text_compare(args, 3)?;
    let needle_cmp = norm_compare(needle, text);
    let mut out: Vec<Variant> = Vec::new();
    for element in array.variant_elements().unwrap_or_default() {
        let s = as_str(&element)?;
        let contains = norm_compare(s, text).contains(&needle_cmp);
        if contains == include {
            out.push(element);
        }
    }
    string_array(out)
}

pub fn join(args: &[Variant]) -> LibResult<Variant> {
    let array = need(args, 0)?
        .as_safearray()
        .ok_or_else(|| LibError::type_mismatch("Join expects an array"))?;
    let delim = match opt(args, 1) {
        Some(v) => as_str(v)?,
        None => " ".to_string(),
    };
    let parts = array.variant_elements().unwrap_or_default();
    let texts: LibResult<Vec<String>> = parts.iter().map(as_str).collect();
    Ok(vstr(texts?.join(&delim)))
}

/// `Replace(expr, find, replace, [start], [count], [compare])`. Per VBA the
/// result is the substring of `expr` FROM `start` (1-based; the prefix before
/// `start` is dropped), with up to `count` replacements (default -1 = all) of
/// `find` by `replace`, matched binary or text per `compare` (default binary).
pub fn replace(args: &[Variant]) -> LibResult<Variant> {
    let s = as_str(need(args, 0)?)?;
    let find = as_str(need(args, 1)?)?;
    let with = as_str(need(args, 2)?)?;
    let start = match numeric_arg(args, 3) {
        Some(v) => as_i32(v)?,
        None => 1,
    };
    let count = match numeric_arg(args, 4) {
        Some(v) => as_i32(v)?,
        None => -1,
    };
    let text = text_compare(args, 5)?;
    if start < 1 || count < -1 {
        // VBA "Invalid procedure call or argument" (error 5).
        return Err(LibError::invalid_call("Replace start/count out of range"));
    }

    // The result begins at `start` (1-based); a start past the end yields "".
    let chars: Vec<char> = s.chars().collect();
    let from = (start - 1) as usize;
    if from >= chars.len() {
        return Ok(vstr(String::new()));
    }
    let hay: Vec<char> = chars[from..].to_vec();

    if find.is_empty() || count == 0 {
        return Ok(vstr(hay.into_iter().collect::<String>()));
    }

    let needle: Vec<char> = find.chars().collect();
    let needle_cmp: Vec<char> = norm_compare(find.clone(), text).chars().collect();
    let limit = if count < 0 {
        usize::MAX
    } else {
        count as usize
    };
    let mut out = String::new();
    let mut i = 0usize;
    let mut done = 0usize;
    while i < hay.len() {
        let window_matches = done < limit
            && i + needle.len() <= hay.len()
            && hay[i..i + needle.len()]
                .iter()
                .zip(&needle_cmp)
                .all(|(h, n)| {
                    if text {
                        h.to_ascii_lowercase() == *n
                    } else {
                        h == n
                    }
                });
        if window_matches {
            out.push_str(&with);
            i += needle.len();
            done += 1;
        } else {
            out.push(hay[i]);
            i += 1;
        }
    }
    Ok(vstr(out))
}

pub fn trim(args: &[Variant], left: bool, right: bool) -> LibResult<Variant> {
    let s = as_str(need(args, 0)?)?;
    let t = match (left, right) {
        (true, true) => s.trim_matches(' '),
        (true, false) => s.trim_start_matches(' '),
        (false, true) => s.trim_end_matches(' '),
        (false, false) => &s,
    };
    Ok(vstr(t.to_string()))
}

/// `StrComp(s1, s2, [compare])` — optional trailing compare mode (0=binary, 1=text).
pub fn str_comp(args: &[Variant]) -> LibResult<Variant> {
    let text = text_compare(args, 2)?;
    let a = compare_units(need(args, 0)?)?;
    let b = compare_units(need(args, 1)?)?;
    Ok(vi32(
        match calcvba_runtime::string_compare::compare_utf16(&a, &b, text) {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        },
    ))
}

/// `Like` — supports `?`, `*`, `#`, `[charlist]` (with `!` negation and `a-z`
/// ranges), and an optional trailing compare mode (0=binary, 1=text).
pub fn like(args: &[Variant]) -> LibResult<Variant> {
    let text = text_compare(args, 2)?;
    let subject = string_units(need(args, 0)?)?;
    let pattern = string_units(need(args, 1)?)?;
    let tokens = like_tokens(&pattern, text)?;
    let subject: Vec<u16> = if text {
        subject.into_iter().map(fold_unit).collect()
    } else {
        subject
    };
    Ok(vbool(like_match(&subject, &tokens)))
}

/// One element of a compiled `Like` pattern. A `Like` runs over UTF-16 code units, so `?` is one
/// unit and a surrogate pair is two.
enum LikeToken {
    Any,
    One,
    Digit,
    Unit(u16),
    Set {
        negate: bool,
        ranges: Vec<(u16, u16)>,
    },
}

/// Case fold of one code unit for text comparison (Unicode lower case when it stays one unit).
fn fold_unit(unit: u16) -> u16 {
    let Some(c) = char::from_u32(u32::from(unit)) else {
        return unit;
    };
    let mut lower = c.to_lowercase();
    match (lower.next(), lower.next()) {
        (Some(l), None) if (l as u32) <= 0xFFFF => l as u32 as u16,
        _ => unit,
    }
}

fn invalid_pattern() -> LibError {
    LibError::new(93, "Invalid pattern string")
}

/// Compile a pattern (folded when `text`). An unclosed `[` or a reversed range is
/// "Invalid pattern string" (93); an empty `[]` matches the empty string.
fn like_tokens(pattern: &[u16], text: bool) -> LibResult<Vec<LikeToken>> {
    let fold = |u: u16| if text { fold_unit(u) } else { u };
    let mut tokens = Vec::new();
    let mut i = 0;
    while i < pattern.len() {
        let unit = pattern[i];
        i += 1;
        match unit {
            0x2A => tokens.push(LikeToken::Any),
            0x3F => tokens.push(LikeToken::One),
            0x23 => tokens.push(LikeToken::Digit),
            0x5B => {
                let close = pattern[i..]
                    .iter()
                    .position(|u| *u == 0x5D)
                    .map(|p| p + i)
                    .ok_or_else(invalid_pattern)?;
                let mut body = &pattern[i..close];
                i = close + 1;
                let negate = body.first() == Some(&0x21) && body.len() > 1;
                if negate {
                    body = &body[1..];
                }
                if body.is_empty() {
                    continue;
                }
                let mut ranges = Vec::new();
                let mut k = 0;
                while k < body.len() {
                    if k + 2 < body.len() && body[k + 1] == 0x2D {
                        let (lo, hi) = (fold(body[k]), fold(body[k + 2]));
                        if lo > hi {
                            return Err(invalid_pattern());
                        }
                        ranges.push((lo, hi));
                        k += 3;
                    } else {
                        ranges.push((fold(body[k]), fold(body[k])));
                        k += 1;
                    }
                }
                tokens.push(LikeToken::Set { negate, ranges });
            }
            other => tokens.push(LikeToken::Unit(fold(other))),
        }
    }
    Ok(tokens)
}

fn like_match(s: &[u16], p: &[LikeToken]) -> bool {
    let Some((first, rest)) = p.split_first() else {
        return s.is_empty();
    };
    match first {
        LikeToken::Any => (0..=s.len()).any(|i| like_match(&s[i..], rest)),
        LikeToken::One => !s.is_empty() && like_match(&s[1..], rest),
        LikeToken::Digit => {
            !s.is_empty() && (0x30..=0x39).contains(&s[0]) && like_match(&s[1..], rest)
        }
        LikeToken::Unit(u) => !s.is_empty() && s[0] == *u && like_match(&s[1..], rest),
        LikeToken::Set { negate, ranges } => {
            !s.is_empty()
                && (ranges.iter().any(|(lo, hi)| (*lo..=*hi).contains(&s[0])) != *negate)
                && like_match(&s[1..], rest)
        }
    }
}

/// Shared code-point decode for `Chr`/`ChrW`: the argument is a VBA `Long` whose
/// low 16 bits are the (wide) character code; negatives wrap via `code as u16`.
/// Rejects the UTF-16 surrogate range and otherwise returns the `char`.
fn wide_char(code: i32) -> LibResult<char> {
    if !(-32768..=65535).contains(&code) {
        return Err(LibError::invalid_call("invalid character code"));
    }
    let value = code as u16;
    char::from_u32(u32::from(value)).ok_or_else(|| LibError::invalid_call("invalid character code"))
}

/// `ChrW(code)` — genuine WIDE: the single UTF-16 CODE UNIT whose value is the
/// argument's low 16 bits (negatives wrap). Any `0..=65535` is valid — INCLUDING a
/// lone surrogate half `0xD800..=0xDFFF` — because a VBA string is a UTF-16 code-unit
/// sequence, not Unicode scalar values. Building the string directly from the unit
/// (not via `char`) is what lets `ChrW(&HD83D) & ChrW(&HDE00)` form a real astral pair.
pub fn chr_w(args: &[Variant]) -> LibResult<Variant> {
    let code = as_i32(need(args, 0)?)?;
    if !(-32768..=65535).contains(&code) {
        return Err(LibError::invalid_call("invalid character code"));
    }
    Ok(Variant::from_utf16_units(&[code as u16]))
}

/// `Chr(code)` — genuine ANSI: codes 0..=255 decode through the live system ANSI
/// code page (`CP_ACP` on Windows, the CP-1252 fallback off-Windows); codes
/// 256..=65535 act WIDE (VBA7 `Chr` is wide above 255); negatives wrap.
pub fn chr(args: &[Variant]) -> LibResult<Variant> {
    let code = as_i32(need(args, 0)?)?;
    let value = if (-32768..=65535).contains(&code) {
        code as u16
    } else {
        return Err(LibError::invalid_call("invalid character code"));
    };
    let ch = if value <= 0xFF {
        // A single ANSI byte decodes to one Unicode char via the shared codec; fall
        // back to the byte's own code point if the host ACP leaves it undefined.
        calcvba_runtime::ansi::ansi_decode(&[value as u8])
            .chars()
            .next()
            .unwrap_or(value as u8 as char)
    } else {
        wide_char(code)?
    };
    Ok(vstr(ch.to_string()))
}

/// `AscW(s)` — genuine WIDE: the first char's Unicode code point, returned as a VBA
/// `Integer` (i16) so code points > 32767 come back negative (matching VBA AscW).
pub fn asc_w(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    // Read the first UTF-16 CODE UNIT directly (so `AscW(ChrW(n))` round-trips ANY
    // 0..=65535, including a lone surrogate half a `char`-based read could not carry).
    // The i16 wrap matches VBA: code points > 32767 come back negative.
    let unit = match value.string_units() {
        Some(units) => units.first().copied(),
        None => as_str(value)?.encode_utf16().next(),
    }
    .ok_or_else(|| LibError::invalid_call("AscW of empty string"))?;
    Ok(Variant::from_i16(unit as i16))
}

/// `Asc(s)` — genuine ANSI: the first char's byte in the live system ANSI code page
/// (`CP_ACP` on Windows, the CP-1252 fallback off-Windows). For the common
/// Western/CP-1252 SBCS case this is the 0..=255 ANSI code; a char the host ACP
/// cannot represent yields 63 (`"?"`, VBA's unmappable best-fit). (DBCS lead-byte /
/// 2-byte `Asc` on a multibyte ACP is an accepted documented edge: we return the
/// first byte only.)
pub fn asc(args: &[Variant]) -> LibResult<Variant> {
    let s = as_str(need(args, 0)?)?;
    let ch = s
        .chars()
        .next()
        .ok_or_else(|| LibError::invalid_call("Asc of empty string"))?;
    let byte = calcvba_runtime::ansi::ansi_encode(&ch.to_string())
        .first()
        .copied()
        .unwrap_or(b'?');
    Ok(Variant::from_i16(i16::from(byte)))
}

pub fn space(args: &[Variant]) -> LibResult<Variant> {
    Ok(vstr(" ".repeat(alloc_count(need(args, 0)?)?)))
}

pub fn string_repeat(args: &[Variant]) -> LibResult<Variant> {
    let count = alloc_count(need(args, 0)?)?;
    let arg = need(args, 1)?;
    let ch = if arg.vtype() == VarType::String {
        // A string supplies its own first character.
        let units = string_units(arg)?;
        *units
            .first()
            .ok_or_else(|| LibError::invalid_call("String character is empty"))?
    } else {
        // A number is a character code: the low byte, read through the ANSI code page.
        let byte = as_i64(arg)?.rem_euclid(256) as u8;
        calcvba_runtime::ansi::ansi_decode(&[byte])
            .encode_utf16()
            .next()
            .unwrap_or(u16::from(byte))
    };
    Ok(Variant::from_utf16_units(&vec![ch; count]))
}

pub fn str_reverse(args: &[Variant]) -> LibResult<Variant> {
    Ok(vstr(
        as_str(need(args, 0)?)?.chars().rev().collect::<String>(),
    ))
}

/// FIDELITY: handles case modes plus the deterministic ANSI byte conversion modes
/// through the shared runtime ANSI codec. East Asian width/kana modes still pass
/// through outside relevant locales.
pub fn str_conv(args: &[Variant]) -> LibResult<Variant> {
    let mode = as_i32(need(args, 1)?)?;
    const VB_UPPER_CASE: i32 = 1;
    const VB_LOWER_CASE: i32 = 2;
    const VB_PROPER_CASE: i32 = 3;
    const VB_UNICODE: i32 = 64;
    const VB_FROM_UNICODE: i32 = 128;
    // Wide, narrow, katakana and hiragana need an East Asian locale.
    const VB_EAST_ASIAN: i32 = 4 | 8 | 16 | 32;

    // Internal mode of the binder's `Byte() = String` store: the UTF-16LE bytes of the text.
    if mode == calcvba_bundle::STRCONV_UTF16_BYTES {
        // From a Variant source, anything but a String (an array) is stored as it is, and the
        // store checks its type.
        if need(args, 0)?.vtype() != VarType::String {
            return Ok(need(args, 0)?.clone());
        }
        let bytes = string_bytes(need(args, 0)?)?
            .into_iter()
            .map(Variant::from_u8)
            .collect();
        let array =
            SafeArray::from_typed_variants(VT_UI1_VALUE, bytes).map_err(LibError::type_mismatch)?;
        return Ok(Variant::from_safearray(array));
    }

    if mode & VB_UNICODE != 0 && mode & VB_FROM_UNICODE != 0 {
        return Err(LibError::invalid_call(
            "vbUnicode and vbFromUnicode are mutually exclusive",
        ));
    }

    if mode & VB_EAST_ASIAN != 0 {
        return Err(LibError::invalid_call(
            "StrConv wide/narrow/kana modes need an East Asian locale",
        ));
    }

    let mut s = if mode & VB_UNICODE != 0 {
        // The string's own bytes are read as ANSI text, one wide character per byte.
        let input = need(args, 0)?;
        if let Some(array) = input.as_safearray() {
            calcvba_runtime::ansi::ansi_decode(&array_bytes(&array)?)
        } else {
            calcvba_runtime::ansi::ansi_decode(&string_bytes(input)?)
        }
    } else {
        as_str(need(args, 0)?)?
    };

    s = match mode & VB_PROPER_CASE {
        VB_UPPER_CASE => s.to_uppercase(),
        VB_LOWER_CASE => s.to_lowercase(),
        VB_PROPER_CASE => proper_case(&s),
        _ => s,
    };

    if mode & VB_FROM_UNICODE != 0 {
        // The ANSI bytes, packed into a String (so `LenB` counts bytes). Storing it in a
        // `Byte()` copies those bytes.
        return Ok(Variant::from_bstr_bytes(
            &calcvba_runtime::ansi::ansi_encode(&s),
        ));
    }

    Ok(vstr(s))
}

fn proper_case(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut word_start = true;
    for c in s.chars() {
        if c.is_whitespace() {
            out.push(c);
            word_start = true;
        } else if word_start {
            out.extend(c.to_uppercase());
            word_start = false;
        } else {
            out.extend(c.to_lowercase());
        }
    }
    out
}

fn array_bytes(array: &SafeArray) -> LibResult<Vec<u8>> {
    array
        .variant_elements()
        .ok_or_else(|| LibError::type_mismatch("expected Byte array"))?
        .into_iter()
        .map(|value| {
            value
                .as_u8()
                .ok_or_else(|| LibError::type_mismatch("expected Byte array"))
        })
        .collect()
}

/// `Format(expr, [mask])` — delegates to the [`crate::format`] engine (named
/// formats, custom numeric masks, custom date/time masks). With no mask, use
/// general text; Null and Empty format as empty text.
pub fn format(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    let mask = match opt(args, 1) {
        Some(v) if !matches!(v.vtype(), VarType::Empty) && !missing_sentinel(v) => as_str(v)?,
        _ => String::new(),
    };
    let first_day = first_day_of_week(args, 2)?;
    let first_week = first_week_of_year(args, 3)?;
    if matches!(value.vtype(), VarType::ArrayVariant | VarType::Error) {
        return Err(LibError::type_mismatch("Format expects a scalar value"));
    }
    Ok(vstr(crate::format::apply_with(
        value, &mask, first_day, first_week,
    )))
}

pub fn format_number(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    if let Some(early) = fixed_format_guard(value)? {
        return Ok(early);
    }
    let options = fixed_format_options(args)?;
    Ok(vstr(crate::format::format_number(
        value,
        options.decimals,
        options.include_leading_digit,
        options.parens_for_negative,
        options.group_digits,
    )))
}

pub fn format_currency(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    if let Some(early) = fixed_format_guard(value)? {
        return Ok(early);
    }
    let options = fixed_format_options(args)?;
    Ok(vstr(crate::format::format_currency(
        value,
        options.decimals,
        options.include_leading_digit,
        options.parens_for_negative,
        options.group_digits,
    )))
}

pub fn format_percent(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    if let Some(early) = fixed_format_guard(value)? {
        return Ok(early);
    }
    let options = fixed_format_options(args)?;
    Ok(vstr(crate::format::format_percent(
        value,
        options.decimals,
        options.include_leading_digit,
        options.parens_for_negative,
        options.group_digits,
    )))
}

pub fn format_date_time(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    let named_format = match opt(args, 1) {
        Some(value) if !matches!(value.vtype(), VarType::Empty) => as_i32(value)?,
        _ => 0,
    };
    crate::format::format_date_time(value, named_format)
        .map(vstr)
        .ok_or_else(|| LibError::invalid_call("FormatDateTime named format must be 0 to 4"))
}

struct FixedFormatOptions {
    decimals: usize,
    include_leading_digit: bool,
    parens_for_negative: bool,
    group_digits: bool,
}

/// `FormatNumber`/`FormatCurrency`/`FormatPercent` of `Null` is the empty string; text that is
/// not a number is a type mismatch.
fn fixed_format_guard(value: &Variant) -> LibResult<Option<Variant>> {
    match value.vtype() {
        VarType::Null => Ok(Some(vstr(""))),
        VarType::String
            if calcvba_runtime::coerce::parse_vba_numeric_string(&as_str(value)?).is_none() =>
        {
            Err(LibError::type_mismatch("expected a numeric value"))
        }
        _ => Ok(None),
    }
}

fn fixed_format_options(args: &[Variant]) -> LibResult<FixedFormatOptions> {
    Ok(FixedFormatOptions {
        decimals: fixed_decimals(args)?,
        include_leading_digit: tri_state_arg(args, 2, true)?,
        parens_for_negative: tri_state_arg(args, 3, false)?,
        group_digits: tri_state_arg(args, 4, true)?,
    })
}

fn fixed_decimals(args: &[Variant]) -> LibResult<usize> {
    let value = match opt(args, 1) {
        Some(value) if !matches!(value.vtype(), VarType::Empty) => as_i32(value)?,
        _ => -1,
    };
    match value {
        -1 => Ok(2),
        0..=99 => Ok(value as usize),
        _ => Err(LibError::invalid_call(
            "NumDigitsAfterDecimal must be -1 or 0 to 99",
        )),
    }
}

fn tri_state_arg(args: &[Variant], index: usize, default: bool) -> LibResult<bool> {
    let value = match opt(args, index) {
        Some(value) if !matches!(value.vtype(), VarType::Empty) => as_i32(value)?,
        _ => -2,
    };
    match value {
        -2 => Ok(default),
        -1 => Ok(true),
        0 => Ok(false),
        _ => Err(LibError::invalid_call(
            "TriState argument must be -2, -1, or 0",
        )),
    }
}

// ── Math ──────────────────────────────────────────────────────────────────────

pub fn math1(args: &[Variant], f: impl Fn(f64) -> f64) -> LibResult<Variant> {
    Ok(vf64(f(as_f64(need(args, 0)?)?)))
}

/// Excel rejects either sign at and above this measured argument magnitude.
/// The adjacent accepted/refused Doubles are 43DFFFFFFDDAD22F / 43DFFFFFFDDAD230.
pub fn trig(args: &[Variant], f: impl Fn(f64) -> f64) -> LibResult<Variant> {
    let value = as_f64(need(args, 0)?)?;
    if value.abs() >= 9.223_372e18 {
        return Err(LibError::invalid_call(
            "trigonometric argument out of range",
        ));
    }
    Ok(vf64(f(value)))
}

pub fn sqr(args: &[Variant]) -> LibResult<Variant> {
    Ok(vf64(sqr_f64(as_f64(need(args, 0)?)?)?))
}
pub fn sqr_f64(x: f64) -> LibResult<f64> {
    if x < 0.0 {
        return Err(LibError::invalid_call("Sqr argument must be non-negative"));
    }
    Ok(x.sqrt())
}

pub fn log(args: &[Variant]) -> LibResult<Variant> {
    let x = as_f64(need(args, 0)?)?;
    if x <= 0.0 {
        return Err(LibError::invalid_call("Log argument must be positive"));
    }
    Ok(vf64(x.ln()))
}

pub fn exp(args: &[Variant]) -> LibResult<Variant> {
    let value = as_f64(need(args, 0)?)?.exp();
    if !value.is_finite() {
        return Err(LibError::overflow("Exp overflow"));
    }
    Ok(vf64(value))
}

/// `Abs` / `Int` / `Fix` preserve the argument's numeric subtype. Abs promotes
/// overflowing Variant integers; the binder keeps a typed integer operand's
/// result type, so its minimum raises error 6. Empty -> Integer 0, Null -> Null,
/// Boolean -> Integer, and a numeric String -> Double.
#[derive(Clone, Copy)]
enum Math1Op {
    Abs,
    Int,
    Fix,
}

pub fn abs(args: &[Variant]) -> LibResult<Variant> {
    math1_typed(need(args, 0)?, Math1Op::Abs)
}
pub fn int_floor(args: &[Variant]) -> LibResult<Variant> {
    math1_typed(need(args, 0)?, Math1Op::Int)
}
pub fn fix_trunc(args: &[Variant]) -> LibResult<Variant> {
    math1_typed(need(args, 0)?, Math1Op::Fix)
}

/// `Sgn` — the sign of the argument as an `Integer` (-1, 0, 1), regardless of
/// the input type. `Null` raises error 94 (the result is `Integer`, which
/// cannot hold `Null`); `Empty` is 0.
pub fn sgn(args: &[Variant]) -> LibResult<Variant> {
    let v = need(args, 0)?;
    if v.vtype() == VarType::Null {
        return Err(LibError::new(94, "invalid use of Null"));
    }
    let x = conv_f64(v)?;
    if x.is_nan() {
        return Err(LibError::invalid_call("Sgn argument must not be NaN"));
    }
    let s: i16 = if x > 0.0 {
        1
    } else if x < 0.0 {
        -1
    } else {
        0
    };
    Ok(Variant::from_i16(s))
}

fn float_op(op: Math1Op, x: f64) -> f64 {
    match op {
        Math1Op::Abs => abs_f64(x),
        Math1Op::Int => x.floor(),
        Math1Op::Fix => x.trunc(),
    }
}
pub fn abs_f64(x: f64) -> f64 {
    x.abs()
}

/// `Abs` of a signed integer, promoting to the next wider type when the
/// magnitude overflows the source type (VBA never wraps or errors here):
/// SignedByte→Integer, Integer→Long, Long→Double, LongLong→Double.
fn abs_signed_promote(x: i64, kind: VarType) -> Variant {
    let mag = x.unsigned_abs();
    match kind {
        VarType::SignedByte if mag <= i8::MAX as u64 => Variant::from_i8(mag as i8),
        VarType::SignedByte => Variant::from_i16(mag as i16),
        VarType::Integer if mag <= i16::MAX as u64 => Variant::from_i16(mag as i16),
        VarType::Integer => Variant::from_i32(mag as i32),
        VarType::Long if mag <= i32::MAX as u64 => Variant::from_i32(mag as i32),
        VarType::Long => Variant::from_f64(mag as f64),
        VarType::LongLong if mag <= i64::MAX as u64 => Variant::from_i64(mag as i64),
        VarType::LongLong => Variant::from_f64(mag as f64),
        _ => Variant::from_f64(mag as f64),
    }
}

/// `Abs`/`Int`/`Fix` of a `Currency` (a fixed 4-dp scaled `i64`), keeping the
/// `Currency` type. `Int` floors toward −∞, `Fix` truncates toward zero.
fn currency_op(op: Math1Op, scaled: i64) -> LibResult<Variant> {
    let result = match op {
        Math1Op::Abs => scaled
            .checked_abs()
            .ok_or_else(|| LibError::overflow("Abs overflows Currency"))?,
        Math1Op::Fix => (scaled / 10_000) * 10_000,
        Math1Op::Int => {
            let q = scaled / 10_000;
            let rem = scaled % 10_000;
            let q = if rem != 0 && scaled < 0 { q - 1 } else { q };
            q * 10_000
        }
    };
    Ok(Variant::from_currency_scaled_i64(result))
}

/// `Abs`/`Int`/`Fix` of a `Decimal`, keeping the `Decimal` type. `Abs` clears
/// the sign; `Int`/`Fix` drop the fractional digits (scale → 0), with `Int`
/// flooring toward −∞ (a negative with a fraction rounds away from zero).
fn decimal_op(op: Math1Op, d: calcvba_runtime::Decimal96) -> Variant {
    use calcvba_runtime::Decimal96;
    match op {
        Math1Op::Abs => {
            Variant::from_decimal96(Decimal96::from_parts(d.lo, d.mid, d.hi, d.scale(), false))
        }
        Math1Op::Int | Math1Op::Fix => {
            let scale = u32::from(d.scale());
            if scale == 0 {
                return Variant::from_decimal96(d);
            }
            let magnitude = d.magnitude_u128();
            let divisor = 10u128.pow(scale.min(38));
            let mut whole = magnitude / divisor;
            let remainder = magnitude % divisor;
            if matches!(op, Math1Op::Int) && d.is_negative() && remainder != 0 {
                whole += 1; // floor of a negative rounds away from zero
            }
            let lo = (whole & 0xFFFF_FFFF) as u32;
            let mid = ((whole >> 32) & 0xFFFF_FFFF) as u32;
            let hi = ((whole >> 64) & 0xFFFF_FFFF) as u32;
            Variant::from_decimal96(Decimal96::from_parts(lo, mid, hi, 0, d.is_negative()))
        }
    }
}

fn math1_typed(v: &Variant, op: Math1Op) -> LibResult<Variant> {
    match v.vtype() {
        VarType::Null => Ok(Variant::null()),
        VarType::Empty => Ok(Variant::from_i16(0)),
        VarType::Boolean => {
            let n: i16 = if v.as_bool().unwrap_or(false) { -1 } else { 0 };
            Ok(Variant::from_i16(match op {
                Math1Op::Abs => n.abs(),
                Math1Op::Int | Math1Op::Fix => n,
            }))
        }
        // Signed integers: Int/Fix are identity; Abs may promote.
        kind @ (VarType::SignedByte | VarType::Integer | VarType::Long | VarType::LongLong) => {
            match op {
                Math1Op::Int | Math1Op::Fix => Ok(v.clone()),
                Math1Op::Abs => {
                    let x = match kind {
                        VarType::SignedByte => i64::from(v.as_i8().unwrap_or(0)),
                        VarType::Integer => i64::from(v.as_i16().unwrap_or(0)),
                        VarType::Long => i64::from(v.as_i32().unwrap_or(0)),
                        _ => v.as_i64().unwrap_or(0),
                    };
                    Ok(abs_signed_promote(x, kind))
                }
            }
        }
        // Unsigned integers are non-negative whole numbers: Abs/Int/Fix all identity.
        VarType::Byte
        | VarType::UnsignedInteger
        | VarType::UnsignedLong
        | VarType::UnsignedInt
        | VarType::UnsignedLongLong => Ok(v.clone()),
        VarType::Single => Ok(Variant::from_f32(
            float_op(op, f64::from(v.as_f32().unwrap_or(0.0))) as f32,
        )),
        VarType::Double => Ok(Variant::from_f64(float_op(op, v.as_f64().unwrap_or(0.0)))),
        VarType::Date => Ok(Variant::from_date_f64(float_op(
            op,
            v.as_date_f64().unwrap_or(0.0),
        ))),
        VarType::Currency => currency_op(op, v.as_currency_scaled_i64().unwrap_or(0)),
        VarType::Decimal => Ok(decimal_op(op, v.as_decimal96().unwrap_or_default())),
        // A numeric String coerces to Double; everything else (Object/Error/…) is a Type mismatch.
        VarType::String => Ok(vf64(float_op(op, conv_f64(v)?))),
        _ => Err(LibError::type_mismatch("Abs/Int/Fix expects a number")),
    }
}

pub fn round(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    let digits = match arg_present(args, 1) {
        Some(v) => as_i32(v)?,
        None => 0,
    };
    if digits < 0 {
        return Err(LibError::invalid_call(
            "Round numdecimalplaces must be non-negative",
        ));
    }
    let digits = digits as u32;
    // The result keeps the argument's type; only a fractional part can change.
    match value.vtype() {
        VarType::Byte | VarType::Integer | VarType::Long | VarType::LongLong => {
            return Ok(value.clone());
        }
        VarType::Boolean => {
            return Ok(Variant::from_i16(if value.as_bool().unwrap_or(false) {
                -1
            } else {
                0
            }));
        }
        VarType::Empty => return Ok(Variant::from_i16(0)),
        VarType::Currency => {
            let scaled = value.as_currency_scaled_i64().unwrap_or(0);
            if digits >= 4 {
                return Ok(value.clone());
            }
            let drop = 4 - digits;
            let magnitude =
                round_half_even_div(u128::from(scaled.unsigned_abs()), drop) * 10u128.pow(drop);
            let magnitude = i64::try_from(magnitude)
                .map_err(|_| LibError::overflow("Round overflows Currency"))?;
            return Ok(Variant::from_currency_scaled_i64(if scaled < 0 {
                -magnitude
            } else {
                magnitude
            }));
        }
        VarType::Decimal => {
            let d = value.as_decimal96().unwrap_or_default();
            let scale = u32::from(d.scale());
            if digits >= scale {
                return Ok(value.clone());
            }
            let magnitude = round_half_even_div(d.magnitude_u128(), scale - digits);
            return Ok(Variant::from_decimal96(decimal96_from_magnitude(
                magnitude,
                digits as u8,
                d.is_negative(),
            )));
        }
        VarType::Single => {
            if digits > 22 {
                return Err(LibError::invalid_call("Round digits must not exceed 22"));
            }
            let x = f64::from(value.as_f32().unwrap_or(0.0));
            return Ok(Variant::from_f32(round_float(x, digits) as f32));
        }
        _ => {}
    }
    if digits > 22 {
        return Err(LibError::invalid_call("Round digits must not exceed 22"));
    }
    let x = as_f64(value)?;
    let rounded = round_float(x, digits);
    if value.vtype() == VarType::Date {
        Ok(Variant::from_date_f64(rounded))
    } else {
        Ok(vf64(rounded))
    }
}

fn round_float(x: f64, digits: u32) -> f64 {
    let factor = 10f64.powi(digits as i32);
    let scaled = x * factor;
    let rounded = scaled.round_ties_even();
    // Excel preserves negative zero at an exact negative half tie, but clears
    // the sign when an ordinary negative value rounds to zero (P_RoundEdges).
    if rounded == 0.0 && scaled != -0.5 {
        0.0
    } else {
        rounded / factor
    }
}

/// `magnitude / 10^drop`, rounded half to even.
fn round_half_even_div(magnitude: u128, drop: u32) -> u128 {
    let divisor = 10u128.pow(drop.min(38));
    let quotient = magnitude / divisor;
    let remainder = magnitude % divisor;
    let half = divisor / 2;
    if remainder > half || (remainder == half && quotient % 2 == 1) {
        quotient + 1
    } else {
        quotient
    }
}

// ── Date / time (serial: days since 1899-12-30) ────────────────────────────────
// Proleptic Gregorian, matching VBA's OLE-Automation `Date`. (The Excel
// worksheet 1900 leap-year quirk is not part of VBA `Date` semantics.) The serial
// ↔ civil math lives in `calcvba_runtime::vba_date` (the single source of truth, so
// the runtime's `Date` display formatters share it); imported above.

/// Sakamoto's algorithm (ported from the legacy VM): 0 = Sunday … 6 = Saturday.
pub(crate) fn day_of_week(year: i32, month: u32, day: u32) -> i32 {
    let table = [0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4];
    let m = month as i64;
    let y = i64::from(year) - i64::from(m < 3);
    (y + y / 4 - y / 100 + y / 400 + table[(m - 1) as usize] + i64::from(day)).rem_euclid(7) as i32
}

fn validate_vba_date_serial(serial: f64) -> LibResult<f64> {
    if serial.is_finite()
        && (VBA_DATE_MIN_SERIAL..VBA_DATE_MAX_EXCLUSIVE_SERIAL).contains(&serial.trunc())
    {
        Ok(serial)
    } else {
        Err(LibError::invalid_call("date out of range"))
    }
}

fn validate_dateserial_arg(value: i64) -> LibResult<i64> {
    if (-32_768..=32_767).contains(&value) {
        Ok(value)
    } else {
        // The arguments are `Integer`: a value outside its range is an Overflow.
        Err(LibError::overflow("DateSerial argument out of range"))
    }
}

fn day_index_weekday(day: i64) -> i32 {
    let (y, m, d) = serial_to_ymd(day as f64);
    day_of_week(y as i32, m as u32, d as u32)
}

fn count_weekday_boundaries_forward(start_day: i64, end_day: i64, target_dow: i32) -> i64 {
    if end_day <= start_day {
        return 0;
    }
    let first_candidate = start_day + 1;
    let offset = (target_dow - day_index_weekday(first_candidate)).rem_euclid(7) as i64;
    let first_hit = first_candidate + offset;
    if first_hit > end_day {
        0
    } else {
        1 + (end_day - first_hit) / 7
    }
}

fn count_weekday_boundaries(start_day: i64, end_day: i64, target_dow: i32) -> i64 {
    if end_day >= start_day {
        count_weekday_boundaries_forward(start_day, end_day, target_dow)
    } else {
        -count_weekday_boundaries_forward(end_day, start_day, target_dow)
    }
}

pub fn date_serial(args: &[Variant]) -> LibResult<Variant> {
    let y = validate_dateserial_arg(as_i64(need(args, 0)?)?)?;
    let m = validate_dateserial_arg(as_i64(need(args, 1)?)?)?;
    let d = validate_dateserial_arg(as_i64(need(args, 2)?)?)?;
    // Negative years are relative to 2000 (Excel: -1 -> 1999, -100 -> 1900).
    // Positive two-digit years use the current default window; host cutoff remains separate.
    let y = match y {
        ..=-1 => y + 2000,
        0..=29 => y + 2000,
        30..=99 => y + 1900,
        100..=9999 => y,
        _ => return Err(LibError::invalid_call("DateSerial year out of range")),
    };
    Ok(Variant::from_date_f64(validate_vba_date_serial(
        ymd_to_serial(y, m, d),
    )?))
}

pub fn time_serial(args: &[Variant]) -> LibResult<Variant> {
    let h = validate_dateserial_arg(as_i64(need(args, 0)?)?)? as f64;
    let m = validate_dateserial_arg(as_i64(need(args, 1)?)?)? as f64;
    let s = validate_dateserial_arg(as_i64(need(args, 2)?)?)? as f64;
    Ok(Variant::from_date_f64(
        (h * 3600.0 + m * 60.0 + s) / 86400.0,
    ))
}

pub fn date_value(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    let serial = if let Some(date) = value
        .as_date_f64()
        .filter(|_| value.vtype() == VarType::Date)
    {
        // Date-to-text coercion rounds seconds, including carry into the next day.
        let (year, month, day) = serial_to_ymd(date);
        validate_vba_date_serial(ymd_to_serial(year, month, day))
            .map_err(|_| LibError::type_mismatch("date out of range"))?
    } else {
        let text = as_str(value)?;
        date_time_text_serial(&text)
            .map_err(|_| LibError::type_mismatch(format!("cannot parse date `{text}`")))?
    };
    Ok(Variant::from_date_f64(serial.trunc()))
}

/// A date argument: a `Date` or numeric serial, or a string read as `CDate` reads it.
fn date_arg(v: &Variant) -> LibResult<f64> {
    let serial = match v.vtype() {
        VarType::Date => Ok(v.as_date_f64().unwrap_or(0.0)),
        // Date text parses as `CDate` parses it; a numeric string is its serial.
        VarType::String => match cdate_from_string(&as_str(v)?) {
            Ok(date) => Ok(date.as_date_f64().unwrap_or(0.0)),
            Err(error) => as_f64(v).map_err(|_| error),
        },
        _ => as_f64(v),
    }?;
    if !calcvba_runtime::coerce::is_date_serial(serial) {
        return Err(LibError::type_mismatch("date argument out of range"));
    }
    Ok(serial)
}

/// `TimeValue` takes a String argument. Date coercion rounds to display seconds;
/// numeric arguments are parsed as text rather than treated as date serials.
pub fn time_value(args: &[Variant]) -> LibResult<Variant> {
    let text = as_str(need(args, 0)?)?;
    Ok(Variant::from_date_f64(parse_time_value_string(&text)?))
}

/// Whether `(y, m, d)` is a real calendar date. The month/day are normalized by
/// `days_from_civil` (Feb 30 rolls into March), so a round-trip through the serial that does
/// NOT return the same `(y, m, d)` means the input was an invalid date — matching VBA, where
/// `CDate("February 30, 2000")` raises error 13 and `IsDate(...)` of it is False.
fn ymd_is_valid(y: i64, m: i64, d: i64) -> bool {
    (1..=12).contains(&m) && d >= 1 && civil_from_days(days_from_civil(y, m, d)) == (y, m, d)
}

fn parse_date(s: &str) -> LibResult<(i64, i64, i64)> {
    let err = || LibError::type_mismatch(format!("cannot parse date `{s}`"));
    let (y, m, d) = parse_date_raw(s).ok_or_else(err)?;
    if !ymd_is_valid(y, m, d) {
        return Err(err());
    }
    Ok((y, m, d))
}

fn parse_time_value_string(s: &str) -> LibResult<f64> {
    let s = s.trim();
    let err = || LibError::type_mismatch(format!("cannot parse time `{s}`"));
    if s.is_empty() {
        return Err(err());
    }
    if let Some((date_part, time_part)) = split_trailing_time(s) {
        parse_date(date_part.trim())?;
        return parse_time_only(&time_part).ok_or_else(err);
    }
    if let Some(time) = parse_time_only(s) {
        return Ok(time);
    }
    // A date with no time of day has the time 0.
    parse_date(s).map(|_| 0.0)
}

fn split_trailing_time(s: &str) -> Option<(&str, String)> {
    let (before, last) = s.trim().rsplit_once(' ')?;
    if last.eq_ignore_ascii_case("am") || last.eq_ignore_ascii_case("pm") {
        let (date_part, time_part) = before.trim_end().rsplit_once(' ')?;
        if time_part.contains([':', '.']) || time_part.bytes().all(|b| b.is_ascii_digit()) {
            return Some((date_part.trim(), format!("{time_part} {last}")));
        }
    } else if last.contains(':') {
        return Some((before.trim(), last.trim().to_string()));
    }
    None
}

fn parse_time_only(s: &str) -> Option<f64> {
    let mut text = s.trim();
    let mut meridian = None;
    let lower = text.to_ascii_lowercase();
    if lower.ends_with("am") || lower.ends_with("pm") {
        meridian = Some(&lower[lower.len() - 2..]);
        text = text[..text.len() - 2].trim_end();
    }

    let parts: Vec<&str> = text.split([':', '.']).collect();
    // `3 PM` is a time; a bare number is not.
    let one_part_with_meridian = parts.len() == 1 && meridian.is_some();
    if !(2..=3).contains(&parts.len()) && !one_part_with_meridian {
        return None;
    }
    if parts.iter().any(|part| {
        let part = part.trim();
        part.is_empty() || !part.bytes().all(|b| b.is_ascii_digit())
    }) {
        return None;
    }
    let mut h = parts[0].trim().parse::<i64>().ok()?;
    let m = match parts.get(1) {
        Some(value) => value.trim().parse::<i64>().ok()?,
        None => 0,
    };
    let sec = match parts.get(2) {
        Some(value) => value.trim().parse::<i64>().ok()?,
        None => 0,
    };
    if !(0..=59).contains(&m) || !(0..=59).contains(&sec) {
        return None;
    }
    match meridian {
        Some("am") => {
            if !(0..=23).contains(&h) {
                return None;
            }
            if h == 12 {
                h = 0;
            }
        }
        Some("pm") => {
            if !(0..=23).contains(&h) {
                return None;
            }
            if h < 12 {
                h += 12;
            }
        }
        _ if !(0..=23).contains(&h) => return None,
        _ => {}
    }
    Some((h as f64 * 3600.0 + m as f64 * 60.0 + sec as f64) / 86400.0)
}

fn parse_date_raw(s: &str) -> Option<(i64, i64, i64)> {
    // Two-digit years use VBA's window: 0-29 -> 2000s, 30-99 -> 1900s.
    fn window(year: i64) -> i64 {
        match year {
            0..=29 => year + 2000,
            30..=99 => year + 1900,
            other => other,
        }
    }
    let s = s.trim();
    let mut month = None;
    let mut nums: Vec<i64> = Vec::new();
    let mut first_token_len = 0;
    for tok in s
        .split([' ', ',', '\t', '/', '-'])
        .filter(|t| !t.is_empty())
    {
        if let Some(m) = month_from_name(tok) {
            if month.replace(m).is_some() {
                return None;
            }
        } else if tok.bytes().all(|b| b.is_ascii_digit()) {
            if nums.is_empty() && month.is_none() {
                first_token_len = tok.len();
            }
            nums.push(tok.parse::<i64>().ok()?);
        } else {
            return None;
        }
    }
    match (month, nums.as_slice()) {
        // Month name, day and year in either order: `31-Jan-95`, `Jan 5, 2000`, `5 January 2000`.
        (Some(m), [a, b]) => {
            let (day, year) = if *a > 31 { (*b, *a) } else { (*a, *b) };
            Some((window(year), m, day))
        }
        // Month name and one number: a day of this year (`Dec 30`) or a year (`Jan 2000`).
        (Some(m), [n]) if *n > 31 => Some((window(*n), m, 1)),
        (Some(m), [n]) => Some((current_year(), m, *n)),
        // All numeric: `y-m-d` when the first part is a year, otherwise US `m/d/y`.
        (None, [a, b, c]) => {
            if first_token_len >= 3 || *a > 31 {
                Some((*a, *b, *c))
            } else {
                Some((window(*c), *a, *b))
            }
        }
        // `m/d`: the day of this year.
        (None, [a, b]) => Some((current_year(), *a, *b)),
        _ => None,
    }
}

/// The current calendar year (UTC).
fn current_year() -> i64 {
    let seconds = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map_or(0, |d| d.as_secs() as i64);
    civil_from_days(seconds.div_euclid(86_400)).0
}

/// The 1-based month number for a month name or its 3-letter prefix (`Jan`,
/// `January`, `jan.`, …). Each English month's 3-letter prefix is unique.
fn month_from_name(tok: &str) -> Option<i64> {
    const ABBR: [&str; 12] = [
        "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec",
    ];
    let t = tok.trim().trim_end_matches('.').to_ascii_lowercase();
    if t.len() < 3 {
        return None;
    }
    ABBR.iter()
        .position(|a| t.starts_with(a))
        .map(|i| i as i64 + 1)
}

pub fn date_add(args: &[Variant]) -> LibResult<Variant> {
    let interval = interval_arg(need(args, 0)?)?;
    // The count is a `Long`; Excel drops its fraction (`DateAdd("d", 1.5, ...)` adds one day).
    let number = as_f64(need(args, 1)?)?.trunc();
    if !number.is_finite() || number.abs() > f64::from(i32::MAX) {
        return Err(LibError::overflow("DateAdd number out of range"));
    }
    let serial = date_arg(need(args, 2)?)?;
    let result = match interval.as_str() {
        "yyyy" | "m" | "q" => {
            let (y, m, d) = serial_to_ymd(serial);
            let months = m - 1
                + match interval.as_str() {
                    "yyyy" => number as i64 * 12,
                    "q" => number as i64 * 3,
                    _ => number as i64,
                };
            let ny = y + months.div_euclid(12);
            let nm = months.rem_euclid(12) + 1;
            // Clamp the day to the target month (DateAdd("m",1,#1/31#) -> end of Feb,
            // not a roll into March), and re-attach the source time of day. A VBA
            // serial carries the time as the fraction |serial - trunc|, applied away
            // from zero, so negative-date serials subtract it.
            let clamped_day = d.min(days_in_month(ny, nm));
            let (h, mi, s) = serial_to_hms(serial);
            let time_frac = (h * 3600 + mi * 60 + s) as f64 / 86_400.0;
            let date_serial = ymd_to_serial(ny, nm, clamped_day);
            if date_serial < 0.0 {
                date_serial - time_frac
            } else {
                date_serial + time_frac
            }
        }
        // Before 1899-12-30 the time of day counts away from zero, so the arithmetic runs on
        // the continuous time line and maps back.
        "d" | "y" | "w" => real_to_serial(serial_to_real(serial) + number),
        "ww" => real_to_serial(serial_to_real(serial) + number * 7.0),
        "h" => real_to_serial(serial_to_real(serial) + number / 24.0),
        "n" => real_to_serial(serial_to_real(serial) + number / 1440.0),
        "s" => real_to_serial(serial_to_real(serial) + number / 86400.0),
        _ => {
            return Err(LibError::invalid_call(format!(
                "unknown DateAdd interval `{interval}`"
            )));
        }
    };
    Ok(Variant::from_date_f64(validate_vba_date_serial(result)?))
}

fn interval_arg(value: &Variant) -> LibResult<String> {
    Ok(as_str(value)?.trim().to_ascii_lowercase())
}

/// The continuous time (days) of a serial: before 1899-12-30 the time of day counts away from
/// zero (`-1.25` is 06:00 on day -1), so it adds to the day instead of subtracting.
fn serial_to_real(serial: f64) -> f64 {
    if serial < 0.0 {
        serial.trunc() + (serial - serial.trunc()).abs()
    } else {
        serial
    }
}

fn real_to_serial(real: f64) -> f64 {
    if real >= 0.0 {
        real
    } else {
        let day = real.floor();
        day - (real - day)
    }
}

pub fn date_diff(args: &[Variant]) -> LibResult<Variant> {
    let interval = interval_arg(need(args, 0)?)?;
    let a = date_arg(need(args, 1)?)?;
    let b = date_arg(need(args, 2)?)?;
    let (ra, rb) = (serial_to_real(a), serial_to_real(b));
    let (a_day, b_day) = (ra.floor() as i64, rb.floor() as i64);
    // Whole seconds on the continuous time line, so an hour or minute boundary is exact.
    let (a_sec, b_sec) = (
        (ra * 86_400.0).round() as i64,
        (rb * 86_400.0).round() as i64,
    );
    let (ya, ma, _) = serial_to_ymd(a);
    let (yb, mb, _) = serial_to_ymd(b);
    let result = match interval.as_str() {
        "yyyy" => yb - ya,
        "q" => (yb * 4 + (mb - 1) / 3) - (ya * 4 + (ma - 1) / 3),
        "m" => (yb - ya) * 12 + (mb - ma),
        "d" | "y" => b_day - a_day,
        "w" => (b_day - a_day) / 7,
        "ww" => {
            let first = first_day_of_week(args, 3)?;
            count_weekday_boundaries(a_day, b_day, first - 1)
        }
        "h" => b_sec.div_euclid(3600) - a_sec.div_euclid(3600),
        "n" => b_sec.div_euclid(60) - a_sec.div_euclid(60),
        "s" => b_sec - a_sec,
        _ => {
            return Err(LibError::invalid_call(format!(
                "unknown DateDiff interval `{interval}`"
            )));
        }
    };
    // The count is a `Long`; one that does not fit (seconds across a century) comes back as a
    // `Double`, as Excel's does.
    Ok(match i32::try_from(result) {
        Ok(count) => vi32(count),
        Err(_) => vf64(result as f64),
    })
}

#[derive(Clone, Copy)]
pub enum DatePart {
    Year,
    Month,
    Day,
    Weekday,
    Hour,
    Minute,
    Second,
}

pub fn date_part(args: &[Variant], part: DatePart) -> LibResult<Variant> {
    let serial = date_arg(need(args, 0)?)?;
    let value = match part {
        DatePart::Year | DatePart::Month | DatePart::Day | DatePart::Weekday => {
            let (y, m, d) = serial_to_ymd(serial);
            match part {
                DatePart::Year => y as i32,
                DatePart::Month => m as i32,
                DatePart::Day => d as i32,
                DatePart::Weekday => {
                    let first = first_day_of_week(args, 1)?;
                    weekday_number(y as i32, m as u32, d as u32, first)
                }
                _ => unreachable!(),
            }
        }
        DatePart::Hour | DatePart::Minute | DatePart::Second => {
            let (h, mi, s) = serial_to_hms(serial);
            match part {
                DatePart::Hour => h as i32,
                DatePart::Minute => mi as i32,
                DatePart::Second => s as i32,
                _ => unreachable!(),
            }
        }
    };
    Ok(Variant::from_i16(value as i16))
}

fn first_day_of_week(args: &[Variant], index: usize) -> LibResult<i32> {
    let first = match arg_present(args, index) {
        Some(v) => as_i32(v)?,
        None => 1,
    };
    let first = if first == 0 { 1 } else { first };
    if !(1..=7).contains(&first) {
        return Err(LibError::invalid_call("firstdayofweek out of range"));
    }
    Ok(first)
}

fn first_week_of_year(args: &[Variant], index: usize) -> LibResult<i32> {
    let first = match arg_present(args, index) {
        Some(v) => as_i32(v)?,
        None => 1,
    };
    let first = if first == 0 { 1 } else { first };
    if !(1..=3).contains(&first) {
        return Err(LibError::invalid_call("firstweekofyear out of range"));
    }
    Ok(first)
}

/// The week of the year (`DatePart "ww"`, `Format "ww"`). `first_day` is `vbSunday..vbSaturday`
/// (1-7); `first_week` is `vbFirstJan1` (1: week 1 holds Jan 1), `vbFirstFourDays` (2: week 1
/// is the first week with four days in the year) or `vbFirstFullWeek` (3). A date before week 1
/// belongs to the last week of the previous year, and (except under `vbFirstJan1`) one on or
/// after week 1 of the next year is week 1.
pub(crate) fn week_of_year(y: i64, m: i64, d: i64, first_day: i32, first_week: i32) -> i64 {
    let day = i128::from(days_from_civil(y, m, d));
    let week_one_start = |year: i64| -> i128 {
        let jan1 = i128::from(days_from_civil(year, 1, 1));
        // Days from the start of the week to Jan 1 (1970-01-01 was a Thursday).
        let before = (jan1 + 4 - (i128::from(first_day) - 1)).rem_euclid(7);
        let week_start = jan1 - before;
        match first_week {
            2 if 7 - before < 4 => week_start + 7,
            3 if before != 0 => week_start + 7,
            _ => week_start,
        }
    };
    // Under vbFirstJan1 the days of the last week stay in this year (Dec 31 can be week 53);
    // under the other two rules they belong to week 1 once it has started.
    if first_week != 1 && day >= week_one_start(y.saturating_add(1)) {
        return 1;
    }
    let start = if day >= week_one_start(y) {
        week_one_start(y)
    } else {
        week_one_start(y.saturating_sub(1))
    };
    ((day - start) / 7 + 1).clamp(i128::from(i64::MIN), i128::from(i64::MAX)) as i64
}

fn weekday_number(year: i32, month: u32, day: u32, first: i32) -> i32 {
    let dow0 = day_of_week(year, month, day);
    (dow0 - (first - 1)).rem_euclid(7) + 1
}

pub fn month_name(args: &[Variant]) -> LibResult<Variant> {
    const NAMES: [&str; 12] = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    let n = as_usize(need(args, 0)?)?;
    let abbreviate = matches!(opt(args, 1), Some(v) if as_bool_lenient(v));
    NAMES
        .get(n.wrapping_sub(1))
        .map(|s| vstr(if abbreviate { &s[..3] } else { s }))
        .ok_or_else(|| LibError::invalid_call("MonthName index out of range"))
}

/// `WeekdayName(weekday, [abbreviate], [firstdayofweek])` — the weekday is
/// relative to `firstdayofweek` (default vbSunday = 1), 1-based.
pub fn weekday_name(args: &[Variant]) -> LibResult<Variant> {
    const NAMES: [&str; 7] = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    let weekday = as_i32(need(args, 0)?)?;
    let abbreviate = matches!(opt(args, 1), Some(v) if as_bool_lenient(v));
    // firstdayofweek: 0 = vbUseSystemDayOfWeek (treat as Sunday), default vbSunday.
    let first = match opt(args, 2) {
        Some(v) => as_i32(v)?,
        None => 1,
    };
    if !(1..=7).contains(&weekday) {
        return Err(LibError::invalid_call("WeekdayName weekday out of range"));
    }
    let first = if first == 0 { 1 } else { first };
    if !(1..=7).contains(&first) {
        return Err(LibError::invalid_call(
            "WeekdayName firstdayofweek out of range",
        ));
    }
    let idx = ((weekday - 1) + (first - 1)).rem_euclid(7) as usize;
    let name = NAMES[idx];
    Ok(vstr(if abbreviate { &name[..3] } else { name }))
}

/// `LenB(value)` — byte count. For strings, VBA reports UTF-16 payload bytes;
/// for scalar variables it reports the storage width of the value's VarType.
pub fn len_b(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    if let Some(bytes) = value.vba_record_memory_len() {
        return Ok(vi32(bytes as i32));
    }
    let bytes = match value.vtype() {
        VarType::String => match value.string_byte_len() {
            Some(bytes) => bytes as usize,
            None => as_str(value)?.encode_utf16().count() * 2,
        },
        VarType::Boolean | VarType::Integer | VarType::UnsignedInteger => 2,
        VarType::Byte | VarType::SignedByte => 1,
        VarType::Long | VarType::UnsignedLong | VarType::UnsignedInt | VarType::Error => 4,
        VarType::LongLong | VarType::UnsignedLongLong => 8,
        VarType::Single => 4,
        VarType::Double | VarType::Currency | VarType::Date => 8,
        VarType::Empty | VarType::Null => 0,
        _ => as_str(value)?.encode_utf16().count() * 2,
    };
    Ok(vi32(bytes as i32))
}

pub fn vba_datepart(args: &[Variant]) -> LibResult<Variant> {
    let interval = interval_arg(need(args, 0)?)?;
    let serial = date_arg(need(args, 1)?)?;
    let (y, m, d) = serial_to_ymd(serial);
    let (h, mi, s) = serial_to_hms(serial);
    let value = match interval.as_str() {
        "yyyy" => y as i32,
        "q" => (((m - 1) / 3) + 1) as i32,
        "m" => m as i32,
        "d" => d as i32,
        "w" => {
            let first = first_day_of_week(args, 2)?;
            weekday_number(y as i32, m as u32, d as u32, first)
        }
        "y" => {
            // Day of the year: this date's ordinal minus Jan 1's.
            (days_from_civil(y, m, d) - days_from_civil(y, 1, 1) + 1) as i32
        }
        "ww" => {
            let first_day = first_day_of_week(args, 2)?;
            let first_week = first_week_of_year(args, 3)?;
            week_of_year(y, m, d, first_day, first_week) as i32
        }
        "h" => h as i32,
        "n" => mi as i32,
        "s" => s as i32,
        other => {
            return Err(LibError::invalid_call(format!(
                "unknown DatePart interval `{other}`"
            )));
        }
    };
    Ok(Variant::from_i16(value as i16))
}

/// `IsMissing(arg)` — True only for an omitted optional `Variant` argument,
/// which the VM materializes as vbError `&H80020004` (DISP_E_PARAMNOTFOUND).
pub fn is_missing(args: &[Variant]) -> LibResult<Variant> {
    const MISSING_ARG: i32 = 0x8002_0004u32 as i32;
    let v = need(args, 0)?;
    let missing = v.vtype() == VarType::Error && v.as_error_code() == Some(MISSING_ARG);
    Ok(vbool(missing))
}

pub fn error_text(args: &[Variant]) -> LibResult<Variant> {
    let Some(value) = arg_present(args, 0) else {
        return Ok(vstr(""));
    };
    let code = as_i32(value)?;
    if code == 0 {
        return Ok(vstr(""));
    }
    // The number is an unsigned 16-bit value; a negative one has no message of its own.
    if code > 65_535 {
        return Err(LibError::overflow("Error number out of range"));
    }
    if code < 0 {
        return Ok(vstr(calcvba_runtime::default_error_message(0)));
    }
    Ok(vstr(calcvba_runtime::default_error_message(code)))
}

/// Lenient truthiness for an optional Boolean-ish flag argument.
fn as_bool_lenient(v: &Variant) -> bool {
    v.as_bool()
        .or_else(|| as_i32(v).ok().map(|n| n != 0))
        .unwrap_or(false)
}

// ── Conversion ─────────────────────────────────────────────────────────────────

fn radix_digits(n: u64, radix: u32) -> String {
    match radix {
        16 => format!("{n:X}"),
        8 => format!("{n:o}"),
        _ => unreachable!("unsupported radix"),
    }
}

fn integer_width_radix(value: &Variant, radix: u32) -> LibResult<String> {
    let digits = match value.vtype() {
        VarType::Boolean => {
            let n: u16 = if value.as_bool().unwrap_or(false) {
                u16::MAX
            } else {
                0
            };
            radix_digits(u64::from(n), radix)
        }
        VarType::SignedByte => radix_digits(value.as_i8().unwrap_or(0) as u8 as u64, radix),
        VarType::Byte => radix_digits(u64::from(value.as_u8().unwrap_or(0)), radix),
        VarType::Integer => radix_digits(value.as_i16().unwrap_or(0) as u16 as u64, radix),
        VarType::UnsignedInteger => radix_digits(u64::from(value.as_u16().unwrap_or(0)), radix),
        VarType::Long => radix_digits(value.as_i32().unwrap_or(0) as u32 as u64, radix),
        VarType::UnsignedLong | VarType::UnsignedInt => {
            radix_digits(u64::from(value.as_u32().unwrap_or(0)), radix)
        }
        VarType::LongLong => radix_digits(value.as_i64().unwrap_or(0) as u64, radix),
        VarType::UnsignedLongLong => radix_digits(value.as_u64().unwrap_or(0), radix),
        _ => radix_digits(as_i64(value)? as u64, radix),
    };
    Ok(digits)
}

pub fn hex(args: &[Variant]) -> LibResult<Variant> {
    Ok(vstr(integer_width_radix(need(args, 0)?, 16)?))
}
pub fn oct(args: &[Variant]) -> LibResult<Variant> {
    Ok(vstr(integer_width_radix(need(args, 0)?, 8)?))
}
pub fn cstr(args: &[Variant]) -> LibResult<Variant> {
    Ok(vstr(as_str(need(args, 0)?)?))
}
pub fn str_fn(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    // A numeric string is converted to its number first; `Empty` is 0.
    let number = match value.vtype() {
        VarType::String => vf64(as_f64(value)?),
        VarType::Empty => vi32(0),
        _ => value.clone(),
    };
    let n = as_f64(&number)?;
    let body = as_str(&number)?;
    let body = if let Some(rest) = body.strip_prefix("0.") {
        format!(".{rest}")
    } else if let Some(rest) = body.strip_prefix("-0.") {
        format!("-.{rest}")
    } else {
        body
    };
    Ok(vstr(if n.is_sign_negative() {
        body
    } else {
        format!(" {body}")
    }))
}
/// FIDELITY: parses a leading numeric prefix.
pub fn val(args: &[Variant]) -> LibResult<Variant> {
    let arg = need(args, 0)?;
    if arg.vtype() == VarType::Error {
        return Err(LibError::type_mismatch(default_error_message(13)));
    }
    let s = as_str(arg)?;
    let compact: String = s
        .chars()
        .filter(|ch| !matches!(ch, ' ' | '\t' | '\n' | '\r' | '\u{3000}'))
        .collect();
    let text = compact.as_str();
    if let Some(value) = val_radix_prefix(text)? {
        return Ok(vf64(value));
    }
    let (value, rest) = val_decimal_prefix(text).unwrap_or((0.0, ""));
    if !value.is_finite() {
        return Err(LibError::overflow(default_error_message(6)));
    }
    if rest.starts_with('%') {
        if value.fract() != 0.0 {
            return Err(LibError::type_mismatch(default_error_message(13)));
        }
        if !(i16::MIN as f64..=i16::MAX as f64).contains(&value) {
            return Err(LibError::overflow(default_error_message(6)));
        }
    }
    Ok(vf64(value))
}

/// Unlike source literals and conversion functions, Val truncates unsuffixed
/// radix text to 32 bits. A leading sign prevents radix recognition altogether.
fn val_radix_prefix(text: &str) -> LibResult<Option<f64>> {
    let Some(rest) = text.strip_prefix('&') else {
        return Ok(None);
    };
    let (radix, digits) = match rest.as_bytes().first() {
        Some(b'H' | b'h') => (16, &rest[1..]),
        Some(b'O' | b'o') => (8, &rest[1..]),
        _ => (8, rest),
    };
    let end = digits
        .bytes()
        .position(|b| !char::from(b).is_digit(radix))
        .unwrap_or(digits.len());
    if end == 0 {
        return Ok(Some(0.0));
    }
    let magnitude = u64::from_str_radix(&digits[..end], radix)
        .map_err(|_| LibError::overflow(default_error_message(6)))?;
    let suffix = digits.as_bytes().get(end).copied();
    let wide = matches!(suffix, Some(b'^' | b'!' | b'#' | b'@'));
    let magnitude = if wide {
        magnitude
    } else {
        magnitude as u32 as u64
    };
    let width_suffix = match suffix {
        Some(b'^' | b'&') => suffix,
        _ => None,
    };
    let width = calcvba_runtime::vba_radix_width(magnitude, width_suffix)
        .ok_or_else(|| LibError::overflow(default_error_message(6)))?;
    let value = calcvba_runtime::vba_radix_signed_value(magnitude, width);
    if suffix == Some(b'%') && i16::try_from(value).is_err() {
        return Err(LibError::overflow(default_error_message(6)));
    }
    Ok(Some(value as f64))
}

fn val_decimal_prefix(text: &str) -> Option<(f64, &str)> {
    let bytes = text.as_bytes();
    let mut pos = 0usize;
    if matches!(bytes.get(pos), Some(b'+' | b'-')) {
        pos += 1;
    }

    let mut saw_digit = false;
    while matches!(bytes.get(pos), Some(b'0'..=b'9')) {
        saw_digit = true;
        pos += 1;
    }
    if matches!(bytes.get(pos), Some(b'.')) {
        pos += 1;
        while matches!(bytes.get(pos), Some(b'0'..=b'9')) {
            saw_digit = true;
            pos += 1;
        }
    }
    if !saw_digit {
        return None;
    }

    let mantissa_end = pos;
    if matches!(bytes.get(pos), Some(b'e' | b'E' | b'd' | b'D')) {
        pos += 1;
        if matches!(bytes.get(pos), Some(b'+' | b'-')) {
            pos += 1;
        }
        let exp_digits_start = pos;
        while matches!(bytes.get(pos), Some(b'0'..=b'9')) {
            pos += 1;
        }
        if exp_digits_start == pos {
            pos = mantissa_end;
        }
    }

    let token = text[..pos].replace(['d', 'D'], "E");
    token.parse::<f64>().ok().map(|value| (value, &text[pos..]))
}
pub fn cdate(args: &[Variant]) -> LibResult<Variant> {
    let v = need(args, 0)?;
    match v.vtype() {
        VarType::Date => Ok(v.clone()),
        VarType::String => cdate_from_string(&as_str(v)?),
        VarType::Error => Err(LibError::type_mismatch("cannot convert Error to Date")),
        _ => {
            // Numeric CDate overflow is 6; text conversion failures are 13.
            let serial = conv_f64(v)?;
            if !calcvba_runtime::coerce::is_date_serial(serial) {
                return Err(LibError::overflow("value does not fit in Date"));
            }
            Ok(Variant::from_date_f64(validate_vba_date_serial(serial)?))
        }
    }
}

pub fn cvdate(args: &[Variant]) -> LibResult<Variant> {
    if need(args, 0)?.vtype() == VarType::Null {
        return Ok(Variant::null());
    }
    cdate(args).map_err(|error| {
        if error.code == 6 {
            LibError::type_mismatch("value does not fit in Date")
        } else {
            error
        }
    })
}

/// CDate/CVDate read time text before numeric text: "0.5" is 00:05,
/// while "43831.5" is a numeric date serial because 43831 is not an hour.
fn cdate_from_string(s: &str) -> LibResult<Variant> {
    let s = s.trim();
    if let Ok(serial) = date_time_text_serial(s) {
        return Ok(Variant::from_date_f64(serial));
    }
    if let Ok(serial) = conv_f64(&vstr(s))
        && calcvba_runtime::coerce::is_date_serial(serial)
    {
        return Ok(Variant::from_date_f64(serial));
    }
    Err(LibError::type_mismatch(format!("cannot parse date `{s}`")))
}

/// Date/time recognition does not include CDate's numeric-string fallback (IsDate).
fn date_time_text_serial(s: &str) -> LibResult<f64> {
    let s = s.trim();
    if let Some(time) = parse_time_only(s) {
        return Ok(time);
    }
    if let Some((date_part, time_part)) = split_trailing_time(s) {
        let (y, m, d) = parse_date(date_part.trim())?;
        let time = parse_time_value_string(&time_part)?;
        let day = ymd_to_serial(y, m, d);
        return validate_vba_date_serial(if day < 0.0 { day - time } else { day + time });
    }
    let (y, m, d) = parse_date(s)?;
    validate_vba_date_serial(ymd_to_serial(y, m, d))
}
pub fn cverr(args: &[Variant]) -> LibResult<Variant> {
    let code = as_i32(need(args, 0)?)?;
    if !(0..=0xFFFF).contains(&code) {
        return Err(LibError::invalid_call(default_error_message(5)));
    }
    Ok(Variant::from_error_code(code))
}

// VBA numeric / type conversions. Each coerces its single argument to the named
// type with VBA banker's rounding and raises Overflow (6) when the rounded value is
// out of the target's range — matching `CLng`/`CInt`/etc.

/// Read a conversion argument as `f64`, parsing a numeric string (VBA `CDbl("3.5")`,
/// `CInt("&H20")`, etc.). `as_f64` alone does not parse strings — only numeric
/// Variants.
fn conv_f64(value: &Variant) -> LibResult<f64> {
    // A conversion function reads an error value as its number (`CLng(CVErr(5))` is 5).
    if let Some(code) = value
        .as_error_code()
        .filter(|_| value.vtype() == VarType::Error)
    {
        return Ok(f64::from(code));
    }
    if value.vtype() == VarType::String {
        let text = as_str(value)?;
        let exponent = text.find(['e', 'E', 'd', 'D']).unwrap_or(text.len());
        let (mantissa, exponent) = text.split_at(exponent);
        // Excel ignores spaces and nonbreaking spaces in the mantissa, but not
        // tabs or spaces inside the exponent ("1 000" works; "1e3 0" fails).
        let text = mantissa.replace([' ', '\u{a0}'], "") + exponent;
        return calcvba_runtime::coerce::parse_vba_number(&text)
            .map_err(|error| LibError::new(error.code(), default_error_message(error.code())));
    }
    as_f64(value)
}
/// Read a conversion argument as `i64` with VBA banker's rounding, parsing a numeric
/// string first. A non-string numeric keeps its exact value via `as_i64`.
fn conv_i64(value: &Variant) -> LibResult<i64> {
    if value.vtype() == VarType::String || value.vtype() == VarType::Error {
        let d = conv_f64(value)?;
        if !d.is_finite() || d.abs() >= 9.223_372_036_854_775e18 {
            return Err(LibError::overflow("integer overflow"));
        }
        return Ok(d.round_ties_even() as i64);
    }
    as_i64(value)
}

fn conv_literal_width_radix_i64(value: &Variant) -> LibResult<Option<i64>> {
    if value.vtype() != VarType::String {
        return Ok(None);
    }
    Ok(parse_vba_prefixed_integer(&as_str(value)?))
}

const DECIMAL96_MAX_MAGNITUDE: u128 = (1u128 << 96) - 1;

fn decimal96_from_magnitude(magnitude: u128, scale: u8, negative: bool) -> Decimal96 {
    Decimal96::from_parts(
        magnitude as u32,
        (magnitude >> 32) as u32,
        (magnitude >> 64) as u32,
        scale,
        negative && magnitude != 0,
    )
}

fn decimal96_zero() -> Decimal96 {
    Decimal96::from_parts(0, 0, 0, 0, false)
}

fn decimal96_from_i64(value: i64) -> Decimal96 {
    decimal96_from_magnitude(u128::from(value.unsigned_abs()), 0, value < 0)
}

fn decimal96_from_currency_scaled(scaled: i64) -> Decimal96 {
    decimal96_from_magnitude(u128::from(scaled.unsigned_abs()), 4, scaled < 0)
}

fn round_decimal_digits(digits: &mut Vec<u8>, discard: usize) {
    if discard == 0 {
        return;
    }
    let keep_len = digits.len().saturating_sub(discard);
    let first_dropped = digits.get(keep_len).copied().unwrap_or(0);
    let rest_nonzero = digits.iter().skip(keep_len + 1).any(|d| *d != 0);
    digits.truncate(keep_len);

    let kept_last_odd = digits.last().is_some_and(|d| d % 2 == 1);
    let round_up = first_dropped > 5 || (first_dropped == 5 && (rest_nonzero || kept_last_odd));
    if !round_up {
        return;
    }

    if digits.is_empty() {
        digits.push(1);
        return;
    }
    for digit in digits.iter_mut().rev() {
        if *digit < 9 {
            *digit += 1;
            return;
        }
        *digit = 0;
    }
    digits.insert(0, 1);
}

fn decimal96_from_digits(
    mut digits: Vec<u8>,
    mut scale: usize,
    negative: bool,
) -> Option<Decimal96> {
    while digits.first() == Some(&0) {
        digits.remove(0);
    }
    if digits.is_empty() {
        return Some(decimal96_zero());
    }

    if scale > 28 {
        round_decimal_digits(&mut digits, scale - 28);
        scale = 28;
        while digits.first() == Some(&0) {
            digits.remove(0);
        }
        if digits.is_empty() {
            return Some(decimal96_zero());
        }
    }

    while scale > 0 && digits.last() == Some(&0) {
        digits.pop();
        scale -= 1;
    }

    let mut magnitude = 0u128;
    for digit in digits {
        magnitude = magnitude.checked_mul(10)?.checked_add(u128::from(digit))?;
        if magnitude > DECIMAL96_MAX_MAGNITUDE {
            return None;
        }
    }
    Some(decimal96_from_magnitude(magnitude, scale as u8, negative))
}

fn parse_decimal96_text(text: &str) -> Option<Decimal96> {
    let t = text.trim();
    if t.is_empty() {
        return None;
    }
    if let Some(value) = parse_vba_prefixed_integer(t) {
        return Some(decimal96_from_i64(value));
    }

    let bytes = t.as_bytes();
    let mut pos = 0;
    let negative = match bytes.get(pos).copied() {
        Some(b'-') => {
            pos += 1;
            true
        }
        Some(b'+') => {
            pos += 1;
            false
        }
        _ => false,
    };

    let mut digits = Vec::new();
    let mut fractional_digits = 0usize;
    let mut saw_digit = false;
    let mut saw_dot = false;
    while let Some(byte) = bytes.get(pos).copied() {
        match byte {
            b'0'..=b'9' => {
                saw_digit = true;
                digits.push(byte - b'0');
                if saw_dot {
                    fractional_digits += 1;
                }
                pos += 1;
            }
            b'.' if !saw_dot => {
                saw_dot = true;
                pos += 1;
            }
            _ => break,
        }
    }
    if !saw_digit {
        return None;
    }

    let mut exponent = 0i32;
    if matches!(bytes.get(pos), Some(b'e' | b'E')) {
        pos += 1;
        let exp_negative = match bytes.get(pos).copied() {
            Some(b'-') => {
                pos += 1;
                true
            }
            Some(b'+') => {
                pos += 1;
                false
            }
            _ => false,
        };
        let exp_start = pos;
        let mut value = 0i32;
        while let Some(byte @ b'0'..=b'9') = bytes.get(pos).copied() {
            value = value.checked_mul(10)?.checked_add(i32::from(byte - b'0'))?;
            pos += 1;
        }
        if pos == exp_start {
            return None;
        }
        exponent = if exp_negative { -value } else { value };
    }
    if pos != bytes.len() {
        return None;
    }

    let scale = i64::try_from(fractional_digits).ok()? - i64::from(exponent);
    if scale < 0 {
        let zero_count = usize::try_from(-scale).ok()?;
        if zero_count > 60 {
            return None;
        }
        digits.extend(std::iter::repeat_n(0, zero_count));
        decimal96_from_digits(digits, 0, negative)
    } else {
        decimal96_from_digits(digits, usize::try_from(scale).ok()?, negative)
    }
}

fn decimal96_from_f64(value: f64) -> LibResult<Decimal96> {
    if !value.is_finite() {
        return Err(LibError::overflow("value does not fit in Decimal"));
    }
    if value == 0.0 {
        return Ok(decimal96_zero());
    }
    // A Double carries 15 significant digits into a Decimal.
    let scientific = format!("{:.14e}", value.abs());
    let (mantissa, exponent) = scientific.split_once('e').unwrap_or((&scientific, "0"));
    let exponent: i32 = exponent.parse().unwrap_or(0);
    let digits = mantissa.replace('.', "");
    let point = exponent + 1; // digits before the decimal point
    let mut text = String::new();
    if value < 0.0 {
        text.push('-');
    }
    if point <= 0 {
        text.push_str("0.");
        text.push_str(&"0".repeat((-point) as usize));
        text.push_str(&digits);
    } else if point as usize >= digits.len() {
        text.push_str(&digits);
        text.push_str(&"0".repeat(point as usize - digits.len()));
    } else {
        text.push_str(&digits[..point as usize]);
        text.push('.');
        text.push_str(&digits[point as usize..]);
    }
    if text.contains('.') {
        text.truncate(text.trim_end_matches('0').len());
    }
    parse_decimal96_text(&text).ok_or_else(|| LibError::overflow("value does not fit in Decimal"))
}

fn decimal96_from_variant(value: &Variant) -> LibResult<Decimal96> {
    match value.vtype() {
        VarType::Decimal => value
            .as_decimal96()
            .ok_or_else(|| LibError::type_mismatch("invalid Decimal variant payload")),
        VarType::String => {
            let text = as_str(value)?;
            match parse_decimal96_text(&text) {
                Some(value) => Ok(value),
                None if parse_vba_numeric_string(&text).is_some() => {
                    Err(LibError::overflow("value does not fit in Decimal"))
                }
                None => Err(LibError::type_mismatch("expected a numeric value")),
            }
        }
        VarType::Currency => Ok(decimal96_from_currency_scaled(
            value.as_currency_scaled_i64().unwrap_or(0),
        )),
        VarType::LongLong => Ok(decimal96_from_i64(value.as_i64().unwrap_or(0))),
        _ => decimal96_from_f64(as_f64(value)?),
    }
}

/// `CDbl` — to `Double` (no overflow; the full f64 range).
pub fn cdbl(args: &[Variant]) -> LibResult<Variant> {
    Ok(vf64(conv_f64(need(args, 0)?)?))
}
/// `CSng` — to `Single`; a value outside the `Single` range overflows.
pub fn csng(args: &[Variant]) -> LibResult<Variant> {
    let value = conv_f64(need(args, 0)?)?;
    let narrowed = value as f32;
    // Values just above MAX can still round to MAX. Check after rounding.
    if value.is_finite() && !narrowed.is_finite() {
        return Err(LibError::overflow("value does not fit in Single"));
    }
    Ok(Variant::from_f32(narrowed))
}
/// `CInt` — to `Integer` (banker's rounding, range −32768..32767).
pub fn cint(args: &[Variant]) -> LibResult<Variant> {
    let arg = need(args, 0)?;
    let value = match conv_literal_width_radix_i64(arg)? {
        Some(value) => value,
        None => conv_i64(arg)?,
    };
    let narrowed =
        i16::try_from(value).map_err(|_| LibError::overflow("value does not fit in Integer"))?;
    Ok(Variant::from_i16(narrowed))
}
/// `CLng` — to `Long` (banker's rounding, range −2^31..2^31−1).
pub fn clng(args: &[Variant]) -> LibResult<Variant> {
    let value = conv_i64(need(args, 0)?)?;
    let narrowed =
        i32::try_from(value).map_err(|_| LibError::overflow("value does not fit in Long"))?;
    Ok(vi32(narrowed))
}
/// `CLngLng` / `CLngPtr` (on the 64-bit runtime) — to `LongLong` (banker's rounding).
pub fn clnglng(args: &[Variant]) -> LibResult<Variant> {
    let arg = need(args, 0)?;
    let overflow = || LibError::overflow("value does not fit in LongLong");
    if let Some(value) = arg.as_i64() {
        return Ok(Variant::from_i64(value));
    }
    let decimal = if arg.vtype() == VarType::String {
        let text = as_str(arg)?;
        let text = text.trim();
        let (negative, body) = match text.as_bytes().first() {
            Some(b'-') => (true, &text[1..]),
            Some(b'+') => (false, &text[1..]),
            _ => (false, text),
        };
        if let Some(body) = body.strip_prefix('&') {
            let (radix, digits) = match body.as_bytes().first() {
                Some(b'h' | b'H') => (16, &body[1..]),
                Some(b'o' | b'O') => (8, &body[1..]),
                _ => (8, body),
            };
            if digits.is_empty() || !digits.chars().all(|c| c.is_digit(radix)) {
                return Err(LibError::type_mismatch("expected a numeric value"));
            }
            // A CLngLng string uses the target's 64-bit two's-complement width,
            // including when the same token would be a signed Long literal.
            let value = u64::from_str_radix(digits, radix).map_err(|_| overflow())? as i64;
            let value = if negative {
                value.checked_neg().ok_or_else(overflow)?
            } else {
                value
            };
            return Ok(Variant::from_i64(value));
        }
        parse_decimal96_text(text)
    } else {
        arg.as_decimal96().or_else(|| {
            arg.as_currency_scaled_i64()
                .map(decimal96_from_currency_scaled)
        })
    };
    if let Some(decimal) = decimal {
        let value = decimal
            .round_scaled_i128(0)
            .and_then(|value| i64::try_from(value).ok())
            .ok_or_else(overflow)?;
        return Ok(Variant::from_i64(value));
    }
    let rounded = conv_f64(arg)?.round_ties_even();
    // -2^63 is representable, whereas +2^63 (also i64::MAX as Double) overflows.
    if !rounded.is_finite()
        || !(-9_223_372_036_854_775_808.0..9_223_372_036_854_775_808.0).contains(&rounded)
    {
        return Err(overflow());
    }
    Ok(Variant::from_i64(rounded as i64))
}
/// `CByte` — to `Byte` (banker's rounding, range 0..255).
pub fn cbyte(args: &[Variant]) -> LibResult<Variant> {
    let arg = need(args, 0)?;
    // `CByte(True)` is 255 (True's low byte).
    if arg.vtype() == VarType::Boolean {
        return Ok(Variant::from_u8(if arg.as_bool().unwrap_or(false) {
            255
        } else {
            0
        }));
    }
    let value = conv_i64(arg)?;
    let narrowed =
        u8::try_from(value).map_err(|_| LibError::overflow("value does not fit in Byte"))?;
    Ok(Variant::from_u8(narrowed))
}
/// `CBool` — to `Boolean` (0 → False, any non-zero → True; `"True"`/`"False"` too).
pub fn cbool(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    if value.vtype() == VarType::String {
        // The `"True"`/`"False"` literals share one recognizer with the implicit Boolean
        // Let-coercion (`arith::coerce_numeric`) so explicit `CBool` and implicit assignment
        // agree; any other string falls through to the numeric path (non-zero → True).
        if let Some(b) = calcvba_runtime::coerce::parse_bool_text(&as_str(value)?) {
            return Ok(vbool(b));
        }
    }
    Ok(vbool(conv_f64(value)? != 0.0))
}
/// `CCur` — to `Currency` (a fixed 4-decimal scaled i64); out of range overflows.
pub fn ccur(args: &[Variant]) -> LibResult<Variant> {
    let arg = need(args, 0)?;
    let overflow = || LibError::overflow("value does not fit in Currency");
    // Exact sources keep every digit; a Double converts from its exact binary value.
    let scaled = match arg.vtype() {
        VarType::Currency => arg.as_currency_scaled_i64(),
        VarType::String => Some(
            calcvba_runtime::coerce::currency_scaled_from_text(&as_str(arg)?).map_err(|error| {
                LibError::new(error.code(), default_error_message(error.code()))
            })?,
        ),
        VarType::Decimal => arg
            .as_decimal96()
            .and_then(|d| d.round_scaled_i128(4))
            .and_then(|n| i64::try_from(n).ok()),
        VarType::LongLong => arg
            .as_i64()
            .and_then(|n| n.checked_mul(calcvba_runtime::CurrencyValue::SCALE)),
        _ => calcvba_runtime::coerce::currency_scaled_from_f64(conv_f64(arg)?),
    };
    Ok(Variant::from_currency_scaled_i64(
        scaled.ok_or_else(overflow)?,
    ))
}
/// `CDec` — to a Variant carrying the Decimal subtype.
pub fn cdec(args: &[Variant]) -> LibResult<Variant> {
    let value = decimal96_from_variant(need(args, 0)?)?;
    Ok(Variant::from_decimal96(value))
}
/// `CVar` — to `Variant`: the value unchanged (everything is already a Variant).
pub fn cvar(args: &[Variant]) -> LibResult<Variant> {
    Ok(need(args, 0)?.clone())
}

// ── Random ───────────────────────────────────────────────────────────────────
//
// VBA's `Rnd` is the VB6/VBA 24-bit linear-congruential generator:
//   x = (x * 0x43FD43FD + 0xC39EC3) mod 2^24,  result = x / 2^24  (as Single).
// `Rnd(0)` repeats the last number; `Rnd(n<0)` reseeds deterministically from
// the Single bit-pattern of `n`; `Randomize [n]` reseeds the high word from the
// Double bit-pattern. State is the 24-bit value in `LibContext::rng_state`.

const RND_MULT: u64 = 0x43FD_43FD;
const RND_INC: u64 = 0x00C3_9EC3;
const RND_MASK: u64 = 0x00FF_FFFF; // 2^24 - 1
const RND_SCALE: f64 = 16_777_216.0; // 2^24

fn rng_step(ctx: &mut LibContext) -> f64 {
    ctx.rng_state = ctx.rng_state.wrapping_mul(RND_MULT).wrapping_add(RND_INC) & RND_MASK;
    ctx.rng_state as f64 / RND_SCALE
}

fn rnd_negative_seed(value: f64) -> u64 {
    let bits = u64::from((value as f32).to_bits());
    ((bits & RND_MASK) + ((bits >> 24) & 0xFF)) & RND_MASK
}

/// `Randomize` XORs the two 16-bit halves of the Double's high 32-bit word.
/// They replace the middle two bytes of the 24-bit state.
fn randomize_seed_word(value: f64) -> u64 {
    let bits = value.to_bits();
    ((bits >> 32) ^ (bits >> 48)) & 0xFFFF
}

/// An optional numeric argument, treating an `Empty` placeholder as omitted.
fn numeric_arg(args: &[Variant], index: usize) -> Option<&Variant> {
    opt(args, index).filter(|v| !matches!(v.vtype(), VarType::Empty) && !missing_sentinel(v))
}

pub fn rnd(args: &[Variant], ctx: &mut LibContext) -> LibResult<Variant> {
    let value = match opt(args, 0) {
        Some(v) => as_f64(v)?,
        None => 1.0, // omitted ⇒ next in sequence
    };
    let result = if value == 0.0 {
        // Rnd(0): return the most recently generated number, unchanged.
        ctx.rng_state as f64 / RND_SCALE
    } else if value < 0.0 {
        // Rnd(n<0): deterministic reseed from the Single bit pattern of n.
        ctx.rng_state = rnd_negative_seed(value);
        rng_step(ctx)
    } else {
        rng_step(ctx)
    };
    Ok(Variant::from_f32(result as f32))
}

pub fn randomize(args: &[Variant], ctx: &mut LibContext) -> LibResult<Variant> {
    // The caller supplies the host timer when the source argument is omitted, before
    // borrowing the contextual state. Explicit Empty coerces to zero.
    let seed = randomize_seed_word(as_f64(need(args, 0)?)?);
    // VBA replaces the high 16 bits of the 24-bit state, preserving the low byte.
    ctx.rng_state = ((ctx.rng_state & 0xFF) | ((seed & 0xFFFF) << 8)) & RND_MASK;
    Ok(vunit())
}

// ── Financial ──────────────────────────────────────────────────────────────────

fn opt_f64(args: &[Variant], index: usize, default: f64) -> LibResult<f64> {
    match arg_present(args, index) {
        Some(v) => as_f64(v),
        None => Ok(default),
    }
}

/// The `due` / `type` argument: any non-zero value means payments at the start of the period.
fn due_flag(args: &[Variant], index: usize) -> LibResult<f64> {
    Ok(if opt_f64(args, index, 0.0)? != 0.0 {
        1.0
    } else {
        0.0
    })
}

/// A financial result; a value that is not a finite number is "Invalid procedure call".
fn finite_result(x: f64) -> LibResult<Variant> {
    if x.is_finite() {
        Ok(vf64(x))
    } else {
        Err(LibError::invalid_call("financial result is not a number"))
    }
}

fn cashflows(v: &Variant) -> LibResult<Vec<f64>> {
    match v.as_safearray() {
        Some(arr) => arr
            .variant_elements()
            .unwrap_or_default()
            .iter()
            .map(as_f64)
            .collect(),
        None => Err(LibError::type_mismatch("expected cash-flow array")),
    }
}

/// Future value of `pv` plus `nper` payments `pmt` at periodic rate `i`
/// (`t` = 0 ordinary / end-of-period, 1 annuity-due / begin-of-period). Shared by
/// `FV` and the `IPmt`/`PPmt` balance computation.
fn fv_value(i: f64, nper: f64, pmt: f64, pv: f64, t: f64) -> f64 {
    if i == 0.0 {
        -(pv + pmt * nper)
    } else {
        let g = (1.0 + i).powf(nper);
        -(pv * g + pmt * (1.0 + i * t) * (g - 1.0) / i)
    }
}

/// Periodic payment for present value `pv` reaching future value `fv` over `nper`
/// periods at rate `i` (`t` as in [`fv_value`]). Shared by `Pmt` and `IPmt`/`PPmt`.
fn pmt_value(i: f64, nper: f64, pv: f64, fv: f64, t: f64) -> f64 {
    if i == 0.0 {
        -(pv + fv) / nper
    } else {
        let g = (1.0 + i).powf(nper);
        -(fv + pv * g) / ((1.0 + i * t) * (g - 1.0) / i)
    }
}

/// The interest portion of the `per`-th payment (the basis for `IPmt`/`PPmt`). The
/// interest is the balance carried into the period times the rate; for an annuity
/// due (`t = 1`) the first period carries no interest and later periods discount by
/// one period — the numpy-financial/Excel convention, live-verified against VBA
/// (`IPmt(0.1/12, 2, 36, 10000, 0, 1) = -80.6667`).
fn ipmt_value(rate: f64, per: f64, nper: f64, pv: f64, fv: f64, t: f64) -> f64 {
    let pmt = pmt_value(rate, nper, pv, fv, t);
    let mut interest = fv_value(rate, per - 1.0, pmt, pv, t) * rate;
    if t == 1.0 {
        if per == 1.0 {
            interest = 0.0;
        } else {
            interest /= 1.0 + rate;
        }
    }
    interest
}

pub fn fv(args: &[Variant]) -> LibResult<Variant> {
    let (i, n, pmt) = (
        as_f64(need(args, 0)?)?,
        as_f64(need(args, 1)?)?,
        as_f64(need(args, 2)?)?,
    );
    let pv = opt_f64(args, 3, 0.0)?;
    let t = due_flag(args, 4)?;
    finite_result(fv_value(i, n, pmt, pv, t))
}

pub fn pv(args: &[Variant]) -> LibResult<Variant> {
    let (i, n, pmt) = (
        as_f64(need(args, 0)?)?,
        as_f64(need(args, 1)?)?,
        as_f64(need(args, 2)?)?,
    );
    let fv = opt_f64(args, 3, 0.0)?;
    let t = due_flag(args, 4)?;
    finite_result(if i == 0.0 {
        -(fv + pmt * n)
    } else {
        let g = (1.0 + i).powf(n);
        -(fv + pmt * (1.0 + i * t) * (g - 1.0) / i) / g
    })
}

pub fn pmt(args: &[Variant]) -> LibResult<Variant> {
    let (i, n, pv) = (
        as_f64(need(args, 0)?)?,
        as_f64(need(args, 1)?)?,
        as_f64(need(args, 2)?)?,
    );
    let fv = opt_f64(args, 3, 0.0)?;
    let t = due_flag(args, 4)?;
    if n == 0.0 {
        return Err(LibError::invalid_call("Pmt requires nper <> 0"));
    }
    finite_result(pmt_value(i, n, pv, fv, t))
}

pub fn ipmt(args: &[Variant]) -> LibResult<Variant> {
    let rate = as_f64(need(args, 0)?)?;
    let per = as_f64(need(args, 1)?)?;
    let nper = as_f64(need(args, 2)?)?;
    let pv = as_f64(need(args, 3)?)?;
    let fv = opt_f64(args, 4, 0.0)?;
    let t = due_flag(args, 5)?;
    if nper <= 0.0 || per < 1.0 || per > nper {
        return Err(LibError::invalid_call("IPmt period out of range"));
    }
    finite_result(ipmt_value(rate, per, nper, pv, fv, t))
}

pub fn ppmt(args: &[Variant]) -> LibResult<Variant> {
    let rate = as_f64(need(args, 0)?)?;
    let per = as_f64(need(args, 1)?)?;
    let nper = as_f64(need(args, 2)?)?;
    let pv = as_f64(need(args, 3)?)?;
    let fv = opt_f64(args, 4, 0.0)?;
    let t = due_flag(args, 5)?;
    if nper <= 0.0 || per < 1.0 || per > nper {
        return Err(LibError::invalid_call("PPmt period out of range"));
    }
    let payment = pmt_value(rate, nper, pv, fv, t);
    finite_result(payment - ipmt_value(rate, per, nper, pv, fv, t))
}

pub fn sln(args: &[Variant]) -> LibResult<Variant> {
    let cost = as_f64(need(args, 0)?)?;
    let salvage = as_f64(need(args, 1)?)?;
    let life = as_f64(need(args, 2)?)?;
    if life == 0.0 {
        return Err(LibError::invalid_call("SLN requires life <> 0"));
    }
    finite_result((cost - salvage) / life)
}

pub fn syd(args: &[Variant]) -> LibResult<Variant> {
    let cost = as_f64(need(args, 0)?)?;
    let salvage = as_f64(need(args, 1)?)?;
    let life = as_f64(need(args, 2)?)?;
    let period = as_f64(need(args, 3)?)?;
    if life <= 0.0 || period <= 0.0 || period > life {
        return Err(LibError::invalid_call("SYD period out of range"));
    }
    finite_result((cost - salvage) * (life - period + 1.0) * 2.0 / (life * (life + 1.0)))
}

/// `DDB(cost, salvage, life, period, [factor = 2])` → double-declining-balance
/// depreciation for `period`. Compute the period's payment and subtract any
/// excess depreciation beyond salvage, in the order measured by `P_DDB`.
pub fn ddb(args: &[Variant]) -> LibResult<Variant> {
    let cost = as_f64(need(args, 0)?)?;
    let salvage = as_f64(need(args, 1)?)?;
    let life = as_f64(need(args, 2)?)?;
    let period = as_f64(need(args, 3)?)?;
    let factor = opt_f64(args, 4, 2.0)?;
    if salvage < 0.0 || period <= 0.0 || period > life || factor <= 0.0 {
        return Err(LibError::invalid_call("DDB argument out of range"));
    }
    if cost <= 0.0 {
        return Ok(vf64(0.0));
    }
    if life < 2.0 || (life == 2.0 && period <= 1.0) {
        return Ok(vf64(cost - salvage));
    }
    if life == 2.0 {
        return Ok(vf64(0.0));
    }
    if period <= 1.0 {
        return Ok(vf64((cost * factor / life).min(cost - salvage)));
    }
    let rate = (life - factor) / life;
    // Excel rounds the powers to Double, then evaluates each payment/total
    // expression in extended precision before storing it as a Double.
    let x = X80::of;
    let growth = rate.powf(period - 1.0);
    let end = rate.powf(period);
    if !growth.is_finite() || !end.is_finite() {
        return Err(LibError::overflow("DDB power is not finite"));
    }
    let payment = x(factor)
        .mul(x(cost))
        .div(x(life))
        .mul(x(growth))
        .to_f64()
        .ok_or_else(|| LibError::overflow("DDB"))?;
    let total = x(cost)
        .mul(x(1.0).sub(x(end)))
        .to_f64()
        .ok_or_else(|| LibError::overflow("DDB"))?;
    let excess = total - cost + salvage;
    let depreciation = payment - excess.max(0.0);
    finite_result(depreciation.max(0.0))
}

pub fn nper(args: &[Variant]) -> LibResult<Variant> {
    let (i, pmt, pv) = (
        as_f64(need(args, 0)?)?,
        as_f64(need(args, 1)?)?,
        as_f64(need(args, 2)?)?,
    );
    let fv = opt_f64(args, 3, 0.0)?;
    let t = due_flag(args, 4)?;
    if i <= -1.0 || (i == 0.0 && pmt == 0.0) {
        return Err(LibError::invalid_call("NPer argument out of range"));
    }
    if i == 0.0 {
        return finite_result(-(pv + fv) / pmt);
    }
    let payment = if t != 0.0 {
        pmt * (1.0 + i) / i
    } else {
        pmt / i
    };
    let (mut future, mut present) = (-fv + payment, pv + payment);
    if future < 0.0 && present < 0.0 {
        future = -future;
        present = -present;
    } else if future <= 0.0 || present <= 0.0 {
        return Err(LibError::invalid_call("NPer logarithm out of range"));
    }
    // Separate logarithms preserve Excel's rounding, including its cancellation
    // at tiny rates (P_NPer.13). ln(future/present) gives a different answer.
    finite_result((future.ln() - present.ln()) / (1.0 + i).ln())
}

pub fn npv(args: &[Variant]) -> LibResult<Variant> {
    let rate = as_f64(need(args, 0)?)?;
    if rate == -1.0 {
        return Err(LibError::invalid_call("NPV requires rate <> -1"));
    }
    let array = need(args, 1)?;
    if array.vtype() == VarType::ArrayVariant && array.as_safearray().is_none() {
        return Err(LibError::invalid_call("NPV requires an allocated array"));
    }
    let flows = cashflows(array)?;
    let x = X80::of;
    let rate = x(rate);
    let (mut denominator, mut acc) = (x(1.0), x(0.0));
    // Excel retains the denominator and sum in extended precision. Its ordered
    // denominator update also differs from multiplication by (1 + rate) near -1.
    for cf in flows {
        denominator = denominator.add(denominator.mul(rate));
        acc = acc.add(x(cf).div(denominator));
    }
    acc.to_f64()
        .map(vf64)
        .ok_or_else(|| LibError::overflow("NPV total"))
}

pub fn irr(args: &[Variant]) -> LibResult<Variant> {
    let no_root = || LibError::invalid_call("IRR failed to converge");
    let array = need(args, 0)?;
    if array.vtype() == VarType::ArrayVariant && array.as_safearray().is_none() {
        return Err(no_root());
    }
    let flows = cashflows(array)?;
    // A rate exists only when the flows change sign.
    if !flows.iter().any(|c| *c > 0.0) || !flows.iter().any(|c| *c < 0.0) {
        return Err(LibError::invalid_call(
            "IRR requires positive and negative flows",
        ));
    }
    // Evaluate each discounted payment in order. The polynomial has the same
    // roots, but Newton's method on it chooses different roots and accepts
    // cases for which Excel reports error 5 (P_IRR.3/.10/.13/.14/.21).
    let npv = |rate: f64| -> LibResult<f64> {
        if rate <= -1.0 || !rate.is_finite() {
            return Err(no_root());
        }
        // VBA keeps the running discount factor and sum in x87 extended precision.
        // Rounding each payment to Double loses the captured IRR.1/.11/.15/.19 bits.
        let x = X80::of;
        let (mut denominator, mut total) = (x(1.0), x(0.0));
        for flow in &flows {
            denominator = denominator.add(denominator.mul(x(rate)));
            total = total.add(x(*flow).div(denominator));
        }
        total.to_f64().ok_or_else(no_root)
    };
    let tolerance = flows.iter().map(|v| v.abs()).fold(0.0, f64::max) * 1e-9;
    let mut rate0 = opt_f64(args, 1, 0.1)?;
    let mut value0 = npv(rate0)?;
    if value0.abs() < tolerance {
        return Ok(vf64(rate0));
    }
    let mut rate1 = rate0 + if value0 > 0.0 { 1e-5 } else { -1e-5 };
    let mut value1 = npv(rate1)?;
    for _ in 0..40 {
        if value0 == value1 {
            rate0 += if rate1 > rate0 { -1e-5 } else { 1e-5 };
            value0 = npv(rate0)?;
            if value0 == value1 {
                return Err(no_root());
            }
        }
        rate0 = rate1 - (rate1 - rate0) * value1 / (value1 - value0);
        value0 = npv(rate0)?;
        if value0.abs() < tolerance && (rate0 - rate1).abs() < 1e-7 {
            return Ok(vf64(rate0));
        }
        std::mem::swap(&mut rate0, &mut rate1);
        std::mem::swap(&mut value0, &mut value1);
    }
    Err(no_root())
}

pub fn mirr(args: &[Variant]) -> LibResult<Variant> {
    let invalid = || LibError::invalid_call("MIRR argument out of range");
    let array = need(args, 0)?;
    if array.vtype() == VarType::ArrayVariant && array.as_safearray().is_none() {
        return Err(invalid());
    }
    let flows = cashflows(array)?;
    let finance = as_f64(need(args, 1)?)?;
    let reinvest = as_f64(need(args, 2)?)?;
    let n = flows.len() as f64;
    if finance == -1.0 || reinvest == -1.0 || n <= 1.0 {
        return Err(invalid());
    }
    let present_value = |rate: f64, positive: bool| {
        let (mut denominator, mut total) = (1.0, 0.0);
        for flow in &flows {
            denominator += denominator * rate;
            if if positive { *flow > 0.0 } else { *flow < 0.0 } {
                total += flow / denominator;
            }
        }
        total
    };
    let negative = present_value(finance, false);
    if negative == 0.0 {
        return Err(LibError::new(11, "MIRR has no negative cash flow"));
    }
    let positive = present_value(reinvest, true);
    let growth = -positive * (reinvest + 1.0).powf(n) / (negative * (finance + 1.0));
    if growth < 0.0 {
        return Err(invalid());
    }
    finite_result(growth.powf(1.0 / (n - 1.0)) - 1.0)
}

/// An x87 80-bit extended value: a 64-bit significand rounded to nearest-even after every
/// operation, `mant * 2^exp` with bit 63 of `mant` set (or `mant == 0` for zero). VBA's `Rate`
/// evaluates its function and its secant step in this precision and rounds to Double only where
/// it stores a variable; Double-only or exact arithmetic each leave one-ulp differences in
/// RateShape (`lin_fv1.5`, `lin_fv2`, `far_n1_12`, `a360_g*`, ...). Finite values only; the
/// exponent is unbounded here, and range is checked when a value is stored as a Double.
#[derive(Clone, Copy)]
struct X80 {
    neg: bool,
    mant: u64,
    exp: i32,
}

/// Round `mant * 2^exp` to `bits` significant bits (nearest, ties to even). Callers keep at
/// least two bits below the target and fold any discarded low bits into the last bit (sticky).
fn round_significand(mant: u128, exp: i32, bits: u32) -> (u128, i32) {
    if mant == 0 {
        return (0, 0);
    }
    let width = 128 - mant.leading_zeros();
    if width <= bits {
        let s = bits - width;
        return (mant << s, exp - s as i32);
    }
    let shift = width - bits;
    let mut m = mant >> shift;
    let rem = mant & ((1u128 << shift) - 1);
    let half = 1u128 << (shift - 1);
    if rem > half || (rem == half && m & 1 == 1) {
        m += 1;
    }
    let mut e = exp + shift as i32;
    if m >> bits != 0 {
        m >>= 1;
        e += 1;
    }
    (m, e)
}

impl X80 {
    fn new(neg: bool, mant: u128, exp: i32) -> X80 {
        let (m, e) = round_significand(mant, exp, 64);
        X80 {
            neg: neg && m != 0,
            mant: m as u64,
            exp: e,
        }
    }

    /// A finite Double, exactly.
    fn of(x: f64) -> X80 {
        let bits = x.to_bits();
        let neg = bits >> 63 != 0;
        let biased = ((bits >> 52) & 0x7ff) as i32;
        let frac = bits & ((1u64 << 52) - 1);
        let (m, e) = if biased == 0 {
            (frac, -1074)
        } else {
            (frac | (1u64 << 52), biased - 1075)
        };
        X80::new(neg, u128::from(m), e)
    }

    fn mul(self, o: X80) -> X80 {
        X80::new(
            self.neg != o.neg,
            u128::from(self.mant) * u128::from(o.mant),
            self.exp + o.exp,
        )
    }

    /// `o` must be non-zero.
    fn div(self, o: X80) -> X80 {
        if self.mant == 0 {
            return self;
        }
        // Two long-division steps give a 125+ bit quotient; the remainder is the sticky bit.
        let d = u128::from(o.mant);
        let n = u128::from(self.mant) << 64;
        let (q1, r1) = (n / d, n % d);
        let n2 = r1 << 62;
        let (q2, r2) = (n2 / d, n2 % d);
        let q = (q1 << 62) | q2 | u128::from(r2 != 0);
        X80::new(self.neg != o.neg, q, self.exp - o.exp - 126)
    }

    fn add(self, o: X80) -> X80 {
        if o.mant == 0 {
            return self;
        }
        if self.mant == 0 {
            return o;
        }
        let (big, small) = if (self.exp, self.mant) >= (o.exp, o.mant) {
            (self, o)
        } else {
            (o, self)
        };
        // 62 guard bits, with the bits shifted out of the smaller operand kept as a sticky bit.
        let a = u128::from(big.mant) << 62;
        let b = u128::from(small.mant) << 62;
        let d = (big.exp - small.exp) as u32;
        let b = if d >= 126 {
            1
        } else {
            (b >> d) | u128::from(b & ((1u128 << d) - 1) != 0)
        };
        let m = if big.neg == small.neg { a + b } else { a - b };
        X80::new(big.neg, m, big.exp - 62)
    }

    fn sub(self, o: X80) -> X80 {
        self.add(X80 { neg: !o.neg, ..o })
    }

    /// The value stored as a Double; `None` when it overflows the Double range.
    fn to_f64(self) -> Option<f64> {
        let sign = u64::from(self.neg) << 63;
        if self.mant == 0 {
            return Some(f64::from_bits(sign));
        }
        if self.exp + 63 >= -1022 {
            let (m, e) = round_significand(u128::from(self.mant), self.exp, 53);
            let top = e + 52;
            if top > 1023 {
                return None;
            }
            let biased = (top + 1023) as u64;
            Some(f64::from_bits(
                sign | (biased << 52) | (m as u64 & ((1u64 << 52) - 1)),
            ))
        } else {
            // Subnormal: round to a multiple of 2^-1074.
            let s = (-(self.exp + 1074)) as u32;
            if s >= 66 {
                return Some(f64::from_bits(sign));
            }
            let mant = u128::from(self.mant);
            let mut m = mant >> s;
            let rem = mant & ((1u128 << s) - 1);
            let half = 1u128 << (s - 1);
            if rem > half || (rem == half && m & 1 == 1) {
                m += 1;
            }
            Some(f64::from_bits(sign | m as u64))
        }
    }
}

/// A double-double value `hi + lo`, for a `pow` accurate well past Double precision.
#[derive(Clone, Copy)]
struct Dd(f64, f64);

fn quick_two_sum(a: f64, b: f64) -> Dd {
    let s = a + b;
    Dd(s, b - (s - a))
}

fn two_sum(a: f64, b: f64) -> Dd {
    let s = a + b;
    let bb = s - a;
    Dd(s, (a - (s - bb)) + (b - bb))
}

impl Dd {
    fn add(self, o: Dd) -> Dd {
        let Dd(s, e) = two_sum(self.0, o.0);
        let Dd(t, f) = two_sum(self.1, o.1);
        let Dd(s, e) = quick_two_sum(s, e + t);
        quick_two_sum(s, e + f)
    }

    fn mul(self, o: Dd) -> Dd {
        let p = self.0 * o.0;
        let e = self.0.mul_add(o.0, -p);
        quick_two_sum(p, e + (self.0 * o.1 + self.1 * o.0))
    }

    fn div_f64(self, d: f64) -> Dd {
        let q = self.0 / d;
        let p = q * d;
        let e = q.mul_add(d, -p);
        quick_two_sum(q, (self.0 - p - e + self.1) / d)
    }

    fn scale(self, k: i32) -> Dd {
        // 2^k in two factors so neither overflows on the way to a representable result.
        let a = 2f64.powi(k / 2);
        let b = 2f64.powi(k - k / 2);
        Dd(self.0 * a * b, self.1 * a * b)
    }
}

/// `exp(a)` in double-double: `a = k ln 2 + r`, `expm1(r / 2^10)` by Taylor series, doubled
/// back with `expm1(2x) = expm1(x) (expm1(x) + 2)`, then scaled by `2^k`.
fn exp_dd(a: Dd) -> Dd {
    const LN2: Dd = Dd(std::f64::consts::LN_2, 2.319_046_813_846_299_6e-17);
    if a.0 > 709.8 {
        return Dd(f64::INFINITY, 0.0);
    }
    if a.0 < -745.2 {
        return Dd(0.0, 0.0);
    }
    let k = (a.0 / LN2.0).round();
    let r = a.add(Dd(-k, 0.0).mul(LN2));
    let s = Dd(r.0 / 1024.0, r.1 / 1024.0);
    let (mut term, mut sum) = (s, s);
    for i in 2..=12 {
        term = term.mul(s).div_f64(f64::from(i));
        sum = sum.add(term);
    }
    for _ in 0..10 {
        sum = sum.mul(sum.add(Dd(2.0, 0.0)));
    }
    sum.add(Dd(1.0, 0.0)).scale(k as i32)
}

/// `ln(x)` for finite `x > 0` in double-double: two Newton steps `y += x e^-y - 1`.
fn ln_dd(x: f64) -> Dd {
    let mut y = Dd(x.ln(), 0.0);
    for _ in 0..2 {
        let t = exp_dd(Dd(-y.0, -y.1)).mul(Dd(x, 0.0));
        y = y.add(t.add(Dd(-1.0, 0.0)));
    }
    y
}

/// `base ^ n` rounded once to Double (RateShape `a3exact` needs the correctly rounded
/// `1.0968289920724807 ^ 3`, which a libm `pow` misses by one ulp). Whole `n` up to 2^20 by
/// binary powering, otherwise `exp(n ln base)`, both in double-double.
fn pow_rounded(base: f64, n: f64) -> f64 {
    if n == 0.0 || base == 1.0 {
        return 1.0;
    }
    if base == 0.0 {
        return if n > 0.0 { 0.0 } else { f64::INFINITY };
    }
    let whole = n.fract() == 0.0;
    if base < 0.0 && !whole {
        return f64::NAN;
    }
    let odd = whole && n.abs() < 9.007_199_254_740_992e15 && (n % 2.0).abs() == 1.0;
    let sign = if base < 0.0 && odd { -1.0 } else { 1.0 };
    let b = base.abs();
    let v = if whole && (1.0..=1_048_576.0).contains(&n) {
        let mut k = n as u32;
        let mut x = Dd(b, 0.0);
        let mut acc = Dd(1.0, 0.0);
        loop {
            if k & 1 == 1 {
                acc = acc.mul(x);
            }
            k >>= 1;
            if k == 0 || !acc.0.is_finite() {
                break;
            }
            x = x.mul(x);
        }
        acc
    } else {
        let l = ln_dd(b);
        let p = Dd(n, 0.0).mul(l);
        exp_dd(p)
    };
    if !v.0.is_finite() {
        return f64::INFINITY * sign;
    }
    sign * (v.0 + v.1)
}

/// `Rate`'s function of the rate `r`, `pv (1+r)^n + pmt (1+r type) ((1+r)^n - 1) / r + fv`,
/// as VBA evaluates it: `1+r` and `(1+r)^n` stored as Doubles, the rest in x87 extended
/// precision, the result stored as a Double. `(1+r)^n` or the result overflowing is error 6
/// (RateShape `a360_far_g10`, `a360_far_g50`); a NaN (a negative base) is error 5.
fn rate_eval(r: f64, nper: f64, pmt: f64, pv: f64, fv: f64, due: bool) -> LibResult<f64> {
    let overflow = || LibError::overflow("Rate");
    let x = X80::of;
    if !r.is_finite() {
        return Err(overflow());
    }
    let v = if r == 0.0 {
        x(pv).add(x(pmt).mul(x(nper))).add(x(fv))
    } else {
        let base = x(r).add(x(1.0)).to_f64().ok_or_else(overflow)?;
        let growth = pow_rounded(base, nper);
        if growth.is_nan() {
            return Err(LibError::invalid_call("Rate failed to converge"));
        }
        if growth.is_infinite() {
            return Err(overflow());
        }
        let g = x(growth);
        let t2 = if due { x(base) } else { x(1.0) };
        let annuity = x(pmt).mul(t2).mul(g.sub(x(1.0))).div(x(r));
        x(pv).mul(g).add(annuity).add(x(fv))
    };
    v.to_f64().ok_or_else(overflow)
}

/// `Rate(nper, pmt, pv, [fv], [type], [guess])`, VBA's secant iteration, pinned bit for bit by
/// `conformance/excel-measurements/rate-shape` (dc-bjt7v):
/// - The second point is `guess / 2` when `f(guess) > 0`, else `guess * 2`. Plain Newton would
///   solve the linear `Rate(1, 0, -1, 20)` exactly; Excel's 18.9999999999997 from guess 0.1 is
///   this one secant step from (0.1, 0.05).
/// - Each step is `r1 - (r1 - r0) f1 / (f1 - f0)` in extended precision, stored as a Double.
/// - It stops when `|f| < 1e-7` (the `lin_fv*`, `far_n1_*` residuals; 1e-6 or 1e-8 miss lines).
/// - When the two latest `f` are equal (guess 0: both points are 0) the older point moves
///   by 1e-5 away from the newer one (`lin19_g0`; 1e-4 and 1e-6 miss); still equal is error 5.
/// - At most 40 steps (`a360_g0` takes 35; 100 would let `noroot`/`zero_rate` converge), then
///   error 5, which is also how an exact root of 0 ends (`zero_rate`, `a5mixed`).
pub fn rate(args: &[Variant]) -> LibResult<Variant> {
    const TOLERANCE: f64 = 1e-7;
    const NUDGE: f64 = 1e-5;
    const MAX_STEPS: usize = 40;
    let nper = as_f64(need(args, 0)?)?;
    let pmt = as_f64(need(args, 1)?)?;
    let pv = as_f64(need(args, 2)?)?;
    let fv = opt_f64(args, 3, 0.0)?;
    let due = due_flag(args, 4)? != 0.0;
    if nper <= 0.0 {
        return Err(LibError::invalid_call("Rate requires nper > 0"));
    }
    let no_root = || LibError::invalid_call("Rate failed to converge");
    let f = |r: f64| rate_eval(r, nper, pmt, pv, fv, due);
    let x = X80::of;
    let mut r0 = opt_f64(args, 5, 0.1)?;
    let mut y0 = f(r0)?;
    let mut r1 = if y0 > 0.0 { r0 / 2.0 } else { r0 * 2.0 };
    let mut y1 = f(r1)?;
    for _ in 0..MAX_STEPS {
        if y1 == y0 {
            let moved = if r1 > r0 {
                x(r0).sub(x(NUDGE))
            } else {
                x(r0).add(x(NUDGE))
            };
            r0 = moved.to_f64().ok_or_else(no_root)?;
            y0 = f(r0)?;
            if y1 == y0 {
                return Err(no_root());
            }
        }
        let step = x(r1).sub(x(r0)).mul(x(y1)).div(x(y1).sub(x(y0)));
        r0 = x(r1)
            .sub(step)
            .to_f64()
            .ok_or_else(|| LibError::overflow("Rate"))?;
        y0 = f(r0)?;
        if y0.abs() < TOLERANCE {
            return Ok(vf64(r0));
        }
        std::mem::swap(&mut r0, &mut r1);
        std::mem::swap(&mut y0, &mut y1);
    }
    Err(no_root())
}

// ── Colour ───────────────────────────────────────────────────────────────────────

pub fn rgb(args: &[Variant]) -> LibResult<Variant> {
    // A component above 255 clamps to 255; a negative one is invalid.
    let component = |index: usize| -> LibResult<i32> {
        let value = as_i32(need(args, index)?)?;
        if value < 0 {
            return Err(LibError::invalid_call("RGB component must not be negative"));
        }
        Ok(value.min(255))
    };
    let (red, green, blue) = (component(0)?, component(1)?, component(2)?);
    Ok(vi32(red + green * 256 + blue * 65_536))
}

/// `QBColor(colorIndex)` → the `Long` packed colour for one of the 16 legacy
/// QuickBasic console colours. The table is live-verified against VBA; an index
/// outside `0..=15` raises "Invalid procedure call or argument" (error 5).
pub fn qb_color(args: &[Variant]) -> LibResult<Variant> {
    // colour-index → packed RGB `Long`, exactly as VBA's `QBColor` returns
    // (live-probed: 0..15 → these values).
    const PALETTE: [i32; 16] = [
        0, 8_388_608, 32_768, 8_421_376, 128, 8_388_736, 32_896, 12_632_256, 8_421_504, 16_711_680,
        65_280, 16_776_960, 255, 16_711_935, 65_535, 16_777_215,
    ];
    let index = as_i32(need(args, 0)?)?;
    if !(0..=15).contains(&index) {
        return Err(LibError::invalid_call("QBColor index must be 0 to 15"));
    }
    Ok(vi32(PALETTE[index as usize]))
}

// ── Interaction ──────────────────────────────────────────────────────────────

/// `Partition(number, start, stop, interval)` identifies the range containing
/// `number`. Integer arguments use VBA banker's rounding; a `Null` argument
/// propagates to a `Null` result.
pub fn partition(args: &[Variant]) -> LibResult<Variant> {
    let number_arg = need(args, 0)?;
    let start_arg = need(args, 1)?;
    let stop_arg = need(args, 2)?;
    let interval_arg = need(args, 3)?;

    if [number_arg, start_arg, stop_arg, interval_arg]
        .iter()
        .any(|value| value.vtype() == VarType::Null)
    {
        return Ok(Variant::null());
    }
    let number = as_i64(number_arg)?;
    let start = as_i64(start_arg)?;
    let stop = as_i64(stop_arg)?;
    let interval = as_i64(interval_arg)?;

    if start < 0 {
        return Err(LibError::invalid_call(
            "Partition start must be greater than or equal to 0",
        ));
    }
    if stop <= start {
        return Err(LibError::invalid_call(
            "Partition stop must be greater than start",
        ));
    }
    if interval < 1 {
        return Err(LibError::invalid_call(
            "Partition interval must be greater than or equal to 1",
        ));
    }

    let after_stop = stop
        .checked_add(1)
        .ok_or_else(|| LibError::overflow("Partition stop overflow"))?;
    let width = after_stop.to_string().len();

    let (lower, upper) = if interval == 1 {
        (Some(number), Some(number))
    } else if number < start {
        (None, Some(start - 1))
    } else if number > stop {
        (Some(after_stop), None)
    } else {
        let offset = number - start;
        let lower = start + (offset / interval) * interval;
        let upper = lower
            .checked_add(interval - 1)
            .unwrap_or(i64::MAX)
            .min(stop);
        (Some(lower), Some(upper))
    };

    Ok(vstr(format!(
        "{}:{}",
        partition_bound(lower, width),
        partition_bound(upper, width)
    )))
}

fn partition_bound(value: Option<i64>, width: usize) -> String {
    match value {
        Some(value) => format!("{value:>width$}"),
        None => " ".repeat(width),
    }
}

// ── Information ──────────────────────────────────────────────────────────────────

pub fn is_vtype(args: &[Variant], pred: impl Fn(VarType) -> bool) -> LibResult<Variant> {
    Ok(vbool(pred(need(args, 0)?.vtype())))
}

pub fn var_type(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    let code = match value.array_element_vartype() {
        Some(element_vartype) => 0x2000 | i32::from(element_vartype),
        None => value.vtype() as i32,
    };
    Ok(vi32(code))
}

pub fn type_name(args: &[Variant]) -> LibResult<Variant> {
    let value = need(args, 0)?;
    if let Some(element_vartype) = value.array_element_vartype() {
        return Ok(vstr(format!(
            "{}()",
            type_name_for_vartype(element_vartype)
        )));
    }
    let name = match value.vtype() {
        VarType::Empty => "Empty",
        VarType::Null => "Null",
        VarType::Integer => "Integer",
        VarType::Long => "Long",
        VarType::Single => "Single",
        VarType::Double => "Double",
        VarType::Currency => "Currency",
        VarType::Date => "Date",
        VarType::String => "String",
        VarType::Object
            if value
                .as_object_ref()
                .map(|object| object.raw())
                .unwrap_or(0)
                == 0 =>
        {
            "Nothing"
        }
        VarType::Object => "Object",
        VarType::Error => "Error",
        VarType::Boolean => "Boolean",
        VarType::Decimal => "Decimal",
        VarType::Byte => "Byte",
        VarType::LongLong => "LongLong",
        VarType::ArrayVariant => "Variant()",
        _ => "Variant",
    };
    Ok(vstr(name))
}

fn type_name_for_vartype(vartype: u16) -> &'static str {
    if vartype == VT_VARIANT_VALUE {
        return "Variant";
    }
    match VarType::from_u16(vartype) {
        Some(VarType::Empty) => "Empty",
        Some(VarType::Null) => "Null",
        Some(VarType::Integer) => "Integer",
        Some(VarType::Long) => "Long",
        Some(VarType::Single) => "Single",
        Some(VarType::Double) => "Double",
        Some(VarType::Currency) => "Currency",
        Some(VarType::Date) => "Date",
        Some(VarType::String) => "String",
        Some(VarType::Object) => "Object",
        Some(VarType::Error) => "Error",
        Some(VarType::Boolean) => "Boolean",
        Some(VarType::Decimal) => "Decimal",
        Some(VarType::Byte) => "Byte",
        Some(VarType::LongLong) => "LongLong",
        _ => "Variant",
    }
}

pub fn is_numeric(args: &[Variant]) -> LibResult<Variant> {
    let v = need(args, 0)?;
    // A *String* is numeric iff its text represents a VBA number — `coerce_to(..,
    // Double)` does not parse strings, so a numeric string like "42" needs the
    // dedicated grammar check. Every other subtype is judged by its `VarType`:
    // VBA's `IsNumeric` is True only for the genuine numeric subtypes and `Empty`,
    // and False for `Boolean`/`Date`/`Null`/`Object`/`Error`/arrays/`Nothing`
    // (`coerce_to(.., Double)` would wrongly accept Boolean and Date).
    let result = match v.vtype() {
        VarType::ByRef => {
            // Cloning a reference reads its target (VBA's `VariantCopyInd`).
            let target = v.clone();
            return is_numeric(std::slice::from_ref(&target));
        }
        VarType::String => is_numeric_string(&as_str(v)?),
        VarType::Empty
        | VarType::Boolean
        | VarType::Integer
        | VarType::Long
        | VarType::Single
        | VarType::Double
        | VarType::Currency
        | VarType::Byte
        | VarType::SignedByte
        | VarType::Decimal
        | VarType::LongLong
        | VarType::UnsignedInteger
        | VarType::UnsignedLong
        | VarType::UnsignedLongLong
        | VarType::UnsignedInt => true,
        VarType::Date
        | VarType::Null
        | VarType::Object
        | VarType::Error
        | VarType::ProcRef
        | VarType::Record
        | VarType::ArrayVariant => false,
    };
    Ok(vbool(result))
}

/// True when `s` represents a VBA number (the `IsNumeric` string grammar). VBA
/// recognises a decimal/sign/exponent literal (optionally surrounded by
/// whitespace) and `&H`/`&O` (or bare `&`) integer literals; it does NOT accept
/// thousands separators, currency symbols (locale-dependent — out of scope), or
/// Rust's `inf`/`nan` spellings.
///
/// This needs a FULL-parse check, not Val's leading-prefix parse: `IsNumeric("12a")`
/// must be False.
fn is_numeric_string(s: &str) -> bool {
    parse_vba_numeric_string(s).is_some()
}

fn parse_vba_numeric_string(s: &str) -> Option<f64> {
    calcvba_runtime::coerce::parse_vba_numeric_string(s)
}

fn parse_vba_prefixed_integer(t: &str) -> Option<i64> {
    let (value, rest) = parse_vba_prefixed_integer_token(t, false, false)?;
    if rest.trim().is_empty() {
        Some(value)
    } else {
        None
    }
}

// `Val` reads a leading token and ignores spaces/tabs/newlines inside it; the
// conversion functions and `IsNumeric` require the whole string to be numeric.
fn parse_vba_prefixed_integer_token(
    t: &str,
    skip_embedded_space: bool,
    allow_suffix: bool,
) -> Option<(i64, &str)> {
    let bytes = t.as_bytes();
    let mut pos = 0;
    let (sign, rest) = if let Some(rest) = t.strip_prefix('-') {
        pos = 1;
        (-1_i64, rest)
    } else if let Some(rest) = t.strip_prefix('+') {
        pos = 1;
        (1_i64, rest)
    } else {
        (1_i64, t)
    };
    // VBA hex (`&H…`) / octal (`&O…` or bare `&` + octal digits) integer literals.
    rest.strip_prefix('&')?;
    pos += 1;
    let (radix, digit_start) = match bytes.get(pos) {
        Some(b'H' | b'h') => (16, pos + 1),
        Some(b'O' | b'o') => (8, pos + 1),
        // Bare `&` followed by octal digits (VBA's terse octal form).
        _ => (8, pos),
    };
    pos = digit_start;

    let mut digits = String::new();
    while let Some(&byte) = bytes.get(pos) {
        if skip_embedded_space && is_val_ignored_space(byte) {
            pos += 1;
            continue;
        }
        let is_digit = match radix {
            16 => byte.is_ascii_hexdigit(),
            8 => (b'0'..=b'7').contains(&byte),
            _ => false,
        };
        if !is_digit {
            break;
        }
        digits.push(char::from(byte));
        pos += 1;
    }
    if digits.is_empty() {
        return None;
    }

    if skip_embedded_space {
        while let Some(&byte) = bytes.get(pos) {
            if !is_val_ignored_space(byte) {
                break;
            }
            pos += 1;
        }
    }
    // `Val` applies the width-based two's-complement sign rule exactly as to a
    // literal and honours an optional `%`/`&`/`^` type character. Conversion
    // functions reject the suffix form (`CLng("&HFFFF&")` is type mismatch).
    let suffix = match bytes.get(pos) {
        Some(&b @ (b'%' | b'&' | b'^')) => Some(b),
        _ => None,
    };
    if suffix.is_some() {
        if !allow_suffix {
            return None;
        }
        pos += 1;
    }
    let magnitude = u64::from_str_radix(&digits, radix).ok()?;
    let width = calcvba_runtime::vba_radix_width(magnitude, suffix)?;
    let value = calcvba_runtime::vba_radix_signed_value(magnitude, width).checked_mul(sign)?;
    Some((value, &t[pos..]))
}

fn is_val_ignored_space(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r')
}

pub fn is_date(args: &[Variant]) -> LibResult<Variant> {
    let v = need(args, 0)?;
    // VBA `IsDate` is True for a Date value or a STRING recognizable as a valid date/time
    // (using the same lenient parser as `CDate`); a raw numeric, Empty, Null, Boolean, or
    // object is NOT a date to `IsDate` (even though `CDate` would convert a number).
    let ok = match v.vtype() {
        VarType::Date => true,
        VarType::String => as_str(v)
            .ok()
            .is_some_and(|s| date_time_text_serial(&s).is_ok()),
        _ => false,
    };
    Ok(vbool(ok))
}

fn truthy(v: &Variant) -> LibResult<bool> {
    match v.vtype() {
        VarType::Null | VarType::Empty => Ok(false),
        // "True" / "False" read as the Booleans; any other string must be a number.
        VarType::String => {
            let text = as_str(v)?;
            match text.trim().to_ascii_lowercase().as_str() {
                "true" => Ok(true),
                "false" => Ok(false),
                _ => Ok(as_f64(v)? != 0.0),
            }
        }
        _ => Ok(as_f64(v)? != 0.0),
    }
}

/// `IIf(cond, truepart, falsepart)` — EAGER: both parts are already evaluated
/// by the caller; this just selects one.
pub fn iif(args: &[Variant]) -> LibResult<Variant> {
    if truthy(need(args, 0)?)? {
        Ok(need(args, 1)?.clone())
    } else {
        Ok(need(args, 2)?.clone())
    }
}

pub fn choose(args: &[Variant]) -> LibResult<Variant> {
    // The index is truncated toward zero (`Choose(1.5, ...)` picks the first choice).
    let index = as_f64(need(args, 0)?)?.trunc();
    if !(1.0..=f64::from(i32::MAX)).contains(&index) {
        return Ok(Variant::null());
    }
    Ok(args
        .get(index as usize)
        .cloned()
        .unwrap_or_else(Variant::null))
}

pub fn switch(args: &[Variant]) -> LibResult<Variant> {
    if !args.len().is_multiple_of(2) {
        return Err(LibError::invalid_call(
            "Switch needs expression/value pairs",
        ));
    }
    let mut i = 0;
    while i + 1 < args.len() {
        if truthy(&args[i])? {
            return Ok(args[i + 1].clone());
        }
        i += 2;
    }
    Ok(Variant::null())
}

/// Mac four-character codes are big-endian ANSI bytes, padded with spaces.
pub fn mac_id(args: &[Variant]) -> LibResult<Variant> {
    let value = as_str(need(args, 0)?)?;
    let mut bytes = [b' '; 4];
    for (slot, byte) in bytes
        .iter_mut()
        .zip(calcvba_runtime::ansi::ansi_encode(&value))
    {
        *slot = byte;
    }
    Ok(vi32(i32::from_be_bytes(bytes)))
}

pub fn mac_script(args: &[Variant]) -> LibResult<Variant> {
    as_str(need(args, 0)?)?;
    Err(LibError::invalid_call("MacScript is unavailable"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn date_parts_reject_out_of_range_serials_before_calendar_conversion() {
        for serial in [
            -657_435.0,
            2_958_466.0,
            -1e30,
            1e30,
            f64::MIN,
            f64::MAX,
            f64::NAN,
            f64::INFINITY,
        ] {
            for part in [
                DatePart::Year,
                DatePart::Month,
                DatePart::Day,
                DatePart::Weekday,
                DatePart::Hour,
                DatePart::Minute,
                DatePart::Second,
            ] {
                assert_eq!(
                    date_part(&[Variant::from_f64(serial)], part)
                        .unwrap_err()
                        .code,
                    13
                );
            }
            // Format uses the same calendar helpers through a separate path.
            let _ = crate::format::apply_with(&Variant::from_f64(serial), "ww y dddd", 1, 1);
        }
    }

    fn vs(s: &str) -> Variant {
        Variant::from_string(s.to_string())
    }
    fn like_(s: &str, p: &str) -> bool {
        like(&[vs(s), vs(p)]).unwrap().as_bool().unwrap()
    }
    fn instr_(args: &[Variant], rev: bool) -> i32 {
        instr(args, rev).unwrap().as_i32().unwrap()
    }
    fn split_(args: &[Variant]) -> Vec<String> {
        split(args)
            .unwrap()
            .as_safearray()
            .unwrap()
            .variant_elements()
            .unwrap_or_default()
            .iter()
            .map(|v| v.as_bstr().map(|b| b.as_str()).unwrap_or_default())
            .collect()
    }
    fn val_(s: &str) -> f64 {
        val(&[vs(s)]).unwrap().as_f64().unwrap()
    }
    fn partition_(args: &[Variant]) -> String {
        partition(args).unwrap().as_bstr().unwrap().as_str()
    }
    fn bstr(result: LibResult<Variant>) -> String {
        result.unwrap().as_bstr().unwrap().as_str()
    }
    fn bstr_bytes(result: LibResult<Variant>) -> Vec<u8> {
        result.unwrap().as_bstr().unwrap().payload_bytes().to_vec()
    }
    fn bstr_units(result: LibResult<Variant>) -> Vec<u16> {
        result.unwrap().as_bstr().unwrap().payload_units().to_vec()
    }

    #[test]
    fn val_parses_vba_radix_prefixes() {
        assert_eq!(val_("&HFFFFFFFF"), -1.0);
        assert_eq!(val_("&HFFFF&"), 65_535.0);
        assert_eq!(val_("&O37777777777"), -1.0);
        assert_eq!(val_("+&H7F"), 0.0);
        assert_eq!(val_("-&O10"), 0.0);
        assert_eq!(val_("&H F F"), 255.0);
        assert_eq!(val_("&H10 trailing"), 16.0);
    }

    #[test]
    fn val_parses_longest_complete_decimal_prefix() {
        assert_eq!(val_("123abc"), 123.0);
        assert_eq!(val_("12-3"), 12.0);
        assert_eq!(val_("1.2.3"), 1.2);
        assert_eq!(val_("1e2"), 100.0);
        assert_eq!(val_("1e"), 1.0);
        assert_eq!(val_("1e+"), 1.0);
        assert_eq!(val_("1e-2"), 0.01);
        assert_eq!(val_("1D2"), 100.0);
        assert_eq!(val_("1 2"), 12.0);
        assert_eq!(val_("- .5"), -0.5);
        assert_eq!(val_("$1"), 0.0);
    }

    #[test]
    fn sgn_rejects_nan_double() {
        assert_eq!(sgn(&[Variant::from_f64(f64::NAN)]).unwrap_err().code, 5);
    }

    #[test]
    fn mid_rejects_start_less_than_one() {
        assert_eq!(
            mid(&[vs("abcdef"), Variant::from_i32(0), Variant::from_i32(2)])
                .unwrap_err()
                .code,
            5
        );
        assert_eq!(
            mid_stmt(&[
                vs("abcdef"),
                Variant::from_i32(0),
                Variant::from_i32(2),
                vs("ZZ"),
            ])
            .unwrap_err()
            .code,
            5
        );
        assert_eq!(
            mid(&[vs("abcdef"), Variant::from_i32(1), Variant::from_i32(2)])
                .unwrap()
                .as_bstr()
                .unwrap()
                .as_str(),
            "ab"
        );
    }

    #[test]
    fn leftb_rightb_slice_raw_bstr_bytes() {
        let text = vs("ABC");

        assert_eq!(
            bstr_bytes(left_b(&[text.clone(), Variant::from_i32(0)])),
            Vec::<u8>::new()
        );
        assert_eq!(
            bstr_bytes(left_b(&[text.clone(), Variant::from_i32(1)])),
            vec![0x41]
        );
        assert_eq!(
            bstr_units(left_b(&[text.clone(), Variant::from_i32(3)])),
            vec![0x0041]
        );
        assert_eq!(
            bstr_bytes(left_b(&[text.clone(), Variant::from_i32(99)])),
            vec![0x41, 0x00, 0x42, 0x00, 0x43, 0x00]
        );

        assert_eq!(
            bstr_bytes(right_b(&[text.clone(), Variant::from_i32(1)])),
            vec![0x00]
        );
        assert_eq!(
            bstr_units(right_b(&[text.clone(), Variant::from_i32(3)])),
            vec![0x4300]
        );
        assert_eq!(
            bstr_bytes(right_b(&[text, Variant::from_i32(99)])),
            vec![0x41, 0x00, 0x42, 0x00, 0x43, 0x00]
        );
    }

    #[test]
    fn leftb_rightb_odd_byte_results_keep_lenb_payload() {
        let left = left_b(&[vs("ABC"), Variant::from_i32(1)]).unwrap();
        assert_eq!(
            len_b(std::slice::from_ref(&left)).unwrap().as_i32(),
            Some(1)
        );
        assert_eq!(left.string_units(), Some(vec![]));

        let right = right_b(&[vs("ABC"), Variant::from_i32(3)]).unwrap();
        assert_eq!(
            len_b(std::slice::from_ref(&right)).unwrap().as_i32(),
            Some(3)
        );
        assert_eq!(right.string_units(), Some(vec![0x4300]));
    }

    #[test]
    fn leftb_rightb_reject_negative_count() {
        assert_eq!(
            left_b(&[vs("ABC"), Variant::from_i32(-1)])
                .unwrap_err()
                .code,
            5
        );
        assert_eq!(
            right_b(&[vs("ABC"), Variant::from_i32(-1)])
                .unwrap_err()
                .code,
            5
        );
    }

    #[test]
    fn format_number_family_honours_defaults_and_tristates() {
        assert_eq!(
            bstr(format_number(&[Variant::from_f64(1234.567)])),
            "1,234.57"
        );
        assert_eq!(
            bstr(format_number(&[
                Variant::from_f64(1234.567),
                Variant::from_i32(0),
                Variant::from_i32(-2),
                Variant::from_i32(0),
                Variant::from_i32(0),
            ])),
            "1235"
        );
        assert_eq!(
            bstr(format_number(&[
                Variant::from_f64(0.5),
                Variant::from_i32(2),
                Variant::from_i32(0),
                Variant::from_i32(0),
                Variant::from_i32(0),
            ])),
            ".50"
        );
        assert_eq!(
            bstr(format_number(&[
                Variant::from_f64(-1234.5),
                Variant::from_i32(2),
                Variant::from_i32(-1),
                Variant::from_i32(-1),
                Variant::from_i32(-1),
            ])),
            "(1,234.50)"
        );
        assert_eq!(
            bstr(format_currency(&[
                Variant::from_f64(-1234.5),
                Variant::from_i32(2),
                Variant::from_i32(-1),
                Variant::from_i32(-1),
                Variant::from_i32(-1),
            ])),
            "($1,234.50)"
        );
        assert_eq!(
            bstr(format_percent(&[
                Variant::from_f64(0.1234),
                Variant::from_i32(1),
                Variant::from_i32(-1),
                Variant::from_i32(0),
                Variant::from_i32(0),
            ])),
            "12.3%"
        );
        assert_eq!(
            bstr(format_percent(&[
                Variant::from_f64(0.005),
                Variant::from_i32(2),
                Variant::from_i32(0),
                Variant::from_i32(0),
                Variant::from_i32(0),
            ])),
            ".50%"
        );
    }

    #[test]
    fn format_date_time_honours_named_formats() {
        let date = Variant::from_date_f64(ymd_to_serial(2020, 1, 15) + 13.5 / 24.0);
        assert_eq!(
            bstr(format_date_time(&[date.clone(), Variant::from_i32(0)])),
            "1/15/2020 1:30:00 PM"
        );
        assert_eq!(
            bstr(format_date_time(&[date.clone(), Variant::from_i32(1)])),
            "Wednesday, January 15, 2020"
        );
        assert_eq!(
            bstr(format_date_time(&[date.clone(), Variant::from_i32(2)])),
            "1/15/2020"
        );
        assert_eq!(
            bstr(format_date_time(&[date.clone(), Variant::from_i32(3)])),
            "1:30:00 PM"
        );
        assert_eq!(
            bstr(format_date_time(&[date, Variant::from_i32(4)])),
            "13:30"
        );
    }

    #[test]
    fn format_family_rejects_invalid_option_values() {
        assert_eq!(
            format_number(&[Variant::from_f64(1.2), Variant::from_i32(-2)])
                .unwrap_err()
                .code,
            5
        );
        assert_eq!(
            format_number(&[
                Variant::from_f64(1.2),
                Variant::from_i32(2),
                Variant::from_i32(1),
            ])
            .unwrap_err()
            .code,
            5
        );
        assert_eq!(
            format_date_time(&[Variant::from_f64(1.2), Variant::from_i32(5)])
                .unwrap_err()
                .code,
            5
        );
    }

    #[test]
    fn error_text_returns_default_messages_and_fallbacks() {
        assert_eq!(bstr(error_text(&[])), "");
        assert_eq!(bstr(error_text(&[Variant::from_i32(0)])), "");
        assert_eq!(
            bstr(error_text(&[Variant::from_i32(11)])),
            "Division by zero"
        );
        assert_eq!(
            bstr(error_text(&[Variant::from_i32(12345)])),
            "Application-defined or object-defined error"
        );
        // A negative number has no message of its own; one past 65535 overflows.
        assert_eq!(
            bstr(error_text(&[Variant::from_i32(-1)])),
            "Application-defined or object-defined error"
        );
        assert_eq!(
            error_text(&[Variant::from_i32(70_000)]).unwrap_err().code,
            6
        );
    }

    #[test]
    fn week_of_year_follows_first_day_and_first_week() {
        // 1/1/2005 is a Saturday.
        assert_eq!(week_of_year(2005, 1, 1, 1, 1), 1);
        assert_eq!(week_of_year(2005, 1, 1, 1, 3), 52);
        assert_eq!(week_of_year(2005, 1, 1, 2, 2), 53);
        // 12/29/2003 (Monday) belongs to week 1 of 2004 under Monday / four-day weeks.
        assert_eq!(week_of_year(2003, 12, 29, 2, 2), 1);
    }

    #[test]
    fn like_reports_bad_patterns_and_matches_empty_brackets() {
        assert_eq!(like(&[vs("["), vs("[")]).unwrap_err().code, 93);
        assert_eq!(like(&[vs("b"), vs("[z-a]")]).unwrap_err().code, 93);
        assert!(like_("ab", "a[]b"));
        assert!(like_("ab", "a[b-c]"));
    }

    #[test]
    fn mid_statement_past_the_end_is_an_error() {
        let past = mid_stmt(&[vs("abc"), Variant::from_i32(5), vs("x")]);
        assert_eq!(past.unwrap_err().code, 5);
        assert_eq!(
            bstr(mid_stmt(&[vs("abc"), Variant::from_i32(2), vs("xyz")])),
            "axy"
        );
        // MidB counts bytes: byte 3 of "abc" (UTF-16LE) is the low byte of "b".
        let out = mid_b_stmt(&[vs("abc"), Variant::from_i32(3), vs("z")]).unwrap();
        assert_eq!(out.as_bstr().unwrap().payload_bytes(), b"a\0z\0c\0");
    }

    #[test]
    fn byte_functions_work_on_utf16_bytes() {
        assert_eq!(asc_b(&[vs("A")]).unwrap().as_u8(), Some(65));
        assert_eq!(
            chr_b(&[Variant::from_i32(65)])
                .unwrap()
                .as_bstr()
                .unwrap()
                .payload_bytes(),
            b"A"
        );
        assert_eq!(instr_b(&[vs("abc"), vs("b")]).unwrap().as_i32(), Some(3));
        assert_eq!(
            mid_b(&[vs("abc"), Variant::from_i32(3), Variant::from_i32(2)])
                .unwrap()
                .as_bstr()
                .unwrap()
                .payload_bytes(),
            b"b\0"
        );
    }

    #[test]
    fn financial_guards_and_rounding() {
        assert_eq!(
            pmt(&[vf64(0.1), vf64(0.0), vf64(1000.0)]).unwrap_err().code,
            5
        );
        assert_eq!(sln(&[vf64(1.0), vf64(0.0), vf64(0.0)]).unwrap_err().code, 5);
        let due = fv(&[vf64(0.1), vf64(10.0), vf64(-100.0), vf64(0.0), vf64(2.0)]).unwrap();
        assert!((due.as_f64().unwrap() - 1753.1167061).abs() < 1e-6);
        // Round keeps its argument type.
        let cur = round(&[Variant::from_currency_scaled_i64(25_000)]).unwrap();
        assert_eq!(cur.as_currency_scaled_i64(), Some(20_000));
        let single = round(&[Variant::from_f32(1.25), Variant::from_i32(1)]).unwrap();
        assert_eq!(single.as_f32(), Some(1.2));
    }

    #[test]
    fn round_excel_types_digit_limit_and_signed_zero() {
        let date = round(&[Variant::from_date_f64(2.5)]).unwrap();
        assert_eq!(date.vtype(), VarType::Date);
        assert_eq!(date.as_date_f64(), Some(2.0));
        assert_eq!(
            round(&[Variant::from_bool(true)]).unwrap(),
            Variant::from_i16(-1)
        );
        assert_eq!(round(&[Variant::empty()]).unwrap(), Variant::from_i16(0));
        for value in [
            vf64(2.5),
            Variant::from_f32(2.5),
            Variant::from_date_f64(2.5),
        ] {
            assert_eq!(round(&[value, vi32(23)]).unwrap_err().code, 5);
        }
        assert_eq!(
            round(&[Variant::from_u8(1), vi32(23)]).unwrap(),
            Variant::from_u8(1)
        );
        for (value, digits, bits) in [(-0.05, 1, 1u64 << 63), (-0.49, 0, 0), (-1e-300, 5, 0)] {
            assert_eq!(
                round(&[vf64(value), vi32(digits)])
                    .unwrap()
                    .as_f64()
                    .unwrap()
                    .to_bits(),
                bits
            );
        }
    }

    #[test]
    fn split_honours_limit_and_compare() {
        // No limit → all substrings.
        assert_eq!(split_(&[vs("a,b,c"), vs(",")]), ["a", "b", "c"]);
        // limit=2 → the last element holds the unsplit remainder.
        assert_eq!(
            split_(&[vs("a,b,c"), vs(","), Variant::from_i32(2)]),
            ["a", "b,c"]
        );
        // limit=1 → the whole string.
        assert_eq!(
            split_(&[vs("a,b,c"), vs(","), Variant::from_i32(1)]),
            ["a,b,c"]
        );
        // limit=0 → an empty array.
        assert!(split_(&[vs("a,b,c"), vs(","), Variant::from_i32(0)]).is_empty());
        // Text compare (mode 1): split on "x" case-insensitively, original case preserved.
        assert_eq!(
            split_(&[
                vs("aXbxC"),
                vs("x"),
                Variant::from_i32(-1),
                Variant::from_i32(1)
            ]),
            ["a", "b", "C"]
        );
        // Binary compare (default): only the lowercase "x" splits.
        assert_eq!(split_(&[vs("aXbxC"), vs("x")]), ["aXb", "C"]);
    }

    #[test]
    fn like_charlist() {
        assert!(like_("abc", "[a-c]bc"));
        assert!(like_("Abc", "[A-Z]*"));
        assert!(like_("x", "[!a-c]"));
        assert!(!like_("b", "[!a-c]"));
        assert!(like_("a]b", "a[]]b")); // literal ']' as first list member
        assert!(like_("3", "#"));
        assert!(!like_("z", "[0-9]"));
    }

    #[test]
    fn instr_start_and_compare() {
        assert_eq!(instr_(&[vs("hello world"), vs("o")], false), 5);
        // Leading numeric `start` = 6 → first 'o' at or after position 6.
        assert_eq!(
            instr_(&[Variant::from_i32(6), vs("hello world"), vs("o")], false),
            8
        );
        // InStrRev finds the last 'o'.
        assert_eq!(instr_(&[vs("hello world"), vs("o")], true), 8);
        // Text compare: per VBA `compare` requires `start`, so the 4-arg form
        // `InStr(1, "ABC", "b", 1)` is case-insensitive → 2.
        assert_eq!(
            instr_(
                &[
                    Variant::from_i32(1),
                    vs("ABC"),
                    vs("b"),
                    Variant::from_i32(1)
                ],
                false
            ),
            2
        );
        // Binary compare (default): no match for differing case.
        assert_eq!(instr_(&[vs("ABC"), vs("b")], false), 0);
    }

    #[test]
    fn instr_start_is_arity_not_type() {
        // 2 args is always `(string1, string2)` — even when `string1` is a numeric string,
        // it is NOT a leading `start`. `InStr("12", "2")` finds '2' at position 2.
        assert_eq!(instr_(&[vs("12"), vs("2")], false), 2);
    }

    #[test]
    fn instr_rev_honours_start_and_compare() {
        // Last 'a' overall is at 7; with start=4 the search is limited to "abca" → 4.
        assert_eq!(instr_(&[vs("abcabca"), vs("a")], true), 7);
        assert_eq!(
            instr_(&[vs("abcabca"), vs("a"), Variant::from_i32(4)], true),
            4
        );
        // compare is arg 4 (after start): text compare finds 'B' case-insensitively.
        assert_eq!(
            instr_(
                &[
                    vs("aBcaBc"),
                    vs("b"),
                    Variant::from_i32(-1),
                    Variant::from_i32(1)
                ],
                true
            ),
            5
        );
    }

    fn isnum(v: Variant) -> bool {
        is_numeric(&[v]).unwrap().as_bool().unwrap()
    }

    #[test]
    fn is_numeric_of_strings_and_values() {
        // Numeric strings (the bug fix): decimal / sign / dot / exponent / whitespace.
        assert!(isnum(vs("42")));
        assert!(isnum(vs("3.14")));
        assert!(isnum(vs("-1.5e3")));
        assert!(isnum(vs("  7  ")));
        assert!(isnum(vs("+1.5")));
        assert!(isnum(vs(".5")));
        // VBA hex/octal literals.
        assert!(isnum(vs("&HFF")));
        assert!(!isnum(vs("&H1F&"))); // conversion grammar rejects type suffixes
        assert!(isnum(vs("&O17")));
        assert!(isnum(vs("&777"))); // bare-& octal
        // Non-numeric strings.
        assert!(!isnum(vs("12a")));
        assert!(!isnum(vs("")));
        assert!(!isnum(vs("abc")));
        assert!(!isnum(vs("nan"))); // Rust spelling VBA rejects
        assert!(!isnum(vs("inf")));
        assert!(!isnum(vs("&HZZ"))); // bad hex digits
        assert!(!isnum(vs("&H"))); // empty hex body
        // Genuine numeric subtypes and Empty are numeric (the VarType check).
        assert!(isnum(Variant::from_i32(42)));
        assert!(isnum(Variant::from_f64(3.5)));
        assert!(isnum(Variant::empty()));
        // Boolean IS numeric (Excel: `IsNumeric(True)`); Date and Null are not.
        assert!(isnum(Variant::from_bool(true)));
        assert!(isnum(Variant::from_bool(false)));
        assert!(!isnum(Variant::from_date_f64(45000.0))); // a Date-typed Variant
        assert!(!isnum(Variant::null()));
    }

    #[test]
    fn cverr_accepts_unsigned_error_code_range_only() {
        assert_eq!(
            cverr(&[Variant::from_i32(0)]).unwrap().as_error_code(),
            Some(0)
        );
        assert_eq!(
            cverr(&[Variant::from_i32(0xFFFF)]).unwrap().as_error_code(),
            Some(0xFFFF)
        );
        let low = cverr(&[Variant::from_i32(-1)]).unwrap_err();
        assert_eq!(low.code, 5);
        assert_eq!(low.message, default_error_message(5));
        let high = cverr(&[Variant::from_i32(0x1_0000)]).unwrap_err();
        assert_eq!(high.code, 5);
        assert_eq!(high.message, default_error_message(5));
    }

    #[test]
    fn partition_formats_documented_ranges() {
        assert_eq!(
            partition_(&[
                Variant::from_i32(-1),
                Variant::from_i32(0),
                Variant::from_i32(99),
                Variant::from_i32(5),
            ]),
            "   : -1"
        );
        assert_eq!(
            partition_(&[
                Variant::from_i32(0),
                Variant::from_i32(0),
                Variant::from_i32(99),
                Variant::from_i32(5),
            ]),
            "  0:  4"
        );
        assert_eq!(
            partition_(&[
                Variant::from_i32(99),
                Variant::from_i32(0),
                Variant::from_i32(99),
                Variant::from_i32(5),
            ]),
            " 95: 99"
        );
        assert_eq!(
            partition_(&[
                Variant::from_i32(100),
                Variant::from_i32(0),
                Variant::from_i32(99),
                Variant::from_i32(5),
            ]),
            "100:   "
        );
        assert_eq!(
            partition_(&[
                Variant::from_i32(1000),
                Variant::from_i32(100),
                Variant::from_i32(1010),
                Variant::from_i32(20),
            ]),
            "1000:1010"
        );
        assert_eq!(
            partition_(&[
                Variant::from_i32(100),
                Variant::from_i32(0),
                Variant::from_i32(1000),
                Variant::from_i32(1),
            ]),
            " 100: 100"
        );
    }

    #[test]
    fn partition_rounds_nulls_and_rejects_invalid_ranges() {
        assert_eq!(
            partition(&[
                Variant::from_f64(2.5),
                Variant::from_i32(0),
                Variant::from_i32(10),
                Variant::from_i32(2),
            ])
            .unwrap()
            .as_bstr()
            .unwrap()
            .as_str(),
            " 2: 3"
        );
        assert_eq!(
            partition(&[
                Variant::from_f64(3.5),
                Variant::from_i32(0),
                Variant::from_i32(10),
                Variant::from_i32(2),
            ])
            .unwrap()
            .as_bstr()
            .unwrap()
            .as_str(),
            " 4: 5"
        );
        assert_eq!(
            partition(&[
                Variant::null(),
                Variant::from_i32(0),
                Variant::from_i32(10),
                Variant::from_i32(2),
            ])
            .unwrap()
            .vtype(),
            VarType::Null
        );
        assert_eq!(
            partition(&[
                Variant::from_i32(1),
                Variant::from_i32(-1),
                Variant::from_i32(10),
                Variant::from_i32(2),
            ])
            .unwrap_err()
            .code,
            5
        );
        assert_eq!(
            partition(&[
                Variant::from_i32(1),
                Variant::from_i32(0),
                Variant::from_i32(0),
                Variant::from_i32(2),
            ])
            .unwrap_err()
            .code,
            5
        );
        assert_eq!(
            partition(&[
                Variant::from_i32(1),
                Variant::from_i32(0),
                Variant::from_i32(10),
                Variant::from_i32(0),
            ])
            .unwrap_err()
            .code,
            5
        );
        assert_eq!(
            partition(&[Variant::null()]).unwrap_err().code,
            5,
            "missing arguments are not masked by an earlier Null"
        );
    }

    fn chr_(id_args: &[Variant], wide: bool) -> String {
        let r = if wide { chr_w(id_args) } else { chr(id_args) };
        as_str(&r.unwrap()).unwrap()
    }
    fn asc_(s: &str, wide: bool) -> i32 {
        let r = if wide { asc_w(&[vs(s)]) } else { asc(&[vs(s)]) };
        let value = r.unwrap();
        value
            .as_i32()
            .or_else(|| value.as_i16().map(i32::from))
            .unwrap()
    }

    #[test]
    fn chr_asc_ansi_vs_wide() {
        // ASCII (the common case) is identical for both variants.
        assert_eq!(chr_(&[Variant::from_i32(65)], false), "A");
        assert_eq!(chr_(&[Variant::from_i32(65)], true), "A");
        assert_eq!(asc_("A", false), 65);
        assert_eq!(asc_("A", true), 65);

        // 128..255: Chr decodes through the system ANSI codec — on the dev/CI host
        // (CP-1252 Windows ACP, or the off-Windows CP-1252 fallback) 150 → en-dash —
        // while ChrW is always the raw code point.
        assert_eq!(chr_(&[Variant::from_i32(150)], false), "\u{2013}"); // –
        assert_eq!(chr_(&[Variant::from_i32(150)], true), "\u{0096}"); // U+0096
        assert_ne!(
            chr_(&[Variant::from_i32(150)], false),
            chr_(&[Variant::from_i32(150)], true)
        );
        // Asc is the system-ANSI byte (150 on a CP-1252 host); AscW is the code point.
        assert_eq!(asc_("\u{2013}", false), 150); // en-dash → ANSI byte 150
        assert_eq!(asc_("\u{2013}", true), 8211); // en-dash → U+2013

        // The wide range (>255) agrees: Chr acts wide above 255.
        assert_eq!(chr_(&[Variant::from_i32(8364)], false), "\u{20AC}"); // €
        assert_eq!(chr_(&[Variant::from_i32(8364)], true), "\u{20AC}");
        // € round-trips: AscW is the code point, Asc is the CP-1252 byte 128.
        assert_eq!(asc_("\u{20AC}", true), 8364);
        assert_eq!(asc_("\u{20AC}", false), 128);

        // AscW returns an Integer (i16): code points > 32767 come back negative.
        assert_eq!(asc_("\u{8000}", true), (0x8000_u16 as i16) as i32);
        // Asc of an unmappable char is 63 ("?").
        assert_eq!(asc_("\u{4E2D}", false), 63);
    }

    #[test]
    fn str_comp_distinguishes_lone_surrogates() {
        // VBA strings are binary UTF-16 code-unit sequences; StrComp must NOT fold distinct
        // lone surrogate halves to U+FFFD (the lossy `as_str` path did, making them compare
        // equal). Binary mode (no third arg).
        let high = Variant::from_utf16_units(&[0xD800]);
        let low = Variant::from_utf16_units(&[0xDC00]);
        let same = Variant::from_utf16_units(&[0xD800]);
        // Distinct halves are unequal; 0xD800 < 0xDC00 by code unit, equal halves compare 0.
        assert_eq!(
            str_comp(&[high.clone(), low.clone()])
                .unwrap()
                .as_i32()
                .unwrap(),
            -1
        );
        assert_eq!(str_comp(&[low, high.clone()]).unwrap().as_i32().unwrap(), 1);
        assert_eq!(str_comp(&[high, same]).unwrap().as_i32().unwrap(), 0);
        // Text mode (compare = 1) still ASCII case-folds without disturbing surrogates.
        let upper = Variant::from_string("A".to_string());
        let lower = Variant::from_string("a".to_string());
        assert_eq!(
            str_comp(&[upper, lower, Variant::from_i32(1)])
                .unwrap()
                .as_i32()
                .unwrap(),
            0
        );
    }

    #[test]
    fn rnd_deterministic_bounded_and_repeatable() {
        let mut a = LibContext::default();
        let mut b = LibContext::default();
        // `Rnd` returns a Single; use the coercing helper to read it.
        let x = as_f64(&rnd(&[], &mut a).unwrap()).unwrap();
        let y = as_f64(&rnd(&[], &mut b).unwrap()).unwrap();
        assert_eq!(x, y, "default seed is deterministic");
        assert_eq!(a.rng_state, 0xB4_9E_C3);
        assert_eq!(x, 0.7055475115776062);
        assert!((0.0..1.0).contains(&x), "result in [0,1): {x}");
        let second = as_f64(&rnd(&[], &mut a).unwrap()).unwrap();
        assert_eq!(a.rng_state, 0x88_8E_7A);
        assert_eq!(second, 0.5334240198135376);
        // Rnd(0) repeats the most recent value without advancing.
        let again = as_f64(&rnd(&[Variant::from_i32(0)], &mut a).unwrap()).unwrap();
        assert_eq!(second, again);
    }

    #[test]
    fn rnd_negative_and_randomize_match_excel_seed_states() {
        let mut ctx = LibContext::default();
        let negative = as_f64(&rnd(&[Variant::from_i32(-1)], &mut ctx).unwrap()).unwrap();
        assert_eq!(ctx.rng_state, 0x39_58_86);
        assert_eq!(negative, 0.2240070104598999);
        let next = as_f64(&rnd(&[], &mut ctx).unwrap()).unwrap();
        assert_eq!(ctx.rng_state, 0x09_2D_31);
        assert_eq!(next, 0.035845816135406494);

        let mut repeatable = LibContext::default();
        let _ = rnd(&[Variant::from_i32(-1)], &mut repeatable).unwrap();
        randomize(&[Variant::from_i32(1)], &mut repeatable).unwrap();
        assert_eq!(repeatable.rng_state, 0x3F_F0_86);
        let first = as_f64(&rnd(&[], &mut repeatable).unwrap()).unwrap();
        assert_eq!(repeatable.rng_state, 0x55_65_31);
        assert_eq!(first, 0.3335753083229065);
        let second = as_f64(&rnd(&[], &mut repeatable).unwrap()).unwrap();
        assert_eq!(repeatable.rng_state, 0x11_73_30);
        assert_eq!(second, 0.06816387176513672);
    }

    #[test]
    fn rnd_matches_excel_full_period() {
        // P_Rnd.34: the next value after the 2^24-state cycle repeats the first.
        let mut ctx = LibContext::default();
        let first = rnd(&[Variant::from_i32(-1)], &mut ctx).unwrap();
        for _ in 1..16_777_216 {
            rng_step(&mut ctx);
        }
        let repeated = rnd(&[], &mut ctx).unwrap();
        assert_eq!(first.as_f32().unwrap().to_bits(), 0x3E65_6218);
        assert_eq!(first.as_f32(), repeated.as_f32());
    }

    #[test]
    fn financial_cashflow_functions_reject_scalar_values() {
        assert_eq!(
            npv(&[Variant::from_i32(0), Variant::from_i32(10)])
                .unwrap_err()
                .code,
            13
        );
        assert_eq!(irr(&[Variant::from_i32(50)]).unwrap_err().code, 13);
        assert_eq!(
            mirr(&[
                Variant::from_i32(70),
                Variant::from_i32(10),
                Variant::from_i32(12),
            ])
            .unwrap_err()
            .code,
            13
        );
    }

    #[test]
    fn nper_excel_logarithm_order_and_domain() {
        for (rate, payment, present, bits) in [
            (0.05 / 12.0, -200.0, 10000.0, 0x404C_1796_D6F4_A714),
            (1e-12, -100.0, 1000.0, 0x4024_01D1_9EB1_60F0),
        ] {
            assert_eq!(
                nper(&[vf64(rate), vf64(payment), vf64(present)])
                    .unwrap()
                    .as_f64()
                    .unwrap()
                    .to_bits(),
                bits
            );
        }
        assert_eq!(
            nper(&[vf64(-1.0), vf64(-100.0), vf64(1000.0)])
                .unwrap_err()
                .code,
            5
        );
        assert_eq!(
            nper(&[vf64(0.0), vf64(0.0), vf64(1000.0)])
                .unwrap_err()
                .code,
            5
        );
    }

    #[test]
    fn irr_excel_convergence_and_error_cases() {
        let flows = |values: &[f64]| {
            Variant::from_safearray(
                SafeArray::from_typed_variants(
                    calcvba_runtime::safe_array::VT_R8_VALUE,
                    values.iter().copied().map(vf64).collect(),
                )
                .unwrap(),
            )
        };
        let series = flows(&[-100.0, 30.0, 40.0, 50.0, 60.0]);
        for (guess, bits) in [(0.5, 0x3FCF_DB68_EA9C_0B5B), (-0.5, 0x3FCF_DB68_EA9C_0995)] {
            assert_eq!(
                irr(&[series.clone(), vf64(guess)])
                    .unwrap()
                    .as_f64()
                    .unwrap()
                    .to_bits(),
                bits
            );
        }
        assert_eq!(irr(&[series.clone(), vf64(100.0)]).unwrap_err().code, 5);
        assert_eq!(irr(&[series.clone(), vf64(-1.0)]).unwrap_err().code, 5);
        assert_eq!(irr(&[series, Variant::null()]).unwrap_err().code, 94);
        for values in [&[-1000.0, 1.0][..], &[-100.0, 0.0001], &[0.0, 0.0]] {
            assert_eq!(irr(&[flows(values)]).unwrap_err().code, 5);
        }
        assert_eq!(
            irr(&[Variant::unallocated_array(
                calcvba_runtime::safe_array::VT_R8_VALUE
            )])
            .unwrap_err()
            .code,
            5
        );
    }

    #[test]
    fn mirr_excel_ordered_cashflows_and_errors() {
        let flows = |values: &[f64]| {
            Variant::from_safearray(SafeArray::from_variants(
                values.iter().copied().map(vf64).collect(),
            ))
        };
        for (values, finance, reinvest, bits) in [
            (
                &[-120000.0, 39000.0, 30000.0, 21000.0, 37000.0, 46000.0][..],
                0.1,
                0.12,
                0x3FC0_23DA_3B11_CED0,
            ),
            (
                &[-100.0, 300.0, -300.0, 150.0],
                0.1,
                0.12,
                0x3FC2_EF92_009E_4178,
            ),
        ] {
            assert_eq!(
                mirr(&[flows(values), vf64(finance), vf64(reinvest)])
                    .unwrap()
                    .as_f64()
                    .unwrap()
                    .to_bits(),
                bits
            );
        }
        let mixed = flows(&[-100.0, 50.0, 60.0]);
        for (finance, reinvest) in [(-1.0, 0.12), (0.1, -1.0)] {
            assert_eq!(
                mirr(&[mixed.clone(), vf64(finance), vf64(reinvest)])
                    .unwrap_err()
                    .code,
                5
            );
        }
        for values in [&[10.0, 20.0, 30.0][..], &[0.0, 0.0, 0.0]] {
            assert_eq!(
                mirr(&[flows(values), vf64(0.1), vf64(0.12)])
                    .unwrap_err()
                    .code,
                11
            );
        }
        assert_eq!(
            mirr(&[
                Variant::unallocated_array(calcvba_runtime::safe_array::VT_R8_VALUE),
                vf64(0.1),
                vf64(0.1)
            ])
            .unwrap_err()
            .code,
            5
        );
    }

    /// Every line of `conformance/excel-measurements/rate-shape/RateShape.excel.txt` (dc-bjt7v):
    /// `Rate` over the probe's grid, by IEEE bits or error number.
    #[test]
    fn rate_matches_excel_rate_shape() {
        const EXCEL: &str =
            include_str!("../../../conformance/excel-measurements/rate-shape/RateShape.excel.txt");
        // (key, nper, pmt, pv, fv, type, guess) as RateShape.bas builds them.
        type Case = (String, f64, f64, f64, f64, f64, Option<f64>);
        let mut cases: Vec<Case> = Vec::new();
        let guesses: [(&str, Option<f64>); 15] = [
            ("", None),
            ("0.1", Some(0.1)),
            ("0", Some(0.0)),
            ("0.5", Some(0.5)),
            ("1", Some(1.0)),
            ("5", Some(5.0)),
            ("10", Some(10.0)),
            ("18", Some(18.0)),
            ("18.9", Some(18.9)),
            ("19", Some(19.0)),
            ("19.1", Some(19.1)),
            ("20", Some(20.0)),
            ("25", Some(25.0)),
            ("-0.5", Some(-0.5)),
            ("100", Some(100.0)),
        ];
        for (k, g) in guesses {
            cases.push((format!("lin19_g{k}"), 1.0, 0.0, -1.0, 20.0, 0.0, g));
        }
        for (k, f) in [
            ("1.1", 1.1),
            ("1.5", 1.5),
            ("2", 2.0),
            ("3", 3.0),
            ("5", 5.0),
            ("11", 11.0),
            ("21", 21.0),
            ("50", 50.0),
            ("100", 100.0),
            ("1000", 1000.0),
            ("10000", 10000.0),
            ("1000000", 1000000.0),
        ] {
            cases.push((format!("lin_fv{k}"), 1.0, 0.0, -1.0, f, 0.0, None));
        }
        for (k, f) in [("1.21", 1.21), ("4", 4.0), ("9", 9.0), ("100", 100.0)] {
            cases.push((format!("n2_fv{k}"), 2.0, 0.0, -1.0, f, 0.0, None));
        }
        cases.push(("n3_fv8".into(), 3.0, 0.0, -1.0, 8.0, 0.0, None));
        cases.push(("n10_fv2".into(), 10.0, 0.0, -1.0, 2.0, 0.0, None));
        for (k, g) in [
            ("", None),
            ("0", Some(0.0)),
            ("-0.004", Some(-0.004)),
            ("-0.0049", Some(-0.0049)),
            ("-0.005", Some(-0.005)),
            ("0.01", Some(0.01)),
            ("0.5", Some(0.5)),
        ] {
            cases.push((format!("a360_g{k}"), 360.0, -1.0, 1000.0, 0.0, 0.0, g));
        }
        for (k, n, pmt, pv, fv, t, g) in [
            ("a10", 10.0, -100.0, 500.0, 0.0, 0.0, None),
            ("a10due", 10.0, -100.0, 500.0, 0.0, 1.0, None),
            ("a12g05", 12.0, -10.0, 100.0, 0.0, 0.0, Some(0.5)),
            ("a48", 48.0, -250.0, 10000.0, 0.0, 0.0, None),
            ("a48fv", 48.0, -250.0, 10000.0, -1000.0, 0.0, None),
            ("a5mixed", 5.0, -100.0, 300.0, 200.0, 0.0, None),
            ("a360due", 360.0, -6.0, 1000.0, 0.0, 1.0, None),
            ("a3exact", 3.0, -40.0, 100.0, 0.0, 0.0, None),
        ] {
            cases.push((k.into(), n, pmt, pv, fv, t, g));
        }
        for i in 1..=12 {
            let fv = 2f64.powi(i);
            cases.push((format!("far_n1_{i}"), 1.0, 0.0, -1.0, fv, 0.0, None));
        }
        for g in [1.0, 2.0, 5.0, 10.0, 50.0] {
            cases.push((
                format!("a360_far_g{g}"),
                360.0,
                -1.0,
                1000.0,
                0.0,
                0.0,
                Some(g),
            ));
        }
        cases.push(("noroot".into(), 10.0, 100.0, 100.0, 100.0, 0.0, None));
        cases.push(("zero_rate".into(), 10.0, -10.0, 100.0, 0.0, 0.0, None));

        let expected: Vec<(&str, &str)> = EXCEL
            .lines()
            .filter(|l| !l.trim().is_empty())
            .map(|l| {
                let (k, v) = l.trim_end().split_once('=').unwrap();
                (k, v.split('|').next().unwrap())
            })
            .collect();
        assert_eq!(expected.len(), cases.len());
        let mut wrong = Vec::new();
        for ((key, n, pmt, pv, fv, t, g), (want_key, want)) in cases.iter().zip(&expected) {
            assert_eq!(key, want_key);
            let mut args = vec![
                Variant::from_f64(*n),
                Variant::from_f64(*pmt),
                Variant::from_f64(*pv),
                Variant::from_f64(*fv),
                Variant::from_f64(*t),
            ];
            if let Some(g) = g {
                args.push(Variant::from_f64(*g));
            }
            let got = match rate(&args) {
                Ok(v) => format!("{:016X}", v.as_f64().unwrap().to_bits()),
                Err(e) => format!("ERR{}", e.code),
            };
            if got != *want {
                wrong.push(format!("{key}: Excel {want}, CalcVBA {got}"));
            }
        }
        assert!(wrong.is_empty(), "{}", wrong.join("\n"));
    }
}
