//! VBA 7.1's regular-expression objects. State and returned objects are owned by
//! the ordinary runtime boxes; both executors call this same dispatcher.
use calcvba_runtime::ObjectRefExt;
use std::sync::atomic::{AtomicI32, Ordering};

use calcvba_bundle::ProjectMemberKind;
use calcvba_runtime::{
    ObjectRef, Variant, object_ref::COMPAT_OBJECT_CLASS_DESCRIPTOR,
    object_ref::RuntimeClassDescriptor, variant::VarType,
};
use regress::{Flags, Match, Regex};

use crate::{LibError, LibResult};

const REGEXP: i32 = i32::MIN + 1;
const MATCH: i32 = i32::MIN + 2;
const MATCHES: i32 = i32::MIN + 3;
const SUBMATCHES: i32 = i32::MIN + 4;
const PATTERN: i32 = 0;
const GLOBAL: i32 = 1;
const IGNORE_CASE: i32 = 2;
const MULTILINE: i32 = 3;

static REGEXP_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
    name: "RegExp",
    ..COMPAT_OBJECT_CLASS_DESCRIPTOR
};
static MATCH_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
    name: "Match",
    ..COMPAT_OBJECT_CLASS_DESCRIPTOR
};
static MATCHES_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
    name: "MatchCollection",
    ..COMPAT_OBJECT_CLASS_DESCRIPTOR
};
static SUBMATCHES_CLASS: RuntimeClassDescriptor = RuntimeClassDescriptor {
    name: "SubMatches",
    ..COMPAT_OBJECT_CLASS_DESCRIPTOR
};

fn object(route: i32, descriptor: &'static RuntimeClassDescriptor) -> ObjectRef {
    static NEXT: AtomicI32 = AtomicI32::new(-1_000_000);
    ObjectRef::from_project_instance(
        NEXT.fetch_sub(1, Ordering::Relaxed),
        route,
        0,
        false,
        descriptor,
    )
}

pub fn new_regexp() -> Variant {
    let o = object(REGEXP, &REGEXP_CLASS);
    o.project_field_set(PATTERN, Variant::from_string(""));
    for token in [GLOBAL, IGNORE_CASE, MULTILINE] {
        o.project_field_set(token, Variant::from_bool(false));
    }
    Variant::from_object_ref(o)
}

pub fn is_object(object: &ObjectRef) -> bool {
    object.is_project_instance() && (REGEXP..=SUBMATCHES).contains(&object.route_key())
}

/// Canonical class name, also used by the binder and assignment validation.
pub fn class_name(name: &str) -> Option<&'static str> {
    let bare = match name.split_once('.') {
        Some((unit, bare)) if unit.eq_ignore_ascii_case("VBA") => bare,
        Some(_) => return None,
        None => name,
    };
    ["RegExp", "Match", "MatchCollection", "SubMatches"]
        .into_iter()
        .find(|n| n.eq_ignore_ascii_case(bare))
}

/// Parameter order applies to both early and late named calls.
fn parameters(name: &str) -> &'static [&'static str] {
    match name {
        "test" | "execute" => &["sourceString"],
        "replace" => &["sourceString", "replaceVar"],
        "item" => &["Index"],
        _ => &[],
    }
}

pub fn dispatch(
    object: &ObjectRef,
    name: &str,
    kind: ProjectMemberKind,
    args: &[Variant],
    names: &[Option<String>],
) -> LibResult<Variant> {
    let name = if name.is_empty() {
        if object.route_key() == MATCH {
            "value"
        } else {
            "item"
        }
    } else {
        name
    };
    let name = name.to_ascii_lowercase();
    let params = parameters(&name);
    let mut ordered: Vec<Option<Variant>> = vec![None; args.len().max(params.len())];
    let mut next = 0;
    for (i, arg) in args.iter().enumerate() {
        let index = if let Some(Some(n)) = names.get(i) {
            params
                .iter()
                .position(|p| p.eq_ignore_ascii_case(n))
                .ok_or_else(|| LibError::new(448, "Named argument not found"))?
        } else {
            let index = next;
            next += 1;
            index
        };
        if index >= ordered.len() || ordered[index].is_some() {
            return Err(LibError::new(450, "Wrong number of arguments"));
        }
        ordered[index] = Some(arg.clone());
    }
    let args: Vec<_> = ordered
        .into_iter()
        .map(|a| a.unwrap_or_else(|| Variant::from_error_code(0x8002_0004u32 as i32)))
        .collect();
    let put = matches!(
        kind,
        ProjectMemberKind::PropertyLet | ProjectMemberKind::PropertySet
    );
    match object.route_key() {
        REGEXP => {
            let property = match name.as_str() {
                "pattern" => Some(PATTERN),
                "global" => Some(GLOBAL),
                "ignorecase" => Some(IGNORE_CASE),
                "multiline" => Some(MULTILINE),
                _ => None,
            };
            if let Some(token) = property {
                if put {
                    arity(&args, 1)?;
                    let value = if token == PATTERN {
                        crate::pure::cstr(&args)?
                    } else {
                        crate::pure::cbool(&args)?
                    };
                    object.project_field_set(token, value);
                    return Ok(Variant::empty());
                }
                arity(&args, 0)?;
                return Ok(object
                    .project_field_get(token)
                    .unwrap_or_else(Variant::empty));
            }
            if put {
                return Err(LibError::new(
                    438,
                    "Object doesn't support this property or method",
                ));
            }
            if !matches!(name.as_str(), "test" | "execute" | "replace") {
                return Err(LibError::new(
                    438,
                    "Object doesn't support this property or method",
                ));
            }
            arity(&args, if name == "replace" { 2 } else { 1 })?;
            let source = units(&args[0])?;
            let pattern = object.project_field_get(PATTERN).unwrap();
            let flags = Flags {
                icase: flag(object, IGNORE_CASE),
                multiline: flag(object, MULTILINE),
                ..Flags::default()
            };
            let regex = compile(&units(&pattern)?, flags)?;
            let mut matches = regex.find_from_ucs2(&source, 0);
            if name == "test" {
                return Ok(Variant::from_bool(matches.next().is_some()));
            }
            let all = flag(object, GLOBAL);
            let matches: Vec<_> = matches.take(if all { usize::MAX } else { 1 }).collect();
            match name.as_str() {
                "execute" => {
                    let values = matches.iter().map(|m| match_object(&source, m)).collect();
                    Ok(Variant::from_object_ref(collection(
                        MATCHES,
                        &MATCHES_CLASS,
                        values,
                    )))
                }
                "replace" => {
                    // replaceVar is a Variant; Null is a type mismatch, rather than
                    // the invalid-use-of-Null error of a BSTR parameter.
                    if matches!(args[1].vtype(), VarType::Null | VarType::Error) {
                        return Err(LibError::type_mismatch("Type mismatch"));
                    }
                    let replacement = units(&args[1])?;
                    let mut out = Vec::new();
                    let mut cursor = 0;
                    for m in matches {
                        out.extend_from_slice(&source[cursor..m.start()]);
                        expand_replacement(&mut out, &replacement, &source, &m);
                        cursor = m.end();
                    }
                    out.extend_from_slice(&source[cursor..]);
                    Ok(Variant::from_utf16_units(&out))
                }
                _ => Err(LibError::new(
                    438,
                    "Object doesn't support this property or method",
                )),
            }
        }
        MATCH if !put => {
            if name == "submatches" && !args.is_empty() {
                let submatches = object
                    .project_field_get(3)
                    .unwrap()
                    .as_object_ref()
                    .unwrap();
                return dispatch(&submatches, "Item", kind, &args, &[]);
            }
            arity(&args, 0)?;
            let token = match name.as_str() {
                "value" => 0,
                "firstindex" => 1,
                "length" => 2,
                "submatches" => 3,
                _ => {
                    return Err(LibError::new(
                        438,
                        "Object doesn't support this property or method",
                    ));
                }
            };
            Ok(object
                .project_field_get(token)
                .unwrap_or_else(Variant::empty))
        }
        MATCHES | SUBMATCHES if !put => match name.as_str() {
            "count" => {
                arity(&args, 0)?;
                Ok(Variant::from_i32(
                    object.with_native_collection(|d| d.count()).unwrap(),
                ))
            }
            "item" => {
                arity(&args, 1)?;
                if args[0].vtype() == VarType::Null {
                    return Err(LibError::type_mismatch("Type mismatch"));
                }
                let index = crate::as_i32(&args[0])?;
                let one_based = index
                    .checked_add(1)
                    .filter(|_| index >= 0)
                    .ok_or_else(|| LibError::invalid_call("Invalid index"))?;
                object
                    .with_native_collection(|d| {
                        d.item(&calcvba_runtime::collection::Selector::Index(one_based))
                    })
                    .unwrap()
                    .map_err(|_| LibError::invalid_call("Invalid index"))
            }
            "_newenum" => {
                arity(&args, 0)?;
                Ok(Variant::from_object_ref(object.clone()))
            }
            _ => Err(LibError::new(
                438,
                "Object doesn't support this property or method",
            )),
        },
        _ => Err(LibError::new(
            438,
            "Object doesn't support this property or method",
        )),
    }
}

fn arity(args: &[Variant], count: usize) -> LibResult<()> {
    if args.len() != count
        || args
            .iter()
            .any(|v| v.as_error_code() == Some(0x8002_0004u32 as i32))
    {
        Err(LibError::new(450, "Wrong number of arguments"))
    } else {
        Ok(())
    }
}
fn flag(object: &ObjectRef, token: i32) -> bool {
    object
        .project_field_get(token)
        .and_then(|v| v.as_bool())
        .unwrap_or(false)
}
fn units(value: &Variant) -> LibResult<Vec<u16>> {
    if let Some(units) = value.string_units() {
        return Ok(units);
    }
    Ok(crate::pure::cstr(std::slice::from_ref(value))?
        .string_units()
        .unwrap())
}
fn collection(
    route: i32,
    descriptor: &'static RuntimeClassDescriptor,
    values: Vec<Variant>,
) -> ObjectRef {
    let o = object(route, descriptor);
    o.with_native_collection(|d| {
        for value in values {
            d.add(value, None, None, None).expect("unkeyed append");
        }
    });
    o
}
fn match_object(source: &[u16], m: &Match) -> Variant {
    let o = object(MATCH, &MATCH_CLASS);
    o.project_field_set(0, Variant::from_utf16_units(&source[m.range.clone()]));
    o.project_field_set(1, Variant::from_i32(m.start() as i32));
    o.project_field_set(2, Variant::from_i32(m.range.len() as i32));
    let captures = m
        .captures
        .iter()
        .map(|r| {
            r.as_ref().map_or_else(Variant::empty, |r| {
                Variant::from_utf16_units(&source[r.clone()])
            })
        })
        .collect();
    o.project_field_set(
        3,
        Variant::from_object_ref(collection(SUBMATCHES, &SUBMATCHES_CLASS, captures)),
    );
    Variant::from_object_ref(o)
}

fn compile(pattern: &[u16], flags: Flags) -> LibResult<Regex> {
    // VBScript's own syntax checks, measured in Excel (P_RegExpEdges): an empty class `[]` is
    // an unbalanced bracket (5019); a `)` with no open group is a syntax error (5017).
    let mut depth = 0i32;
    let mut in_class = false;
    let mut j = 0;
    while j < pattern.len() {
        match pattern[j] {
            92 => j += 1, // backslash escape
            91 if !in_class => {
                if pattern.get(j + 1) == Some(&93) {
                    return Err(LibError::new(5019, "Unbalanced bracket"));
                }
                in_class = true;
            }
            93 if in_class => in_class = false,
            40 if !in_class => depth += 1,
            41 if !in_class => {
                depth -= 1;
                if depth < 0 {
                    return Err(LibError::new(5017, "Syntax error in regular expression"));
                }
            }
            _ => {}
        }
        j += 1;
    }
    let mut translated = Vec::new();
    let mut class = false;
    let mut i = 0;
    while i < pattern.len() {
        let c = pattern[i];
        if c == b'\\' as u16 {
            translated.push(u32::from(c));
            i += 1;
            if let Some(c) = pattern.get(i) {
                translated.push(u32::from(*c));
            }
        } else {
            if !class
                && c == b'(' as u16
                && pattern.get(i + 1) == Some(&(b'?' as u16))
                && !matches!(pattern.get(i + 2), Some(58 | 61 | 33))
            {
                return Err(LibError::new(5017, "Syntax error in regular expression"));
            }
            if !class && c == b'.' as u16 {
                // Excel's dot consumes CR (P_Match.17), but not LF.
                translated.extend("[^\\n]".chars().map(u32::from));
            } else {
                translated.push(u32::from(c));
            }
            if c == b'[' as u16 {
                class = true;
            }
            if c == b']' as u16 {
                class = false;
            }
        }
        i += 1;
    }
    Regex::from_unicode(translated.into_iter(), flags).map_err(|e| {
        let code = match e.text.as_str() {
            "Unbalanced parenthesis" => 5020,
            "Unbalanced bracket" => 5019,
            "Quantifier not allowed here" | "Invalid atom character" => 5018,
            "Invalid character range"
            | "Range values reversed, start char code is greater than end char code." => 5021,
            _ => 5017,
        };
        LibError::new(code, e.text)
    })
}

fn expand_replacement(out: &mut Vec<u16>, replacement: &[u16], source: &[u16], m: &Match) {
    let mut i = 0;
    while i < replacement.len() {
        if replacement[i] != b'$' as u16 || i + 1 == replacement.len() {
            out.push(replacement[i]);
            i += 1;
            continue;
        }
        let next = replacement[i + 1];
        let slice = match next {
            36 => Some(&replacement[i..i + 1]),   // $$
            38 => Some(&source[m.range.clone()]), // $&
            96 => Some(&source[..m.start()]),     // $`
            39 => Some(&source[m.end()..]),       // $'
            _ => None,
        };
        if let Some(slice) = slice {
            out.extend_from_slice(slice);
            i += 2;
            continue;
        }
        // `$0d` (d 1-9) is group d when that group exists; `$0` and `$00` stay literal.
        if next == b'0' as u16
            && let Some(&d) = replacement.get(i + 2).filter(|&&c| (49..=57).contains(&c))
        {
            {
                let group = (d - 48) as usize;
                if group <= m.captures.len() {
                    if let Some(range) = &m.captures[group - 1] {
                        out.extend_from_slice(&source[range.clone()]);
                    }
                    i += 3;
                    continue;
                }
            }
        }
        if (b'1' as u16..=b'9' as u16).contains(&next) {
            let mut group = (next - b'0' as u16) as usize;
            let mut consumed = 2;
            if let Some(second) = replacement.get(i + 2).filter(|&&c| (48..=57).contains(&c)) {
                let two = group * 10 + (*second - 48) as usize;
                if two <= m.captures.len() {
                    group = two;
                    consumed = 3;
                }
            }
            if group <= m.captures.len() {
                if let Some(range) = &m.captures[group - 1] {
                    out.extend_from_slice(&source[range.clone()]);
                }
                i += consumed;
                continue;
            }
        }
        out.push(replacement[i]);
        i += 1;
    }
}
