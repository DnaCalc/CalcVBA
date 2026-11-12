//! `Scripting.Dictionary`, served by the engine itself.
//!
//! The Dictionary is safe: it touches nothing outside its own memory. Instead of activating
//! the COM class (which needs Windows and the `com.activate` capability), the binder entry
//! points call [`with_builtin_scripting`], which rewrites a manifest that uses it:
//!
//! - `CreateObject("Scripting.Dictionary")` becomes `DnaScripting.NewDictionary()`;
//! - the type name `Scripting.Dictionary` (and a bare `Dictionary` after `As` or `New` when
//!   the manifest references the Scripting library and has no module of that name) becomes
//!   `DnaScriptingDictionary`;
//! - two modules are added: the class `DnaScriptingDictionary` (plain VBA source, below) and
//!   the factory module `DnaScripting`.
//!
//! Both executors therefore run the same VBA code on every platform. A ProgID computed at run
//! time is not recognised and still goes through the COM activation policy.

use std::borrow::Cow;

use crate::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectReference, SymbolProjectManifest,
};

const CLASS_NAME: &str = "DnaScriptingDictionary";
const FACTORY_NAME: &str = "DnaScripting";
const PROGID: &str = "Scripting.Dictionary";

/// Only the engine's injected implementation has a public name distinct from its
/// private source name. A user class with the same name retains its own identity.
pub fn runtime_type_name(module: &ModuleUnit) -> Option<&'static str> {
    (module.module_name == CLASS_NAME && module.source == DICTIONARY_CLASS_SOURCE)
        .then_some("Dictionary")
}

/// The Dictionary class, as VBA source.
pub const DICTIONARY_CLASS_SOURCE: &str = r#"Option Explicit
Private pKeys() As Variant
Private pItems() As Variant
Private pCount As Long
Private pMode As Long

Private Sub Class_Initialize()
    ReDim pKeys(0 To 7)
    ReDim pItems(0 To 7)
End Sub

Public Property Get CompareMode() As Long
    CompareMode = pMode
End Property

Public Property Let CompareMode(ByVal Value As Long)
    If pCount > 0 Then Err.Raise 5, "Scripting.Dictionary", "Invalid procedure call or argument"
    If Value < 0 Or Value > 2 Then Err.Raise 5, "Scripting.Dictionary", "Invalid procedure call or argument"
    pMode = Value
End Property

Public Property Get Count() As Long
    Count = pCount
End Property

Public Property Get Item(ByVal Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim i As Long
    i = FindKey(Key)
    If i < 0 Then
        Insert Key, Empty
        Exit Property
    End If
    If IsObject(pItems(i)) Then
        Set Item = pItems(i)
    Else
        Item = pItems(i)
    End If
End Property

Public Property Let Item(ByVal Key As Variant, ByVal Value As Variant)
    Dim i As Long
    i = FindKey(Key)
    If i < 0 Then
        Insert Key, Value
    Else
        pItems(i) = Value
    End If
End Property

Public Property Set Item(ByVal Key As Variant, ByVal Value As Object)
    Dim i As Long
    i = FindKey(Key)
    If i < 0 Then
        Insert Key, Value
    Else
        Set pItems(i) = Value
    End If
End Property

Public Property Let Key(ByVal OldKey As Variant, ByVal NewKey As Variant)
    Dim i As Long
    i = FindKey(OldKey)
    If i < 0 Then Err.Raise 32811, "Scripting.Dictionary", "Application-defined or object-defined error"
    If FindKey(NewKey) >= 0 Then Err.Raise 457, "Scripting.Dictionary", "This key is already associated with an element of this collection"
    pKeys(i) = NewKey
End Property

Public Sub Add(ByVal Key As Variant, ByVal Item As Variant)
    If FindKey(Key) >= 0 Then Err.Raise 457, "Scripting.Dictionary", "This key is already associated with an element of this collection"
    Insert Key, Item
End Sub

Public Function Exists(ByVal Key As Variant) As Boolean
    Exists = (FindKey(Key) >= 0)
End Function

Public Function Keys() As Variant
    Dim r() As Variant
    Dim i As Long, n As Long
    n = pCount
    If n = 0 Then
        Keys = Array()
        Exit Function
    End If
    ReDim r(0 To n - 1)
    For i = 0 To n - 1
        If IsObject(pKeys(i)) Then
            Set r(i) = pKeys(i)
        Else
            r(i) = pKeys(i)
        End If
    Next i
    Keys = r
End Function

Public Function Items() As Variant
    Dim r() As Variant
    Dim i As Long, n As Long
    n = pCount
    If n = 0 Then
        Items = Array()
        Exit Function
    End If
    ReDim r(0 To n - 1)
    For i = 0 To n - 1
        If IsObject(pItems(i)) Then
            Set r(i) = pItems(i)
        Else
            r(i) = pItems(i)
        End If
    Next i
    Items = r
End Function

Public Sub Remove(ByVal Key As Variant)
    Dim i As Long, last As Long
    i = FindKey(Key)
    If i < 0 Then Err.Raise 32811, "Scripting.Dictionary", "Application-defined or object-defined error"
    last = pCount - 1
    Do While i < last
        pKeys(i) = pKeys(i + 1)
        If IsObject(pItems(i + 1)) Then
            Set pItems(i) = pItems(i + 1)
        Else
            pItems(i) = pItems(i + 1)
        End If
        i = i + 1
    Loop
    pKeys(last) = Empty
    pItems(last) = Empty
    pCount = last
End Sub

Public Sub RemoveAll()
    ReDim pKeys(0 To 7)
    ReDim pItems(0 To 7)
    pCount = 0
End Sub

Public Function HashVal(ByVal Key As Variant) As Variant
    HashVal = FindKey(Key)
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Dim c As New VBA.Collection
    Dim i As Long
    For i = 0 To pCount - 1
        c.Add pKeys(i)
    Next i
    Set NewEnum = c.[_NewEnum]
End Property

Private Sub Insert(ByVal Key As Variant, ByVal Value As Variant)
    Dim n As Long, cap As Long
    n = pCount
    cap = UBound(pKeys)
    If n > cap Then
        cap = n * 2 + 7
        ReDim Preserve pKeys(0 To cap)
        ReDim Preserve pItems(0 To cap)
    End If
    If IsObject(Key) Then
        Set pKeys(n) = Key
    Else
        pKeys(n) = Key
    End If
    If IsObject(Value) Then
        Set pItems(n) = Value
    Else
        pItems(n) = Value
    End If
    pCount = n + 1
End Sub

Private Function FindKey(ByVal Key As Variant) As Long
    Dim i As Long
    For i = 0 To pCount - 1
        If SameKey(pKeys(i), Key) Then
            FindKey = i
            Exit Function
        End If
    Next i
    FindKey = -1
End Function

Private Function SameKey(ByVal A As Variant, ByVal B As Variant) As Boolean
    Dim ta As Long, tb As Long
    ta = VarType(A)
    tb = VarType(B)
    If ta = vbString Or tb = vbString Then
        If ta = vbString And tb = vbString Then
            If pMode = 0 Then
                SameKey = (StrComp(A, B, vbBinaryCompare) = 0)
            Else
                SameKey = (StrComp(A, B, vbTextCompare) = 0)
            End If
        End If
    ElseIf IsObject(A) Or IsObject(B) Then
        If IsObject(A) And IsObject(B) Then SameKey = (A Is B)
    ElseIf IsNumeric(A) And IsNumeric(B) Then
        SameKey = (CDbl(A) = CDbl(B))
    Else
        SameKey = (ta = tb)
    End If
End Function
"#;

const FACTORY_SOURCE: &str = r#"Option Explicit
Public Function NewDictionary() As Object
    Set NewDictionary = New DnaScriptingDictionary
End Function
"#;

/// The Scripting library's CompareMethod constants, added when a project uses them unqualified
/// and declares none of its own.
const COMPARE_METHOD_SOURCE: &str = r#"Public Enum CompareMethod
    BinaryCompare = 0
    TextCompare = 1
    DatabaseCompare = 2
End Enum
"#;

/// The manifest with built-in Scripting.Dictionary uses rewritten, or the manifest itself when
/// it has none.
pub fn with_builtin_scripting(manifest: &SymbolProjectManifest) -> Cow<'_, SymbolProjectManifest> {
    let has_scripting_ref = manifest.references.iter().any(|reference| match reference {
        ProjectReference::TypeLibrary { name, guid, .. } => {
            name.eq_ignore_ascii_case("Scripting")
                || guid.as_deref().is_some_and(|g| {
                    g.eq_ignore_ascii_case("{420B2830-E718-11CF-893D-00A0C9054228}")
                })
        }
        _ => false,
    });
    let own_dictionary = manifest
        .modules
        .iter()
        .any(|m| m.module_name.eq_ignore_ascii_case("Dictionary"));
    let bare = has_scripting_ref && !own_dictionary;
    let mut changed = false;
    let modules: Vec<ModuleUnit> = manifest
        .modules
        .iter()
        .map(|module| {
            let source = rewrite_source(&module.source, bare);
            if source != module.source {
                changed = true;
                ModuleUnit {
                    source,
                    ..module.clone()
                }
            } else {
                module.clone()
            }
        })
        .collect();
    if !changed {
        return Cow::Borrowed(manifest);
    }
    let mut out = manifest.clone();
    out.modules = modules;
    if !out.modules.iter().any(|m| m.module_name == CLASS_NAME) {
        out.modules.push(ModuleUnit {
            module_name: CLASS_NAME.into(),
            module_kind: ModuleKind::Class,
            attributes: ModuleAttributes::named(CLASS_NAME),
            source: DICTIONARY_CLASS_SOURCE.into(),
        });
        let lowered: Vec<String> = out
            .modules
            .iter()
            .map(|m| m.source.to_ascii_lowercase())
            .collect();
        let uses_constants = lowered
            .iter()
            .any(|s| s.contains("textcompare") || s.contains("binarycompare"));
        let declares_constants = lowered.iter().any(|s| s.contains("enum comparemethod"));
        let mut factory = FACTORY_SOURCE.to_string();
        if uses_constants && !declares_constants {
            factory.push_str(COMPARE_METHOD_SOURCE);
        }
        out.modules.push(ModuleUnit {
            module_name: FACTORY_NAME.into(),
            module_kind: ModuleKind::Procedural,
            attributes: ModuleAttributes::named(FACTORY_NAME),
            source: factory,
        });
    }
    // The Scripting library is then needed only if something else of it is still used.
    let still_used = out.modules.iter().any(|m| {
        m.module_name != CLASS_NAME && m.source.to_ascii_lowercase().contains("scripting.")
    });
    if !still_used {
        out.references.retain(|reference| {
            !matches!(reference, ProjectReference::TypeLibrary { name, .. } if name.eq_ignore_ascii_case("Scripting"))
        });
    }
    Cow::Owned(out)
}

fn is_ident(c: char) -> bool {
    c.is_alphanumeric() || c == '_'
}

fn starts_with_ci(chars: &[char], at: usize, word: &str) -> bool {
    let w: Vec<char> = word.chars().collect();
    at + w.len() <= chars.len()
        && chars[at..at + w.len()]
            .iter()
            .zip(&w)
            .all(|(a, b)| a.eq_ignore_ascii_case(b))
}

/// The last word of `out` (ignoring trailing blanks), lowercase.
fn previous_word(out: &str) -> String {
    let trimmed = out.trim_end_matches([' ', '\t']);
    let start = trimmed
        .char_indices()
        .rev()
        .find(|(_, c)| !is_ident(*c))
        .map(|(i, c)| i + c.len_utf8())
        .unwrap_or(0);
    trimmed[start..].to_ascii_lowercase()
}

fn rewrite_source(source: &str, bare: bool) -> String {
    let mut result = String::with_capacity(source.len());
    for line in source.split_inclusive('\n') {
        result.push_str(&rewrite_line(line, bare));
    }
    result
}

fn rewrite_line(line: &str, bare: bool) -> String {
    let chars: Vec<char> = line.chars().collect();
    let mut out = String::with_capacity(line.len() + 16);
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            out.push(c);
            i += 1;
            while i < chars.len() {
                out.push(chars[i]);
                if chars[i] == '"' {
                    if chars.get(i + 1) == Some(&'"') {
                        out.push('"');
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        if c == '\'' {
            out.extend(&chars[i..]);
            break;
        }
        let boundary = i == 0 || !is_ident(chars[i - 1]);
        if boundary && is_ident(c) {
            let after_dot = out.trim_end_matches([' ', '\t']).ends_with('.');
            let after_vba = out.to_ascii_lowercase().ends_with("vba.");
            if !after_dot || after_vba {
                if starts_with_ci(&chars, i, "Rem")
                    && chars.get(i + 3).is_none_or(|n| !is_ident(*n))
                    && out.trim().is_empty()
                {
                    out.extend(&chars[i..]);
                    break;
                }
                if starts_with_ci(&chars, i, "Scripting.Dictionary")
                    && chars.get(i + 20).is_none_or(|n| !is_ident(*n))
                {
                    out.push_str(CLASS_NAME);
                    i += 20;
                    continue;
                }
                if starts_with_ci(&chars, i, "CreateObject")
                    && chars.get(i + 12).is_none_or(|n| !is_ident(*n))
                    && let Some(end) = create_object_end(&chars, i + 12)
                {
                    if after_vba {
                        out.truncate(out.len() - 4);
                    }
                    out.push_str("DnaScripting.NewDictionary()");
                    i = end;
                    continue;
                }
                if bare
                    && !after_dot
                    && starts_with_ci(&chars, i, "Dictionary")
                    && chars.get(i + 10).is_none_or(|n| !is_ident(*n) && *n != '.')
                    && matches!(previous_word(&out).as_str(), "as" | "new")
                {
                    out.push_str(CLASS_NAME);
                    i += 10;
                    continue;
                }
            }
            // Copy the whole identifier.
            while i < chars.len() && is_ident(chars[i]) {
                out.push(chars[i]);
                i += 1;
            }
            continue;
        }
        out.push(c);
        i += 1;
    }
    out
}

/// After `CreateObject`: `( "Scripting.Dictionary" )`, returning the index past the `)`.
fn create_object_end(chars: &[char], mut i: usize) -> Option<usize> {
    let skip_ws = |i: &mut usize| {
        while chars.get(*i).is_some_and(|c| *c == ' ' || *c == '\t') {
            *i += 1;
        }
    };
    skip_ws(&mut i);
    if chars.get(i) != Some(&'(') {
        return None;
    }
    i += 1;
    skip_ws(&mut i);
    if chars.get(i) != Some(&'"') {
        return None;
    }
    i += 1;
    let text: String = chars.get(i..i + PROGID.len())?.iter().collect();
    if !text.eq_ignore_ascii_case(PROGID) {
        return None;
    }
    i += PROGID.len();
    if chars.get(i) != Some(&'"') {
        return None;
    }
    i += 1;
    skip_ws(&mut i);
    (chars.get(i) == Some(&')')).then_some(i + 1)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn public_name_applies_only_to_the_injected_dictionary_implementation() {
        let mut module = ModuleUnit {
            module_name: CLASS_NAME.into(),
            module_kind: ModuleKind::Class,
            attributes: ModuleAttributes::named(CLASS_NAME),
            source: DICTIONARY_CLASS_SOURCE.into(),
        };
        assert_eq!(runtime_type_name(&module), Some("Dictionary"));
        module.source = "Option Explicit\n".into();
        assert_eq!(runtime_type_name(&module), None);
    }

    #[test]
    fn rewrites_progid_and_type_names_outside_strings_and_comments() {
        let src = "Dim d As Scripting.Dictionary\nSet d = CreateObject( \"scripting.dictionary\" )\nx = \"Scripting.Dictionary\" ' CreateObject(\"Scripting.Dictionary\")\nSet e = VBA.CreateObject(\"Scripting.Dictionary\")\n";
        let out = rewrite_source(src, false);
        assert_eq!(
            out,
            "Dim d As DnaScriptingDictionary\nSet d = DnaScripting.NewDictionary()\nx = \"Scripting.Dictionary\" ' CreateObject(\"Scripting.Dictionary\")\nSet e = DnaScripting.NewDictionary()\n"
        );
    }

    #[test]
    fn bare_dictionary_type_needs_as_or_new() {
        let out = rewrite_source(
            "Dim a As Dictionary, b As New Dictionary\nx = Dictionary.Count\n",
            true,
        );
        assert_eq!(
            out,
            "Dim a As DnaScriptingDictionary, b As New DnaScriptingDictionary\nx = Dictionary.Count\n"
        );
    }
}
