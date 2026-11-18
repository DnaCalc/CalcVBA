//! The VBA base-library provider: `vb*` value constants, the intrinsic catalog,
//! the user-facing structural intrinsics (`VarPtr`/`StrPtr`/`ObjPtr`), the special
//! forms (`Array`/`IIf`/…), and the predeclared `Debug`/`Err` objects. Constants
//! are ported verbatim from the legacy `frontend_library`.
//!
//! The built-in `Collection` is *not* a predeclared object: it is a class of the
//! `VBA` library bundle, resolved here as a creatable extern coclass
//! (`resolve_extern_coclass`) whose members route cross-bundle to that unit
//! (`resolve_member` → `ExternMember`), exactly like a referenced project's class.

use calcvba_bundle::ProjectMemberKind;
use calcvba_bundle::coreir::CoreConst;

use crate::binding::{Binding, DispatchRoute, SpecialForm};
use crate::catalog::{intrinsic_entry, name_to_intrinsic};
use crate::model::{PredeclaredObjectId, fold_identifier};
use crate::predeclared::{predeclared_member, predeclared_object};
use crate::provider::Provider;
use crate::signature::VarTypeRef;
use crate::structural::StructuralIntrinsic;

/// The unit name of the built-in library bundle (see `calcvba-bundle/vba_library`).
const VBA_UNIT: &str = "VBA";

pub struct VbaLibraryProvider;

impl Provider for VbaLibraryProvider {
    fn resolve(&self, name: &str) -> Option<Binding> {
        // A `vb*` value constant resolves through the SAME shared route a referenced
        // project's `Public Const`/`Enum` member uses: the folded literal carried in
        // `ConstValue`. The binder inlines it identically (see `expr::const_type`),
        // so this is a pure routing change — the values are unchanged.
        if let Some(value) = library_constant(name) {
            return Some(Binding::new(None, DispatchRoute::ConstValue(value)));
        }
        if let Some(object) = predeclared_object(name) {
            return Some(Binding::new(None, DispatchRoute::PredeclaredObject(object)));
        }
        if let Some(id) = name_to_intrinsic(name) {
            // A migrated library function (the `Strings`/`Math`/`DateTime`/
            // `Conversion`/`Random`/`Financial` modules, the `Information` predicates,
            // the `Interaction` by-name functions, and the `FileIo` by-name functions —
            // exported under the `FileSystem` module) resolves like a referenced
            // project's hidden-module free function: a cross-bundle `ExternMember` (no
            // receiver) against the synthetic `VBA` unit's owning module. The binder
            // lowers it to a `ModuleFunc` import + `ExternProc` call, and the VM runs
            // the linked native-bodied proc through `calcvba-lib` — the same `calcvba-lib`
            // body the bespoke `Native` route used. Host-sensitive migrated functions
            // still receive the host via `invoke_native_lib`, so behaviour is
            // unchanged. The `(owner, member)`
            // location is `NativeImplId::library_member`, the single source of truth
            // shared with the bundle's export tokens, so the import links to the export
            // by construction. The non-migrated intrinsics — the `Information` special
            // forms (`IIf`/`Choose`/`Switch`), `CreateObject` + the `Com*` machinery,
            // the FileIo STATEMENT forms (P4), and Diagnostics — keep the `Native(id)`
            // route.
            if let Some((owner, member)) = id.library_member_for_name(name) {
                let entry = intrinsic_entry(id);
                return Some(Binding::new(
                    None,
                    DispatchRoute::ExternMember {
                        unit: VBA_UNIT.to_string(),
                        owner: owner.to_string(),
                        member: member.to_string(),
                        kind: ProjectMemberKind::Method,
                        param_types: Vec::new(),
                        param_names: intrinsic_param_names(&entry),
                        param_optional: intrinsic_param_optional(&entry),
                        param_optional_defaults: intrinsic_param_optional_defaults(&entry),
                        variadic: entry.sig.param_array,
                        has_receiver: false,
                    },
                ));
            }
            return Some(Binding::new(None, DispatchRoute::Native(id)));
        }
        if let Some(structural) = user_structural_intrinsic(name) {
            return Some(Binding::new(None, DispatchRoute::Structural(structural)));
        }
        if let Some(form) = special_form(name) {
            return Some(Binding::new(None, DispatchRoute::SpecialForm(form)));
        }
        None
    }

    fn resolve_member(
        &self,
        recv: &VarTypeRef,
        name: &str,
        want: Option<ProjectMemberKind>,
    ) -> Option<Binding> {
        let VarTypeRef::Object(type_name) = recv else {
            return None;
        };
        // The built-in `Collection` resolves like a referenced coclass: its members
        // bind cross-bundle to the `VBA` library class (late dispatch by name), not
        // via a predeclared/`NativeImplId` route.
        if is_collection_type(type_name) {
            return collection_member(name);
        }
        if let Some(binding) = regexp_member(type_name, name, want) {
            return Some(binding);
        }
        if is_err_object_type(type_name) {
            return err_object_member(name, want);
        }
        let object = predeclared_object(type_name)?;
        let route = predeclared_member(object, name)?;
        Some(Binding::new(None, route))
    }

    fn resolve_default_member(&self, recv: &VarTypeRef) -> Option<Binding> {
        let VarTypeRef::Object(type_name) = recv else {
            return None;
        };
        if is_collection_type(type_name) {
            let mut binding = collection_member("Item")?;
            binding.is_default = true;
            return Some(binding);
        }
        if is_err_object_type(type_name) {
            let mut binding = err_object_member("Number", Some(ProjectMemberKind::PropertyGet))?;
            binding.is_default = true;
            return Some(binding);
        }
        let member = if type_name.rsplit('.').next()?.eq_ignore_ascii_case("Match") {
            "Value"
        } else {
            "Item"
        };
        let mut binding = regexp_member(type_name, member, None)?;
        binding.is_default = true;
        Some(binding)
    }

    fn resolve_qualified(&self, parts: &[&str]) -> Option<Binding> {
        if parts.len() == 2 && fold_identifier(parts[0]) == "vba" {
            return self.resolve(parts[1]);
        }
        // `Math.Abs(x)`, `Strings.Left(s, 2)`: a VBA library module named without the `VBA.` prefix. A
        // project module of that name is found first (the project provider comes before this one).
        // `VbVarType.vbString`: a library constant named through its enum, without the `VBA.` prefix.
        if parts.len() == 2 && is_constant_owner(parts[0]) {
            let binding = self.resolve(parts[1])?;
            return matches!(binding.route, DispatchRoute::ConstValue(_)).then_some(binding);
        }
        if parts.len() == 2 && is_vba_library_module(parts[0]) {
            let owner = fold_identifier(parts[0]);
            let binding = self.resolve(parts[1])?;
            return binding_owner_matches(&owner, &binding).then_some(binding);
        }
        if parts.len() == 3 && fold_identifier(parts[0]) == "vba" {
            let owner = fold_identifier(parts[1]);
            let binding = self.resolve(parts[2])?;
            let matches_owner = match &binding.route {
                DispatchRoute::ConstValue(_) => is_constant_owner(&owner),
                _ => binding_owner_matches(&owner, &binding),
            };
            return matches_owner.then_some(binding);
        }
        None
    }

    fn resolve_coclass(&self, name: &str) -> Option<String> {
        let lower = fold_identifier(name);
        matches!(lower.as_str(), "regexp" | "vba.regexp").then(|| "VBScript.RegExp".to_string())
    }

    fn resolve_extern_coclass(&self, name: &str) -> Option<(String, String)> {
        // `New Collection` (or `New VBA.Collection`) instantiates the built-in
        // `VBA.Collection` class via the same cross-bundle `NewExtern` path a
        // referenced project's coclass uses.
        let folded = fold_identifier(name);
        let bare = folded.strip_prefix("vba.").unwrap_or(&folded);
        (bare == "collection").then(|| (VBA_UNIT.to_string(), "Collection".to_string()))
    }
}

/// The names of the VBA library's modules (`VBA.Math`, `VBA.Strings`, ...), usable as qualifiers.
fn is_vba_library_module(name: &str) -> bool {
    matches!(
        fold_identifier(name).as_str(),
        "math"
            | "strings"
            | "conversion"
            | "datetime"
            | "interaction"
            | "information"
            | "filesystem"
            | "financial"
            | "fileio"
            | "random"
            | "vbeglobal"
    )
}

fn binding_owner_matches(owner: &str, binding: &Binding) -> bool {
    use calcvba_bundle::{LibraryModule, NativeImplId};
    match &binding.route {
        DispatchRoute::Native(id) => match id {
            NativeImplId::UserForms | NativeImplId::Load | NativeImplId::Unload => {
                owner == "vbeglobal"
            }
            _ => match id.module() {
                LibraryModule::Interaction => owner == "interaction",
                LibraryModule::FileIo => owner == "filesystem" || owner == "fileio",
                _ => id
                    .library_member()
                    .is_some_and(|(module, _)| module.eq_ignore_ascii_case(owner)),
            },
        },
        DispatchRoute::ExternMember {
            owner: module,
            member,
            ..
        } => {
            module.eq_ignore_ascii_case(owner)
                || (module == "Random" && owner == "math")
                || matches!(
                    (owner, member.to_ascii_lowercase().as_str()),
                    ("conversion", "error" | "error$" | "fix" | "int")
                        | ("filesystem", "dir")
                        | ("strings", "monthname" | "weekdayname")
                )
        }
        DispatchRoute::PredeclaredObject(PredeclaredObjectId::Err) => owner == "information",
        DispatchRoute::SpecialForm(SpecialForm::Erl) => owner == "information",
        DispatchRoute::SpecialForm(
            SpecialForm::IIf | SpecialForm::Choose | SpecialForm::Switch | SpecialForm::CallByName,
        ) => owner == "interaction",
        _ => false,
    }
}

fn is_constant_owner(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "vbvartype"
            | "vbappwinstyle"
            | "vbcalendar"
            | "vbcalltype"
            | "vbcomparemethod"
            | "vbdatetimeformat"
            | "vbdayofweek"
            | "vbfileattribute"
            | "vbfirstweekofyear"
            | "vbimestatus"
            | "vbmsgboxresult"
            | "vbmsgboxstyle"
            | "vbqueryclose"
            | "vbstrconv"
            | "vbtristate"
            | "constants"
            | "colorconstants"
            | "formshowconstants"
            | "keycodeconstants"
            | "systemcolorconstants"
    )
}

fn is_collection_type(type_name: &str) -> bool {
    let folded = fold_identifier(type_name);
    let bare = folded.strip_prefix("vba.").unwrap_or(&folded);
    bare == "collection"
}

/// Resolve a `Collection` member to its cross-bundle [`DispatchRoute::ExternMember`]
/// route against the `VBA` unit. The accessor kinds must match the VBA bundle's
/// `ClassMethod` kinds (`Count` is a property-get; the rest are methods) so the
/// late-dispatch kind check matches; they are fixed by member name, not by the
/// caller's `want` hint.
fn collection_member(name: &str) -> Option<Binding> {
    let (member, kind) = match fold_identifier(name).as_str() {
        "add" => ("Add", ProjectMemberKind::Method),
        "item" => ("Item", ProjectMemberKind::Method),
        "count" => ("Count", ProjectMemberKind::PropertyGet),
        "remove" => ("Remove", ProjectMemberKind::Method),
        "_newenum" => ("_NewEnum", ProjectMemberKind::PropertyGet),
        _ => return None,
    };
    Some(Binding::new(
        None,
        DispatchRoute::ExternMember {
            unit: VBA_UNIT.to_string(),
            owner: "Collection".to_string(),
            member: member.to_string(),
            kind,
            param_types: Vec::new(),
            param_names: Vec::new(),
            param_optional: Vec::new(),
            param_optional_defaults: Vec::new(),
            variadic: false,
            has_receiver: true,
        },
    ))
}

/// VBA 7.1 regex class members use the native object's ordinary late dispatch.
fn regexp_member(type_name: &str, name: &str, want: Option<ProjectMemberKind>) -> Option<Binding> {
    let lower = fold_identifier(type_name);
    let class = lower.strip_prefix("vba.").unwrap_or(&lower);
    let member = fold_identifier(name);
    let (kind, params): (ProjectMemberKind, &[&str]) = match (class, member.as_str()) {
        ("regexp", "pattern" | "global" | "ignorecase" | "multiline") => {
            (want.unwrap_or(ProjectMemberKind::PropertyGet), &[])
        }
        ("regexp", "test" | "execute") => (ProjectMemberKind::Method, &["sourceString"]),
        ("regexp", "replace") => (ProjectMemberKind::Method, &["sourceString", "replaceVar"]),
        ("match", "value" | "firstindex" | "length" | "submatches") => {
            (ProjectMemberKind::PropertyGet, &[])
        }
        ("matchcollection" | "submatches", "count" | "_newenum") => {
            (ProjectMemberKind::PropertyGet, &[])
        }
        ("matchcollection" | "submatches", "item") => (ProjectMemberKind::PropertyGet, &["Index"]),
        _ => return None,
    };
    Some(Binding::new(
        None,
        DispatchRoute::ExternMember {
            unit: VBA_UNIT.to_string(),
            owner: type_name.to_string(),
            member: name.to_string(),
            kind,
            param_types: Vec::new(),
            param_names: params.iter().map(|p| (*p).to_string()).collect(),
            param_optional: vec![false; params.len()],
            param_optional_defaults: Vec::new(),
            variadic: false,
            has_receiver: true,
        },
    ))
}

fn is_err_object_type(type_name: &str) -> bool {
    type_name.eq_ignore_ascii_case("ErrObject") || type_name.eq_ignore_ascii_case("VBA.ErrObject")
}

fn err_object_member(name: &str, want: Option<ProjectMemberKind>) -> Option<Binding> {
    let (kind, params): (ProjectMemberKind, &[&str]) = match name.to_ascii_lowercase().as_str() {
        "raise" => (
            ProjectMemberKind::Method,
            &["Number", "Source", "Description", "HelpFile", "HelpContext"],
        ),
        "clear" => (ProjectMemberKind::Method, &[]),
        "number" | "description" | "source" | "helpfile" | "helpcontext" => {
            (want.unwrap_or(ProjectMemberKind::PropertyGet), &[])
        }
        "lastdllerror" => (ProjectMemberKind::PropertyGet, &[]),
        _ => return None,
    };
    Some(Binding::new(
        None,
        DispatchRoute::ExternMember {
            unit: VBA_UNIT.into(),
            owner: "ErrObject".into(),
            member: name.into(),
            kind,
            param_types: Vec::new(),
            param_names: params.iter().map(|p| (*p).into()).collect(),
            param_optional: (0..params.len()).map(|i| i > 0).collect(),
            param_optional_defaults: vec![None; params.len()],
            variadic: false,
            has_receiver: true,
        },
    ))
}

/// The user-facing structural pointer intrinsics (the rest of the structural set
/// is compiler-internal and not resolvable by source name).
fn user_structural_intrinsic(name: &str) -> Option<StructuralIntrinsic> {
    Some(match name.to_ascii_lowercase().as_str() {
        "varptr" => StructuralIntrinsic::VarPtr,
        "strptr" => StructuralIntrinsic::StrPtr,
        "objptr" => StructuralIntrinsic::ObjPtr,
        _ => return None,
    })
}

fn special_form(name: &str) -> Option<SpecialForm> {
    Some(match name.to_ascii_lowercase().as_str() {
        "array" => SpecialForm::Array,
        "iif" => SpecialForm::IIf,
        "choose" => SpecialForm::Choose,
        "switch" => SpecialForm::Switch,
        "callbyname" => SpecialForm::CallByName,
        "typeof" => SpecialForm::TypeOf,
        "addressof" => SpecialForm::AddressOf,
        "ubound" => SpecialForm::UBound,
        "lbound" => SpecialForm::LBound,
        "erl" => SpecialForm::Erl,
        _ => return None,
    })
}

fn intrinsic_param_names(entry: &crate::catalog::IntrinsicEntry) -> Vec<String> {
    if !entry.param_names.is_empty() {
        return entry
            .param_names
            .iter()
            .map(|name| (*name).to_string())
            .collect();
    }
    match entry.sig.max_args {
        Some(max) => vec![String::new(); usize::from(max)],
        None => Vec::new(),
    }
}

fn intrinsic_param_optional(entry: &crate::catalog::IntrinsicEntry) -> Vec<bool> {
    match entry.sig.max_args {
        Some(max) => (0..max).map(|index| index >= entry.sig.min_args).collect(),
        None => Vec::new(),
    }
}

fn intrinsic_param_optional_defaults(
    entry: &crate::catalog::IntrinsicEntry,
) -> Vec<Option<CoreConst>> {
    match entry.sig.max_args {
        Some(max) => {
            let mut defaults = vec![None; usize::from(max)];
            // Omitted Include is True; an explicitly supplied Empty is False.
            // Recording the declared default also preserves gaps before a named Compare.
            if entry.id == calcvba_bundle::NativeImplId::Filter {
                defaults[2] = Some(CoreConst::Bool(true));
            }
            defaults
        }
        None => Vec::new(),
    }
}

/// The predeclared-object lookup helper (also used by `build`).
pub fn library_predeclared_object(name: &str) -> Option<PredeclaredObjectId> {
    predeclared_object(name)
}

/// Resolve a `vb*` base-library value constant by case-insensitive name, as the
/// folded `CoreConst` literal the binder inlines. Ported verbatim from
/// `frontend_library`. Integer constants are `I32` (VBA `Long`/`Integer`), string
/// constants are `Str`; this is exactly what a referenced project's `Public Const`
/// folds to, so a `vb*` constant binds identically to a published one.
/// `vbNullString` is `NullStr` (a String with a null BSTR). `Nothing`/`Empty`/`Null`
/// are absent: they are structural intrinsics, not value constants.
pub fn library_constant(name: &str) -> Option<CoreConst> {
    // `int(i32)` → `CoreConst::I32` (VBA `Long`); `str(&str)` → `CoreConst::Str`.
    let int = CoreConst::I32;
    let str = |s: &str| CoreConst::Str(s.to_string());
    Some(match name.to_ascii_lowercase().as_str() {
        "vbcr" => str("\r"),
        "vblf" => str("\n"),
        "vbcrlf" | "vbnewline" => str("\r\n"),
        "vbtab" => str("\t"),
        // VbCallType (for CallByName's third argument).
        "vbmethod" => int(1),
        "vbget" => int(2),
        "vblet" => int(4),
        "vbset" => int(8),
        "vbnullchar" => str("\0"),
        // `""` in every operation, but its BSTR is null: `StrPtr` 0, `NULL` to a `Declare`.
        "vbnullstring" => CoreConst::NullStr,
        "vbback" => str("\u{0008}"),
        "vbformfeed" => str("\u{000C}"),
        "vbverticaltab" => str("\u{000B}"),
        "vbbinarycompare" => int(0),
        "vbtextcompare" => int(1),
        "vbdatabasecompare" => int(2),
        "vbuppercase" => int(1),
        "vblowercase" => int(2),
        "vbpropercase" => int(3),
        "vbwide" => int(4),
        "vbnarrow" => int(8),
        "vbkatakana" => int(16),
        "vbhiragana" => int(32),
        "vbunicode" => int(64),
        "vbfromunicode" => int(128),
        "vbempty" => int(0),
        "vbnull" => int(1),
        "vbinteger" => int(2),
        "vblong" => int(3),
        "vbsingle" => int(4),
        "vbdouble" => int(5),
        "vbcurrency" => int(6),
        "vbdate" => int(7),
        "vbstring" => int(8),
        "vbobject" => int(9),
        "vberror" => int(10),
        "vbboolean" => int(11),
        "vbvariant" => int(12),
        "vbdataobject" => int(13),
        "vbdecimal" => int(14),
        "vbbyte" => int(17),
        "vblonglong" => int(20),
        "vbuserdefinedtype" => int(36),
        "vbarray" => int(8192),
        "vbtrue" => int(-1),
        "vbfalse" => int(0),
        "vbusedefault" => int(-2),
        "vbokonly" => int(0),
        "vbokcancel" => int(1),
        "vbabortretryignore" => int(2),
        "vbyesnocancel" => int(3),
        "vbyesno" => int(4),
        "vbretrycancel" => int(5),
        "vbcritical" => int(16),
        "vbquestion" => int(32),
        "vbexclamation" => int(48),
        "vbinformation" => int(64),
        "vbdefaultbutton1" => int(0),
        "vbdefaultbutton2" => int(256),
        "vbdefaultbutton3" => int(512),
        "vbdefaultbutton4" => int(768),
        "vbapplicationmodal" => int(0),
        "vbsystemmodal" => int(4096),
        "vbmodeless" => int(0),
        "vbmodal" => int(1),
        "vbmsgboxhelpbutton" => int(16384),
        "vbmsgboxsetforeground" => int(65536),
        "vbmsgboxright" => int(524288),
        "vbmsgboxrtlreading" => int(1048576),
        "vbok" => int(1),
        "vbcancel" => int(2),
        "vbabort" => int(3),
        "vbretry" => int(4),
        "vbignore" => int(5),
        "vbyes" => int(6),
        "vbno" => int(7),
        "vbblack" => int(0),
        "vbred" => int(255),
        "vbgreen" => int(65280),
        "vbyellow" => int(65535),
        "vbblue" => int(16711680),
        "vbmagenta" => int(16711935),
        "vbcyan" => int(16776960),
        "vbwhite" => int(16777215),
        "vbgeneraldate" => int(0),
        "vblongdate" => int(1),
        "vbshortdate" => int(2),
        "vblongtime" => int(3),
        "vbshorttime" => int(4),
        "vbusesystemdayofweek" => int(0),
        "vbsunday" => int(1),
        "vbmonday" => int(2),
        "vbtuesday" => int(3),
        "vbwednesday" => int(4),
        "vbthursday" => int(5),
        "vbfriday" => int(6),
        "vbsaturday" => int(7),
        "vbusesystem" => int(0),
        "vbfirstjan1" => int(1),
        "vbfirstfourdays" => int(2),
        "vbfirstfullweek" => int(3),
        "vbnormal" => int(0),
        "vbreadonly" => int(1),
        "vbhidden" => int(2),
        "vbsystem" => int(4),
        "vbvolume" => int(8),
        "vbdirectory" => int(16),
        "vbarchive" => int(32),
        "vbobjecterror" => int(-2147221504),
        // VbAppWinStyle (the `Shell` window styles).
        "vbhide" => int(0),
        "vbnormalfocus" => int(1),
        "vbminimizedfocus" => int(2),
        "vbmaximizedfocus" => int(3),
        "vbnormalnofocus" => int(4),
        "vbminimizednofocus" => int(6),
        // VbCalendar.
        "vbcalgreg" => int(0),
        "vbcalhijri" => int(1),
        // VbFileAttribute's file alias.
        "vbalias" => int(64),
        // VbQueryClose.
        "vbformcontrolmenu" => int(0),
        "vbformcode" => int(1),
        "vbappwindows" => int(2),
        "vbapptaskmanager" => int(3),
        "vbformmdiform" => int(4),
        // VbIMEStatus.
        "vbimenoop" | "vbimemodenocontrol" => int(0),
        "vbimeon" | "vbimemodeon" => int(1),
        "vbimeoff" | "vbimemodeoff" => int(2),
        "vbimemodedisable" | "vbimedisable" => int(3),
        "vbimemodehiragana" | "vbimehiragana" => int(4),
        "vbimemodekatakana" | "vbimekatakanadbl" => int(5),
        "vbimemodekatakanahalf" | "vbimekatakanasng" => int(6),
        "vbimemodealphafull" | "vbimealphadbl" => int(7),
        "vbimemodealpha" | "vbimealphasng" => int(8),
        "vbimemodehangulfull" => int(9),
        "vbimemodehangul" => int(10),
        "vbmsgbox" => int(-2147483625),
        "vbmsgboxtext" => int(-2147483624),
        // SystemColorConstants.
        "vbscrollbars" => int(-2147483648),
        "vbdesktop" => int(-2147483647),
        "vbactivetitlebar" => int(-2147483646),
        "vbinactivetitlebar" => int(-2147483645),
        "vbmenubar" => int(-2147483644),
        "vbwindowbackground" => int(-2147483643),
        "vbwindowframe" => int(-2147483642),
        "vbmenutext" => int(-2147483641),
        "vbwindowtext" => int(-2147483640),
        "vbtitlebartext" => int(-2147483639),
        "vbactiveborder" => int(-2147483638),
        "vbinactiveborder" => int(-2147483637),
        "vbapplicationworkspace" => int(-2147483636),
        "vbhighlight" => int(-2147483635),
        "vbhighlighttext" => int(-2147483634),
        "vbbuttonface" | "vb3dface" => int(-2147483633),
        "vbbuttonshadow" | "vb3dshadow" => int(-2147483632),
        "vbgraytext" => int(-2147483631),
        "vbbuttontext" => int(-2147483630),
        "vbinactivecaptiontext" => int(-2147483629),
        "vb3dhighlight" => int(-2147483628),
        "vb3ddkshadow" => int(-2147483627),
        "vb3dlight" => int(-2147483626),
        "vbinfotext" => int(-2147483625),
        "vbinfobackground" => int(-2147483624),
        other => return key_code_constant(other).map(CoreConst::I16),
    })
}

/// `KeyCodeConstants`: the `vbKey*` virtual-key codes, which are `Integer` (not `Long`).
fn key_code_constant(lower: &str) -> Option<i16> {
    let key = lower.strip_prefix("vbkey")?;
    let named = match key {
        "lbutton" => 1,
        "rbutton" => 2,
        "cancel" => 3,
        "mbutton" => 4,
        "back" => 8,
        "tab" => 9,
        "clear" => 12,
        "return" => 13,
        "shift" => 16,
        "control" => 17,
        "menu" => 18,
        "pause" => 19,
        "capital" => 20,
        "escape" => 27,
        "space" => 32,
        "pageup" => 33,
        "pagedown" => 34,
        "end" => 35,
        "home" => 36,
        "left" => 37,
        "up" => 38,
        "right" => 39,
        "down" => 40,
        "select" => 41,
        "print" => 42,
        "execute" => 43,
        "snapshot" => 44,
        "insert" => 45,
        "delete" => 46,
        "help" => 47,
        "numlock" => 144,
        "multiply" => 106,
        "add" => 107,
        "separator" => 108,
        "subtract" => 109,
        "decimal" => 110,
        "divide" => 111,
        _ => {
            // A single letter or digit is its character code; `numpadN` is 96 + N; `fN` is 111 + N.
            let bytes = key.as_bytes();
            return match bytes {
                [c @ b'a'..=b'z'] => Some(i16::from(c.to_ascii_uppercase())),
                [c @ b'0'..=b'9'] => Some(i16::from(*c)),
                _ => {
                    if let Some(n) = key
                        .strip_prefix("numpad")
                        .and_then(|n| n.parse::<i16>().ok())
                    {
                        (0..=9).contains(&n).then_some(96 + n)
                    } else {
                        key.strip_prefix('f')
                            .and_then(|n| n.parse::<i16>().ok())
                            .filter(|n| (1..=16).contains(n))
                            .map(|n| 111 + n)
                    }
                }
            };
        }
    };
    Some(named)
}
