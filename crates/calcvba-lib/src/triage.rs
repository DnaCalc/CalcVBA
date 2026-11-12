//! Predicted capabilities of a VBA program, from bound OxIR (or, when a project does not bind,
//! from a lexical scan of its source).
//!
//! [`capabilities_of`] is the one table from a library built-in to the capabilities its host
//! call needs. It is an exhaustive `match` with no wildcard, so a new `NativeImplId` fails to
//! compile until someone classifies it. The runtime policy (`calcvba_hal::PolicyHost`) gates the
//! same operations; this table is how triage names them before anything runs.
//!
//! The prediction is a static over-approximation: it lists what the code mentions, not what a
//! run reaches. Compare it with the audit log as the design record says: observed minus
//! predicted is a triage bug, predicted minus observed is only unconfirmed.

use std::collections::{BTreeMap, BTreeSet};

use calcvba_bundle::NativeImplId;
use calcvba_hal::Capability;
use calcvba_oxir::OxProgram;
use calcvba_oxir::inst::{OxAsNew, OxInst};
use calcvba_oxir::value::{OxArg, OxCallArg, OxConst, OxNativeCallee, OxOperand};

/// The capabilities one built-in needs. `FileOpen` can be either, until its mode is known.
pub const fn capabilities_of(id: NativeImplId) -> &'static [Capability] {
    use Capability as C;
    use NativeImplId::*;
    match id {
        // fs.read
        FileLen | FileCurDir | FileGetAttr | FileDateTime | Dir => &[C::FsRead],
        // fs.write
        FileKill | FileMkDir | FileRmDir | FileChDir | FileCopy | FileSetAttr | FileChDrive
        | FileRename => &[C::FsWrite],
        FileOpen => &[C::FsRead, C::FsWrite],
        Shell => &[C::Process],
        Environ | Command => &[C::Env],
        GetSetting | GetAllSettings | SaveSetting | DeleteSetting => &[C::Settings],
        SendKeys | AppActivate => &[C::UiKeys],
        CreateObject | GetObject => &[C::ComActivate],
        // Not gated: pure functions, handle operations after an open, console, UI prompts,
        // time, diagnostics and the events of an already activated object.
        Len
        | LenB
        | Left
        | LeftB
        | Right
        | RightB
        | Mid
        | MidStmt
        | LSetStmt
        | RSetStmt
        | InStr
        | InStrRev
        | LCase
        | UCase
        | Split
        | Join
        | Replace
        | Trim
        | LTrim
        | RTrim
        | StrComp
        | Like
        | Chr
        | Asc
        | ChrW
        | AscW
        | Space
        | StringRepeat
        | StrReverse
        | StrConv
        | Format
        | FormatNumber
        | FormatCurrency
        | FormatPercent
        | FormatDateTime
        | Filter
        | Abs
        | Int
        | Fix
        | Sgn
        | Round
        | Sqr
        | Sin
        | Cos
        | Log
        | Exp
        | Atn
        | Tan
        | DateSerial
        | TimeSerial
        | DateValue
        | TimeValue
        | DateAdd
        | DateDiff
        | Year
        | Month
        | Day
        | Weekday
        | Hour
        | Minute
        | Second
        | MonthName
        | WeekdayName
        | DatePart
        | DateNow
        | TimeNow
        | Now
        | Timer
        | Hex
        | Oct
        | CStr
        | Str
        | Val
        | CDate
        | CVErr
        | CBool
        | CByte
        | CInt
        | CLng
        | CLngLng
        | CLngPtr
        | CSng
        | CDbl
        | CCur
        | CDec
        | CVar
        | Rnd
        | Randomize
        | Fv
        | Pv
        | Pmt
        | Npv
        | Irr
        | Mirr
        | Rate
        | NPer
        | IPmt
        | PPmt
        | Sln
        | Syd
        | Ddb
        | IsArray
        | VarType
        | TypeName
        | IsNumeric
        | IsError
        | IsDate
        | IsObject
        | IsNull
        | IsEmpty
        | IsMissing
        | ErrorText
        | IIf
        | Choose
        | Switch
        | Rgb
        | QbColor
        | FreeFile
        | FileClose
        | FileRead
        | FileWrite
        | FilePrint
        | ConsolePrint
        | FileInput
        | ConsoleInput
        | FileLineInput
        | ConsoleLineInput
        | FileEof
        | FileLof
        | FileSeek
        | FileLoc
        | FilePut
        | FileGetInto
        | FileWidth
        | FileLock
        | FileUnlock
        | Partition
        | MsgBox
        | InputBox
        | Beep
        | DoEvents
        | ComSubscribeEvent
        | ComUnsubscribeEvent
        | ComEventCallbackSubscription
        | ComEventCallbackArg
        | ComReleaseEventCallback
        | DebugPrint
        | FileInputChars => &[],
    }
}

/// How a prediction was made.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PredictionSource {
    /// From bound OxIR.
    Bound,
    /// From a scan of the source text, when the project did not bind.
    Lexical,
}

impl PredictionSource {
    pub const fn name(self) -> &'static str {
        match self {
            PredictionSource::Bound => "bound",
            PredictionSource::Lexical => "lexical",
        }
    }
}

/// The predicted capabilities, with the details (ProgIDs, `lib!symbol`, mentioned operations).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Prediction {
    pub source: PredictionSource,
    pub capabilities: BTreeSet<Capability>,
    /// Per capability: ProgIDs, `lib!symbol` names or operation names.
    pub details: BTreeMap<Capability, BTreeSet<String>>,
    /// The code dispatches by a runtime name or procedure reference (`CallByName`,
    /// `CallProcRef`, late-bound COM): a marker, not a capability.
    pub dynamic: bool,
    /// A COM activation whose ProgID is not a constant.
    pub com_progid_unknown: bool,
    /// How many procedures the prediction looked at (zero for a lexical scan).
    pub procedures: usize,
}

impl Prediction {
    fn new(source: PredictionSource) -> Self {
        Self {
            source,
            capabilities: BTreeSet::new(),
            details: BTreeMap::new(),
            dynamic: false,
            com_progid_unknown: false,
            procedures: 0,
        }
    }

    fn add(&mut self, capability: Capability, detail: impl Into<String>) {
        self.capabilities.insert(capability);
        self.details
            .entry(capability)
            .or_default()
            .insert(detail.into());
    }

    /// Fold another prediction (another program of the same project) into this one.
    pub fn merge(&mut self, other: &Prediction) {
        self.capabilities.extend(other.capabilities.iter().copied());
        for (cap, details) in &other.details {
            self.details
                .entry(*cap)
                .or_default()
                .extend(details.iter().cloned());
        }
        self.dynamic |= other.dynamic;
        self.com_progid_unknown |= other.com_progid_unknown;
        self.procedures += other.procedures;
        if other.source == PredictionSource::Lexical {
            self.source = PredictionSource::Lexical;
        }
    }
}

/// One argument of a built-in call, reduced to what triage can read.
enum Arg<'a> {
    Operand(&'a OxOperand),
    Const(i32),
    Other,
}

fn call_args(args: &[OxCallArg]) -> Vec<Arg<'_>> {
    args.iter()
        .map(|arg| match arg {
            OxCallArg::Operand(op) => Arg::Operand(op),
            OxCallArg::Const(n) => Arg::Const(*n),
            _ => Arg::Other,
        })
        .collect()
}

fn proc_args(args: &[OxArg]) -> Vec<Arg<'_>> {
    args.iter()
        .map(|arg| match arg {
            OxArg::ByVal(op) => Arg::Operand(op),
            _ => Arg::Other,
        })
        .collect()
}

fn const_str<'a>(args: &[Arg<'a>], index: usize) -> Option<&'a str> {
    match args.get(index) {
        Some(Arg::Operand(OxOperand::Const(OxConst::Str(text)))) => Some(text),
        _ => None,
    }
}

/// A constant `Open` mode: `Some(false)` is `For Input`, `Some(true)` can write.
fn open_mode(args: &[Arg<'_>], index: usize) -> Option<bool> {
    let code = match args.get(index) {
        Some(Arg::Operand(OxOperand::Const(OxConst::I32(n)))) => i64::from(*n),
        Some(Arg::Operand(OxOperand::Const(OxConst::I64(n)))) => *n,
        Some(Arg::Operand(OxOperand::Const(OxConst::I16(n)))) => i64::from(*n),
        Some(Arg::Const(n)) => i64::from(*n),
        _ => return None,
    };
    Some(code & 0xFFFF != 0)
}

fn note_builtin(prediction: &mut Prediction, id: NativeImplId, args: &[Arg<'_>]) {
    match id {
        NativeImplId::FileOpen => match open_mode(args, 1) {
            Some(true) => prediction.add(Capability::FsWrite, "Open"),
            Some(false) => prediction.add(Capability::FsRead, "Open"),
            None => {
                prediction.add(Capability::FsRead, "Open");
                prediction.add(Capability::FsWrite, "Open");
            }
        },
        NativeImplId::CreateObject | NativeImplId::GetObject => {
            // CreateObject(progid); GetObject(path, class).
            let progid = if id == NativeImplId::CreateObject {
                const_str(args, 0)
            } else {
                const_str(args, 1).or_else(|| const_str(args, 0))
            };
            match progid {
                Some(text) if !text.is_empty() => prediction.add(Capability::ComActivate, text),
                _ => {
                    prediction.com_progid_unknown = true;
                    prediction.add(Capability::ComActivate, "<dynamic>");
                }
            }
        }
        _ => {
            for cap in capabilities_of(id) {
                prediction.add(*cap, format!("{id:?}"));
            }
        }
    }
}

/// Predict the capabilities of one bound program.
pub fn predicted_capabilities(program: &OxProgram) -> Prediction {
    let mut prediction = Prediction::new(PredictionSource::Bound);
    prediction.procedures = program.funcs.len();
    let note_as_new = |binding: &OxAsNew, prediction: &mut Prediction| {
        if let OxAsNew::ComClass { prog_id } = binding {
            prediction.add(Capability::ComActivate, prog_id.clone());
        }
    };
    let library = calcvba_bundle::vba_library_bundle();
    for class in &program.classes {
        for field in &class.as_new_fields {
            note_as_new(&field.binding, &mut prediction);
        }
    }
    for func in &program.funcs {
        for block in &func.blocks {
            for inst in &block.instrs {
                match inst {
                    OxInst::AsNew { binding, .. } => note_as_new(binding, &mut prediction),
                    OxInst::CallNative { callee, args, .. } => match callee {
                        OxNativeCallee::Builtin(id) => {
                            note_builtin(&mut prediction, *id, &call_args(args))
                        }
                        OxNativeCallee::Declare { descriptor_id, .. } => {
                            let detail = program
                                .external_calls
                                .iter()
                                .find(|call| call.descriptor_id == *descriptor_id)
                                .map(|call| format!("{}!{}", call.library, call.alias))
                                .unwrap_or_else(|| format!("descriptor#{descriptor_id}"));
                            prediction.add(Capability::Ffi, detail);
                        }
                    },
                    // A migrated library built-in is a cross-bundle call into the `VBA` unit.
                    OxInst::CallExtern { import, args, .. } => {
                        let Some(imp) = program.imports.get(import.0) else {
                            continue;
                        };
                        if !imp.unit.eq_ignore_ascii_case("VBA") {
                            continue;
                        }
                        let id = library
                            .exports
                            .iter()
                            .find(|e| e.token.matches(&imp.token))
                            .and_then(|export| match export.target {
                                calcvba_bundle::ExportTarget::Proc(p) => {
                                    library.procedures.get(p).and_then(|proc| proc.native)
                                }
                                _ => None,
                            });
                        if let Some(calcvba_bundle::NativeBody::Library(id)) = id {
                            note_builtin(&mut prediction, id, &proc_args(args));
                        }
                    }
                    OxInst::ComCallLate { .. }
                    | OxInst::CallByName { .. }
                    | OxInst::CallProcRef { .. } => prediction.dynamic = true,
                    _ => {}
                }
            }
        }
    }
    prediction
}

// ── Lexical fallback ─────────────────────────────────────────────────────────────────────────

/// Words that name a gated operation when they start a statement or appear as a call.
/// Matching is on the identifier alone, case-insensitive; an identifier after a `.` is a member
/// of some object, not the library function, and is skipped.
fn word_capability(word: &str) -> Option<Capability> {
    use Capability as C;
    Some(match word {
        "kill" | "mkdir" | "rmdir" | "filecopy" | "setattr" | "chdir" | "chdrive" => C::FsWrite,
        "dir" | "filelen" | "getattr" | "filedatetime" | "curdir" => C::FsRead,
        "shell" => C::Process,
        "environ" | "command" => C::Env,
        "getsetting" | "getallsettings" | "savesetting" | "deletesetting" => C::Settings,
        "sendkeys" | "appactivate" => C::UiKeys,
        "createobject" | "getobject" => C::ComActivate,
        _ => return None,
    })
}

#[derive(Debug, PartialEq)]
enum Tok {
    Word(String),
    Str(String),
    Dot,
    Colon,
    Other,
    Newline,
}

fn tokenize(source: &str) -> Vec<Tok> {
    let chars: Vec<char> = source.chars().collect();
    let mut out = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '\n' {
            out.push(Tok::Newline);
            i += 1;
        } else if c == '\'' {
            while i < chars.len() && chars[i] != '\n' {
                i += 1;
            }
        } else if c == '"' {
            let mut text = String::new();
            i += 1;
            while i < chars.len() {
                if chars[i] == '"' {
                    if chars.get(i + 1) == Some(&'"') {
                        text.push('"');
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                if chars[i] == '\n' {
                    break;
                }
                text.push(chars[i]);
                i += 1;
            }
            out.push(Tok::Str(text));
        } else if c.is_alphabetic() || c == '_' {
            let start = i;
            while i < chars.len() && (chars[i].is_alphanumeric() || chars[i] == '_') {
                i += 1;
            }
            let mut word: String = chars[start..i].iter().collect();
            // `Environ$`, `Dir$` and friends.
            if i < chars.len() && chars[i] == '$' {
                i += 1;
            }
            word.make_ascii_lowercase();
            if word == "rem" {
                while i < chars.len() && chars[i] != '\n' {
                    i += 1;
                }
            } else {
                out.push(Tok::Word(word));
            }
        } else if c == ':' {
            out.push(Tok::Colon);
            i += 1;
        } else if c == '.' {
            out.push(Tok::Dot);
            i += 1;
        } else {
            out.push(Tok::Other);
            i += 1;
        }
    }
    out
}

/// Scan source text for `Declare` statements and the gated built-ins by name.
pub fn predicted_capabilities_lexical(sources: &[&str]) -> Prediction {
    let mut prediction = Prediction::new(PredictionSource::Lexical);
    for source in sources {
        let toks = tokenize(source);
        let mut at_start = true;
        let mut i = 0;
        while i < toks.len() {
            let after_dot = i > 0 && toks[i - 1] == Tok::Dot;
            let vba_qualified = after_dot && {
                let mut start = i;
                while start >= 2
                    && toks[start - 1] == Tok::Dot
                    && matches!(toks[start - 2], Tok::Word(_))
                {
                    start -= 2;
                }
                matches!(&toks[start], Tok::Word(w) if w == "vba")
            };
            if let Tok::Word(word) = &toks[i]
                && (!after_dot || vba_qualified)
            {
                {
                    if word == "new" {
                        // Both `Set x = New Class` and deferred `As New Class`.
                        // Project classes may add false positives; fallback is conservative.
                        prediction.add(Capability::ComActivate, "New");
                    } else if word == "declare" {
                        // Declare [PtrSafe] Function|Sub name Lib "lib" [Alias "sym"]
                        let line_end = toks[i..]
                            .iter()
                            .position(|t| matches!(t, Tok::Newline | Tok::Colon))
                            .map_or(toks.len(), |p| i + p);
                        let strings: Vec<&str> = toks[i..line_end]
                            .iter()
                            .filter_map(|t| match t {
                                Tok::Str(s) => Some(s.as_str()),
                                _ => None,
                            })
                            .collect();
                        let name = toks[i..line_end].iter().find_map(|t| match t {
                            Tok::Word(w)
                                if !matches!(
                                    w.as_str(),
                                    "declare"
                                        | "ptrsafe"
                                        | "function"
                                        | "sub"
                                        | "private"
                                        | "public"
                                ) =>
                            {
                                Some(w.as_str())
                            }
                            _ => None,
                        });
                        let lib = strings.first().copied().unwrap_or("?");
                        let sym = strings.get(1).copied().or(name).unwrap_or("?");
                        prediction.add(Capability::Ffi, format!("{lib}!{sym}"));
                    } else if word == "open" && at_start {
                        let line_end = toks[i..]
                            .iter()
                            .position(|t| matches!(t, Tok::Newline | Tok::Colon))
                            .map_or(toks.len(), |p| i + p);
                        let mut cap = Capability::FsWrite;
                        let words: Vec<&str> = toks[i..line_end]
                            .iter()
                            .filter_map(|t| match t {
                                Tok::Word(w) => Some(w.as_str()),
                                _ => None,
                            })
                            .collect();
                        for w in words.windows(2) {
                            if w[0] == "for" && w[1] == "input" {
                                cap = Capability::FsRead;
                            }
                        }
                        prediction.add(cap, "Open");
                    } else if word == "name" && at_start {
                        // `Name a As b` as a statement; a variable called Name is not followed by As.
                        let line_end = toks[i..]
                            .iter()
                            .position(|t| matches!(t, Tok::Newline | Tok::Colon))
                            .map_or(toks.len(), |p| i + p);
                        if toks[i..line_end].contains(&Tok::Word("as".into())) {
                            prediction.add(Capability::FsWrite, "Name");
                        }
                    } else if let Some(cap) = word_capability(word) {
                        let detail = word.clone();
                        prediction.add(cap, detail);
                    }
                }
            }
            at_start = matches!(toks[i], Tok::Newline | Tok::Colon)
                || (matches!(&toks[i], Tok::Other) && at_start)
                || matches!(&toks[i], Tok::Word(w) if w == "then" || w == "else");
            i += 1;
        }
    }
    prediction
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lexical_scan_finds_declares_and_builtins() {
        let source = "Declare PtrSafe Function GetTick Lib \"kernel32\" Alias \"GetTickCount\" () As Long\n\
            Sub A()\n  Open \"a\" For Input As #1\n  x = Environ(\"HOME\") ' Shell here is a comment\n  Set o = CreateObject(\"X.Y\")\n  rng.Shell\nEnd Sub\n";
        let p = predicted_capabilities_lexical(&[source]);
        let caps: Vec<_> = p.capabilities.iter().map(|c| c.name()).collect();
        assert_eq!(caps, ["fs.read", "env", "ffi", "com.activate"]);
        assert!(p.details[&Capability::Ffi].contains("kernel32!GetTickCount"));
    }

    #[test]
    fn lexical_scan_tracks_qualified_builtins_new_and_statement_boundaries() {
        for (source, cap) in [
            ("x = VBA.Interaction.Shell(\"x\")", Capability::Process),
            (
                "Dim x As New WinHttp.WinHttpRequest",
                Capability::ComActivate,
            ),
            ("x = 1: Open \"a\" For Input As #1", Capability::FsRead),
            (
                "If True Then Open \"a\" For Output As #1",
                Capability::FsWrite,
            ),
        ] {
            let p = predicted_capabilities_lexical(&[source]);
            assert!(p.capabilities.contains(&cap), "{source}");
        }
        assert!(
            predicted_capabilities_lexical(&["x.Shell \"x\" ' VBA.Interaction.Shell"])
                .capabilities
                .is_empty()
        );
    }

    #[test]
    fn every_capability_has_a_builtin_or_a_special_form() {
        // The table is exhaustive by construction; spot-check the classification.
        assert_eq!(capabilities_of(NativeImplId::Shell), &[Capability::Process]);
        assert!(capabilities_of(NativeImplId::Len).is_empty());
    }
}
