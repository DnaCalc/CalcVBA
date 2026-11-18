//! The exhaustive intrinsic catalog: every `NativeImplId` gets a name set, an
//! arity-shaped signature, and a [`CallShape`] (the quirky-syntax classification).
//!
//! `intrinsic_entry` is a `match` over `NativeImplId` **with no `_` arm** — adding
//! a variant to `calcvba-bundle` fails to compile here until it is classified. That
//! is the no-partial guarantee.

use calcvba_bundle::NativeImplId;

use crate::signature::CallShape;

/// An arity-shaped intrinsic signature. Intrinsics are Variant-flexible, so the
/// binding-relevant facts are arity + variadic + call shape (full per-param types
/// are not needed to bind a library call).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct IntrinsicSig {
    pub min_args: u8,
    /// `None` = no fixed upper bound (variadic / `ParamArray`-like).
    pub max_args: Option<u8>,
    pub param_array: bool,
}

#[derive(Debug, Clone, Copy)]
pub struct IntrinsicEntry {
    pub id: NativeImplId,
    /// User-facing names (with `$` variants). Empty = not resolvable by bare name
    /// (statement-form, predeclared-member, or internal machinery).
    pub names: &'static [&'static str],
    pub sig: IntrinsicSig,
    pub param_names: &'static [&'static str],
    pub call_shape: CallShape,
}

const fn sig(min: u8, max: u8) -> IntrinsicSig {
    IntrinsicSig {
        min_args: min,
        max_args: Some(max),
        param_array: false,
    }
}
const fn variadic(min: u8) -> IntrinsicSig {
    IntrinsicSig {
        min_args: min,
        max_args: None,
        param_array: true,
    }
}
const fn e(
    id: NativeImplId,
    names: &'static [&'static str],
    sig: IntrinsicSig,
    call_shape: CallShape,
) -> IntrinsicEntry {
    IntrinsicEntry {
        id,
        names,
        sig,
        param_names: &[],
        call_shape,
    }
}
const fn e_params(
    id: NativeImplId,
    names: &'static [&'static str],
    sig: IntrinsicSig,
    param_names: &'static [&'static str],
    call_shape: CallShape,
) -> IntrinsicEntry {
    IntrinsicEntry {
        id,
        names,
        sig,
        param_names,
        call_shape,
    }
}

/// Total classification of every intrinsic. No wildcard arm — exhaustiveness is
/// the coverage guarantee.
pub const fn intrinsic_entry(id: NativeImplId) -> IntrinsicEntry {
    use CallShape::{AssignmentTarget, DebugMember, FileStatement, Ordinary, SpecialForm};
    use NativeImplId::*;
    match id {
        // ── Strings ──
        Len => e_params(Len, &["Len"], sig(1, 1), &["Expression"], Ordinary),
        LenB => e_params(LenB, &["LenB"], sig(1, 1), &["Expression"], Ordinary),
        Left => e_params(
            Left,
            &["Left", "Left$"],
            sig(2, 2),
            &["String", "Length"],
            Ordinary,
        ),
        LeftB => e_params(
            LeftB,
            &["LeftB", "LeftB$"],
            sig(2, 2),
            &["String", "Length"],
            Ordinary,
        ),
        AscB => e_params(AscB, &["AscB"], sig(1, 1), &["String"], Ordinary),
        ChrB => e_params(ChrB, &["ChrB", "ChrB$"], sig(1, 1), &["CharCode"], Ordinary),
        InStrB => e(InStrB, &["InStrB"], sig(2, 4), Ordinary),
        MidB => e_params(
            MidB,
            &["MidB", "MidB$"],
            sig(2, 3),
            &["String", "Start", "Length"],
            Ordinary,
        ),
        Right => e_params(
            Right,
            &["Right", "Right$"],
            sig(2, 2),
            &["String", "Length"],
            Ordinary,
        ),
        RightB => e_params(
            RightB,
            &["RightB", "RightB$"],
            sig(2, 2),
            &["String", "Length"],
            Ordinary,
        ),
        Mid => e_params(
            Mid,
            &["Mid", "Mid$"],
            sig(2, 3),
            &["String", "Start", "Length"],
            Ordinary,
        ),
        MidStmt => e(MidStmt, &[], sig(2, 3), AssignmentTarget),
        MidBStmt => e(MidBStmt, &[], sig(2, 3), AssignmentTarget),
        LSetStmt => e(LSetStmt, &[], sig(2, 2), AssignmentTarget),
        RSetStmt => e(RSetStmt, &[], sig(2, 2), AssignmentTarget),
        InStr => e(InStr, &["InStr"], sig(2, 4), Ordinary),
        InStrRev => e_params(
            InStrRev,
            &["InStrRev"],
            sig(2, 4),
            &["StringCheck", "StringMatch", "Start", "Compare"],
            Ordinary,
        ),
        LCase => e_params(
            LCase,
            &["LCase", "LCase$"],
            sig(1, 1),
            &["String"],
            Ordinary,
        ),
        UCase => e_params(
            UCase,
            &["UCase", "UCase$"],
            sig(1, 1),
            &["String"],
            Ordinary,
        ),
        Split => e_params(
            Split,
            &["Split"],
            sig(1, 4),
            &["Expression", "Delimiter", "Limit", "Compare"],
            Ordinary,
        ),
        Join => e_params(
            Join,
            &["Join"],
            sig(1, 2),
            &["SourceArray", "Delimiter"],
            Ordinary,
        ),
        Replace => e_params(
            Replace,
            &["Replace"],
            sig(3, 6),
            &["Expression", "Find", "Replace", "Start", "Count", "Compare"],
            Ordinary,
        ),
        Trim => e_params(Trim, &["Trim", "Trim$"], sig(1, 1), &["String"], Ordinary),
        LTrim => e_params(
            LTrim,
            &["LTrim", "LTrim$"],
            sig(1, 1),
            &["String"],
            Ordinary,
        ),
        RTrim => e_params(
            RTrim,
            &["RTrim", "RTrim$"],
            sig(1, 1),
            &["String"],
            Ordinary,
        ),
        StrComp => e(StrComp, &["StrComp"], sig(2, 3), Ordinary),
        Like => e(Like, &[], sig(2, 2), Ordinary), // the `Like` operator
        Chr => e_params(Chr, &["Chr", "Chr$"], sig(1, 1), &["CharCode"], Ordinary),
        Asc => e_params(Asc, &["Asc"], sig(1, 1), &["String"], Ordinary),
        ChrW => e_params(ChrW, &["ChrW", "ChrW$"], sig(1, 1), &["CharCode"], Ordinary),
        AscW => e_params(AscW, &["AscW"], sig(1, 1), &["String"], Ordinary),
        Space => e_params(
            Space,
            &["Space", "Space$"],
            sig(1, 1),
            &["Number"],
            Ordinary,
        ),
        StringRepeat => e_params(
            StringRepeat,
            &["String", "String$"],
            sig(2, 2),
            &["Number", "Character"],
            Ordinary,
        ),
        StrReverse => e_params(
            StrReverse,
            &["StrReverse"],
            sig(1, 1),
            &["Expression"],
            Ordinary,
        ),
        StrConv => e_params(
            StrConv,
            &["StrConv"],
            sig(2, 3),
            &["String", "Conversion", "LocaleID"],
            Ordinary,
        ),
        Format => e_params(
            Format,
            &["Format", "Format$"],
            sig(1, 4),
            &["Expression", "Format", "FirstDayOfWeek", "FirstWeekOfYear"],
            Ordinary,
        ),
        FormatNumber => e_params(
            FormatNumber,
            &["FormatNumber"],
            sig(1, 5),
            &[
                "Expression",
                "NumDigitsAfterDecimal",
                "IncludeLeadingDigit",
                "UseParensForNegativeNumbers",
                "GroupDigits",
            ],
            Ordinary,
        ),
        FormatCurrency => e_params(
            FormatCurrency,
            &["FormatCurrency"],
            sig(1, 5),
            &[
                "Expression",
                "NumDigitsAfterDecimal",
                "IncludeLeadingDigit",
                "UseParensForNegativeNumbers",
                "GroupDigits",
            ],
            Ordinary,
        ),
        FormatPercent => e_params(
            FormatPercent,
            &["FormatPercent"],
            sig(1, 5),
            &[
                "Expression",
                "NumDigitsAfterDecimal",
                "IncludeLeadingDigit",
                "UseParensForNegativeNumbers",
                "GroupDigits",
            ],
            Ordinary,
        ),
        FormatDateTime => e_params(
            FormatDateTime,
            &["FormatDateTime"],
            sig(1, 2),
            &["Expression", "NamedFormat"],
            Ordinary,
        ),
        Filter => e_params(
            Filter,
            &["Filter"],
            sig(2, 4),
            &["SourceArray", "Match", "Include", "Compare"],
            Ordinary,
        ),

        // ── Math ──
        Abs => e(Abs, &["Abs"], sig(1, 1), Ordinary),
        Int => e(Int, &["Int"], sig(1, 1), Ordinary),
        Fix => e(Fix, &["Fix"], sig(1, 1), Ordinary),
        Sgn => e(Sgn, &["Sgn"], sig(1, 1), Ordinary),
        Round => e_params(
            Round,
            &["Round"],
            sig(1, 2),
            &["Number", "NumDigitsAfterDecimal"],
            Ordinary,
        ),
        Sqr => e_params(Sqr, &["Sqr"], sig(1, 1), &["Number"], Ordinary),
        Sin => e_params(Sin, &["Sin"], sig(1, 1), &["Number"], Ordinary),
        Cos => e_params(Cos, &["Cos"], sig(1, 1), &["Number"], Ordinary),
        Log => e_params(Log, &["Log"], sig(1, 1), &["Number"], Ordinary),
        Exp => e_params(Exp, &["Exp"], sig(1, 1), &["Number"], Ordinary),
        Atn => e_params(Atn, &["Atn"], sig(1, 1), &["Number"], Ordinary),
        Tan => e_params(Tan, &["Tan"], sig(1, 1), &["Number"], Ordinary),

        // ── DateTime ──
        Calendar => e(Calendar, &["Calendar"], sig(0, 0), Ordinary),
        CalendarLet => e(CalendarLet, &[], sig(1, 1), Ordinary),
        MacID => e_params(MacID, &["MacID"], sig(1, 1), &["Constant"], Ordinary),
        MacScript => e_params(MacScript, &["MacScript"], sig(1, 1), &["Script"], Ordinary),
        IMEStatus => e(IMEStatus, &["IMEStatus"], sig(0, 0), Ordinary),
        ErrObjectValue => e(ErrObjectValue, &[], sig(0, 0), Ordinary),
        UserForms => e(UserForms, &["UserForms"], sig(0, 1), Ordinary),
        Load => e_params(Load, &["Load"], sig(1, 1), &["Object"], Ordinary),
        Unload => e_params(Unload, &["Unload"], sig(1, 1), &["Object"], Ordinary),
        FileInputBytes => e_params(
            FileInputBytes,
            &["InputB", "InputB$"],
            sig(2, 2),
            &["Number", "FileNumber"],
            Ordinary,
        ),
        FileReset => e(FileReset, &["Reset"], sig(0, 0), Ordinary),
        FileWidthFunction => e_params(
            FileWidthFunction,
            &["Width"],
            sig(2, 2),
            &["FileNumber", "Width"],
            Ordinary,
        ),
        DateSerial => e_params(
            DateSerial,
            &["DateSerial"],
            sig(3, 3),
            &["Year", "Month", "Day"],
            Ordinary,
        ),
        TimeSerial => e_params(
            TimeSerial,
            &["TimeSerial"],
            sig(3, 3),
            &["Hour", "Minute", "Second"],
            Ordinary,
        ),
        DateValue => e_params(DateValue, &["DateValue"], sig(1, 1), &["Date"], Ordinary),
        TimeValue => e_params(TimeValue, &["TimeValue"], sig(1, 1), &["Time"], Ordinary),
        DateAdd => e_params(
            DateAdd,
            &["DateAdd"],
            sig(3, 3),
            &["Interval", "Number", "Date"],
            Ordinary,
        ),
        DateDiff => e_params(
            DateDiff,
            &["DateDiff"],
            sig(3, 5),
            &[
                "Interval",
                "Date1",
                "Date2",
                "FirstDayOfWeek",
                "FirstWeekOfYear",
            ],
            Ordinary,
        ),
        Year => e_params(Year, &["Year"], sig(1, 1), &["Date"], Ordinary),
        Month => e_params(Month, &["Month"], sig(1, 1), &["Date"], Ordinary),
        Day => e_params(Day, &["Day"], sig(1, 1), &["Date"], Ordinary),
        Weekday => e_params(
            Weekday,
            &["Weekday"],
            sig(1, 2),
            &["Date", "FirstDayOfWeek"],
            Ordinary,
        ),
        Hour => e_params(Hour, &["Hour"], sig(1, 1), &["Time"], Ordinary),
        Minute => e_params(Minute, &["Minute"], sig(1, 1), &["Time"], Ordinary),
        Second => e_params(Second, &["Second"], sig(1, 1), &["Time"], Ordinary),
        MonthName => e_params(
            MonthName,
            &["MonthName"],
            sig(1, 2),
            &["Month", "Abbreviate"],
            Ordinary,
        ),
        WeekdayName => e_params(
            WeekdayName,
            &["WeekdayName"],
            sig(1, 3),
            &["Weekday", "Abbreviate", "FirstDayOfWeek"],
            Ordinary,
        ),
        DatePart => e_params(
            DatePart,
            &["DatePart"],
            sig(2, 4),
            &["Interval", "Date", "FirstDayOfWeek", "FirstWeekOfYear"],
            Ordinary,
        ),
        DateNow => e(DateNow, &["Date"], sig(0, 0), Ordinary),
        TimeNow => e(TimeNow, &["Time"], sig(0, 0), Ordinary),
        DateNowText => e(DateNowText, &["Date$"], sig(0, 0), Ordinary),
        TimeNowText => e(TimeNowText, &["Time$"], sig(0, 0), Ordinary),
        Now => e(Now, &["Now"], sig(0, 0), Ordinary),
        Timer => e(Timer, &["Timer"], sig(0, 0), Ordinary),

        // ── Conversion ──
        Hex => e_params(Hex, &["Hex", "Hex$"], sig(1, 1), &["Number"], Ordinary),
        Oct => e_params(Oct, &["Oct", "Oct$"], sig(1, 1), &["Number"], Ordinary),
        CStr => e_params(CStr, &["CStr"], sig(1, 1), &["Expression"], Ordinary),
        Str => e_params(Str, &["Str", "Str$"], sig(1, 1), &["Number"], Ordinary),
        Val => e_params(Val, &["Val"], sig(1, 1), &["String"], Ordinary),
        CDate => e_params(
            CDate,
            &["CDate", "CVDate"],
            sig(1, 1),
            &["Expression"],
            Ordinary,
        ),
        CVErr => e_params(CVErr, &["CVErr"], sig(1, 1), &["Expression"], Ordinary),
        CBool => e_params(CBool, &["CBool"], sig(1, 1), &["Expression"], Ordinary),
        CByte => e_params(CByte, &["CByte"], sig(1, 1), &["Expression"], Ordinary),
        CInt => e_params(CInt, &["CInt"], sig(1, 1), &["Expression"], Ordinary),
        CLng => e_params(CLng, &["CLng"], sig(1, 1), &["Expression"], Ordinary),
        CLngLng => e_params(CLngLng, &["CLngLng"], sig(1, 1), &["Expression"], Ordinary),
        CLngPtr => e_params(CLngPtr, &["CLngPtr"], sig(1, 1), &["Expression"], Ordinary),
        CSng => e_params(CSng, &["CSng"], sig(1, 1), &["Expression"], Ordinary),
        CDbl => e_params(CDbl, &["CDbl"], sig(1, 1), &["Expression"], Ordinary),
        CCur => e_params(CCur, &["CCur"], sig(1, 1), &["Expression"], Ordinary),
        CDec => e_params(CDec, &["CDec"], sig(1, 1), &["Expression"], Ordinary),
        CVar => e_params(CVar, &["CVar"], sig(1, 1), &["Expression"], Ordinary),

        // ── Random ──
        Rnd => e_params(Rnd, &["Rnd"], sig(0, 1), &["Number"], Ordinary),
        Randomize => e_params(Randomize, &["Randomize"], sig(0, 1), &["Number"], Ordinary),

        // ── Financial ──
        Fv => e_params(
            Fv,
            &["FV"],
            sig(3, 5),
            &["Rate", "NPer", "Pmt", "PV", "Due"],
            Ordinary,
        ),
        Pv => e_params(
            Pv,
            &["PV"],
            sig(3, 5),
            &["Rate", "NPer", "Pmt", "FV", "Due"],
            Ordinary,
        ),
        Pmt => e_params(
            Pmt,
            &["Pmt"],
            sig(3, 5),
            &["Rate", "NPer", "PV", "FV", "Due"],
            Ordinary,
        ),
        Npv => e_params(Npv, &["NPV"], sig(2, 2), &["Rate", "ValueArray"], Ordinary),
        Irr => e_params(Irr, &["IRR"], sig(1, 2), &["ValueArray", "Guess"], Ordinary),
        Mirr => e_params(
            Mirr,
            &["MIRR"],
            sig(3, 3),
            &["ValueArray", "FinanceRate", "ReinvestRate"],
            Ordinary,
        ),
        Rate => e_params(
            Rate,
            &["Rate"],
            sig(3, 6),
            &["NPer", "Pmt", "PV", "FV", "Due", "Guess"],
            Ordinary,
        ),
        NPer => e_params(
            NPer,
            &["NPer"],
            sig(3, 5),
            &["Rate", "Pmt", "PV", "FV", "Due"],
            Ordinary,
        ),
        IPmt => e_params(
            IPmt,
            &["IPmt"],
            sig(4, 6),
            &["Rate", "Per", "NPer", "PV", "FV", "Due"],
            Ordinary,
        ),
        PPmt => e_params(
            PPmt,
            &["PPmt"],
            sig(4, 6),
            &["Rate", "Per", "NPer", "PV", "FV", "Due"],
            Ordinary,
        ),
        Sln => e_params(
            Sln,
            &["SLN"],
            sig(3, 3),
            &["Cost", "Salvage", "Life"],
            Ordinary,
        ),
        Syd => e_params(
            Syd,
            &["SYD"],
            sig(4, 4),
            &["Cost", "Salvage", "Life", "Period"],
            Ordinary,
        ),
        Ddb => e_params(
            Ddb,
            &["DDB"],
            sig(4, 5),
            &["Cost", "Salvage", "Life", "Period", "Factor"],
            Ordinary,
        ),

        // ── Information ──
        IsArray => e_params(IsArray, &["IsArray"], sig(1, 1), &["VarName"], Ordinary),
        VarType => e_params(VarType, &["VarType"], sig(1, 1), &["VarName"], Ordinary),
        TypeName => e_params(TypeName, &["TypeName"], sig(1, 1), &["VarName"], Ordinary),
        IsNumeric => e_params(
            IsNumeric,
            &["IsNumeric"],
            sig(1, 1),
            &["Expression"],
            Ordinary,
        ),
        IsError => e_params(IsError, &["IsError"], sig(1, 1), &["Expression"], Ordinary),
        IsDate => e_params(IsDate, &["IsDate"], sig(1, 1), &["Expression"], Ordinary),
        IsObject => e_params(
            IsObject,
            &["IsObject"],
            sig(1, 1),
            &["Expression"],
            Ordinary,
        ),
        IsNull => e_params(IsNull, &["IsNull"], sig(1, 1), &["Expression"], Ordinary),
        IsEmpty => e_params(IsEmpty, &["IsEmpty"], sig(1, 1), &["Expression"], Ordinary),
        IsMissing => e_params(IsMissing, &["IsMissing"], sig(1, 1), &["ArgName"], Ordinary),
        ErrorText => e_params(
            ErrorText,
            &["Error", "Error$"],
            sig(0, 1),
            &["ErrorNumber"],
            Ordinary,
        ),
        // IIf/Choose/Switch resolve by name through the `vba_library` special-form
        // route, not the intrinsic-name map — so they carry no names here; the
        // entry exists only to give the lowering target a shape.
        IIf => e_params(
            IIf,
            &[],
            sig(3, 3),
            &["Expression", "TruePart", "FalsePart"],
            SpecialForm,
        ),
        Choose => e(Choose, &[], variadic(1), SpecialForm),
        Switch => e(Switch, &[], variadic(2), SpecialForm),
        Rgb => e_params(
            Rgb,
            &["RGB"],
            sig(3, 3),
            &["Red", "Green", "Blue"],
            Ordinary,
        ),
        QbColor => e_params(QbColor, &["QBColor"], sig(1, 1), &["Color"], Ordinary),

        // ── Interaction ──
        Command => e(Command, &["Command", "Command$"], sig(0, 0), Ordinary),
        GetSetting => e_params(
            GetSetting,
            &["GetSetting"],
            sig(3, 4),
            &["AppName", "Section", "Key", "Default"],
            Ordinary,
        ),
        GetAllSettings => e_params(
            GetAllSettings,
            &["GetAllSettings"],
            sig(2, 2),
            &["AppName", "Section"],
            Ordinary,
        ),
        SaveSetting => e_params(
            SaveSetting,
            &["SaveSetting"],
            sig(4, 4),
            &["AppName", "Section", "Key", "Setting"],
            FileStatement,
        ),
        DeleteSetting => e_params(
            DeleteSetting,
            &["DeleteSetting"],
            sig(2, 3),
            &["AppName", "Section", "Key"],
            FileStatement,
        ),
        Partition => e_params(
            Partition,
            &["Partition"],
            sig(4, 4),
            &["Number", "Start", "Stop", "Interval"],
            Ordinary,
        ),

        // ── File / Console I/O ──
        FreeFile => e_params(
            FreeFile,
            &["FreeFile"],
            sig(0, 1),
            &["RangeNumber"],
            Ordinary,
        ),
        FileOpen => e(FileOpen, &[], variadic(0), FileStatement),
        FileClose => e(FileClose, &[], variadic(0), FileStatement),
        // `Kill pathname` deletes files. Unlike Open/Close/Print#/Name/Lock/…, `Kill`
        // is not a lexer keyword, so it parses as an ordinary statement-call and must
        // resolve by name (a 1-argument native).
        FileKill => e_params(FileKill, &["Kill"], sig(1, 1), &["PathName"], FileStatement),
        // Like `Kill`, the directory statements are not lexer keywords — they
        // resolve by name as 1-argument statement-calls.
        FileMkDir => e_params(FileMkDir, &["MkDir"], sig(1, 1), &["Path"], FileStatement),
        FileRmDir => e_params(FileRmDir, &["RmDir"], sig(1, 1), &["Path"], FileStatement),
        FileCurDir => e_params(
            FileCurDir,
            &["CurDir", "CurDir$"],
            sig(0, 1),
            &["Drive"],
            Ordinary,
        ),
        FileChDir => e_params(FileChDir, &["ChDir"], sig(1, 1), &["Path"], FileStatement),
        FileLen => e_params(FileLen, &["FileLen"], sig(1, 1), &["PathName"], Ordinary),
        FileCopy => e_params(
            FileCopy,
            &["FileCopy"],
            sig(2, 2),
            &["Source", "Destination"],
            FileStatement,
        ),
        FileGetAttr => e_params(
            FileGetAttr,
            &["GetAttr"],
            sig(1, 1),
            &["PathName"],
            Ordinary,
        ),
        FileSetAttr => e_params(
            FileSetAttr,
            &["SetAttr"],
            sig(2, 2),
            &["PathName", "Attributes"],
            FileStatement,
        ),
        FileChDrive => e_params(
            FileChDrive,
            &["ChDrive"],
            sig(1, 1),
            &["Drive"],
            FileStatement,
        ),
        FileDateTime => e_params(
            FileDateTime,
            &["FileDateTime"],
            sig(1, 1),
            &["PathName"],
            Ordinary,
        ),
        FileRead => e(FileRead, &[], variadic(0), Ordinary),
        FileWrite => e(FileWrite, &[], variadic(0), FileStatement),
        FilePrint => e(FilePrint, &[], variadic(0), FileStatement),
        ConsolePrint => e(ConsolePrint, &["Print"], variadic(0), FileStatement),
        FileInput => e(FileInput, &[], variadic(0), FileStatement),
        FileInputChars => e_params(
            FileInputChars,
            &["Input", "Input$"],
            sig(2, 2),
            &["Number", "FileNumber"],
            Ordinary,
        ),
        ConsoleInput => e(ConsoleInput, &[], variadic(0), FileStatement),
        FileLineInput => e(FileLineInput, &[], variadic(0), FileStatement),
        ConsoleLineInput => e(ConsoleLineInput, &[], variadic(0), FileStatement),
        FileEof => e_params(FileEof, &["EOF"], sig(1, 1), &["FileNumber"], Ordinary),
        FileLof => e_params(FileLof, &["LOF"], sig(1, 1), &["FileNumber"], Ordinary),
        FileSeek => e_params(FileSeek, &["Seek"], sig(1, 1), &["FileNumber"], Ordinary),
        FileAttr => e_params(
            FileAttr,
            &["FileAttr"],
            sig(2, 2),
            &["FileNumber", "ReturnType"],
            Ordinary,
        ),
        // Random-access / misc file statements (routed from the parser, not by name).
        FilePut => e(FilePut, &[], variadic(0), FileStatement),
        FileGetInto => e(FileGetInto, &[], variadic(0), FileStatement),
        FileWidth => e(FileWidth, &[], variadic(0), FileStatement),
        FileRename => e(FileRename, &[], variadic(0), FileStatement),
        FileLock => e(FileLock, &[], variadic(0), FileStatement),
        FileUnlock => e(FileUnlock, &[], variadic(0), FileStatement),
        FileLoc => e_params(FileLoc, &["Loc"], sig(1, 1), &["FileNumber"], Ordinary),

        // ── Interaction / host ──
        MsgBox => e_params(
            MsgBox,
            &["MsgBox"],
            sig(1, 5),
            &["Prompt", "Buttons", "Title", "HelpFile", "Context"],
            Ordinary,
        ),
        InputBox => e_params(
            InputBox,
            &["InputBox"],
            sig(1, 7),
            &[
                "Prompt", "Title", "Default", "XPos", "YPos", "HelpFile", "Context",
            ],
            Ordinary,
        ),
        Beep => e(Beep, &["Beep"], sig(0, 0), Ordinary),
        DoEvents => e(DoEvents, &["DoEvents"], sig(0, 0), Ordinary),
        SendKeys => e_params(
            SendKeys,
            &["SendKeys"],
            sig(1, 2),
            &["String", "Wait"],
            FileStatement,
        ),
        AppActivate => e_params(
            AppActivate,
            &["AppActivate"],
            sig(1, 2),
            &["Title", "Wait"],
            FileStatement,
        ),
        Shell => e_params(
            Shell,
            &["Shell"],
            sig(1, 2),
            &["PathName", "WindowStyle"],
            Ordinary,
        ),
        Environ => e_params(
            Environ,
            &["Environ", "Environ$"],
            sig(1, 1),
            &["Expression"],
            Ordinary,
        ),
        Dir => e_params(
            Dir,
            &["Dir", "Dir$"],
            sig(0, 2),
            &["PathName", "Attributes"],
            Ordinary,
        ),
        CreateObject => e_params(
            CreateObject,
            &["CreateObject"],
            sig(1, 2),
            &["Class", "ServerName"],
            SpecialForm,
        ),
        GetObject => e_params(
            GetObject,
            &["GetObject"],
            sig(1, 2),
            &["PathName", "Class"],
            SpecialForm,
        ),
        ComSubscribeEvent => e(ComSubscribeEvent, &[], variadic(0), SpecialForm),
        ComUnsubscribeEvent => e(ComUnsubscribeEvent, &[], variadic(0), SpecialForm),
        ComEventCallbackSubscription => {
            e(ComEventCallbackSubscription, &[], variadic(0), SpecialForm)
        }
        ComEventCallbackArg => e(ComEventCallbackArg, &[], variadic(0), SpecialForm),
        ComReleaseEventCallback => e(ComReleaseEventCallback, &[], variadic(0), SpecialForm),
        AppObject => e(AppObject, &[], sig(1, 1), SpecialForm),
        EvaluateForeign => e(EvaluateForeign, &[], sig(2, 2), SpecialForm),

        // ── Diagnostics ──
        DebugPrint => e(DebugPrint, &[], variadic(0), DebugMember),
    }
}

/// Every `NativeImplId`, generated from the enum declaration itself
/// ([`NativeImplId::ALL`]). The exhaustive `intrinsic_entry` match is the other
/// half of the no-partial guarantee: a new variant fails to compile until it has
/// an entry, and it joins this slice automatically — nothing to hand-maintain.
pub const ALL_INTRINSICS: &[NativeImplId] = NativeImplId::ALL;

/// Resolve a bare intrinsic name (case-insensitive) to its `NativeImplId`.
pub fn name_to_intrinsic(name: &str) -> Option<NativeImplId> {
    ALL_INTRINSICS.iter().copied().find(|id| {
        intrinsic_entry(*id)
            .names
            .iter()
            .any(|candidate| candidate.eq_ignore_ascii_case(name))
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn catalog_covers_every_native_impl_id() {
        // `intrinsic_entry` is a wildcard-free match, so name/signature/kind
        // coverage is enforced at compile time and there is no count to maintain.
        // This guards the one property the compiler can't: no callable name is
        // claimed by two variants, which would make `name_to_intrinsic` (a
        // first-match lookup) silently shadow one of them. Statement/assignment
        // forms such as `MidStmt` carry no name and are dispatched structurally.
        let mut seen: Vec<String> = Vec::new();
        for id in ALL_INTRINSICS.iter().copied() {
            for name in intrinsic_entry(id).names {
                let lower = name.to_ascii_lowercase();
                assert!(
                    !seen.contains(&lower),
                    "name {name:?} is claimed by more than one NativeImplId"
                );
                seen.push(lower);
            }
        }
    }

    #[test]
    fn names_round_trip_through_reverse_map() {
        for id in ALL_INTRINSICS.iter().copied() {
            for name in intrinsic_entry(id).names {
                assert_eq!(name_to_intrinsic(name), Some(id), "name {name}");
            }
        }
    }

    #[test]
    fn migrated_library_member_name_matches_catalog_primary() {
        // Ordinary functions use the source primary as their bundle binding.
        // Input expressions need a separate binding because FileSystem.Input
        // already names the CSV statement in existing compiled packages.
        for id in ALL_INTRINSICS.iter().copied() {
            let Some((_module, member)) = id.library_member() else {
                continue;
            };
            let primary = intrinsic_entry(id)
                .names
                .first()
                .unwrap_or_else(|| panic!("migrated id {id:?} has no catalog name"));
            let binding = if id == NativeImplId::FileInputChars {
                "InputString"
            } else {
                primary
            };
            assert!(
                member.eq_ignore_ascii_case(binding),
                "library_member name {member:?} != binding {binding:?} for {id:?}",
            );
        }
    }

    #[test]
    fn input_expression_and_statement_keep_distinct_library_bindings() {
        use NativeImplId::{FileInput, FileInputChars};
        assert_eq!(intrinsic_entry(FileInputChars).names, &["Input", "Input$"]);
        assert_eq!(
            FileInputChars.library_member(),
            Some(("FileSystem", "InputString"))
        );
        assert_eq!(
            FileInput.library_statement_member(),
            Some(("FileSystem", "Input"))
        );
        assert_eq!(name_to_intrinsic("Input"), Some(FileInputChars));
        assert_eq!(name_to_intrinsic("Input$"), Some(FileInputChars));
        assert_eq!(name_to_intrinsic("InputString"), None);
    }

    #[test]
    fn quirky_call_shapes_are_classified() {
        assert_eq!(
            intrinsic_entry(NativeImplId::MidStmt).call_shape,
            CallShape::AssignmentTarget
        );
        assert_eq!(
            intrinsic_entry(NativeImplId::FileOpen).call_shape,
            CallShape::FileStatement
        );
        assert_eq!(
            intrinsic_entry(NativeImplId::FilePrint).call_shape,
            CallShape::FileStatement
        );
        assert_eq!(
            intrinsic_entry(NativeImplId::SendKeys).call_shape,
            CallShape::FileStatement
        );
        assert_eq!(
            intrinsic_entry(NativeImplId::AppActivate).call_shape,
            CallShape::FileStatement
        );
        assert_eq!(
            intrinsic_entry(NativeImplId::DebugPrint).call_shape,
            CallShape::DebugMember
        );
        assert_eq!(
            intrinsic_entry(NativeImplId::CreateObject).call_shape,
            CallShape::SpecialForm
        );
    }
}
