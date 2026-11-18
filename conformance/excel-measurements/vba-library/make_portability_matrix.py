"""Writes portability-matrix.tsv: every VBA library procedure (functions.tsv) plus the
statements and language features that touch the host, classified for running Windows VBA on
Linux. See docs/design/linux-portability.md for the classes and modes.

Classes:
  pure      result depends only on the arguments; portable when the implementation is
            platform-independent (see the note column for current exceptions)
  locale    depends on regional settings; portable once the locale is configured, not taken
            from the Linux host
  shim      Windows behaviour can be emulated on Linux (paths, attributes, registry, encodings)
  host      the value legitimately depends on the machine (clock, environment, processes);
            portable in shape, not in value
  windows   needs Windows itself (Win32 DLLs, COM servers, window manager): reject or replace

python -I make_portability_matrix.py
"""

import csv
import os

HERE = os.path.dirname(os.path.abspath(__file__))

# (class, CalcVBA on Linux today, shim or check)
FILE = "shim"
PATH_NOTE = "Windows paths passed to the OS unchanged (backslash and C:\\ are filename characters)"
PATH_SHIM = "path mapping (drive letters to roots, \\ to /), case-insensitive lookup"
C = {
    # Strings / Conversion that touch the code page
    "Asc": ("shim", "ANSI table hard-coded Windows-1252 off Windows", "configured ANSI code page (the Windows machine's ACP)"),
    "AscB": ("pure", "", ""),
    "Chr": ("shim", "Windows-1252 table", "configured ANSI code page"),
    "ChrB": ("pure", "", ""),
    "StrConv": ("shim", "vbUnicode/vbFromUnicode via Windows-1252; vbWide/vbNarrow/Kana by table", "configured code page and LCID"),
    "String": ("shim", "character codes 128-255 via Windows-1252", "configured ANSI code page"),
    "StrComp": ("locale", "vbTextCompare uses Rust lower-casing, CompareStringW on Windows", "a portable text-compare table matching CompareStringW for the configured LCID"),
    "InStr": ("locale", "vbTextCompare as StrComp", "as StrComp"),
    "InStrB": ("locale", "vbTextCompare as StrComp", "as StrComp"),
    "InStrRev": ("locale", "vbTextCompare as StrComp", "as StrComp"),
    "Replace": ("locale", "vbTextCompare as StrComp", "as StrComp"),
    "Split": ("locale", "vbTextCompare as StrComp", "as StrComp"),
    "Filter": ("locale", "vbTextCompare as StrComp", "as StrComp"),
    "LCase": ("locale", "Rust Unicode case mapping", "match LCMapStringW for the configured LCID"),
    "UCase": ("locale", "Rust Unicode case mapping", "match LCMapStringW for the configured LCID"),
    "Format": ("locale", "fixed en-US conventions on every host", "locale model (date order, separators, names)"),
    "FormatCurrency": ("locale", "fixed conventions", "locale model"),
    "FormatDateTime": ("locale", "fixed en-US date text", "locale model"),
    "FormatNumber": ("locale", "fixed conventions", "locale model"),
    "FormatPercent": ("locale", "fixed conventions", "locale model"),
    "MonthName": ("locale", "English names", "locale model"),
    "WeekdayName": ("locale", "English names", "locale model"),
    "CStr": ("locale", "dates as en-US M/D/YYYY", "locale model"),
    "Str": ("pure", "", ""),
    "Val": ("pure", "", ""),
    "CDate": ("locale", "ambiguous text dates refused instead of read by the locale", "locale model"),
    "CVDate": ("locale", "as CDate", "locale model"),
    "DateValue": ("locale", "as CDate", "locale model"),
    "TimeValue": ("locale", "fixed AM/PM forms", "locale model"),
    "IsDate": ("locale", "as CDate", "locale model"),
    "IsNumeric": ("locale", "period decimal only", "locale model"),
    "CDbl": ("locale", "period decimal only for strings", "locale model"),
    "CSng": ("locale", "as CDbl", "locale model"),
    "CCur": ("locale", "as CDbl", "locale model"),
    "CDec": ("locale", "as CDbl", "locale model"),
    "Calendar": ("locale", "not implemented", "locale model (Gregorian/Hijri)"),
    # Math
    "Sin": ("pure", "calls the platform libm: last bit differs Linux vs Windows (dc-q6brh.112)", "one Excel-exact implementation"),
    "Cos": ("pure", "platform libm", "one Excel-exact implementation"),
    "Tan": ("pure", "platform libm", "one Excel-exact implementation"),
    "Atn": ("pure", "platform libm", "one Excel-exact implementation"),
    "Exp": ("pure", "platform libm", "one Excel-exact implementation"),
    "Log": ("pure", "platform libm", "one Excel-exact implementation"),
    "Sqr": ("pure", "IEEE sqrt is exact on every platform", ""),
    # DateTime
    "Now": ("host", "local time via localtime_r", "time zone and clock are the machine's; a fixed clock in deterministic mode"),
    "Date": ("host", "local date", "as Now"),
    "Time": ("host", "local time", "as Now"),
    "Timer": ("host", "local seconds since midnight", "as Now"),
    # FileSystem
    "ChDir": (FILE, PATH_NOTE, PATH_SHIM),
    "ChDrive": (FILE, "does nothing on Linux", "drive map: ChDrive selects a mapped root"),
    "CurDir": (FILE, "the process directory; drive argument ignored", "report the mapped Windows path (C:\\...)"),
    "Dir": (FILE, "wildcards case-sensitive; hidden/system/archive never reported", PATH_SHIM + "; attribute store"),
    "EOF": ("pure", "", ""),
    "FileAttr": ("pure", "", ""),
    "FileCopy": (FILE, PATH_NOTE, PATH_SHIM),
    "FileDateTime": (FILE, "mtime in local time", PATH_SHIM),
    "FileLen": (FILE, PATH_NOTE, PATH_SHIM),
    "FreeFile": ("pure", "", ""),
    "GetAttr": (FILE, "only vbReadOnly and vbDirectory", "attribute store (xattr or sidecar) for hidden/system/archive"),
    "SetAttr": (FILE, "only the read-only bit", "attribute store"),
    "Kill": (FILE, "wildcards case-sensitive", PATH_SHIM),
    "Loc": ("pure", "", ""),
    "LOF": ("pure", "", ""),
    "MkDir": (FILE, PATH_NOTE, PATH_SHIM),
    "RmDir": (FILE, PATH_NOTE, PATH_SHIM),
    "Reset": ("pure", "", ""),
    "Seek": ("pure", "", ""),
    "Input": ("shim", "ANSI via Windows-1252", "configured code page"),
    "InputB": ("pure", "", ""),
    "Width": ("pure", "", ""),
    # Interaction
    "AppActivate": ("windows", "error 5 off Windows", "reject in strict mode; no shim (window manager)"),
    "Beep": ("host", "does nothing", "terminal bell or nothing"),
    "CallByName": ("pure", "", ""),
    "Command": ("host", "process arguments", ""),
    "CreateObject": ("windows", "only Scripting.Dictionary (engine class); other ProgIDs unsupported off Windows", "portable classes for common ProgIDs (Scripting.FileSystemObject, VBScript.RegExp, ADODB.Stream?); reject the rest"),
    "GetObject": ("windows", "unsupported off Windows", "reject; no running-object table"),
    "DeleteSetting": ("shim", "in memory only (on Windows too)", "settings store: a file under XDG config, the registry on Windows"),
    "GetAllSettings": ("shim", "in memory only", "settings store"),
    "GetSetting": ("shim", "in memory only", "settings store"),
    "SaveSetting": ("shim", "in memory only", "settings store"),
    "DoEvents": ("pure", "yield", ""),
    "Environ": ("shim", "names case-sensitive; no TEMP/TMP/USERPROFILE/APPDATA; Environ(n) order differs", "case-insensitive names, synthesised Windows variables mapped to Linux equivalents, VBA's 1-based order"),
    "InputBox": ("host", "prompt to stderr, returns the default on linux-stdio", "UI host decides (scripted responses, a dialog in a GUI host)"),
    "MsgBox": ("host", "prompt to stderr, returns the default", "UI host decides"),
    "MacScript": ("windows", "Mac-only in VBA: error on Windows", "same error everywhere"),
    "SendKeys": ("windows", "error 5 everywhere", "reject in strict mode"),
    "Shell": ("windows", "runs sh -c with the text: a Windows command line will not run", "reject in strict mode, or a command map for known programs"),
    "IMEStatus": ("host", "not implemented", "0 (no IME)"),
    "MacID": ("windows", "not implemented", "the same error as Windows Excel"),
    # Information
    "Err": ("pure", "", ""),
    "Erl": ("pure", "", ""),
    # Hidden module
    "VarPtr": ("pure", "real addresses", "values differ per run on any host"),
    "StrPtr": ("pure", "real addresses", "values differ per run on any host"),
    "ObjPtr": ("pure", "real addresses", "values differ per run on any host"),
    # VBEGlobal
    "Load": ("host", "not implemented", "needs a forms host (CalcForms)"),
    "Unload": ("host", "not implemented", "needs a forms host (CalcForms)"),
    "UserForms": ("host", "not implemented", "needs a forms host (CalcForms)"),
}
DEFAULT = ("pure", "", "")

# Statements and language features outside the library list.
EXTRA = [
    ("Language", "Declare (Win32 DLLs: kernel32, user32, ...)", "windows",
     "dlopen of lib<name>.so (libkernel32.so does not exist); integer arguments only, max 6",
     "emulated table for common kernel32/user32 calls (GetTickCount, Sleep, RtlMoveMemory, QueryPerformanceCounter, GetTempPath...), reject the rest in strict mode"),
    ("Language", "Declare (C runtime: msvcrt, ucrtbase)", "shim", "mapped to libc.so.6", "keep; check argument types"),
    ("Language", "AddressOf (callbacks to Win32 APIs)", "windows", "callback thunks are Windows x64 only", "reject in strict mode"),
    ("Language", "#If Win64 / Win32 / Mac / VBA7", "shim", "Linux is treated as Win64 Windows (Win64 = Win32 = True)", "keep as Windows (the code was written for it), configurable per project"),
    ("Language", "LongPtr", "pure", "follows the target pointer width (64-bit)", ""),
    ("Language", "Early-bound COM references (Scripting, ADODB, MSXML, Excel...)", "windows", "type libraries need Windows COM", "portable classes for the common libraries, reject others"),
    ("Statements", "Open / Close / Get / Put", FILE, PATH_NOTE + "; Open creates missing parent folders (VBA raises 76)", PATH_SHIM),
    ("Statements", "Print # / Write # / Line Input #", "shim", "CRLF on every platform (as VBA); text written as UTF-8, not ANSI", "write and read text files in the configured ANSI code page"),
    ("Statements", "Input #", "shim", "quoted fields Latin-1, unquoted UTF-8", "configured ANSI code page"),
    ("Statements", "Name ... As", FILE, PATH_NOTE + "; creates the target's parent folders", PATH_SHIM),
    ("Statements", "Lock / Unlock", "shim", "check", "advisory locks"),
    ("Statements", "Like with Option Compare Text", "locale", "Rust case folding", "as StrComp"),
    ("Statements", "Option Compare Text (=, <, ... on strings)", "locale", "Rust case folding", "as StrComp"),
    ("Statements", "Date literals #m/d/yyyy#", "pure", "fixed month/day/year grammar (MS-VBAL)", ""),
    ("Runtime", "Floating point arithmetic + - * /", "pure", "IEEE; same on every platform", ""),
    ("Runtime", "^ operator", "pure", "platform libm pow", "one Excel-exact implementation"),
]


def main():
    rows = list(csv.DictReader(open(os.path.join(HERE, "functions.tsv"), encoding="utf-8"), delimiter="\t"))
    out = []
    for r in rows:
        if r["module"] in ("Collection", "ErrObject", "RegExp", "Match", "MatchCollection", "SubMatches"):
            cls, today, shim = DEFAULT
        else:
            cls, today, shim = C.get(r["name"], DEFAULT)
        out.append((r["module"], r["name"] + ("/$" if r["dollar_form"] else ""), cls, today, shim))
    out += EXTRA
    with open(os.path.join(HERE, "portability-matrix.tsv"), "w", encoding="utf-8", newline="\n") as f:
        f.write("area\tfeature\tclass\tcalcvba_on_linux_today\tshim_or_check\n")
        for row in out:
            f.write("\t".join(row) + "\n")
    from collections import Counter
    print(Counter(r[2] for r in out))


if __name__ == "__main__":
    main()
