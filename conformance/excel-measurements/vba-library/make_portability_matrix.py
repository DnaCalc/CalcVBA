"""Writes portability-matrix.tsv: every VBA library procedure (functions.tsv) plus the
statements and language features that touch the host, for the two VBA dialects (Windows Excel,
Mac Excel) and the hosts CalcVBA runs on (Windows, macOS, Linux, wasm WASI, wasm browser).
See docs/design/portability.md.

Columns:
  class        what the feature depends on, relative to the Windows dialect:
               pure | locale | shim | host | windows
  mac_dialect  how Excel for Mac's VBA differs ("to measure" until measured on a Mac)
  windows, macos, linux, wasi, browser
               what CalcVBA has to do on that host to give the reference dialect's answer:
                 native   the host's own facility gives it directly
                 model    CalcVBA's own model (locale, maths): the same code on every host
                 shim     emulation (path map, stores, code page, ...)
                 host     a machine value (clock, environment, UI host)
                 replace  only through a portable replacement (portable class, emulated Declare)
                 reject   cannot work there: refused at load with the reason
  today        what CalcVBA does now (surveyed 2026-11-18)

python -I make_portability_matrix.py
"""

import csv
import os

HERE = os.path.dirname(os.path.abspath(__file__))
HOSTS = ("windows", "macos", "linux", "wasi", "browser")

# Verdict rows per class, (windows, macos, linux, wasi, browser).
BY_CLASS = {
    "pure": ("native", "native", "native", "native", "native"),
    "locale": ("model", "model", "model", "model", "model"),
    "file": ("native", "shim", "shim", "shim", "shim"),
    "store": ("native", "shim", "shim", "shim", "shim"),
    "codepage": ("model", "model", "model", "model", "model"),
    "envvar": ("native", "shim", "shim", "shim", "shim"),
    "clock": ("host", "host", "host", "host", "host"),
    "ui": ("host", "host", "host", "host", "host"),
    "forms": ("host", "host", "host", "reject", "host"),
    "math": ("model", "model", "model", "model", "model"),
    "win32": ("native", "replace", "replace", "replace", "replace"),
    "com": ("native", "replace", "replace", "replace", "replace"),
    "winonly": ("native", "reject", "reject", "reject", "reject"),
    "process": ("native", "shim", "shim", "reject", "reject"),
}
CLASS_OF = {"file": "shim", "store": "shim", "codepage": "shim", "envvar": "shim", "clock": "host",
            "ui": "host", "forms": "host", "math": "pure", "win32": "windows", "com": "windows",
            "winonly": "windows", "process": "windows", "pure": "pure", "locale": "locale"}

PATH_TODAY = "Windows paths go to the OS unchanged; wildcards case-sensitive off Windows"
MAC_PATHS = "POSIX paths with / (Excel 2016+); sandboxed file access (to measure)"
L = "locale"
# name -> (kind, mac_dialect, today)
F = {
    "Asc": ("codepage", "code page for 128-255 may be Mac Roman (to measure)", "Windows-1252 table off Windows"),
    "Chr": ("codepage", "as Asc (to measure)", "Windows-1252 table off Windows"),
    "String": ("codepage", "as Asc (to measure)", "Windows-1252 table off Windows"),
    "StrConv": ("codepage", "vbUnicode/vbFromUnicode code page; vbWide/Kana (to measure)", "Windows-1252 table off Windows"),
    "Input": ("codepage", "as Asc (to measure)", "Windows-1252 table off Windows"),
    "StrComp": (L, "text compare rules of the Mac locale (to measure)", "Rust case folding off Windows, CompareStringW on Windows"),
    "InStr": (L, "as StrComp", "as StrComp"), "InStrB": (L, "as StrComp", "as StrComp"),
    "InStrRev": (L, "as StrComp", "as StrComp"), "Replace": (L, "as StrComp", "as StrComp"),
    "Split": (L, "as StrComp", "as StrComp"), "Filter": (L, "as StrComp", "as StrComp"),
    "LCase": (L, "to measure", "Rust Unicode case mapping"), "UCase": (L, "to measure", "Rust Unicode case mapping"),
    "Format": (L, "Mac regional settings (to measure)", "fixed en-US conventions on every host"),
    "FormatCurrency": (L, "Mac regional settings", "fixed conventions"),
    "FormatDateTime": (L, "Mac regional settings", "fixed en-US"),
    "FormatNumber": (L, "Mac regional settings", "fixed conventions"),
    "FormatPercent": (L, "Mac regional settings", "fixed conventions"),
    "MonthName": (L, "Mac language", "English"), "WeekdayName": (L, "Mac language", "English"),
    "CStr": (L, "Mac regional settings for dates", "en-US M/D/YYYY dates"),
    "CDate": (L, "Mac regional settings", "ambiguous text dates refused"),
    "CVDate": (L, "as CDate", "as CDate"), "DateValue": (L, "as CDate", "as CDate"),
    "TimeValue": (L, "as CDate", "fixed AM/PM forms"), "IsDate": (L, "as CDate", "as CDate"),
    "IsNumeric": (L, "Mac decimal separator", "period decimal only"),
    "CDbl": (L, "Mac decimal separator", "period decimal only"), "CSng": (L, "as CDbl", "as CDbl"),
    "CCur": (L, "as CDbl", "as CDbl"), "CDec": (L, "as CDbl", "as CDbl"),
    "Calendar": (L, "to measure", "not implemented"),
    "Sin": ("math", "Mac Excel's maths library may differ from Windows Excel in the last bit (to measure)", "platform libm (Linux and Windows already differ: dc-q6brh.112)"),
    "Cos": ("math", "as Sin", "platform libm"), "Tan": ("math", "as Sin", "platform libm"),
    "Atn": ("math", "as Sin", "platform libm"), "Exp": ("math", "as Sin", "platform libm"),
    "Log": ("math", "as Sin", "platform libm"),
    "Now": ("clock", "same", "local time"), "Date": ("clock", "same", "local date"),
    "Time": ("clock", "same", "local time"), "Timer": ("clock", "resolution differs (to measure)", "local seconds since midnight, f32"),
    "ChDir": ("file", MAC_PATHS, PATH_TODAY), "ChDrive": ("file", "no drives: error or no-op (to measure)", "no-op off Windows"),
    "CurDir": ("file", MAC_PATHS, "process directory; drive argument ignored"),
    "Dir": ("file", "POSIX paths; MacID file-type filters; case-insensitive on default APFS (to measure)", PATH_TODAY + "; no hidden/system/archive"),
    "FileCopy": ("file", MAC_PATHS, PATH_TODAY), "FileDateTime": ("file", MAC_PATHS, "mtime, local time"),
    "FileLen": ("file", MAC_PATHS, PATH_TODAY), "GetAttr": ("file", "attribute bits on Mac (to measure)", "only read-only and directory off Windows"),
    "SetAttr": ("file", "to measure", "only the read-only bit off Windows"), "Kill": ("file", MAC_PATHS, PATH_TODAY),
    "MkDir": ("file", MAC_PATHS, PATH_TODAY), "RmDir": ("file", MAC_PATHS, PATH_TODAY),
    "SaveSetting": ("store", "stored in a preferences plist, not a registry (to measure)", "in memory on every host"),
    "GetSetting": ("store", "as SaveSetting", "in memory"), "GetAllSettings": ("store", "as SaveSetting", "in memory"),
    "DeleteSetting": ("store", "as SaveSetting", "in memory"),
    "Environ": ("envvar", "Unix variables (HOME, USER, TMPDIR); no TEMP or USERPROFILE", "case-sensitive names off Windows; Environ(n) order differs"),
    "Command": ("clock", "to measure", "process arguments"),
    "Beep": ("ui", "same", "no-op"), "MsgBox": ("ui", "Mac dialog; some button styles differ (to measure)", "prompt to stderr on linux-stdio"),
    "InputBox": ("ui", "as MsgBox", "default returned on linux-stdio"), "DoEvents": ("pure", "same", "yield"),
    "IMEStatus": ("ui", "to measure", "not implemented"),
    "AppActivate": ("winonly", "activates by application name (to measure)", "error 5 off Windows"),
    "SendKeys": ("winonly", "not supported on Mac (to measure)", "error 5 everywhere"),
    "Shell": ("process", "runs a POSIX command (to measure)", "sh -c off Windows, cmd /C on Windows"),
    "CreateObject": ("com", "no COM on Mac: Scripting.Dictionary and other ProgIDs fail (to measure the error)", "only Scripting.Dictionary off Windows"),
    "GetObject": ("com", "no COM on Mac (to measure)", "unsupported off Windows"),
    "MacID": ("pure", "Mac file-type IDs (works on Mac)", "not implemented"),
    "MacScript": ("winonly", "AppleScript; deprecated in sandboxed Excel 2016+ for AppleScriptTask (to measure)", "not implemented"),
    "Load": ("forms", "UserForms on Mac (to measure)", "not implemented"),
    "Unload": ("forms", "as Load", "not implemented"), "UserForms": ("forms", "as Load", "not implemented"),
    "VarPtr": ("pure", "same", "real addresses"), "StrPtr": ("pure", "same", "real addresses"),
    "ObjPtr": ("pure", "same", "real addresses"),
}
# Mac-native verdict overrides: things the Mac dialect itself has (used when the target is Mac VBA
# on a macOS host); the per-host verdicts above assume Windows-dialect code.
EXTRA = [
    ("Language", "Declare into Win32 DLLs (kernel32, user32, ...)", "win32",
     "absent on Mac: Mac Declares name dylibs", "dlopen lib<name>.so off Windows (fails); integers only"),
    ("Language", "Declare into Mac dylibs (libc.dylib, ...)", "pure",
     "native on Mac", "dlopen off Windows; integer arguments only"),
    ("Language", "Declare into the C runtime (msvcrt, ucrtbase)", "win32",
     "absent", "mapped to libc.so.6 off Windows"),
    ("Language", "AddressOf (callbacks)", "winonly", "absent for Win32 APIs (to measure for dylibs)",
     "callback thunks on Windows x64 only"),
    ("Language", "#If Win64 / Win32 / Mac / VBA7", "pure", "Mac = True, Win32 = Win64 = False, VBA7 = True",
     "Linux and wasm are treated as 64-bit Windows; macOS profile as Mac"),
    ("Language", "LongPtr / pointer width", "pure", "64-bit", "the target pointer width; wasm32 is 32-bit (to decide: keep 64-bit VBA semantics)"),
    ("Language", "Early-bound references (Scripting, ADODB, MSXML, ...)", "com", "absent on Mac", "type libraries need Windows COM"),
    ("Statements", "Open / Close / Get / Put", "file", MAC_PATHS, PATH_TODAY + "; Open creates missing parent folders (VBA: 76)"),
    ("Statements", "Print # / Write # / Line Input #", "codepage", "line end and code page on Mac (to measure)",
     "CRLF everywhere; text as UTF-8, not ANSI"),
    ("Statements", "Name ... As", "file", MAC_PATHS, PATH_TODAY + "; creates the target's parent folders"),
    ("Statements", "Option Compare Text, Like", L, "Mac locale rules (to measure)", "Rust case folding off Windows"),
    ("Runtime", "^ operator", "math", "as Sin", "platform libm pow"),
    ("Runtime", "Execution engine (JIT)", "pure", "",
     "VM3 everywhere; the JIT needs runtime code generation: not in wasm, MAP_JIT on Apple Silicon"),
    ("Runtime", "Blocking calls (MsgBox, InputBox, DoEvents, Wait-style loops)", "ui", "same as Windows",
     "synchronous; a browser main thread cannot block: run VBA in a worker with a synchronous host channel"),
]


def verdicts(kind):
    v = list(BY_CLASS[kind])
    return v


def main():
    rows = list(csv.DictReader(open(os.path.join(HERE, "functions.tsv"), encoding="utf-8"), delimiter="\t"))
    out = []
    for r in rows:
        if r["module"] in ("Collection", "ErrObject", "RegExp", "Match", "MatchCollection", "SubMatches"):
            kind, mac, today = "pure", "same; RegExp in VBA 7.1 for Mac (to measure)" if r["module"] in ("RegExp", "Match", "MatchCollection", "SubMatches") else "same", ""
        else:
            kind, mac, today = F.get(r["name"], ("pure", "same (to measure)", ""))
        out.append((r["module"], r["name"] + ("/$" if r["dollar_form"] else ""), CLASS_OF[kind], mac, *verdicts(kind), today))
    for area, feature, kind, mac, today in EXTRA:
        v = verdicts(kind)
        if feature.startswith("Declare into Mac"):
            v = ["reject", "native", "replace", "reject", "reject"]
        if feature.startswith("Execution engine"):
            v = ["native", "shim", "native", "shim", "shim"]
        out.append((area, feature, CLASS_OF[kind], mac, *v, today))
    with open(os.path.join(HERE, "portability-matrix.tsv"), "w", encoding="utf-8", newline="\n") as f:
        f.write("area\tfeature\tclass\tmac_dialect\t" + "\t".join(HOSTS) + "\ttoday\n")
        for row in out:
            f.write("\t".join(row) + "\n")
    from collections import Counter
    print("classes:", Counter(r[2] for r in out))
    for i, h in enumerate(HOSTS):
        print(h, Counter(r[4 + i] for r in out))


if __name__ == "__main__":
    main()
