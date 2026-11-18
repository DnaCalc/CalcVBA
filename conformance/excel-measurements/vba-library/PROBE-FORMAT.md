# VBA library probes: module format

One probe module per library function: `probes/P_<Name>.bas` (for a class, one per class, `P_Collection.bas`). Each module's `Run()` returns one `key=value` line per case.

- Excel VBA runs every module in one workbook, and Excel's answers go to `probes/<Module>.excel.txt`.
- CalcVBA runs each module on its own (with `ProbeFmt.bas`) on VM3 and the JIT, so a function CalcVBA lacks breaks only its own module.
- The complete list of functions is `functions.tsv`, built from VBE7.DLL's type library and MS-VBAL §6.

```vb
Attribute VB_Name = "P_Left"
Option Explicit
' Left / Left$ (Strings; MS-VBAL 6.1.2.11.1.x).
Private Const N As Long = 9

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Left("abcdef", 3)                ' typical
        Case 2: v = Left("abc", 0)
        Case 3: v = Left("abc", 10)                  ' past the end
        Case 4: v = Left("abc", -1)                  ' negative length: error 5
        Case 5: v = Left(Null, 2)                    ' Null propagates
        Case 6: v = Left("abc", Null)                ' Null length: error 94
        Case 7: v = Left$("abc", 2)                  ' $ form: String
        Case 8: v = Left(String:="abc", Length:=2)   ' named arguments (VBE7 parameter names)
        Case 9: v = Left(12345, "2")                 ' numeric argument and numeric string
    End Select
    T = "Left." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Left." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
```

## Rules
- **Formatting:** `Sh(v)` comes from `ProbeFmt.bas`; don't redefine it. It prints the type and the exact value, including the IEEE bits for Double, Single and Date, so `v` can be any type. For an object result use `Set v = ...` and then `Sh(v)`.
- **Keys:** `<FunctionName>.<k>`, numbered 1..N. Keep a short comment on each case saying what it tests.
- **Coverage, about 8-15 cases per function, more where the behaviour is rich (Format, DateAdd, StrConv, Round...):**
  - typical values;
  - boundaries (0, negative, empty string, past the end, the largest and smallest values of the type);
  - the result type;
  - Null and Empty arguments;
  - numeric strings for numeric arguments, and numbers for string arguments;
  - error cases (5, 6, 13, 94);
  - every optional argument, and named arguments using the VBE7 parameter names (in `functions.tsv`);
  - the `$` form where one exists.
- **Locale:** this Excel uses South African regional settings (short date yyyy/mm/dd, period as the decimal separator). Results that depend on regional settings are fine to include; mark them `' locale` in the comment. Prefer locale-free inputs: `DateSerial` or `#m/d/yyyy#` literals rather than date strings, except in cases that test string parsing on purpose.
- **Determinism:** never print a value that changes from run to run (Now, Timer, Date, Time, Rnd without a seed, Shell task ids, StrPtr addresses, file times). Print its TypeName, or a relation such as `Now >= Date`, `StrPtr(vbNullString) = 0` or `VarPtr(x) <> 0`, instead.
- **Safety (these run in the owner's Excel):**
  - Never show UI: no MsgBox or InputBox calls that would display, no SendKeys, no AppActivate on a real window.
  - Never launch a program. For Shell, only error cases such as a missing file.
  - Beep is fine.
  - CreateObject/GetObject: only `"Scripting.Dictionary"` and error cases (a bad ProgID).
  - Registry: SaveSetting/GetSetting/DeleteSetting/GetAllSettings only with AppName `"CalcVbaProbe"`, and the module deletes what it saved.
  - Files only under `Environ("TEMP") & "\cvp-lib"`. Create the folder in the case and clean it up afterwards (Close, Kill, RmDir), with `On Error Resume Next` in the cleanup.
- **Syntax:**
  - The module must compile in Excel VBA. Keep to plain VBA7 syntax, and declare everything (Option Explicit).
  - A case that would be a compile error (a wrong argument count, a named argument that doesn't exist) does not belong here; leave it out.
  - Use `CVErr`, `Array`, `DateSerial` and so on freely.
- **Order:** one function per module, cases in the order typical → boundaries → Null/Empty/types → errors → named/$ forms.
