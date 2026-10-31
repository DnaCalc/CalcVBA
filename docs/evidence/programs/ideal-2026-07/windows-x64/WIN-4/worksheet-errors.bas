Option Explicit

Private Declare PtrSafe Sub CopyProbeBytes Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As LongPtr)

Public Function RaiseMismatch() As Variant
    Err.Raise 13
End Function

Public Function DivideByZero(ByVal divisor As Double) As Double
    DivideByZero = 1 / divisor
End Function

Public Function UntypedPlus(value)
    UntypedPlus = value + 1
End Function

Public Function UntypedTimes(value)
    UntypedTimes = value * 2
End Function

Public Function UntypedIdentity(value) As String
    UntypedIdentity = TypeName(value) & ":" & CStr(VarType(value))
End Function

Public Function DoubleOverflowCode() As Long
    On Error GoTo Failed
    Err.Clear
    Dim first As Double, second As Double, result As Double
    first = 1E+308
    second = 2
    result = first * second
    DoubleOverflowCode = 0
    Exit Function
Failed:
    DoubleOverflowCode = Err.Number
End Function

Public Function DoubleOverflowCompare() As String
    On Error Resume Next
    Err.Clear
    Dim x As Double
    x = 1E+308 * 1E+308
    Dim probeError As Long, parts(0 To 1) As Long
    probeError = Err.Number
    CopyProbeBytes parts(0), x, 8
    DoubleOverflowCompare = CStr(probeError) & ":" & CStr(x = x) & ":" & CStr(x > 0) & ":" & Right$("00000000" & Hex$(parts(1)), 8) & Right$("00000000" & Hex$(parts(0)), 8)
End Function

' Exists only for supervisor macro-presence checks. The worker never runs it:
' worksheet formulas are calculated from COM while VBA is idle.
Public Function RunProbe() As String
    RunProbe = "worksheet-idle-probe"
End Function
