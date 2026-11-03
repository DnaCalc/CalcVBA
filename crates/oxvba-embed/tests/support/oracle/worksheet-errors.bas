Option Explicit

Private Declare PtrSafe Sub CopyProbeBytes Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As LongPtr)
Private objectKept As Range

Private Function KeepStaticObject(ByVal value As Range, ByVal replace As Boolean) As Range
    Static saved As Range
    If replace Then Set saved = value
    Set KeepStaticObject = saved
End Function

Public Function OwningObjectProbe() As String
    Dim original As Range, returned As Range, current As Object, cleared As Range
    Set original = ThisWorkbook.Worksheets(1).Range("D1")
    original.Value2 = 41
    Set objectKept = original
    Set current = objectKept
    Set returned = KeepStaticObject(original, True)
    Set original = Nothing
    Set objectKept = Nothing
    returned.Value2 = 43
    Set objectKept = KeepStaticObject(Nothing, False)
    OwningObjectProbe = TypeName(objectKept) & "|" & TypeName(current) & "|" & _
        CStr(IsObject(objectKept)) & "|" & CStr(objectKept Is current) & "|" & CStr(objectKept.Value2)
    Set cleared = KeepStaticObject(Nothing, True)
    Set objectKept = Nothing
    OwningObjectProbe = OwningObjectProbe & "|" & TypeName(returned) & "|" & CStr(returned Is current)
    Dim wrong As Object, unused As Double, parameterError As Long
    Set wrong = ThisWorkbook.Worksheets(1)
    On Error Resume Next
    Err.Clear
    unused = UnusedTypedObject(wrong)
    parameterError = Err.Number
    On Error GoTo 0
    OwningObjectProbe = OwningObjectProbe & "|" & CStr(parameterError) & "|" & _
        CStr(UnusedTypedObject(Nothing)) & "|" & CStr(UnusedTypedObject(returned))
End Function

Private Function UnusedTypedObject(ByVal value As Range) As Double
    UnusedTypedObject = 7
End Function

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
