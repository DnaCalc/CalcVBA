Option Explicit
Private ordinaryCalls As Long
Private volatileCalls As Long
Public Function StateUdf(x) As Long
    ordinaryCalls = ordinaryCalls + 1
    StateUdf = ordinaryCalls
End Function
Public Function VolStateUdf(x) As Long
    Application.Volatile
    volatileCalls = volatileCalls + 1
    VolStateUdf = volatileCalls
End Function
Public Function OptionalProbe(Optional a As Variant, Optional n As Long = 7) As String
    OptionalProbe = CStr(IsMissing(a)) & ":" & CStr(VarType(a)) & ":" & CStr(n)
End Function
Public Function PackProbe(ParamArray a() As Variant) As String
    Dim i As Long, text As String
    text = CStr(LBound(a)) & ":" & CStr(UBound(a))
    For i = LBound(a) To UBound(a)
        text = text & ":" & CStr(VarType(a(i)))
    Next i
    PackProbe = text
End Function
Public Function RangeProbe(r As Range) As String
    RangeProbe = CStr(r.Count) & ":" & CStr(r.Cells(1, 1).Value2) & ":" & Application.Caller.Address
End Function
Public Function WfProbe(r As Range) As Double
    WfProbe = Application.WorksheetFunction.Sum(r)
End Function
Public Function WfFault() As Long
    On Error GoTo Failed
    WfFault = Application.WorksheetFunction.Match("missing", Array(1, 2), 0)
    Exit Function
Failed:
    WfFault = Err.Number
End Function
Public Function FailProbe()
    Err.Raise 5
End Function
Public Function ErrorProbe()
    ErrorProbe = CVErr(2042)
End Function
Public Function HandledFault()
    On Error Resume Next
    Err.Raise 5
    HandledFault = Err.Number
End Function
Public Function ReadErr()
    ReadErr = Err.Number
End Function
Public Function RunProbe() As String
    On Error GoTo RunFailed
    Dim stage As String
    stage = "start"
    Dim ws As Worksheet, text As String, i As Long, oldCalc As Long
    Set ws = ThisWorkbook.Worksheets(1)
    stage = "calculation"
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    stage = "write inputs"
    ws.Range("A1").Value2 = 40
    ws.Range("B1").Formula = "=StateUdf(A1)"
    ws.Range("C1").Formula = "=VolStateUdf(A1)"
    stage = "calculate"
    Application.Calculate
    text = "state=" & ws.Range("B1").Value2 & ":" & ws.Range("C1").Value2
    Application.Calculate
    text = text & "," & ws.Range("B1").Value2 & ":" & ws.Range("C1").Value2
    ws.Range("A2").Value2 = 12
    Application.Calculate
    text = text & "," & ws.Range("B1").Value2 & ":" & ws.Range("C1").Value2
    ws.Range("A1").Value2 = 41
    Application.Calculate
    text = text & "," & ws.Range("B1").Value2 & ":" & ws.Range("C1").Value2
    ws.Range("B1").Calculate
    text = text & ",forced=" & ws.Range("B1").Value2
    Dim formulas As Variant
    formulas = Array("=OptionalProbe()", "=OptionalProbe(,9)", "=OptionalProbe(A3)", "=OptionalProbe(2)", "=PackProbe()", "=PackProbe(1,2,3)", "=PackProbe(,2)", "=RangeProbe(A1:A2)", "=WfProbe(A1:A2)", "=WfFault()", "=ErrorProbe()", "=HandledFault()", "=ReadErr()")
    For i = 0 To UBound(formulas)
        stage = "formula " & CStr(i)
        ws.Cells(i + 1, 4).Formula = formulas(i)
        ws.Cells(i + 1, 4).Calculate
        text = text & "|" & CStr(ws.Cells(i + 1, 4).Value2)
    Next i
    Application.Calculation = oldCalc
    RunProbe = text
    Exit Function
RunFailed:
    RunProbe = "fault:" & CStr(Err.Number) & ":" & Err.Description & ":" & stage & ":" & text
End Function
