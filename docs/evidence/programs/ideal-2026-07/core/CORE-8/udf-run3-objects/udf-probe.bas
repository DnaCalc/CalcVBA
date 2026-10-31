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
Public Function ArgType(a) As String
    ArgType = TypeName(a) & ":" & CStr(IsObject(a))
End Function
Public Function EchoUdf(a)
    EchoUdf = a
End Function
Public Function EchoSetUdf(a)
    Set EchoSetUdf = a
End Function
Public Function TypedLongUdf(x As Long) As Long
    TypedLongUdf = x
End Function
Public Function TypedDoubleUdf(x As Double) As Double
    TypedDoubleUdf = x
End Function
Public Function WfNamed(r As Range, which As String)
    WfNamed = CallByName(Application.WorksheetFunction, which, VbMethod, r)
End Function
Public Function WfShape(which As String)
    On Error GoTo Failed
    Dim wf As Object
    Set wf = Application.WorksheetFunction
    WfShape = CallByName(wf, which, VbMethod, Array(-1, -2))
    Exit Function
Failed:
    WfShape = Err.Number
End Function
Public Function IndexProbe(r As Range, x As Double) As String
    IndexProbe = CStr(r.Cells(x, 1).Row) & ":" & CStr(r.Cells(x, 1).Value2)
End Function
Public Function NumericVolatile()
    Application.Volatile 1
    NumericVolatile = 42
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
    formulas = Array( _
        "=OptionalProbe()", "=OptionalProbe(,9)", "=OptionalProbe(A3)", "=OptionalProbe(2)", "=PackProbe()", _
        "=PackProbe(1,2,3)", "=PackProbe(,2)", "=RangeProbe(A1:A2)", "=WfProbe(A1:A2)", "=WfFault()", _
        "=ErrorProbe()", "=HandledFault()", "=ReadErr()", "=ArgType(A1)", "=ArgType(4)", _
        "=EchoUdf(A1)", "=EchoSetUdf(A1)", "=TypedLongUdf(3.5)", "=TypedLongUdf(NA())", "=TypedDoubleUdf(NA())", _
        "=WfNamed(A1:A2,""MIN"")", "=WfNamed(A1:A2,""MAX"")", "=WfNamed(A1:A2,""AVERAGE"")", "=WfShape(""ABS"")", "=WfShape(""LN"")", _
        "=WfShape(""IF"")", "=IndexProbe(A1:A2,1.5)", "=IndexProbe(A1:A2,2.5)", "=NumericVolatile()")
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
