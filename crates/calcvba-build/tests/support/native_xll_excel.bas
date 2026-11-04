Option Explicit
Private Declare PtrSafe Function OutstandingResults Lib "__XLL__" () As Long
Private Declare PtrSafe Function OwnedBuffers Lib "__XLL__" () As Long
Private Declare PtrSafe Function CloseXll Lib "__XLL__" Alias "xlAutoClose" () As Long
Private Declare PtrSafe Function AddDouble Lib "__XLL__" (ByVal a As Double, ByVal b As Double) As Double
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As LongLong) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As LongLong) As Long
Public Function ProbeNativeExports() As String
    Dim loaded As Boolean, observed As String, i As Long, value As Double
    Dim started As LongLong, finished As LongLong, frequency As LongLong
    Dim nativeCost As Double, excelCost As Double, operCost As Double
    Dim formulas() As String, expected() As Variant, caseIndex As Long, cellValue As Variant
    loaded = Application.RegisterXLL("__XLL__")
    If Not loaded Then Err.Raise 5, "XllProbe", "RegisterXLL returned False"
    Range("A1").Formula = "=CalcVbaAdd(2.25,3.5)"
    Range("A2").Formula = "=CalcVbaTwice(4.5)"
    Application.Calculate
    observed = CStr(Range("A1").Value2) & "|" & CStr(Range("A2").Value2) & "|" & CStr(OutstandingResults())
    If observed <> "5.75|9|0" Then Err.Raise 5, "XllProbe", observed
    Range("A1").Value2 = 7#
    Range("A2").ClearContents
    '__CASES__
    For caseIndex = LBound(formulas) To UBound(formulas)
        Range("B3").Formula2 = formulas(caseIndex)
        Range("B3").Calculate
        cellValue = Range("B3").Value2
        If IsError(expected(caseIndex)) Then
            If Not IsError(cellValue) Then Err.Raise 5, "XllProbe", "expected error: " & formulas(caseIndex)
            If CStr(cellValue) <> CStr(expected(caseIndex)) Then Err.Raise 5, "XllProbe", "error mismatch: " & formulas(caseIndex)
        Else
            If IsError(cellValue) Then Err.Raise 5, "XllProbe", "unexpected error: " & formulas(caseIndex)
            If VarType(expected(caseIndex)) = vbString Then
                If CStr(cellValue) <> CStr(expected(caseIndex)) Then Err.Raise 5, "XllProbe", "text mismatch: " & formulas(caseIndex)
            ElseIf cellValue <> expected(caseIndex) Then
                Err.Raise 5, "XllProbe", "value mismatch: " & formulas(caseIndex)
            End If
        End If
        If OutstandingResults() <> 0& Or OwnedBuffers() <> 0& Then Err.Raise 5, "XllProbe", "unreleased buffer: " & formulas(caseIndex)
    Next caseIndex
    QueryPerformanceFrequency frequency
    QueryPerformanceCounter started
    For i = 1 To 10000
        value = AddDouble(2.25, 3.5)
    Next
    QueryPerformanceCounter finished
    nativeCost = CDbl(finished - started) * 1000000# / CDbl(frequency) / 10000#
    QueryPerformanceCounter started
    For i = 1 To 1000
        value = Application.Run("CalcVbaAdd", 2.25, 3.5)
    Next
    QueryPerformanceCounter finished
    excelCost = CDbl(finished - started) * 1000000# / CDbl(frequency) / 1000#
    QueryPerformanceCounter started
    For i = 1 To 1000
        value = Application.Run("CalcVbaTwice", 4.5)
    Next
    QueryPerformanceCounter finished
    operCost = CDbl(finished - started) * 1000000# / CDbl(frequency) / 1000#
    If OutstandingResults() <> 0 Then Err.Raise 5, "XllProbe", "unreleased XLOPER12"
    If CloseXll() <> 1& Then Err.Raise 5, "XllProbe", "xlAutoClose failed"
    ProbeNativeExports = observed & "|Declare_us=" & CStr(nativeCost) & "|ExcelRun_us=" & CStr(excelCost) & "|ExcelOper_us=" & CStr(operCost) & "|cases=" & CStr(UBound(formulas) + 1&) & "|buffers=" & CStr(OwnedBuffers()) & "|closed=1"
End Function
