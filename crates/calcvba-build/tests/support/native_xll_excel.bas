Option Explicit
Private Declare PtrSafe Function OutstandingResults Lib "__XLL__" () As Long
Private Declare PtrSafe Function AddDouble Lib "__XLL__" (ByVal a As Double, ByVal b As Double) As Double
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As LongLong) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As LongLong) As Long
Public Function ProbeNativeExports() As String
    Dim loaded As Boolean, observed As String, i As Long, value As Double
    Dim started As LongLong, finished As LongLong, frequency As LongLong
    Dim nativeCost As Double, excelCost As Double, operCost As Double
    loaded = Application.RegisterXLL("__XLL__")
    If Not loaded Then Err.Raise 5, "XllProbe", "RegisterXLL returned False"
    Range("A1").Formula = "=CalcVbaAdd(2.25,3.5)"
    Range("A2").Formula = "=CalcVbaTwice(4.5)"
    Application.Calculate
    observed = CStr(Range("A1").Value2) & "|" & CStr(Range("A2").Value2) & "|" & CStr(OutstandingResults())
    If observed <> "5.75|9|0" Then Err.Raise 5, "XllProbe", observed
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
    ProbeNativeExports = observed & "|Declare_us=" & CStr(nativeCost) & "|ExcelRun_us=" & CStr(excelCost) & "|ExcelOper_us=" & CStr(operCost)
End Function
