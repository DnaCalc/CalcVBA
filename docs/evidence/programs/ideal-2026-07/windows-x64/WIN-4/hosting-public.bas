Option Explicit
Public Function HostLong(ByVal x As Long) As Long
    HostLong = x
End Function
Private Function HostPrivate() As Long
    HostPrivate = 73
End Function
Public Sub HostSub()
End Sub
Public Function HostError() As Variant
    Err.Raise 13, "HostingProbe", "controlled type error"
End Function
Public Function HostCVErr() As Variant
    HostCVErr = CVErr(2042)
End Function
Public Function HostKind(x As Variant) As String
    HostKind = TypeName(x) & ":" & CStr(VarType(x))
End Function
Private Function CellObservation(ws As Worksheet, ByVal expression As String) As String
    ws.Range("B1").Formula = expression
    ws.Range("B1").Calculate
    CellObservation = CStr(ws.Range("B1").Value2)
End Function
Private Function RunObservation(ByVal proc As String, Optional x As Variant) As String
    On Error GoTo Failed
    Dim v As Variant
    If IsMissing(x) Then
        v = Application.Run("'" & ThisWorkbook.Name & "'!OracleSelfTest." & proc)
    Else
        v = Application.Run("'" & ThisWorkbook.Name & "'!OracleSelfTest." & proc, x)
    End If
    RunObservation = TypeName(v) & ":" & CStr(v)
    Exit Function
Failed:
    RunObservation = "err=" & CStr(Err.Number) & ":" & Err.Description
End Function
Public Function RunProbe() As String
    On Error GoTo Failed
    Dim ws As Worksheet, other As Workbook, s As String
    Application.VBE.MainWindow.Visible = False
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value2 = 2.5
    s = "cell-public=" & CellObservation(ws, "=HostLong(2.5)")
    s = s & "|cell-private=" & CellObservation(ws, "=HostPrivate()")
    s = s & "|cell-sub=" & CellObservation(ws, "=HostSub()")
    s = s & "|cell-text=" & CellObservation(ws, "=HostLong(""3"")")
    s = s & "|cell-badtext=" & CellObservation(ws, "=HostLong(""bad"")")
    s = s & "|cell-cverr=" & CellObservation(ws, "=HostCVErr()")
    s = s & "|cell-range=" & CellObservation(ws, "=HostKind(A1)")
    s = s & "|cell-number=" & CellObservation(ws, "=HostKind(2.5)")
    s = s & "|run-long=" & RunObservation("HostLong", 2.5)
    s = s & "|run-private=" & RunObservation("HostPrivate")
    s = s & "|run-sub=" & RunObservation("HostSub")
    s = s & "|run-cverr=" & RunObservation("HostCVErr")
    Set other = Application.Workbooks.Add
    s = s & "|external-public=" & CellObservation(other.Worksheets(1), "='" & ThisWorkbook.Name & "'!HostLong(2.5)")
    s = s & "|external-private=" & CellObservation(other.Worksheets(1), "='" & ThisWorkbook.Name & "'!HostPrivate()")
    other.Close False
    RunProbe = s
    Exit Function
Failed:
    RunProbe = "FAILED:" & CStr(Err.Number) & ":" & Err.Description
End Function
