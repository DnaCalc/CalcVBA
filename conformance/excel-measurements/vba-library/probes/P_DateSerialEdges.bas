Attribute VB_Name = "P_DateSerialEdges"
Option Explicit
' Excel-backed relative negative years and normalized month/day boundaries.
Private Const N As Long = 23
Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = DateSerial(-1, 1, 1)
        Case 2: v = DateSerial(-2, 1, 1)
        Case 3: v = DateSerial(-29, 1, 1)
        Case 4: v = DateSerial(-30, 1, 1)
        Case 5: v = DateSerial(-49, 1, 1)
        Case 6: v = DateSerial(-50, 1, 1)
        Case 7: v = DateSerial(-99, 1, 1)
        Case 8: v = DateSerial(-100, 1, 1)
        Case 9: v = DateSerial(-1899, 1, 1)
        Case 10: v = DateSerial(-1900, 1, 1)
        Case 11: v = DateSerial(-1901, 1, 1)
        Case 12: v = DateSerial(-1999, 1, 1)
        Case 13: v = DateSerial(-2000, 1, 1)
        Case 14: v = DateSerial(-2001, 1, 1)
        Case 15: v = DateSerial(-32768, 1, 1)
        Case 16: v = DateSerial(0, 1, 1)
        Case 17: v = DateSerial(29, 1, 1)
        Case 18: v = DateSerial(99, 1, 1)
        Case 19: v = DateSerial(100, 1, 1)
        Case 20: v = DateSerial(-1, 13, 1)
        Case 21: v = DateSerial(-1, 0, 0)
        Case 22: v = DateSerial(True, 1, 1)
        Case 23: v = DateSerial(-1900, 1, 0)
    End Select
    T = "DateSerialEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "DateSerialEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
