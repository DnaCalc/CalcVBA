Attribute VB_Name = "P_TimeSerial"
Option Explicit
' TimeSerial (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 26

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = TimeSerial(12, 30, 45)   ' typical
        Case 2: v = TimeSerial(0, 0, 0)   ' zero
        Case 3: v = TimeSerial(24, 0, 0)   ' 24 hours rolls over
        Case 4: v = TimeSerial(23, 59, 59)   ' last second of the day
        Case 5: v = TimeSerial(23, 59, 60)   ' second overflow
        Case 6: v = TimeSerial(0, 90, 0)   ' minute overflow
        Case 7: v = TimeSerial(0, 0, 3600)   ' second overflow to an hour
        Case 8: v = TimeSerial(0, 0, 86400)   ' a whole day of seconds
        Case 9: v = TimeSerial(-1, 0, 0)   ' negative hour
        Case 10: v = TimeSerial(0, -30, 0)   ' negative minute
        Case 11: v = TimeSerial(0, 0, -1)   ' negative second
        Case 12: v = TimeSerial(-24, 0, 0)   ' negative whole day
        Case 13: v = TimeSerial(12, -30, 0)   ' mixed signs
        Case 14: v = TimeSerial(48, 0, 0)   ' two days of hours
        Case 15: a = 32767: v = TimeSerial(a, 59, 59)   ' largest Integer hours
        Case 16: a = -32768: v = TimeSerial(a, 0, 0)   ' smallest Integer hours
        Case 17: a = 32768: v = TimeSerial(a, 0, 0)   ' beyond Integer: overflow error 6
        Case 18: a = 12.5: v = TimeSerial(a, 0, 0)   ' fractional hour rounds to even
        Case 19: a = 13.5: v = TimeSerial(a, 0.5, 0.5)   ' fractional parts round to even
        Case 20: a = "12": b = "30": c = "15": v = TimeSerial(a, b, c)   ' numeric strings
        Case 21: a = Empty: v = TimeSerial(a, 1, 1)   ' Empty hour is 0
        Case 22: a = Null: v = TimeSerial(a, 1, 1)   ' Null argument: error 94
        Case 23: a = "abc": v = TimeSerial(a, 1, 1)   ' non-numeric string: error 13
        Case 24: v = TimeSerial(Hour:=6, Minute:=15, Second:=30)   ' named arguments
        Case 25: v = TimeSerial(Second:=30, Minute:=15, Hour:=6)   ' named arguments out of order
        Case 26: v = TimeSerial(True, 0, 0)   ' Boolean argument is -1
    End Select
    T = "TimeSerial." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "TimeSerial." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
