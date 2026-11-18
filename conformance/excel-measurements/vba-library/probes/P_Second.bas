Attribute VB_Name = "P_Second"
Option Explicit
' Second (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 25

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Second(TimeSerial(23, 59, 59))   ' last second
        Case 2: v = Second(0)   ' serial 0
        Case 3: v = Second(0.5)   ' noon
        Case 4: v = Second(0.99999999)   ' just under a day: rounds to the second
        Case 5: v = Second(-0.25)   ' negative fraction
        Case 6: v = Second(1.5)   ' date with time
        Case 7: v = Second(Null)   ' Null propagates
        Case 8: v = Second(Empty)   ' Empty is 0
        Case 9: v = Second("13:45:30")   ' string time ' locale
        Case 10: v = Second("1:45:20 PM")   ' string time with PM ' locale
        Case 11: v = Second("10:00:75")   ' invalid string: error 13
        Case 12: v = Second("abc")   ' invalid string: error 13
        Case 13: v = Second("")   ' empty string: error 13
        Case 14: v = Second("0.5")   ' numeric string
        Case 15: v = Second(True)   ' Boolean -1
        Case 16: v = Second(2958466)   ' beyond upper limit: error 5
        Case 17: v = Second(-657435)   ' below lower limit: error 5
        Case 18: v = Second(Time:=TimeSerial(6, 15, 30))   ' named argument
        Case 19: v = TypeName(Second(0.5))   ' result type
        Case 20: v = TypeName(Second(Null))   ' result type for Null
        Case 21: v = Second(#3:04:05 PM#)   ' time literal
        Case 22: v = Second(0.0000115740740740741)   ' exactly one second as a serial
        Case 23: v = Second(0.0000115)   ' just under one second
        Case 24: v = Second(0.0000116)   ' just over one second
        Case 25: v = Second(CDate(43831.0001157407))   ' about 10 seconds
    End Select
    T = "Second." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "Second." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
