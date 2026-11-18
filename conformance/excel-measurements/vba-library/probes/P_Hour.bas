Attribute VB_Name = "P_Hour"
Option Explicit
' Hour (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 25

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Hour(TimeSerial(23, 59, 59))   ' last second
        Case 2: v = Hour(0)   ' serial 0
        Case 3: v = Hour(0.5)   ' noon
        Case 4: v = Hour(0.99999999)   ' just under a day: rounds to the second
        Case 5: v = Hour(-0.25)   ' negative fraction
        Case 6: v = Hour(1.5)   ' date with time
        Case 7: v = Hour(24)   ' whole number
        Case 8: v = Hour(Null)   ' Null propagates
        Case 9: v = Hour(Empty)   ' Empty is 0
        Case 10: v = Hour("13:45")   ' string time ' locale
        Case 11: v = Hour("1:45 PM")   ' string time with PM ' locale
        Case 12: v = Hour("25:00")   ' invalid string: error 13
        Case 13: v = Hour("abc")   ' invalid string: error 13
        Case 14: v = Hour("")   ' empty string: error 13
        Case 15: v = Hour("0.5")   ' numeric string
        Case 16: v = Hour(True)   ' Boolean -1
        Case 17: v = Hour(2958466)   ' beyond upper limit: error 5
        Case 18: v = Hour(-657435)   ' below lower limit: error 5
        Case 19: v = Hour(Time:=TimeSerial(6, 15, 30))   ' named argument
        Case 20: v = TypeName(Hour(0.5))   ' result type
        Case 21: v = TypeName(Hour(Null))   ' result type for Null
        Case 22: v = Hour(#3:04:05 PM#)   ' time literal
        Case 23: v = Hour(0.0416666667)   ' just above 1 AM
        Case 24: v = Hour(0.0416666)   ' just under 1 AM
        Case 25: v = Hour(CDate(43831.5))   ' date and time
    End Select
    T = "Hour." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "Hour." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
