Attribute VB_Name = "P_Minute"
Option Explicit
' Minute (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 24

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Minute(TimeSerial(23, 59, 59))   ' last minute
        Case 2: v = Minute(0)   ' serial 0
        Case 3: v = Minute(0.5)   ' noon
        Case 4: v = Minute(0.99999999)   ' just under a day: rounds to the second
        Case 5: v = Minute(-0.25)   ' negative fraction
        Case 6: v = Minute(1.5)   ' date with time
        Case 7: v = Minute(Null)   ' Null propagates
        Case 8: v = Minute(Empty)   ' Empty is 0
        Case 9: v = Minute("13:45")   ' string time ' locale
        Case 10: v = Minute("1:45 PM")   ' string time with PM ' locale
        Case 11: v = Minute("10:75")   ' invalid string: error 13
        Case 12: v = Minute("abc")   ' invalid string: error 13
        Case 13: v = Minute("")   ' empty string: error 13
        Case 14: v = Minute("0.5")   ' numeric string
        Case 15: v = Minute(True)   ' Boolean -1
        Case 16: v = Minute(2958466)   ' beyond upper limit: error 5
        Case 17: v = Minute(-657435)   ' below lower limit: error 5
        Case 18: v = Minute(Time:=TimeSerial(6, 15, 30))   ' named argument
        Case 19: v = TypeName(Minute(0.5))   ' result type
        Case 20: v = TypeName(Minute(Null))   ' result type for Null
        Case 21: v = Minute(#3:04:05 PM#)   ' time literal
        Case 22: v = Minute(TimeSerial(0, 59, 30) + 0.0000001)   ' just past 59 minutes 30 seconds
        Case 23: v = Minute(TimeSerial(0, 0, 59.5))   ' seconds fraction rounds in TimeSerial
        Case 24: v = Minute(CDate(43831.0104166667))   ' 00:15
    End Select
    T = "Minute." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "Minute." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
