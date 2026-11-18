Attribute VB_Name = "P_TimeValue"
Option Explicit
' TimeValue (VBA.DateTime; MS-VBAL 6.1.2.x).
Private Const N As Long = 26

Private Function T(ByVal k As Long) As String
    Dim v As Variant, a As Variant, b As Variant, c As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = TimeValue("10:30")   ' hours and minutes
        Case 2: v = TimeValue("10:30:45")   ' with seconds
        Case 3: v = TimeValue("1:30 PM")   ' PM ' locale
        Case 4: v = TimeValue("12:00 AM")   ' midnight
        Case 5: v = TimeValue("12:00 PM")   ' noon
        Case 6: v = TimeValue("12:30:15 AM")   ' just after midnight
        Case 7: v = TimeValue("11:59:59 PM")   ' last second
        Case 8: v = TimeValue("24:00")   ' hour 24: error 13
        Case 9: v = TimeValue("10:75")   ' minute 75: error 13
        Case 10: v = TimeValue("13:00 PM")   ' 24-hour with PM
        Case 11: v = TimeValue("10")   ' hour only: error 13
        Case 12: v = TimeValue("")   ' empty string: error 13
        Case 13: v = TimeValue("abc")   ' invalid text: error 13
        Case 14: v = TimeValue(Null)   ' Null argument
        Case 15: v = TimeValue(Empty)   ' Empty argument
        Case 16: v = TimeValue("2020/01/02 10:30:00")   ' date and time, date dropped ' locale
        Case 17: v = TimeValue("0:0:0")   ' all zero
        Case 18: v = TimeValue("10:30:45.5")   ' fractional second
        Case 19: v = TimeValue("10 AM")   ' hour with AM
        Case 20: v = TimeValue("10:30 pm")   ' lower-case pm
        Case 21: v = TimeValue("  10:30  ")   ' surrounding spaces
        Case 22: v = TimeValue("1/1/2020")   ' date-only text
        Case 23: v = TimeValue("23:59:59")   ' 24-hour last second
        Case 24: v = TypeName(TimeValue("10:30"))   ' result type
        Case 25: v = TimeValue(Time:="6:15:30")   ' named argument
        Case 26: a = TimeSerial(10, 0, 0) + DateSerial(2020, 2, 29): v = TimeValue(a)   ' Date argument converted to text and back
    End Select
    T = "TimeValue." & k & "=" & Sh(v)
    Exit Function
fail:
    Dim en As Long
    en = Err.Number
    T = "TimeValue." & k & "=ERR" & en
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
