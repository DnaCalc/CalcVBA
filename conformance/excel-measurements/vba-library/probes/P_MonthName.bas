Attribute VB_Name = "P_MonthName"
Option Explicit
' MonthName (Strings; names are English on this machine).
Private Const N As Long = 17

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = MonthName(1)                      ' locale
        Case 2: v = MonthName(12)                     ' locale
        Case 3: v = MonthName(1, True)                ' abbreviated; locale
        Case 4: v = MonthName(9, True)                ' Sep; locale
        Case 5: v = MonthName(5, True)                ' May: abbreviation equals the name; locale
        Case 6: v = MonthName(0)                      ' error 5
        Case 7: v = MonthName(13)                     ' error 5
        Case 8: v = MonthName(-1)                     ' error 5
        Case 9: v = MonthName(Null)                   ' error 94
        Case 10: v = MonthName(3, Null)               ' Null abbreviate: error 94
        Case 11: v = MonthName("3")                   ' numeric string
        Case 12: v = MonthName(2.5)                   ' fractional month rounds (banker's); locale
        Case 13: v = MonthName(3.5)                   ' fractional month rounds to 4
        Case 14: v = MonthName(Month:=2, Abbreviate:=False)  ' named arguments
        Case 15: v = MonthName(6, 1)                  ' numeric Abbreviate (nonzero = True)
        Case 16: v = MonthName(6, 0)                  ' numeric Abbreviate 0
        Case 17: v = MonthName(Empty)                 ' Empty month = 0: error 5
    End Select
    T = "MonthName." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "MonthName." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
