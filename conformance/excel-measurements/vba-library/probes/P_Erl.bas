Attribute VB_Name = "P_Erl"
Option Explicit
' Erl (Information; line number of the last executed numbered line when an error occurred).
Private Const N As Long = 11

Private Function E1() As Long
    Dim x As Long
10  On Error GoTo h
30  x = 1 \ x
40  E1 = -1
    Exit Function
h:
    E1 = Erl
End Function

Private Function E2() As Long
    Dim x As Long
    On Error GoTo h
    x = 1 \ x
    E2 = -1
    Exit Function
h:
    E2 = Erl
End Function

Private Function E3() As Long
100 On Error GoTo h
200 Err.Raise 5
300 E3 = -1
    Exit Function
h:
    E3 = Erl
End Function

Private Function E4() As Long
    Dim x As Long
1   On Error GoTo h
2   x = 0
    x = 1 \ x                 ' unnumbered line fails: Erl keeps the last numbered line
    E4 = -1
    Exit Function
h:
    E4 = Erl
End Function

Private Function E5() As Long
65535 On Error GoTo h
    Err.Raise 5
    E5 = -1
    Exit Function
h:
    E5 = Erl
End Function

Private Function E6() As String
    Dim x As Long
7   On Error Resume Next
8   x = 1 \ x
9   E6 = TypeName(Erl) & ":" & Erl & ":" & Err.Number
End Function

Private Function E7() As Long
11  On Error GoTo h
12  Err.Raise 5
    Exit Function
h:
    Resume Next
End Function

Private Function E8() As Long
    Dim x As Long
    On Error Resume Next
1   x = 1 \ x
2   x = 0
    E8 = Erl                  ' Err cleared by line 2? Erl after a later numbered statement
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = E1()                      ' numbered lines: error at line 30
        Case 2: v = E2()                      ' no line numbers: 0
        Case 3: v = E3()                      ' Err.Raise on line 200
        Case 4: v = E4()                      ' unnumbered failing line after numbered ones
        Case 5: v = E5()                      ' largest line number 65535
        Case 6: v = E6()                      ' Erl type and value with Resume Next
        Case 7: v = E7()                      ' Resume Next after error: function returns default
        Case 8: v = E8()                      ' Erl after the error was handled
        Case 9: v = TypeName(Erl)             ' type of Erl (outside any error)
        Case 10: v = Erl                      ' Erl with no error: 0
        Case 11
            Err.Clear
20          v = Erl                           ' Erl on a numbered line with no error
    End Select
    T = "Erl." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Erl." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
