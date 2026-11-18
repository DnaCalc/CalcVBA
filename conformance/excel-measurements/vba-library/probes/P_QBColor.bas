Attribute VB_Name = "P_QBColor"
Option Explicit
' QBColor (Information; 0-15 map to RGB values, others error 5).
Private Const N As Long = 31

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim i As Integer, j As Long, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = QBColor(0)                        ' black
        Case 2: v = QBColor(1)                        ' blue
        Case 3: v = QBColor(2)
        Case 4: v = QBColor(3)
        Case 5: v = QBColor(4)
        Case 6: v = QBColor(7)                        ' white (grey)
        Case 7: v = QBColor(8)                        ' grey
        Case 8: v = QBColor(9)
        Case 9: v = QBColor(12)
        Case 10: v = QBColor(14)
        Case 11: v = QBColor(15)                       ' bright white
        Case 12: For j = 0 To 15: s = s & QBColor(j) & ",": Next: v = s   ' all sixteen in order
        Case 13: v = QBColor(-1)                       ' error 5
        Case 14: v = QBColor(16)                       ' error 5
        Case 15: v = QBColor(255)                      ' error 5
        Case 16: v = QBColor(32767)                    ' largest Integer: error 5
        Case 17: v = QBColor(32768)                    ' beyond Integer: overflow 6
        Case 18: v = QBColor(2.5)                      ' banker's rounding: 2
        Case 19: v = QBColor(3.5)                      ' rounds to 4
        Case 20: v = QBColor(15.4)                     ' rounds to 15
        Case 21: v = QBColor(15.5)                     ' rounds to 16: error 5
        Case 22: v = QBColor(-0.4)                     ' rounds to 0
        Case 23: v = QBColor("3")                      ' numeric string
        Case 24: v = QBColor("x")                      ' error 13
        Case 25: v = QBColor(Null)                     ' error 94
        Case 26: v = QBColor(Empty)                    ' 0
        Case 27: v = QBColor(True)                     ' -1: error 5
        Case 28: v = QBColor(False)                    ' 0
        Case 29: i = 9: v = QBColor(i)                 ' Integer variable
        Case 30: v = TypeName(QBColor(1))              ' result type: Long
        Case 31: v = QBColor(Color:=4)                 ' named argument
    End Select
    T = "QBColor." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "QBColor." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
