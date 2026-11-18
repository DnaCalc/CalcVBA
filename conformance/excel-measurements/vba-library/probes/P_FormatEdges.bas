Attribute VB_Name = "P_FormatEdges"
Option Explicit
' UNMEASURED Empty/Null precedence, mask alignment and UTF16 source characters.
Private Const N As Long = 16
Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Format(Empty, "@;empty")
        Case 2: v = Format(Empty, "0;-0;zero;null")
        Case 3: v = Format$(Null)
        Case 4: v = Format(Null, Null)
        Case 5: v = Format$(Null, Null)
        Case 6: v = Format(5, , 99)
        Case 7: v = Format(5, , , 99)
        Case 8: v = Format(Empty, "0", 99)
        Case 9: v = Format(CVErr(5))
        Case 10: v = Format(CVErr(5), Null)
        Case 11: v = Format("x", " @ ")
        Case 12: v = Format("abc", "!@@-@@")
        Case 13: v = Format("abc", "\@")
        Case 14: v = Format(ChrW(&HD83D) & ChrW(&HDE00), "@@@")
        Case 15: v = Format(ChrW(&HD800), "@@")
        Case 16: v = Format$(Null, "0;-0;zero;""NULL""")
    End Select
    T = "FormatEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "FormatEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
