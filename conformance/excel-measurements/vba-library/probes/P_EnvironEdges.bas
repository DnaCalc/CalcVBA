Attribute VB_Name = "P_EnvironEdges"
Option Explicit
' UNMEASURED index boundaries and name trimming. Values are never printed.
Private Const N As Long = 22
Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = TypeName(Environ(255))
        Case 2: v = TypeName(Environ(256))
        Case 3: v = TypeName(Environ(1024))
        Case 4: v = TypeName(Environ(32767))
        Case 5: v = TypeName(Environ(32768))
        Case 6: v = TypeName(Environ(2147483647))
        Case 7: v = TypeName(Environ(2147483648#))
        Case 8: v = TypeName(Environ(0.5))
        Case 9: v = Environ(1.5) = Environ(2)
        Case 10: v = TypeName(Environ(-0.5))
        Case 11: v = Environ(" TEMP") = Environ("TEMP")
        Case 12: v = Environ("TEMP ") = Environ("TEMP")
        Case 13: v = Environ(vbTab & "TEMP") = Environ("TEMP")
        Case 14: v = Environ(ChrW(&HA0) & "TEMP") = Environ("TEMP")
        Case 15: v = TypeName(Environ("256"))
        Case 16: v = TypeName(Environ(CVErr(5)))
        Case 17: v = Environ(CDec(1.5)) = Environ(2)
        Case 18: v = TypeName(Environ(CDate(1)))
        Case 19: v = TypeName(Environ(False))
        Case 20: v = Environ(ChrW(&H3000) & "TEMP") = Environ("TEMP")
        Case 21: v = Environ("TEMP" & vbNullChar & "SUFFIX") = Environ("TEMP")
        Case 22: v = TypeName(Environ("="))
    End Select
    T = "EnvironEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "EnvironEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
