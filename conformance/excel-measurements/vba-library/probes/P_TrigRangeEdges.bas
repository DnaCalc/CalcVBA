Attribute VB_Name = "P_TrigRangeEdges"
Option Explicit
' Excel-backed trig admission/error boundaries; numeric precision remains in the function probes.
Private Const N As Long = 40
Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = TypeName(Sin((2# ^ 63) - 1024#))
        Case 2: v = TypeName(Sin(2# ^ 63))
        Case 3: v = TypeName(Sin((2# ^ 63) + 2048#))
        Case 4: v = TypeName(Sin(-(2# ^ 63) + 1024#))
        Case 5: v = TypeName(Sin(-(2# ^ 63)))
        Case 6: v = TypeName(Sin(-(2# ^ 63) - 2048#))
        Case 7: v = TypeName(Sin("9223372036854775808"))
        Case 8: v = TypeName(Cos((2# ^ 63) - 1024#))
        Case 9: v = TypeName(Cos(2# ^ 63))
        Case 10: v = TypeName(Cos((2# ^ 63) + 2048#))
        Case 11: v = TypeName(Cos(-(2# ^ 63) + 1024#))
        Case 12: v = TypeName(Cos(-(2# ^ 63)))
        Case 13: v = TypeName(Cos(-(2# ^ 63) - 2048#))
        Case 14: v = TypeName(Cos("9223372036854775808"))
        Case 15: v = TypeName(Tan((2# ^ 63) - 1024#))
        Case 16: v = TypeName(Tan(2# ^ 63))
        Case 17: v = TypeName(Tan((2# ^ 63) + 2048#))
        Case 18: v = TypeName(Tan(-(2# ^ 63) + 1024#))
        Case 19: v = TypeName(Tan(-(2# ^ 63)))
        Case 20: v = TypeName(Tan(-(2# ^ 63) - 2048#))
        Case 21: v = TypeName(Tan("9223372036854775808"))
        Case 22: v = TypeName(Atn(1E300))
        Case 23: v = TypeName(Sin(9.223372E18 - 1024#))
        Case 24: v = TypeName(Sin(9.223372E18))
        Case 25: v = TypeName(Sin(9.223372E18 + 1024#))
        Case 26: v = TypeName(Sin(-9.223372E18 + 1024#))
        Case 27: v = TypeName(Sin(-9.223372E18))
        Case 28: v = TypeName(Sin(-9.223372E18 - 1024#))
        Case 29: v = TypeName(Cos(9.223372E18 - 1024#))
        Case 30: v = TypeName(Cos(9.223372E18))
        Case 31: v = TypeName(Cos(9.223372E18 + 1024#))
        Case 32: v = TypeName(Cos(-9.223372E18 + 1024#))
        Case 33: v = TypeName(Cos(-9.223372E18))
        Case 34: v = TypeName(Cos(-9.223372E18 - 1024#))
        Case 35: v = TypeName(Tan(9.223372E18 - 1024#))
        Case 36: v = TypeName(Tan(9.223372E18))
        Case 37: v = TypeName(Tan(9.223372E18 + 1024#))
        Case 38: v = TypeName(Tan(-9.223372E18 + 1024#))
        Case 39: v = TypeName(Tan(-9.223372E18))
        Case 40: v = TypeName(Tan(-9.223372E18 - 1024#))
    End Select
    T = "TrigRangeEdges." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "TrigRangeEdges." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
