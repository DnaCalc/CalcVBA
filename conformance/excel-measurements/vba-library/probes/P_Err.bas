Attribute VB_Name = "P_Err"
Option Explicit
' Err (the function returning the ErrObject; Number, Description, Raise, Clear).
Private Const N As Long = 16

Private Function R1() As String
    On Error Resume Next
    Err.Raise 11
    R1 = Err.Number & "|" & Err.Description
End Function

Private Function R2() As String
    On Error Resume Next
    Err.Raise 1000
    R2 = Err.Number & "|" & Err.Description & "|" & (Len(Err.Source) > 0)
End Function

Private Function R3() As String
    On Error Resume Next
    Err.Raise 1000, "MySrc", "My description", "h.hlp", 77
    R3 = Err.Number & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & Err.HelpContext
End Function

Private Function R4() As String
    On Error Resume Next
    Err.Raise 5
    Err.Clear
    R4 = Err.Number & "|" & Err.Description & "|" & Err.Source & "|" & Err.HelpContext
End Function

Private Function R5() As String
    Dim x As Long
    On Error Resume Next
    x = 1 \ x
    R5 = Err.Number & "|" & Err.Description
End Function

Private Function R6() As String
    Dim a(1 To 2) As Long, x As Long
    On Error Resume Next
    x = a(3)
    R6 = Err.Number & "|" & Err.Description
End Function

Private Function R7() As String
    On Error Resume Next
    Err.Raise vbObjectError + 1000
    R7 = Err.Number & "|" & Err.Description
End Function

Private Function R8() As String
    On Error Resume Next
    Err.Description = "set by hand"
    R8 = Err.Number & "|" & Err.Description
End Function

Private Function R9() As String
    On Error Resume Next
    Err.Raise 0
    R9 = Err.Number & "|" & Err.Description
End Function

Private Function R10() As String
    On Error Resume Next
    Err.Raise 70000
    R10 = Err.Number & "|" & Err.Description
End Function

Private Function R11() As String
    On Error Resume Next
    Err.Raise 13
    R11 = Err.Number & "|" & Err.Description & "|" & Err.LastDllError
End Function

Private Function T(ByVal k As Long) As String
    Dim v As Variant, o As Object
    On Error GoTo fail
    Select Case k
        Case 1: v = R1()                        ' Err.Raise 11: number and description
        Case 2: v = R2()                        ' application-defined error text, Source non-empty
        Case 3: v = R3()                        ' all Raise arguments
        Case 4: v = R4()                        ' Clear resets everything
        Case 5: v = R5()                        ' runtime division by zero
        Case 6: v = R6()                        ' subscript out of range
        Case 7: v = R7()                        ' vbObjectError + 1000
        Case 8: v = R8()                        ' Description set directly
        Case 9: v = R9()                        ' Raise 0: error 5
        Case 10: v = R10()                      ' Raise 70000
        Case 11: v = R11()                      ' type mismatch text and LastDllError
        Case 12: v = TypeName(Err)              ' ErrObject
        Case 13
            Err.Clear
            v = Err                             ' default property is Number
        Case 14
            Err.Clear
            v = Err.Number                      ' Long 0 with no error
        Case 15
            Set o = Err                         ' Err assigned to an Object variable
            v = TypeName(o)
        Case 16
            Err.Clear
            v = TypeName(Err.Description) & "|" & TypeName(Err.Source) & "|" & TypeName(Err.HelpContext)
    End Select
    T = "Err." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Err." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
