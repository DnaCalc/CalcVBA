Attribute VB_Name = "CmpEnumDbl"
Enum E1
    A = 1.5
    B
End Enum
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "enumDbl=" & Sh(A & " " & B) & vbLf: Exit Function
fail:
    Run = "enumDbl=ERR" & Err.Number & vbLf
End Function
