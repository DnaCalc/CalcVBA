Attribute VB_Name = "CmpEnumOver"
Enum E2
    A = 2147483647
    B
End Enum
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "enumOver=" & Sh(A) & vbLf: Exit Function
fail:
    Run = "enumOver=ERR" & Err.Number & vbLf
End Function
