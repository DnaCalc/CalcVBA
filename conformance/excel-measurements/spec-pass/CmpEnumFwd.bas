Attribute VB_Name = "CmpEnumFwd"
Enum E3
    X = Y
    Y = 1
End Enum
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "enumFwd=" & Sh(X) & vbLf: Exit Function
fail:
    Run = "enumFwd=ERR" & Err.Number & vbLf
End Function
