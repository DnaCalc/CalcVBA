Attribute VB_Name = "CmpDateYY"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "dateYY=" & Sh(#1/1/99#) & vbLf: Exit Function
fail:
    Run = "dateYY=ERR" & Err.Number & vbLf
End Function
