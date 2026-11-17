Attribute VB_Name = "CmpDate3PM"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "date3PM=" & Sh(#3 PM#) & vbLf: Exit Function
fail:
    Run = "date3PM=ERR" & Err.Number & vbLf
End Function
