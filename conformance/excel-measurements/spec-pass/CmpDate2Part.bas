Attribute VB_Name = "CmpDate2Part"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "date2Part=" & Sh((Year(#1/15#) = Year(Date)) & " " & Month(#1/15#) & "/" & Day(#1/15#)) & vbLf: Exit Function
fail:
    Run = "date2Part=ERR" & Err.Number & vbLf
End Function
