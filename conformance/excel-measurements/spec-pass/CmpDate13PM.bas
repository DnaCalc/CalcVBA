Attribute VB_Name = "CmpDate13PM"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "date13PM=" & Sh(#1/1/2026 13:00 PM#) & vbLf: Exit Function
fail:
    Run = "date13PM=ERR" & Err.Number & vbLf
End Function
