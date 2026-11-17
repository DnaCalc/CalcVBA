Attribute VB_Name = "CmpNumDot"
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "numDot=" & Sh(.5) & vbLf: Exit Function
fail:
    Run = "numDot=ERR" & Err.Number & vbLf
End Function
