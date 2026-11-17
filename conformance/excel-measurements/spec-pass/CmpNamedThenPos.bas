Attribute VB_Name = "CmpNamedThenPos"
Private Function PA(a, ParamArray r()) As Long
    PA = UBound(r)
End Function
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "namedThenPos=" & Sh(PA(a:=1, 2)) & vbLf: Exit Function
fail:
    Run = "namedThenPos=ERR" & Err.Number & vbLf
End Function
