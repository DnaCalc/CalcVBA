Attribute VB_Name = "CmpPropOptional"
Private m As Long
Property Get Q(Optional i As Long) As Long
    Q = m
End Property
Property Let Q(i As Long, ByVal v As Long)
    m = v
End Property
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "propOptional=" & Sh(Q(1)) & vbLf: Exit Function
fail:
    Run = "propOptional=ERR" & Err.Number & vbLf
End Function
