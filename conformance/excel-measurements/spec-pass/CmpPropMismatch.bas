Attribute VB_Name = "CmpPropMismatch"
Private m As Long
Property Get P(ByVal i As Long) As Long
    P = m + i
End Property
Property Let P(ByRef i As Long, ByVal v As Long)
    m = v
End Property
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "propMismatch=" & Sh(P(1)) & vbLf: Exit Function
fail:
    Run = "propMismatch=ERR" & Err.Number & vbLf
End Function
