Attribute VB_Name = "CmpGotoFor"
Private Function H() As Long
    Dim i As Long, n As Long
    GoTo L
    For i = 1 To 3
L:      n = n + 1
    Next
    H = n
End Function
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "gotoFor=" & Sh(H()) & vbLf: Exit Function
fail:
    Run = "gotoFor=ERR" & Err.Number & vbLf
End Function
