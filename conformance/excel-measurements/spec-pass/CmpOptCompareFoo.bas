Attribute VB_Name = "CmpOptCompareFoo"
Option Compare Foo
Private Function Sh(ByVal v As Variant) As String
    If VarType(v) = vbDate Then Sh = "Date:" & CStr(CDbl(v)) Else Sh = TypeName(v) & ":" & CStr(v)
End Function
Public Function Run() As String
    On Error GoTo fail
    Run = "optCompareFoo=" & Sh(("a" = "A")) & vbLf: Exit Function
fail:
    Run = "optCompareFoo=ERR" & Err.Number & vbLf
End Function
