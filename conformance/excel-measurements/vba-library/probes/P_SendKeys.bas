Attribute VB_Name = "P_SendKeys"
Option Explicit
' SendKeys (Interaction; never called, only the key constants a SendKeys string would use).
Private Const N As Long = 3

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = Asc(vbTab) & "," & Asc(vbCr) & "," & Asc(vbLf)                        ' tab, return, line feed codes
        Case 2: v = Len(vbCrLf)                                                           ' vbCrLf length
        Case 3: v = TypeName(vbTab)                                                       ' String constant
    End Select
    T = "SendKeys." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "SendKeys." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
