Attribute VB_Name = "P_UserForms"
Option Explicit
' UserForms collection (VBA global): no form is ever loaded here.
Private Const N As Long = 8

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, n As Long
    On Error GoTo fail
    Select Case k
        Case 1: v = TypeName(UserForms)                                    ' class name
        Case 2: v = UserForms.Count                                        ' no form loaded
        Case 3: v = TypeName(UserForms.Count)
        Case 4: v = IsObject(UserForms)
        Case 5: v = (UserForms Is Nothing)
        Case 6: n = 0: For Each x In UserForms: n = n + 1: Next: v = n     ' For Each over the empty collection
        Case 7: v = UserForms(0).Name                                      ' index into an empty collection
        Case 8: v = VarType(UserForms)
    End Select
    T = "UserForms." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "UserForms." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
