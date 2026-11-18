Attribute VB_Name = "P_Ptr"
Option Explicit
' VarPtr, StrPtr, ObjPtr: only types and relations (addresses vary).
Private Const N As Long = 24

Private Function T(ByVal k As Long) As String
    Dim v As Variant, x As Variant, y As Variant, s As String, s2 As String, c As Collection, c2 As Collection
    Dim av(1) As Variant, al(1) As Long, ad(1) As Double, ab(1) As Byte, ai(1) As Integer
    On Error GoTo fail
    Set c = New Collection
    Set c2 = c
    s = "abc"
    x = 1
    Select Case k
        Case 1: v = (VarPtr(x) <> 0)                                       ' VarPtr nonzero
        Case 2: v = TypeName(VarPtr(x))                                    ' Long on 32-bit, LongLong on 64-bit
        Case 3: v = VarType(VarPtr(x))
        Case 4: v = (StrPtr(vbNullString) = 0)                             ' null string pointer
        Case 5: v = (StrPtr("") <> 0)                                      ' empty string has a BSTR
        Case 6: v = (StrPtr("abc") <> 0)
        Case 7: v = TypeName(StrPtr("a"))
        Case 8: v = (ObjPtr(Nothing) = 0)                                  ' Nothing
        Case 9: v = (ObjPtr(c) <> 0)
        Case 10: v = TypeName(ObjPtr(c))
        Case 11: v = (ObjPtr(c) = ObjPtr(c2))                              ' two references, one object
        Case 12: Set c2 = New Collection: v = (ObjPtr(c) <> ObjPtr(c2))    ' two objects
        Case 13: v = (VarPtr(x) = VarPtr(x))                               ' same variable
        Case 14: v = (VarPtr(x) <> VarPtr(y))                              ' two variables
        Case 15: v = CLng(VarPtr(av(1)) - VarPtr(av(0)))                   ' Variant element size
        Case 16: v = CLng(VarPtr(al(1)) - VarPtr(al(0)))                   ' Long element size
        Case 17: v = CLng(VarPtr(ad(1)) - VarPtr(ad(0)))                   ' Double element size
        Case 18: v = CLng(VarPtr(ab(1)) - VarPtr(ab(0)))                   ' Byte element size
        Case 19: v = CLng(VarPtr(ai(1)) - VarPtr(ai(0)))                   ' Integer element size
        Case 20: v = (StrPtr(s) = StrPtr(s))                               ' same string variable
        Case 21: v = (StrPtr(s2) = 0)                                      ' never-assigned String variable
        Case 22: s2 = "": v = (StrPtr(s2) <> 0)                            ' assigned empty string
        Case 23: v = ((StrPtr(s) Mod 4) = 0)                               ' alignment of string data
        Case 24: v = (ObjPtr(Err) <> 0) & "," & (VarPtr(s) <> 0)           ' Err object; String variable address
    End Select
    T = "Ptr." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Ptr." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
