Attribute VB_Name = "P_IMEStatus"
Option Explicit
' IMEStatus (Information; the IME mode of the host, a VbIMEStatus value).
Private Const N As Long = 9

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    Dim l As Long
    On Error GoTo fail
    Select Case k
        Case 1: v = IMEStatus                                                             ' value (IME dependent)
        Case 2: v = TypeName(IMEStatus)                                                   ' result type
        Case 3: v = VarType(IMEStatus)
        Case 4: v = (IMEStatus = vbIMENoOp)                                               ' no IME installed
        Case 5: l = IMEStatus: v = l                                                      ' assignment to Long
        Case 6: v = (IMEStatus = IMEStatus)                                               ' stable between calls
        Case 7: v = vbIMENoOp & "," & vbIMEOn & "," & vbIMEOff & "," & vbIMEDisable & "," & vbIMEHiragana     ' constants
        Case 8: v = vbIMEKatakanaDbl & "," & vbIMEKatakanaSng & "," & vbIMEAlphaDbl & "," & vbIMEAlphaSng     ' constants
        Case 9: v = IMEStatus()                                                           ' empty parentheses
    End Select
    T = "IMEStatus." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "IMEStatus." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
