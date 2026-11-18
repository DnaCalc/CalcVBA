Attribute VB_Name = "P_Chr"
Option Explicit
' Chr / Chr$ (Strings). Chr returns Variant(String), Chr$ returns String; code page 1252.
Private Const N As Long = 24

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Chr(65)                          ' typical
        Case 2: v = Chr(0)                           ' NUL
        Case 3: v = Chr(128)                         ' euro in 1252 ' locale
        Case 4: v = Chr(255)                         ' y-diaeresis
        Case 5: v = Chr(256)                         ' error 5
        Case 6: v = Chr(-1)                          ' error 5
        Case 7: v = Chr(Null)                        ' Null: error 94
        Case 8: v = Chr("65")                        ' numeric string
        Case 9: v = Chr(65.5)                        ' rounds to even: 66
        Case 10: v = Chr(64.5)                       ' rounds to even: 64
        Case 11: v = Chr(Empty)                      ' Empty is 0
        Case 12: v = Chr$(65)                        ' $ form: String
        Case 13: v = Chr(CharCode:=66)               ' named argument
        Case 14: v = Chr("x")                        ' non-numeric string: error 13
        Case 15: v = Chr(129)                        ' undefined in 1252 ' locale
        Case 16: v = Chr(141)                        ' undefined in 1252 ' locale
        Case 17: v = Chr(150)                        ' en dash in 1252 ' locale
        Case 18: v = Chr(160)                        ' no-break space
        Case 19: v = Chr(13) & Chr(10)               ' CRLF
        Case 20: v = Chr(65536)                      ' error 5
        Case 21: v = Chr$(CharCode:=67)              ' $ form, named
        Case 22: v = Chr$(Null)                      ' $ form with Null: error 94
        Case 23: v = Chr(-32768)                     ' smallest Integer
        Case 24: v = Chr(233) & Chr(223)             ' e-acute, sharp s
    End Select
    T = "Chr." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Chr." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
