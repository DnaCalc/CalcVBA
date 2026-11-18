Attribute VB_Name = "P_Asc"
Option Explicit
' Asc (Strings; MS-VBAL 6.1.2.11.1.x). Returns Integer; code page 1252 on this machine.
Private Const N As Long = 16

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = Asc("A")                         ' typical
        Case 2: v = Asc("ABC")                       ' first character only
        Case 3: v = Asc("")                          ' empty: error 5
        Case 4: v = Asc(Null)                        ' Null: error 94
        Case 5: v = Asc(ChrW(233))                   ' e-acute, in code page 1252
        Case 6: v = Asc(ChrW(8364))                  ' euro sign: 128 in 1252 ' locale
        Case 7: v = Asc(ChrW(26085))                 ' CJK, not in 1252 ' locale
        Case 8: v = Asc(Chr(200))                    ' round trip of Chr
        Case 9: v = Asc(123)                         ' number as string: "1" = 49
        Case 10: v = Asc(String:="x")                ' named argument
        Case 11: v = Asc(ChrW(160))                  ' no-break space
        Case 12: v = Asc(Chr(128))                   ' euro round trip ' locale
        Case 13: v = Asc(Empty)                      ' Empty is "": error 5
        Case 14: v = Asc(ChrW(55357))                ' lone surrogate ' locale
        Case 15: v = Asc(ChrW(223))                  ' sharp s
        Case 16: v = Asc(ChrW(8211))                 ' en dash: 150 in 1252 ' locale
    End Select
    T = "Asc." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Asc." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
