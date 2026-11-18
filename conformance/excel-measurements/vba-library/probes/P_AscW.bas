Attribute VB_Name = "P_AscW"
Option Explicit
' AscW (Strings). Returns Integer: the UTF-16 code unit, signed.
Private Const N As Long = 16

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = AscW("A")                        ' typical
        Case 2: v = AscW("ABC")                      ' first character only
        Case 3: v = AscW("")                         ' empty: error 5
        Case 4: v = AscW(Null)                       ' Null: error 94
        Case 5: v = AscW(ChrW(233))                  ' e-acute
        Case 6: v = AscW(ChrW(8364))                 ' euro
        Case 7: v = AscW(ChrW(-32768))               ' smallest: -32768
        Case 8: v = AscW(ChrW(65535))                ' &HFFFF reads as -1
        Case 9: v = AscW(ChrW(55357))                ' high surrogate: negative
        Case 10: v = AscW(ChrW(56832))               ' low surrogate: negative
        Case 11: v = AscW(ChrW(26085))               ' CJK
        Case 12: v = AscW(123)                       ' number as string: "1"
        Case 13: v = AscW(Empty)                     ' Empty: error 5
        Case 14: v = AscW(String:="z")               ' named argument
        Case 15: v = AscW(Chr(128))                  ' 1252 euro -> 8364 ' locale
        Case 16: v = AscW(ChrW(160))                 ' no-break space
    End Select
    T = "AscW." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "AscW." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
