Attribute VB_Name = "P_LCase"
Option Explicit
' LCase / LCase$ (Strings). LCase returns Variant(String) or Null; LCase$ returns String.
Private Const N As Long = 22

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = LCase("ABC Def")                 ' typical
        Case 2: v = LCase("")                        ' empty
        Case 3: v = LCase(Null)                      ' Null propagates
        Case 4: v = LCase(123)                       ' number: "123"
        Case 5: v = LCase(Empty)                     ' Empty: ""
        Case 6: v = LCase(ChrW(201))                 ' E-acute -> e-acute
        Case 7: v = LCase(ChrW(223))                 ' sharp s unchanged
        Case 8: v = LCase(ChrW(304))                 ' dotted capital I
        Case 9: v = LCase(ChrW(7838))                ' capital sharp s
        Case 10: v = LCase(ChrW(931))                ' Greek capital sigma
        Case 11: v = LCase(ChrW(1050))               ' Cyrillic capital K
        Case 12: v = LCase(ChrW(8486))               ' Ohm sign -> omega
        Case 13: v = LCase(ChrW(65313))              ' fullwidth A
        Case 14: v = LCase$("ABC")                   ' $ form: String
        Case 15: v = LCase$(Null)                    ' $ form with Null: error 94
        Case 16: v = LCase(String:="ABC")            ' named argument
        Case 17: v = LCase(ChrW(8544))               ' Roman numeral one
        Case 18: v = LCase(True)                     ' Boolean: "true" ' locale
        Case 19: v = LCase(1.5E+20)                  ' Double text: "1.5e+20" ' locale
        Case 20: v = LCase(ChrW(55357) & ChrW(56832)) ' surrogate pair unchanged
        Case 21: v = LCase(Chr(138))                 ' S-caron in 1252 ' locale
        Case 22: v = LCase$(String:="Q" & ChrW(197)) ' $ form, named, A-ring
    End Select
    T = "LCase." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "LCase." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
