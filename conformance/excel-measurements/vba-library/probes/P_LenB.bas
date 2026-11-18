Attribute VB_Name = "P_LenB"
Option Explicit
' LenB (Strings). Bytes of a string (2 per UTF-16 unit) or of a variable in memory.
' Named arguments are not used: LenB is a language intrinsic and Expression:= may not compile.
Private Const N As Long = 24

Private Type UPt
    x As Long
    y As Long
End Type

Private Type URec
    n As Integer
    s As String
    f As String * 4
    d As Double
End Type

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    Dim vL As Variant, vD As Variant, vS As Variant
    Dim lg As Long, ig As Integer, db As Double, bo As Boolean
    Dim fs As String * 5
    Dim up As UPt, ur As URec
    On Error GoTo fail
    vL = CLng(5): vD = 1.5: vS = "hello"
    Select Case k
        Case 1: v = LenB("abc")                      ' typical: 6
        Case 2: v = LenB("")                         ' empty
        Case 3: v = "compile-refused"
        Case 4: v = LenB(Null)                       ' Null: Null
        Case 5: v = LenB(Empty)                      ' Empty
        Case 6: v = LenB(vL)                         ' Variant holding a Long
        Case 7: v = LenB(lg)                         ' Long variable: 4
        Case 8: v = LenB(ig)                         ' Integer variable: 2
        Case 9: v = LenB(up)                         ' UDT of two Longs: 8
        Case 10: v = LenB(fs)                        ' String * 5 in memory: 10
        Case 11: v = LenB(ChrW(233))                 ' one UTF-16 unit: 2
        Case 12: v = LenB(ChrW(55357) & ChrW(56832)) ' surrogate pair: 4
        Case 13: v = LenB(ChrB(65))                  ' one byte string: 1
        Case 14: v = LenB(vD)                        ' Variant holding a Double
        Case 15: v = LenB(vbNullString)              ' vbNullString
        Case 16: v = LenB(ur)                        ' UDT with String and String * 4
        Case 17: v = LenB(db)                        ' Double variable: 8
        Case 18: v = LenB(bo)                        ' Boolean variable: 2
        Case 19: v = LenB(vS)                        ' Variant holding a String
        Case 20: v = LenB(ChrB(65) & ChrB(66) & ChrB(67)) ' odd byte count: 3
        Case 21: v = "compile-refused"
        Case 22: v = LenB(CVErr(5))                  ' error value
        Case 23: v = LenB(Array(1, 2))               ' array: error 13
        Case 24: v = "compile-refused"
    End Select
    T = "LenB." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "LenB." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
