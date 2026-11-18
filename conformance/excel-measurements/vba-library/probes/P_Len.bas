Attribute VB_Name = "P_Len"
Option Explicit
' Len (Strings). Characters of a string; byte size of a numeric variable or UDT.
' Named arguments are not used: Len is a language intrinsic and Expression:= may not compile.
Private Const N As Long = 26

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
    Dim vL As Variant, vD As Variant, vB As Variant, vS As Variant
    Dim lg As Long, ig As Integer, db As Double, sg As Single, cu As Currency, dt As Date, bo As Boolean, by As Byte
    Dim fs As String * 5
    Dim up As UPt, ur As URec
    On Error GoTo fail
    vL = CLng(5): vD = 1.5: vB = CByte(7): vS = "hello"
    Select Case k
        Case 1: v = Len("abc")                       ' typical
        Case 2: v = Len("")                          ' empty
        Case 3: v = "compile-refused"
        Case 4: v = Len(Null)                        ' Null: Null
        Case 5: v = Len(Empty)                       ' Empty: 0
        Case 6: v = Len(vL)                          ' Variant holding a Long
        Case 7: v = Len(lg)                          ' Long variable: 4
        Case 8: v = Len(ig)                          ' Integer variable: 2
        Case 9: v = Len(up)                          ' UDT of two Longs: 8
        Case 10: v = Len(fs)                         ' fixed-length String * 5: 5
        Case 11: v = "compile-refused"
        Case 12: v = Len(vD)                         ' Variant holding a Double
        Case 13: v = Len(ChrW(55357) & ChrW(56832))  ' surrogate pair counts 2
        Case 14: v = Len(vbNullString)               ' vbNullString
        Case 15: v = Len(CVErr(5))                   ' error value
        Case 16: v = Len(Array(1, 2))                ' array: error 13
        Case 17: v = "compile-refused"
        Case 18: v = Len(bo)                         ' Boolean variable: 2
        Case 19: v = Len(db)                         ' Double variable: 8
        Case 20: v = Len(dt)                         ' Date variable: 8
        Case 21: v = Len(cu)                         ' Currency variable: 8
        Case 22: v = Len(sg)                         ' Single variable: 4
        Case 23: v = Len(by)                         ' Byte variable: 1
        Case 24: v = Len(ur)                         ' UDT with a String and String * 4
        Case 25: v = Len(vB)                         ' Variant holding a Byte
        Case 26: v = Len(vS)                         ' Variant holding a String
    End Select
    T = "Len." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "Len." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
