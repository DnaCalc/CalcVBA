Attribute VB_Name = "P_StrConv"
Option Explicit
' StrConv (Strings; MS-VBAL 6.1.2.11.1.x). vbUpperCase=1 vbLowerCase=2 vbProperCase=3 vbWide=4 vbNarrow=8 vbKatakana=16 vbHiragana=32 vbUnicode=64 vbFromUnicode=128.
Private Const N As Long = 50

Private Function T(ByVal k As Long) As String
    Dim v As Variant, s As String
    On Error GoTo fail
    Select Case k
        Case 1: v = StrConv("abc def", vbUpperCase)               ' upper
        Case 2: v = StrConv("ABC DEF", vbLowerCase)               ' lower
        Case 3: v = StrConv("hello woRLD", vbProperCase)          ' proper
        Case 4: v = StrConv("a" & vbTab & "b", vbProperCase)      ' tab separator
        Case 5: v = StrConv("a" & Chr(0) & "b", vbProperCase)     ' Chr(0) separator
        Case 6: v = StrConv("don't stop", vbProperCase)           ' apostrophe
        Case 7: v = StrConv("well-known fact", vbProperCase)      ' hyphen
        Case 8: v = StrConv("3rd place_x", vbProperCase)          ' digits and underscore
        Case 9: v = StrConv("a.b,c;d", vbProperCase)              ' punctuation separators
        Case 10: v = StrConv("a" & vbCrLf & "b", vbProperCase)    ' line break separator
        Case 11: v = StrConv("abc", vbUnicode)                    ' ANSI bytes to Unicode
        Case 12: v = Len(StrConv("abc", vbUnicode))               ' length of the vbUnicode result
        Case 13: v = StrConv("abc", vbFromUnicode)                ' to ANSI bytes
        Case 14: v = Len(StrConv("abc", vbFromUnicode))           ' character length of the byte string
        Case 15: v = LenB(StrConv("abc", vbFromUnicode))          ' byte length
        Case 16: v = StrConv(Bz_StrConv(), vbUnicode)             ' Byte() argument to vbUnicode
        Case 17: v = StrConv(Bz_StrConv(), vbFromUnicode)         ' Byte() argument to vbFromUnicode
        Case 18: v = StrConv("a" & ChrW(&H4E2D), vbFromUnicode)   ' non-ANSI char; code page dependent. locale
        Case 19: v = StrConv("abc", vbWide)                       ' error on this locale
        Case 20: v = StrConv("abc", vbNarrow)                     ' error on this locale
        Case 21: v = StrConv("abc", vbKatakana)                   ' error on this locale
        Case 22: v = StrConv("abc", vbHiragana)                   ' error on this locale
        Case 23: v = StrConv("abc", vbUpperCase, 1033)            ' LCID en-US
        Case 24: v = StrConv("i", vbUpperCase, 1055)              ' LCID Turkish: dotted capital I? locale
        Case 25: v = StrConv("I", vbLowerCase, 1055)              ' LCID Turkish: dotless i? locale
        Case 26: v = StrConv("abc", vbUpperCase + vbWide)         ' combined with wide
        Case 27: v = StrConv("abc", vbUpperCase + vbLowerCase)    ' 3 = vbProperCase
        Case 28: v = StrConv("abc", vbUnicode + vbFromUnicode)    ' invalid combination: error 5
        Case 29: v = StrConv("abc", vbUpperCase + vbUnicode)      ' invalid combination: error 5
        Case 30: v = StrConv("abc", vbWide + vbNarrow)            ' invalid combination: error 5
        Case 31: v = StrConv("abc", vbKatakana + vbHiragana)      ' invalid combination: error 5
        Case 32: v = StrConv("abc", 0)                            ' conversion 0: unchanged
        Case 33: v = StrConv("abc", -1)                           ' invalid: error 5
        Case 34: v = StrConv("abc", 256)                          ' invalid: error 5
        Case 35: v = StrConv(Null, vbUpperCase)                   ' Null propagates
        Case 36: v = StrConv(Empty, vbUpperCase)                  ' Empty
        Case 37: v = StrConv(123, vbUpperCase)                    ' number as string argument
        Case 38: v = StrConv("", vbProperCase)                    ' empty string
        Case 39: v = StrConv(ChrW(228) & ChrW(223) & ChrW(255), vbUpperCase)  ' non-ASCII upper
        Case 40: v = StrConv(ChrW(196) & ChrW(214), vbLowerCase)  ' non-ASCII lower
        Case 41: v = StrConv(ChrW(228) & "bc " & ChrW(246) & "x", vbProperCase)  ' non-ASCII proper
        Case 42: v = StrConv(String:="abc", Conversion:=vbUpperCase, LocaleID:=1033)  ' named arguments
        Case 43: v = StrConv("abc", vbUpperCase, 0)               ' LCID 0
        Case 44: v = StrConv("abc", vbUpperCase, 99999)           ' invalid LCID
        Case 45: v = TypeName(StrConv("abc", vbUpperCase))        ' result type
        Case 46: v = TypeName(StrConv("abc", vbFromUnicode))      ' result type of vbFromUnicode
        Case 47: v = StrConv("abc", "1")                          ' numeric string conversion
        Case 48: v = StrConv(Null, vbUnicode)                     ' Null with vbUnicode
        Case 49: v = StrConv("abc", vbProperCase, 1033)           ' proper with LCID
        Case 50: v = StrConv("ABC", vbProperCase + vbWide)        ' proper plus wide
    End Select
    T = "StrConv." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "StrConv." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function

Private Function Bz_StrConv() As Byte()
    Dim b(0 To 3) As Byte
    b(0) = 65: b(1) = 66: b(2) = 67: b(3) = 0
    Bz_StrConv = b
End Function
