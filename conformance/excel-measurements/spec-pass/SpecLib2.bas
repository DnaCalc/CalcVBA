Attribute VB_Name = "SpecLib2"
Option Explicit
' MS-VBAL §6 Information, Interaction, Math, Strings, Collection: open points from the dc-ez3rd
' code inspection.

Private Function Esc(ByVal s As String) As String
    Dim i As Long, c As Long
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 32 Or c > 126 Then Esc = Esc & "{" & c & "}" Else Esc = Esc & Mid$(s, i, 1)
    Next
End Function

Private Function Sh(ByVal v As Variant) As String
    If IsObject(v) Then Sh = "obj " & TypeName(v): Exit Function
    If IsArray(v) Then Sh = TypeName(v) & "[" & LBound(v) & "," & UBound(v) & "]": Exit Function
    Select Case VarType(v)
        Case vbNull: Sh = "Null"
        Case vbEmpty: Sh = "Empty"
        Case vbDate: Sh = "Date:" & CStr(CDbl(v))
        Case vbString: Sh = "String:[" & Esc(v) & "]"
        Case Else: Sh = TypeName(v) & ":" & CStr(v)
    End Select
End Function

Private Function T(ByVal k As String) As String
    Dim v As Variant, c As Collection, x As Single
    Set c = New Collection
    On Error GoTo fail
    Select Case k
        Case "chrNull": v = Chr(Null)
        Case "chrwNull": v = ChrW(Null)
        Case "ascwNull": v = AscW(Null)
        Case "spaceNull": v = Space(Null)
        Case "leftStrNull": v = Left(Null, 1)
        Case "leftLenNull": v = Left("a", Null)
        Case "midStartNull": v = Mid("a", Null)
        Case "instrStartNull": v = InStr(Null, "a", "b")
        Case "instrFindNull": v = InStr(1, "a", Null)
        Case "replaceNull": v = Replace(Null, "a", "b")
        Case "strcompCmpNull": v = StrComp("a", "b", Null)
        Case "instrrevNull": v = InStrRev(Null, "a")
        Case "sqrStr": v = Sqr("4")
        Case "leftStrLen": v = Left("abc", "2")
        Case "iifStr": v = IIf("True", 1, 2)
        Case "spaceStr": v = Space("3")
        Case "chrStr": v = Chr("65")
        Case "roundStr": v = Round("2.5")
        Case "ascwType": v = AscW("a")
        Case "instrEmpty": v = InStr("", "")
        Case "instrEmpty1": v = InStr(1, "", "")
        Case "instrEmpty2": v = InStr(2, "", "")
        Case "instrrevBigStart": v = InStrRev("abc", "c", 10)
        Case "splitEmpty": v = Split("")
        Case "splitType": v = TypeName(Split("a,b", ","))
        Case "filterType": v = TypeName(Filter(Split("a,b", ","), "a"))
        Case "splitLimM2": v = Split("a,b", ",", -2)
        Case "propTab": v = StrConv("a" & vbTab & "b c-d", vbProperCase)
        Case "uniText": v = StrConv("abc", vbUnicode)
        Case "fromUniType": v = VarType(StrConv("abc", vbFromUnicode))
        Case "fromUniLen": v = LenB(StrConv("abc", vbFromUnicode))
        Case "strconvWide": v = StrConv("abc", vbWide)
        Case "string128": v = AscW(String(1, 128))
        Case "stringNeg": v = String(2, -1)
        Case "stringEmpty": v = String(2, "")
        Case "roundCur": v = Round(CCur(2.5))
        Case "roundInt": v = Round(CInt(2))
        Case "roundLng": v = Round(CLng(7))
        Case "roundNull": v = Round(Null)
        Case "roundDec": v = Round(CDec("2.55"), 1)
        Case "roundSng": x = 1.25: v = Round(x, 1)
        Case "isnumTrue": v = IsNumeric(True)
        Case "isnumDate": v = IsNumeric(#1/1/2000#)
        Case "switchOdd": v = Switch(False, 1, True)
        Case "collMissingKey": c.Add 1, "a": v = c("zz")
        Case "collRemoveMissing": c.Add 1, "a": c.Remove "zz": v = c.Count
        Case "collNonStrKey": c.Add 1, 5: v = c.Count
        Case "collBeforeOut": c.Add 1: c.Add 2: c.Add 9, , 5: v = c.Count
        Case "collUmlautKeys": c.Add 1, ChrW(196): c.Add 2, ChrW(228): v = c.Count
        Case "collIndexOut": c.Add 1: c.Add 2: v = c(3)
        Case "fmtStrNum": v = Format("abc", "0.00")
        Case "fmtNullNum": v = Format(Null, "0.00")
        Case "fmtNullSec": v = Format(Null, "0;0;0;""nil""")
        Case "fmtNumAbc": v = FormatNumber("abc")
        Case "fmtNumNull": v = FormatNumber(Null)
        Case "fmtAt": v = Format("abc", "@@@@@")
        Case "fmtUpper": v = Format("abc", ">")
        Case "fmtLit": v = Format(5, """x""0")
        Case "fmtWeek": v = Format(#1/5/2000#, "w", vbMonday)
        Case "fmtWW": v = Format(#1/1/2005#, "ww", vbMonday, vbFirstFourDays)
        Case "rndNeg11": v = Rnd(-1.1)
        Case "rndRandomize11": v = Rnd(-1): Randomize 1.1: v = Rnd
        Case "rndRandomize2": v = Rnd(-1): Randomize 2: v = Rnd
        Case "partW": v = Partition(1000, 100, 1010, 20)
        Case "partW2": v = Partition(100, 0, 1000, 1)
        Case "partW3": v = Partition(5, 0, 99, 10)
        Case "choose15": v = Choose(1.5, "a", "b", "c")
        Case "choose25": v = Choose(2.5, "a", "b", "c")
        Case "chrwMin": v = AscW(ChrW(-32768))
        Case "rgbNeg": v = RGB(-1, 0, 0)
        Case "wdName8": v = WeekdayName(1, False, 8)
        Case "ascB": v = AscB("A")
        Case "chrB": v = LenB(ChrB(65))
        Case "instrB": v = InStrB("abc", "b")
        Case "midB": v = MidB("abc", 3, 2)
        Case Else: v = "unknown case"
    End Select
    T = k & "=" & Sh(v): Exit Function
fail:
    T = k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim keys As Variant, i As Long, s As String
    keys = Split("chrNull chrwNull ascwNull spaceNull leftStrNull leftLenNull midStartNull instrStartNull instrFindNull replaceNull strcompCmpNull instrrevNull sqrStr leftStrLen iifStr spaceStr chrStr roundStr ascwType instrEmpty instrEmpty1 instrEmpty2 instrrevBigStart splitEmpty splitType filterType splitLimM2 propTab uniText fromUniType fromUniLen strconvWide string128 stringNeg stringEmpty roundCur roundInt roundLng roundNull roundDec roundSng isnumTrue isnumDate switchOdd collMissingKey collRemoveMissing collNonStrKey collBeforeOut collUmlautKeys collIndexOut fmtStrNum fmtNullNum fmtNullSec fmtNumAbc fmtNumNull fmtAt fmtUpper fmtLit fmtWeek fmtWW rndNeg11 rndRandomize11 rndRandomize2 partW partW2 partW3 choose15 choose25 chrwMin rgbNeg wdName8 ascB chrB instrB midB", " ")
    For i = 0 To UBound(keys)
        s = s & T(keys(i)) & vbLf
    Next
    Run = s
End Function
