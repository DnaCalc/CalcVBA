Attribute VB_Name = "SpecLib1"
Option Explicit
' MS-VBAL §6 Conversion, DateTime, Financial: open points from the dc-ez3rd code inspection.

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

Private Function ErrAfterClear() As Variant
    On Error Resume Next
    Err.Raise 1000, , "xx"
    Err.Clear
    ErrAfterClear = Error()
End Function

Private Function ErrAfterRaise() As Variant
    On Error Resume Next
    Err.Raise 11
    ErrAfterRaise = Error()
End Function

Private Function T(ByVal k As String) As String
    Dim v As Variant, cf(2) As Double, cf2(1) As Double, d As Date
    On Error GoTo fail
    Select Case k
        Case "dayNull": v = Day(Null)
        Case "yearNull": v = Year(Null)
        Case "weekdayNull": v = Weekday(Null)
        Case "hourNull": v = Hour(Null)
        Case "hexNull": v = Hex(Null)
        Case "octNull": v = Oct(Null)
        Case "strNull": v = Str(Null)
        Case "dayStr": v = Day("1/15/2000")
        Case "dateAddStr": v = DateAdd("m", 1, "31-Jan-95")
        Case "dateDiffStr": v = DateDiff("d", "1/1/2000", "1/3/2000")
        Case "datePartStr": v = DatePart("m", "3/1/2000")
        Case "hexStr": v = Hex("255")
        Case "strStr": v = Str("12")
        Case "strEmpty": v = Str(Empty)
        Case "strDate": v = Str(#1/1/2000#)
        Case "strTrue": v = Str(True)
        Case "strErr": v = Str(CVErr(5))
        Case "dsYear5": v = DateSerial(5, 1, 1)
        Case "dsYear99": v = DateSerial(99, 1, 1)
        Case "dsMon15": v = DateSerial(2000, 15, 1)
        Case "dsMonM10": v = DateSerial(2000, -10, 1)
        Case "dsDay40000": v = DateSerial(2000, 1, 40000)
        Case "dsBigYear": v = DateSerial(40000, 1, 1)
        Case "dpWMon": v = DatePart("w", #1/5/2000#, vbMonday)
        Case "dpWW4": v = DatePart("ww", #12/29/2003#, vbMonday, vbFirstFourDays)
        Case "dpWWFull": v = DatePart("ww", #1/1/2005#, vbSunday, vbFirstFullWeek)
        Case "dpWW4b": v = DatePart("ww", #1/1/2005#, vbMonday, vbFirstFourDays)
        Case "ddQ": v = DateDiff("q", #3/31/2000#, #4/1/2000#)
        Case "ddWneg": v = DateDiff("w", #1/9/2023#, #1/8/2023#)
        Case "ddWWMon": v = DateDiff("ww", #1/1/2023#, #1/31/2023#, vbMonday)
        Case "ddHours": v = DateDiff("h", #1/1/2000 10:59:00 PM#, #1/1/2000 11:01:00 PM#)
        Case "dvNoYear": d = DateValue("12/30"): v = (Year(d) = Year(Date)) & " " & Month(d) & "/" & Day(d)
        Case "dvMonName": d = DateValue("Dec 30"): v = (Year(d) = Year(Date)) & " " & Month(d) & "/" & Day(d)
        Case "dvDmy": v = DateValue("31-Jan-95")
        Case "tvDateOnly": v = TimeValue("1/1/2000")
        Case "tv3pm": v = TimeValue("3 PM")
        Case "dateDollar": v = TypeName(Date$) & " " & Len(Date$) & " " & Mid$(Date$, 3, 1)
        Case "timeDollar": v = TypeName(Time$) & " " & Len(Time$)
        Case "irrPos": cf(0) = 1: cf(1) = 2: cf(2) = 3: v = IRR(cf)
        Case "irrTwo": cf(0) = -100: cf(1) = 230: cf(2) = -132: v = IRR(cf)
        Case "rateBig": v = Rate(1, 0, -1, 20)
        Case "rate360": v = Rate(360, -1, 1000)
        Case "nperBad": v = NPer(0.1, 0, 1000)
        Case "pmtN0": v = Pmt(0.1, 0, 1000)
        Case "npvM1": cf2(0) = 1: cf2(1) = 2: v = NPV(-1, cf2)
        Case "ddbP6": v = DDB(1000, 100, 5, 6)
        Case "slnL0": v = SLN(1, 0, 0)
        Case "sydP6": v = SYD(1000, 100, 5, 6)
        Case "fvDue2": v = FV(0.1, 10, -100, 0, 2)
        Case "ipmtN0": v = IPmt(0.1, 0, 10, 1000)
        Case "errNeg": v = Error(-1)
        Case "errBig": v = Error(70000)
        Case "errNoArgClear": v = ErrAfterClear()
        Case "errNoArgRaise": v = ErrAfterRaise()
        Case "tsNeg": v = CDbl(TimeSerial(-1, 0, 0))
        Case "daddFrac": v = DateAdd("d", 1.5, #1/1/2000#)
        Case "daddMonFrac": v = DateAdd("m", 1.5, #1/31/2000#)
        Case "daddNull": v = DateAdd("d", 1, Null)
        Case "daddHourPre": v = CDbl(DateAdd("h", 1, CDate(-1.25)))
        Case "daddDayPre": v = CDbl(DateAdd("d", 2, CDate(-1.25)))
        Case "freeFileType": v = TypeName(FreeFile)
        Case "cdecDbl": v = CDec(123.456789012345678#)
        Case "cdecTiny": v = CDec(1E-20)
        Case "cboolErr": v = CBool(CVErr(5))
        Case "clngErr": v = CLng(CVErr(5))
        Case Else: v = "unknown case"
    End Select
    T = k & "=" & Sh(v): Exit Function
fail:
    T = k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim keys As Variant, i As Long, s As String
    keys = Split("dayNull yearNull weekdayNull hourNull hexNull octNull strNull dayStr dateAddStr dateDiffStr datePartStr hexStr strStr strEmpty strDate strTrue strErr dsYear5 dsYear99 dsMon15 dsMonM10 dsDay40000 dsBigYear dpWMon dpWW4 dpWWFull dpWW4b ddQ ddWneg ddWWMon ddHours dvNoYear dvMonName dvDmy tvDateOnly tv3pm dateDollar timeDollar irrPos irrTwo rateBig rate360 nperBad pmtN0 npvM1 ddbP6 slnL0 sydP6 fvDue2 ipmtN0 errNeg errBig errNoArgClear errNoArgRaise tsNeg daddFrac daddMonFrac daddNull daddHourPre daddDayPre freeFileType cdecDbl cdecTiny cboolErr clngErr", " ")
    For i = 0 To UBound(keys)
        s = s & T(keys(i)) & vbLf
    Next
    Run = s
End Function
