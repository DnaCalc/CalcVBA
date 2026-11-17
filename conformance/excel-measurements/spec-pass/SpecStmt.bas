Attribute VB_Name = "SpecStmt"
Option Explicit
' MS-VBAL §5.4 statements and §6.1.2.5 FileSystem: open points from the dc-ez3rd code
' inspection. Files go under %TEMP%\cvp-spec, recreated per case.
Private bumps As Long

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

Private Function Root() As String
    Root = Environ("TEMP") & "\cvp-spec"
End Function

Private Sub Wipe()
    Dim f As String
    On Error Resume Next
    Close
    SetAttr Root & "\d\h.txt", vbNormal
    Kill Root & "\d\*.*"
    RmDir Root & "\d\sub"
    RmDir Root & "\d"
    Kill Root & "\*.*"
    MkDir Root
End Sub

Private Function ReadAll(ByVal path As String) As String
    Dim n As Integer, b() As Byte
    n = FreeFile
    Open path For Binary As #n
    If LOF(n) > 0 Then
        ReDim b(LOF(n) - 1)
        Get #n, , b
        ReadAll = StrConv(b, vbUnicode)
    End If
    Close #n
End Function

Private Function Bump() As Long
    bumps = bumps + 1
    Bump = 1
End Function

Private Function OnGo(ByVal n As Variant) As String
    On Error GoTo fail
    On n GoTo L1, L2
    OnGo = "fell": Exit Function
L1: OnGo = "L1": Exit Function
L2: OnGo = "L2": Exit Function
fail: OnGo = "ERR" & Err.Number
End Function

Private Function DirList(ByVal pattern As String, ByVal attrs As Integer) As String
    Dim n As String
    n = Dir(pattern, attrs)
    Do While Len(n) > 0
        DirList = DirList & n & ";"
        n = Dir()
    Loop
End Function

Private Function T(ByVal k As String) As String
    Dim v As Variant, vv As Variant, f As String, d As String, s As String, i As Long, n As Long
    Dim bt As Byte
    Wipe
    f = Root & "\a.txt": d = Root & "\d"
    On Error GoTo fail
    Select Case k
        Case "openMissing": Open Root & "\none.txt" For Input As #1: v = "ok"
        Case "openTwice": Open f For Output As #1: Open Root & "\b.txt" For Output As #1: v = "ok"
        Case "openHash0": Open f For Output As #0: v = "ok"
        Case "openHash600": Open f For Output As #600: v = "ok"
        Case "seekZero": Open f For Binary As #1: Seek #1, 0: v = "ok"
        Case "putOverLen": Open f For Random As #1 Len = 2: Put #1, 1, "abcdef": v = "ok"
        Case "eofBad": v = EOF(9)
        Case "printEmpty": Open f For Output As #1: Print #1, "a", , "b": Close #1: v = ReadAll(f)
        Case "printLeadComma": Open f For Output As #1: Print #1, , "x": Close #1: v = ReadAll(f)
        Case "writeEmpty": Open f For Output As #1: Write #1, "a", , "b": Close #1: v = ReadAll(f)
        Case "writeTrailSemi": Open f For Output As #1: Write #1, "a";: Write #1, "b": Close #1: v = ReadAll(f)
        Case "printSpcTab": Open f For Output As #1: Print #1, "a"; Spc(-2); "b"; Tab(0); "c": Close #1: v = ReadAll(f)
        Case "printWidth": Open f For Output As #1: Width #1, 5: Print #1, "abcdef": Close #1: v = ReadAll(f)
        Case "printDbl": Open f For Output As #1: Print #1, 0.1 + 0.2: Close #1: v = ReadAll(f)
        Case "writeDbl": Open f For Output As #1: Write #1, 0.1 + 0.2: Close #1: v = ReadAll(f)
        Case "printBig": Open f For Output As #1: Print #1, 1E+20; 0.000015; -1 / 3: Close #1: v = ReadAll(f)
        Case "midStmtPast": s = "abc": Mid(s, 5) = "x": v = s
        Case "onGo300": v = OnGo(300)
        Case "onGo40000": v = OnGo(40000)
        Case "onGoNeg": v = OnGo(-1)
        Case "onGoZero": v = OnGo(0)
        Case "forOrder": i = 10: For i = 1 To i + 2: n = n + 1: Next: v = n
        Case "forVarType": For vv = 1 To 2.5: s = s & TypeName(vv) & " ": Next: v = s
        Case "forNull": For vv = 1 To Null: n = n + 1: Next: v = n
        Case "forStr": For vv = 1 To "abc": n = n + 1: Next: v = n
        Case "selShort": bumps = 0: Select Case 1: Case 1, Bump(): End Select: v = bumps
        Case "selShortDiv": Select Case 1: Case 1, 1 / 0: v = "matched": End Select
        Case "redimScalar": vv = 5: ReDim vv(2): v = UBound(vv)
        Case "eraseScalar": vv = 5: Erase vv: v = vv
        Case "randomDefLen": Open f For Random As #1: Put #1, 2, "x": v = LOF(1)
        Case "randomLenBad": Open f For Random As #1 Len = 40000: v = "ok"
        Case "killMissing": Kill Root & "\none.txt": v = "ok"
        Case "mkdirExisting": MkDir d: MkDir d: v = "ok"
        Case "rmdirNonEmpty": MkDir d: Open d & "\x.txt" For Output As #1: Close #1: RmDir d: v = "ok"
        Case "chdirMissing": ChDir Root & "\nodir": v = "ok"
        Case "fileLenMissing": v = FileLen(Root & "\none.txt")
        Case "dirAttrs"
            MkDir d: MkDir d & "\sub"
            Open d & "\a.txt" For Output As #1: Close #1
            Open d & "\h.txt" For Output As #1: Close #1
            SetAttr d & "\h.txt", vbHidden
            v = DirList(d & "\*", vbNormal) & "|" & DirList(d & "\*", vbDirectory) & "|" & DirList(d & "\*", vbHidden)
        Case "dirFolder"
            MkDir d: Open d & "\a.txt" For Output As #1: Close #1
            v = Dir(d) & "|" & Dir(d & "\")
        Case "fileAttrMode": Open f For Binary As #2: v = FileAttr(2, 1)
        Case "eofBinary": Open f For Binary As #2: Put #2, , CByte(1): Seek #2, 1: Get #2, , bt: v = EOF(2)
        Case "getAttrType": Open f For Output As #1: Close #1: v = TypeName(GetAttr(f))
        Case "curDirDollar": v = (CurDir$ = CurDir)
        Case Else: v = "unknown case"
    End Select
    T = k & "=" & Sh(v)
    Wipe
    Exit Function
fail:
    T = k & "=ERR" & Err.Number
    Wipe
End Function

Public Function Run() As String
    Dim keys As Variant, i As Long, s As String
    keys = Split("openMissing openTwice openHash0 openHash600 seekZero putOverLen eofBad printEmpty printLeadComma writeEmpty writeTrailSemi printSpcTab printWidth printDbl writeDbl printBig midStmtPast onGo300 onGo40000 onGoNeg onGoZero forOrder forVarType forNull forStr selShort selShortDiv redimScalar eraseScalar randomDefLen randomLenBad killMissing mkdirExisting rmdirNonEmpty chdirMissing fileLenMissing dirAttrs dirFolder fileAttrMode eofBinary getAttrType curDirDollar", " ")
    For i = 0 To UBound(keys)
        s = s & T(keys(i)) & vbLf
    Next
    Run = s
End Function
