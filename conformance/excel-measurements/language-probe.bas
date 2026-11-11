Option Explicit
' Recording helper for CalcVBA's language rows (gen_language_rows.py). A prepared fixture calls
' DnaRec "<name>", <name> at the end of Main; DnaRunAll runs Main and returns the records.
' Each record is  name=<VarType>:<payload>  on its own line. Doubles, Singles and Dates are the
' 16 hex digits of the IEEE double; strings are 4 hex digits per UTF-16 unit; an array is
' <VarType>:<element VarType>:<rank>:<lbound,ubound>...|<element record>|... (rank 1 and 2).

Private Type DblBits
    d As Double
End Type
Private Type ByteBits
    b(0 To 7) As Byte
End Type

Public DnaOut As String
Public DnaCaptured As Boolean

Private Function HexD(ByVal d As Double) As String
    Dim x As DblBits, y As ByteBits, i As Long, s As String
    x.d = d
    LSet y = x
    For i = 7 To 0 Step -1
        s = s & Right$("0" & Hex$(y.b(i)), 2)
    Next i
    HexD = s
End Function

Private Function EncStr(ByVal s As String) As String
    Dim i As Long, r As String
    For i = 1 To Len(s)
        r = r & Right$("000" & Hex$(AscW(Mid$(s, i, 1)) And &HFFFF&), 4)
    Next i
    EncStr = r
End Function

Private Function Enc(ByVal v As Variant) As String
    Dim t As Long
    t = VarType(v)
    If t >= 8192 Then
        Enc = EncArray(v)
        Exit Function
    End If
    Select Case t
    Case 0: Enc = "0:"
    Case 1: Enc = "1:"
    Case 2, 3, 17, 20: Enc = t & ":" & CStr(v)
    Case 4, 5, 7: Enc = t & ":" & HexD(CDbl(v))
    Case 6, 14: Enc = t & ":" & Trim$(Str$(v))
    Case 8: Enc = "8:" & EncStr(CStr(v))
    Case 10
        Dim n As Long
        On Error Resume Next
        n = CLng(v)
        On Error GoTo 0
        Enc = "10:" & CStr(n)
    Case 11
        If v Then Enc = "11:true" Else Enc = "11:false"
    Case 9, 13: Enc = t & ":" & TypeName(v)
    Case Else: Enc = t & ":"
    End Select
End Function

Private Function EncArray(ByVal v As Variant) As String
    Dim nd As Long, i As Long, j As Long, s As String
    Dim lb(1 To 8) As Long, ub(1 To 8) As Long
    On Error Resume Next
    Do
        Err.Clear
        lb(nd + 1) = LBound(v, nd + 1)
        ub(nd + 1) = UBound(v, nd + 1)
        If Err.Number <> 0 Then Exit Do
        nd = nd + 1
        If nd = 8 Then Exit Do
    Loop
    Err.Clear
    On Error GoTo 0
    s = VarType(v) & ":" & (VarType(v) - 8192) & ":" & nd
    For i = 1 To nd
        s = s & ":" & lb(i) & "," & ub(i)
    Next i
    If nd = 1 Then
        For i = lb(1) To ub(1)
            s = s & "|" & Enc(v(i))
        Next i
    ElseIf nd = 2 Then
        For i = lb(1) To ub(1)
            For j = lb(2) To ub(2)
                s = s & "|" & Enc(v(i, j))
            Next j
        Next i
    End If
    EncArray = s
End Function

Public Sub DnaRec(ByVal name As String, ByVal v As Variant)
    DnaOut = DnaOut & name & "=" & Enc(v) & vbLf
End Sub

' "OK" + records, or "ERR|<Err.Number>|<Err.Description>|<Err.Source>" when Main raised.
Public Function DnaRunAll() As String
    On Error GoTo failed
    DnaOut = ""
    DnaCaptured = False
    Main
    If DnaCaptured Then
        DnaRunAll = "OK" & vbLf & DnaOut
    Else
        DnaRunAll = "NOCAPTURE" & vbLf & DnaOut
    End If
    Exit Function
failed:
    DnaRunAll = "ERR|" & Err.Number & "|" & Replace(Err.Description, vbLf, " ") & "|" & Err.Source
End Function
