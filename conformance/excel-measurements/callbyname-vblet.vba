Function Run() As String
  Dim h As Holder, out As String
  Set h = New Holder
  On Error Resume Next
  Err.Clear: CallByName h, "PI", VbLet, "12": out = out & "I.txt=" & Err.Number & ":" & CStr(CallByName(h, "PI", VbGet)) & vbLf
  Err.Clear: CallByName h, "PI", VbLet, "1.5": out = out & "I.txtfrac=" & Err.Number & ":" & CStr(CallByName(h, "PI", VbGet)) & vbLf
  Err.Clear: CallByName h, "PI", VbLet, 7: out = out & "I.int=" & Err.Number & ":" & CStr(CallByName(h, "PI", VbGet)) & vbLf
  Err.Clear: CallByName h, "PI", VbLet, 2.5: out = out & "I.dbl=" & Err.Number & ":" & CStr(CallByName(h, "PI", VbGet)) & vbLf
  Err.Clear: CallByName h, "PL", VbLet, "12": out = out & "L.txt=" & Err.Number & ":" & CStr(CallByName(h, "PL", VbGet)) & vbLf
  Err.Clear: CallByName h, "PL", VbLet, "1.5": out = out & "L.txtfrac=" & Err.Number & ":" & CStr(CallByName(h, "PL", VbGet)) & vbLf
  Err.Clear: CallByName h, "PL", VbLet, 7: out = out & "L.int=" & Err.Number & ":" & CStr(CallByName(h, "PL", VbGet)) & vbLf
  Err.Clear: CallByName h, "PL", VbLet, 2.5: out = out & "L.dbl=" & Err.Number & ":" & CStr(CallByName(h, "PL", VbGet)) & vbLf
  Err.Clear: CallByName h, "PS", VbLet, "12": out = out & "S.txt=" & Err.Number & ":" & CStr(CallByName(h, "PS", VbGet)) & vbLf
  Err.Clear: CallByName h, "PS", VbLet, "1.5": out = out & "S.txtfrac=" & Err.Number & ":" & CStr(CallByName(h, "PS", VbGet)) & vbLf
  Err.Clear: CallByName h, "PS", VbLet, 7: out = out & "S.int=" & Err.Number & ":" & CStr(CallByName(h, "PS", VbGet)) & vbLf
  Err.Clear: CallByName h, "PS", VbLet, 2.5: out = out & "S.dbl=" & Err.Number & ":" & CStr(CallByName(h, "PS", VbGet)) & vbLf
  Err.Clear: CallByName h, "PD", VbLet, "12": out = out & "D.txt=" & Err.Number & ":" & CStr(CallByName(h, "PD", VbGet)) & vbLf
  Err.Clear: CallByName h, "PD", VbLet, "1.5": out = out & "D.txtfrac=" & Err.Number & ":" & CStr(CallByName(h, "PD", VbGet)) & vbLf
  Err.Clear: CallByName h, "PD", VbLet, 7: out = out & "D.int=" & Err.Number & ":" & CStr(CallByName(h, "PD", VbGet)) & vbLf
  Err.Clear: CallByName h, "PD", VbLet, 2.5: out = out & "D.dbl=" & Err.Number & ":" & CStr(CallByName(h, "PD", VbGet)) & vbLf
  Err.Clear: CallByName h, "PB", VbLet, "12": out = out & "B.txt=" & Err.Number & ":" & CStr(CallByName(h, "PB", VbGet)) & vbLf
  Err.Clear: CallByName h, "PB", VbLet, "1.5": out = out & "B.txtfrac=" & Err.Number & ":" & CStr(CallByName(h, "PB", VbGet)) & vbLf
  Err.Clear: CallByName h, "PB", VbLet, 7: out = out & "B.int=" & Err.Number & ":" & CStr(CallByName(h, "PB", VbGet)) & vbLf
  Err.Clear: CallByName h, "PB", VbLet, 2.5: out = out & "B.dbl=" & Err.Number & ":" & CStr(CallByName(h, "PB", VbGet)) & vbLf
  Err.Clear: CallByName h, "PC", VbLet, "12": out = out & "C.txt=" & Err.Number & ":" & CStr(CallByName(h, "PC", VbGet)) & vbLf
  Err.Clear: CallByName h, "PC", VbLet, "1.5": out = out & "C.txtfrac=" & Err.Number & ":" & CStr(CallByName(h, "PC", VbGet)) & vbLf
  Err.Clear: CallByName h, "PC", VbLet, 7: out = out & "C.int=" & Err.Number & ":" & CStr(CallByName(h, "PC", VbGet)) & vbLf
  Err.Clear: CallByName h, "PC", VbLet, 2.5: out = out & "C.dbl=" & Err.Number & ":" & CStr(CallByName(h, "PC", VbGet)) & vbLf
  Err.Clear: CallByName h, "PBo", VbLet, "12": out = out & "Bo.txt=" & Err.Number & ":" & CStr(CallByName(h, "PBo", VbGet)) & vbLf
  Err.Clear: CallByName h, "PBo", VbLet, "1.5": out = out & "Bo.txtfrac=" & Err.Number & ":" & CStr(CallByName(h, "PBo", VbGet)) & vbLf
  Err.Clear: CallByName h, "PBo", VbLet, 7: out = out & "Bo.int=" & Err.Number & ":" & CStr(CallByName(h, "PBo", VbGet)) & vbLf
  Err.Clear: CallByName h, "PBo", VbLet, 2.5: out = out & "Bo.dbl=" & Err.Number & ":" & CStr(CallByName(h, "PBo", VbGet)) & vbLf
  Err.Clear: CallByName h, "PSt", VbLet, "12": out = out & "St.txt=" & Err.Number & ":" & CStr(CallByName(h, "PSt", VbGet)) & vbLf
  Err.Clear: CallByName h, "PSt", VbLet, "1.5": out = out & "St.txtfrac=" & Err.Number & ":" & CStr(CallByName(h, "PSt", VbGet)) & vbLf
  Err.Clear: CallByName h, "PSt", VbLet, 7: out = out & "St.int=" & Err.Number & ":" & CStr(CallByName(h, "PSt", VbGet)) & vbLf
  Err.Clear: CallByName h, "PSt", VbLet, 2.5: out = out & "St.dbl=" & Err.Number & ":" & CStr(CallByName(h, "PSt", VbGet)) & vbLf
  Err.Clear: CallByName h, "PDt", VbLet, "12": out = out & "Dt.txt=" & Err.Number & ":" & CStr(CallByName(h, "PDt", VbGet)) & vbLf
  Err.Clear: CallByName h, "PDt", VbLet, "1.5": out = out & "Dt.txtfrac=" & Err.Number & ":" & CStr(CallByName(h, "PDt", VbGet)) & vbLf
  Err.Clear: CallByName h, "PDt", VbLet, 7: out = out & "Dt.int=" & Err.Number & ":" & CStr(CallByName(h, "PDt", VbGet)) & vbLf
  Err.Clear: CallByName h, "PDt", VbLet, 2.5: out = out & "Dt.dbl=" & Err.Number & ":" & CStr(CallByName(h, "PDt", VbGet)) & vbLf
  Run = out
End Function
