Function DblB() As String
  Dim x As Double, s As String, t As String
  On Error Resume Next
  Err.Clear: x = 1E-03: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-03=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-03: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-03=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-03: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-03=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-03: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-03=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-04: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-04=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-04: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-04=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-04: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-04=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-04: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-04=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-05: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-05=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-05: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-05=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-05: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-05=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-05: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-05=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-06: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-06=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-06: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-06=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-06: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-06=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-06: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-06=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-07: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-07=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-07: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-07=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-07: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-07=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-07: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-07=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-08: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-08=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-08: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-08=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-08: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-08=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-08: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-08=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-09: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-09=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-09: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-09=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-09: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-09=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-09: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-09=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-10: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-10=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-10: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-10=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-10: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-10=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-10: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-10=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-11: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-11=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-11: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-11=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-11: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-11=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-11: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-11=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-12: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-12=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-12: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-12=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-12: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-12=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-12: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-12=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-13: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-13=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-13: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-13=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-13: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-13=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-13: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-13=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-14: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-14=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-14: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-14=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-14: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-14=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-14: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-14=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-15: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-15=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-15: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-15=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-15: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-15=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-15: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-15=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-16: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-16=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-16: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-16=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-16: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-16=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-16: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-16=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-17: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-17=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-17: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-17=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-17: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-17=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-17: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-17=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-18: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-18=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-18: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-18=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-18: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-18=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-18: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-18=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-19: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-19=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-19: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-19=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-19: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-19=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-19: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-19=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-20: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-20=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-20: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-20=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-20: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-20=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-20: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-20=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-21: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-21=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-21: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-21=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-21: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-21=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-21: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-21=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-22: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-22=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-22: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-22=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-22: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-22=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-22: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-22=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-23: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E-23=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-23: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E-23=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-23: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E-23=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E-23: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E-23=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+13: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E+13=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+13: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E+13=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+13: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E+13=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E+13: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E+13=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+14: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E+14=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+14: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E+14=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+14: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E+14=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E+14: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E+14=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+15: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E+15=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+15: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E+15=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+15: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E+15=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E+15: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E+15=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+16: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E+16=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+16: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E+16=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+16: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E+16=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E+16: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E+16=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+17: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E+17=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+17: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E+17=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+17: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E+17=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E+17: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E+17=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+18: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E+18=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+18: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E+18=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+18: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E+18=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E+18: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E+18=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+19: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1E+19=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+19: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.5E+19=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+19: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.2345E+19=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.23456789012345E+19: s = CStr(x): t = Str$(x)
  DblB = DblB & "DblB.1.23456789012345E+19=" & s & "|" & t & "|" & Err.Number & vbLf
End Function
Function SglB() As String
  Dim x As Single, s As String, t As String
  On Error Resume Next
  Err.Clear: x = 1E-03: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E-03=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-03: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E-03=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-03: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E-03=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E-03: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E-03=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-04: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E-04=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-04: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E-04=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-04: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E-04=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E-04: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E-04=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-05: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E-05=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-05: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E-05=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-05: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E-05=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E-05: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E-05=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-06: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E-06=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-06: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E-06=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-06: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E-06=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E-06: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E-06=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-07: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E-07=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-07: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E-07=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-07: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E-07=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E-07: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E-07=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-08: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E-08=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-08: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E-08=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-08: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E-08=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E-08: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E-08=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-09: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E-09=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-09: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E-09=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-09: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E-09=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E-09: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E-09=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-10: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E-10=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-10: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E-10=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-10: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E-10=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E-10: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E-10=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-11: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E-11=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-11: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E-11=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-11: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E-11=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E-11: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E-11=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-12: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E-12=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-12: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E-12=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-12: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E-12=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E-12: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E-12=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E-13: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E-13=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E-13: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E-13=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E-13: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E-13=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E-13: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E-13=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+04: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E+04=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+04: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E+04=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+04: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E+04=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E+04: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E+04=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+05: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E+05=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+05: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E+05=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+05: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E+05=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E+05: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E+05=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+06: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E+06=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+06: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E+06=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+06: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E+06=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E+06: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E+06=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+07: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E+07=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+07: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E+07=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+07: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E+07=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E+07: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E+07=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+08: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E+08=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+08: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E+08=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+08: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E+08=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E+08: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E+08=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+09: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E+09=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+09: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E+09=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+09: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E+09=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E+09: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E+09=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+10: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E+10=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+10: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E+10=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+10: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E+10=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E+10: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E+10=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1E+11: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1E+11=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.5E+11: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.5E+11=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.2345E+11: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.2345E+11=" & s & "|" & t & "|" & Err.Number & vbLf
  Err.Clear: x = 1.234567E+11: s = CStr(x): t = Str$(x)
  SglB = SglB & "SglB.1.234567E+11=" & s & "|" & t & "|" & Err.Number & vbLf
End Function
Function ZeroB() As String
  Dim x As Double, y As Double, z As Double
  On Error Resume Next
  x = 0: y = -1: z = x * y
  ZeroB = "negzero=" & CStr(z) & "|" & Str$(z) & "|" & ("" & z) & "|" & Err.Number & vbLf
  Err.Clear: z = -x
  ZeroB = ZeroB & "negzero2=" & CStr(z) & "|" & Str$(z) & "|" & Err.Number & vbLf
  Err.Clear: z = CDbl(-0#)
  ZeroB = ZeroB & "negzero3=" & CStr(z) & "|" & Str$(z) & "|" & Err.Number & vbLf
End Function
Function RunAll() As String
  RunAll = DblB() & SglB() & ZeroB()
End Function
