Attribute VB_Name = "LogProbe"
Option Explicit
Function Run() As String
    Dim out As String, v As Variant
    v = CByte(1) And CByte(3): out = out & "CByte(1) And CByte(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CInt(1) And CByte(3): out = out & "CInt(1) And CByte(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = True And CByte(3): out = out & "True And CByte(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = True And CInt(3): out = out & "True And CInt(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CInt(1) And CInt(3): out = out & "CInt(1) And CInt(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CInt(1) And CLng(3): out = out & "CInt(1) And CLng(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CLng(1) And CLng(3): out = out & "CLng(1) And CLng(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CLng(1) And CLngLng(3): out = out & "CLng(1) And CLngLng(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = 1.5 And CInt(3): out = out & "1.5 And CInt(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CSng(1) And CInt(3): out = out & "CSng(1) And CInt(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CCur(1) And CInt(3): out = out & "CCur(1) And CInt(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = "1" And CInt(3): out = out & """1"" And CInt(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CDate(1) And CInt(3): out = out & "CDate(1) And CInt(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = Empty And CInt(3): out = out & "Empty And CInt(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = Empty And CByte(3): out = out & "Empty And CByte(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = Empty And Empty: out = out & "Empty And Empty=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = True And False: out = out & "True And False=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = True And Empty: out = out & "True And Empty=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = Not CByte(1): out = out & "Not CByte(1)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = Not CInt(1): out = out & "Not CInt(1)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = Not True: out = out & "Not True=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = Not CLng(1): out = out & "Not CLng(1)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = Not Empty: out = out & "Not Empty=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = Not CLngLng(1): out = out & "Not CLngLng(1)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CInt(1) Or CInt(2): out = out & "CInt(1) Or CInt(2)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CByte(1) Xor CByte(3): out = out & "CByte(1) Xor CByte(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CInt(5) Eqv CInt(3): out = out & "CInt(5) Eqv CInt(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CInt(5) Imp CInt(3): out = out & "CInt(5) Imp CInt(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    v = CByte(1) Imp CByte(3): out = out & "CByte(1) Imp CByte(3)=" & TypeName(v) & ":" & CStr(v) & vbLf
    Run = out
End Function
