Attribute VB_Name = "Ops3Probe"
Option Explicit
Private Type TPair
    f As Long
    g As Variant
End Type
Function MakeV() As Variant
    Set MakeV = New Seven
End Function
Private Function Show(ByVal v As Variant) As String
    If IsObject(v) Then Show = "obj " & TypeName(v) Else Show = TypeName(v) & ":" & CStr(v)
End Function
Private Function Try(ByVal which As String) As String
    Dim w As Variant, s As Variant, l1 As Long, v As Variant, a(0 To 1) As Variant, u As TPair
    Set w = New Dflt: Set s = New Seven: l1 = 1
    a(0) = 1: Set a(1) = s: u.f = 1: Set u.g = s
    On Error GoTo fail
    Select Case which
        Case "g0": v = w + (1)
        Case "g1": v = w + (l1)
        Case "g2": v = w + CVar(1)
        Case "g3": v = w + Abs(1)
        Case "g4": v = w + CInt(l1)
        Case "g5": v = w + l1 * 1
        Case "g6": v = w + -1
        Case "g7": v = w + a(0)
        Case "g8": v = w + a(1)
        Case "g9": v = w + (s)
        Case "g10": v = w + CVar(s)
        Case "g11": v = w + u.f
        Case "g12": v = w + u.g
        Case "g13": v = (w) + 1
        Case "g14": v = a(1) + 1
        Case "g15": v = u.g + 1
        Case "g16": v = CVar(w) + 1
        Case "g17": v = MakeV() + MakeV()
        Case "g18": v = MakeV() + w
        Case "g19": v = (s) + w
        Case "g20": v = w Mod s
        Case "g21": v = s Mod w
        Case "g22": v = w \ s
        Case "g23": v = w ^ s
        Case "g24": v = w / s
        Case "g25": v = w And s
        Case "g26": v = w Like s
        Case "g27": v = w Mod 5
        Case "g28": v = w \ (5)
        Case "g33": v = -s
        Case "g34": v = Not s
        Case "g35": v = -(w + 1)
    End Select
    Try = Show(v)
    Exit Function
fail:
    Try = "raised " & Err.Number
End Function
Function Run() As String
    Dim out As String
    out = out & "w + (1)=" & Try("g0") & vbLf
    out = out & "w + (l1)=" & Try("g1") & vbLf
    out = out & "w + CVar(1)=" & Try("g2") & vbLf
    out = out & "w + Abs(1)=" & Try("g3") & vbLf
    out = out & "w + CInt(l1)=" & Try("g4") & vbLf
    out = out & "w + l1 * 1=" & Try("g5") & vbLf
    out = out & "w + -1=" & Try("g6") & vbLf
    out = out & "w + a(0)=" & Try("g7") & vbLf
    out = out & "w + a(1)=" & Try("g8") & vbLf
    out = out & "w + (s)=" & Try("g9") & vbLf
    out = out & "w + CVar(s)=" & Try("g10") & vbLf
    out = out & "w + u.f=" & Try("g11") & vbLf
    out = out & "w + u.g=" & Try("g12") & vbLf
    out = out & "(w) + 1=" & Try("g13") & vbLf
    out = out & "a(1) + 1=" & Try("g14") & vbLf
    out = out & "u.g + 1=" & Try("g15") & vbLf
    out = out & "CVar(w) + 1=" & Try("g16") & vbLf
    out = out & "MakeV() + MakeV()=" & Try("g17") & vbLf
    out = out & "MakeV() + w=" & Try("g18") & vbLf
    out = out & "(s) + w=" & Try("g19") & vbLf
    out = out & "w Mod s=" & Try("g20") & vbLf
    out = out & "s Mod w=" & Try("g21") & vbLf
    out = out & "w \ s=" & Try("g22") & vbLf
    out = out & "w ^ s=" & Try("g23") & vbLf
    out = out & "w / s=" & Try("g24") & vbLf
    out = out & "w And s=" & Try("g25") & vbLf
    out = out & "w Like s=" & Try("g26") & vbLf
    out = out & "w Mod 5=" & Try("g27") & vbLf
    out = out & "w \ (5)=" & Try("g28") & vbLf
    out = out & "-s=" & Try("g33") & vbLf
    out = out & "Not s=" & Try("g34") & vbLf
    out = out & "-(w + 1)=" & Try("g35") & vbLf
    Run = out
End Function
