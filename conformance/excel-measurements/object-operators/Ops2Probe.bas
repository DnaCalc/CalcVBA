Attribute VB_Name = "Ops2Probe"
Option Explicit
Function MakeV() As Variant
    Set MakeV = New Seven
End Function
Private Function Show(ByVal v As Variant) As String
    If IsObject(v) Then Show = "obj " & TypeName(v) Else Show = TypeName(v) & ":" & CStr(v)
End Function
Private Function Try(ByVal which As String) As String
    Dim w As Variant, s As Variant, o As Object, t As Object, d As Dflt, x1 As Variant, l1 As Long, v As Variant
    Set w = New Dflt: Set s = New Seven: Set o = w: Set t = s: Set d = w: x1 = 1: l1 = 1
    On Error GoTo fail
    Select Case which
        Case "f0": v = w + 1
        Case "f1": v = w + l1
        Case "f2": v = w + x1
        Case "f3": v = w + 1#
        Case "f4": v = w + "1"
        Case "f5": v = w + True
        Case "f6": v = w + Empty
        Case "f7": v = w + Null
        Case "f8": v = w + o
        Case "f9": v = o + w
        Case "f10": v = w + (o)
        Case "f11": v = w + t
        Case "f12": v = t + w
        Case "f13": v = w + d
        Case "f14": v = d + w
        Case "f15": v = o + t
        Case "f16": v = o + 1
        Case "f17": v = w + s
        Case "f18": v = s + w
        Case "f19": v = w * s
        Case "f20": v = s * w
        Case "f21": v = w - s
        Case "f22": v = s - w
        Case "f23": v = w & s
        Case "f24": v = s & w
        Case "f25": v = w + w
        Case "f26": v = s + s
        Case "f27": v = w + MakeV()
        Case "f28": v = MakeV() + 1
        Case "f29": v = MakeV() + s
        Case "f30": v = w.Value + 1
        Case "f31": v = d + 1
        Case "f32": v = w + 1 + s
        Case "f33": v = CStr(w + 1)
        Case "f34": v = CStr(w + s)
        Case "f35": v = w = s
        Case "f36": v = w > s
        Case "f37": v = s < w
        Case "f38": v = w = 42
        Case "f39": v = w = o
    End Select
    Try = Show(v)
    Exit Function
fail:
    Try = "raised " & Err.Number
End Function
Function Run() As String
    Dim out As String
    out = out & "w + 1=" & Try("f0") & vbLf
    out = out & "w + l1=" & Try("f1") & vbLf
    out = out & "w + x1=" & Try("f2") & vbLf
    out = out & "w + 1#=" & Try("f3") & vbLf
    out = out & "w + ""1""=" & Try("f4") & vbLf
    out = out & "w + True=" & Try("f5") & vbLf
    out = out & "w + Empty=" & Try("f6") & vbLf
    out = out & "w + Null=" & Try("f7") & vbLf
    out = out & "w + o=" & Try("f8") & vbLf
    out = out & "o + w=" & Try("f9") & vbLf
    out = out & "w + (o)=" & Try("f10") & vbLf
    out = out & "w + t=" & Try("f11") & vbLf
    out = out & "t + w=" & Try("f12") & vbLf
    out = out & "w + d=" & Try("f13") & vbLf
    out = out & "d + w=" & Try("f14") & vbLf
    out = out & "o + t=" & Try("f15") & vbLf
    out = out & "o + 1=" & Try("f16") & vbLf
    out = out & "w + s=" & Try("f17") & vbLf
    out = out & "s + w=" & Try("f18") & vbLf
    out = out & "w * s=" & Try("f19") & vbLf
    out = out & "s * w=" & Try("f20") & vbLf
    out = out & "w - s=" & Try("f21") & vbLf
    out = out & "s - w=" & Try("f22") & vbLf
    out = out & "w & s=" & Try("f23") & vbLf
    out = out & "s & w=" & Try("f24") & vbLf
    out = out & "w + w=" & Try("f25") & vbLf
    out = out & "s + s=" & Try("f26") & vbLf
    out = out & "w + MakeV()=" & Try("f27") & vbLf
    out = out & "MakeV() + 1=" & Try("f28") & vbLf
    out = out & "MakeV() + s=" & Try("f29") & vbLf
    out = out & "w.Value + 1=" & Try("f30") & vbLf
    out = out & "d + 1=" & Try("f31") & vbLf
    out = out & "w + 1 + s=" & Try("f32") & vbLf
    out = out & "CStr(w + 1)=" & Try("f33") & vbLf
    out = out & "CStr(w + s)=" & Try("f34") & vbLf
    out = out & "w = s=" & Try("f35") & vbLf
    out = out & "w > s=" & Try("f36") & vbLf
    out = out & "s < w=" & Try("f37") & vbLf
    out = out & "w = 42=" & Try("f38") & vbLf
    out = out & "w = o=" & Try("f39") & vbLf
    Run = out
End Function
