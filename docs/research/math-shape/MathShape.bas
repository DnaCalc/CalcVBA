Attribute VB_Name = "MathShape"
Option Explicit
' Shape probe for dc-q6brh.112: Excel VBA's Sin/Cos/Tan/Atn/Exp/Log over many arguments, as hex,
' to tell which implementation VBE7 uses (CRT, correctly rounded, x87 kernel, ...).
Private Type TD
    v As Double
End Type
Private Type TL
    lo As Long
    hi As Long
End Type

Private Function H(ByVal x As Double) As String
    Dim d As TD, l As TL
    d.v = x
    LSet l = d
    H = Right$("0000000" & Hex$(l.hi), 8) & Right$("0000000" & Hex$(l.lo), 8)
End Function

Private Function Line1(ByVal f As String, ByVal x As Double) As String
    Dim y As Double
    On Error GoTo bad
    Select Case f
        Case "sin": y = Sin(x)
        Case "cos": y = Cos(x)
        Case "tan": y = Tan(x)
        Case "atn": y = Atn(x)
        Case "exp": y = Exp(x)
        Case "log": y = Log(x)
    End Select
    Line1 = f & " " & H(x) & " " & H(y)
    Exit Function
bad:
    Line1 = f & " " & H(x) & " ERR" & Err.Number
End Function

Private Function Args(ByVal f As String) As Variant
    Dim a() As Double, n As Long, k As Long
    ReDim a(1 To 2000)
    Select Case f
        Case "sin", "cos", "tan"
            For k = 1 To 1000: n = n + 1: a(n) = k * 0.0107: Next k
            For k = 1 To 500: n = n + 1: a(n) = k * 1.37: Next k
            For k = 1 To 200: n = n + 1: a(n) = 10# ^ (k / 10#) * 1.1: Next k
        Case "atn"
            For k = 1 To 1000: n = n + 1: a(n) = k * 0.0107: Next k
            For k = 1 To 500: n = n + 1: a(n) = k * 1.37: Next k
        Case "exp"
            For k = 1 To 1500: n = n + 1: a(n) = -20# + k * 0.0271: Next k
        Case "log"
            For k = 1 To 1000: n = n + 1: a(n) = k * 0.0107: Next k
            For k = 1 To 500: n = n + 1: a(n) = k * 13.7: Next k
    End Select
    ReDim Preserve a(1 To n)
    Args = a
End Function

Public Function Run() As String
    Dim fs As Variant, f As Variant, a As Variant, i As Long, parts() As String, n As Long
    fs = Array("sin", "cos", "tan", "atn", "exp", "log")
    ReDim parts(1 To 12000)
    For Each f In fs
        a = Args(f)
        For i = LBound(a) To UBound(a)
            n = n + 1
            parts(n) = Line1(f, a(i))
        Next i
    Next f
    ReDim Preserve parts(1 To n)
    Run = Join(parts, vbLf)
End Function
