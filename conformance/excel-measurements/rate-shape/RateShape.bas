Attribute VB_Name = "RateShape"
Option Explicit
' VBA Rate: the shape of its iteration (dc-bjt7v). Each line is key=<IEEE bits hex>|<CStr> or
' key=ERRn, for Rate(nper, pmt, pv, fv, type, guess) over a grid chosen to tell candidate
' algorithms apart: a linear case (n=1, pmt=0) that plain Newton solves in one step, guesses on
' both sides of the root, slow-converging and non-converging inputs.
Private Type DblT
    d As Double
End Type
Private Type BitsT
    lo As Long
    hi As Long
End Type

Private Function Bits(ByVal d As Double) As String
    Dim a As DblT, b As BitsT
    a.d = d
    LSet b = a
    Bits = Right$("00000000" & Hex$(b.hi), 8) & Right$("00000000" & Hex$(b.lo), 8)
End Function

Private Function R(ByVal key As String, ByVal n As Double, ByVal pmt As Double, ByVal pv As Double, _
                   ByVal fv As Double, ByVal typ As Double, ByVal guess As Variant) As String
    Dim v As Double
    On Error GoTo fail
    If IsMissing(guess) Or IsEmpty(guess) Then
        v = Rate(n, pmt, pv, fv, typ)
    Else
        v = Rate(n, pmt, pv, fv, typ, guess)
    End If
    R = key & "=" & Bits(v) & "|" & CStr(v) & vbLf
    Exit Function
fail:
    R = key & "=ERR" & Err.Number & vbLf
End Function

Public Function Run() As String
    Dim s As String, g As Variant, f As Variant, i As Long
    ' A: linear case, root 19, various guesses
    For Each g In Array(Empty, 0.1, 0, 0.5, 1, 5, 10, 18, 18.9, 19, 19.1, 20, 25, -0.5, 100)
        s = s & R("lin19_g" & CStr(g), 1, 0, -1, 20, 0, g)
    Next
    ' B: linear case, other roots, default guess
    For Each f In Array(1.1, 1.5, 2, 3, 5, 11, 21, 50, 100, 1000, 10000, 1000000)
        s = s & R("lin_fv" & CStr(f), 1, 0, -1, f, 0, Empty)
    Next
    ' C: two-period and higher pure growth (roots exact: fv^(1/n)-1)
    For Each f In Array(1.21, 4, 9, 100)
        s = s & R("n2_fv" & CStr(f), 2, 0, -1, f, 0, Empty)
    Next
    s = s & R("n3_fv8", 3, 0, -1, 8, 0, Empty)
    s = s & R("n10_fv2", 10, 0, -1, 2, 0, Empty)
    ' D: annuities
    For Each g In Array(Empty, 0, -0.004, -0.0049, -0.005, 0.01, 0.5)
        s = s & R("a360_g" & CStr(g), 360, -1, 1000, 0, 0, g)
    Next
    s = s & R("a10", 10, -100, 500, 0, 0, Empty)
    s = s & R("a10due", 10, -100, 500, 0, 1, Empty)
    s = s & R("a12g05", 12, -10, 100, 0, 0, 0.5)
    s = s & R("a48", 48, -250, 10000, 0, 0, Empty)
    s = s & R("a48fv", 48, -250, 10000, -1000, 0, Empty)
    s = s & R("a5mixed", 5, -100, 300, 200, 0, Empty)
    s = s & R("a360due", 360, -6, 1000, 0, 1, Empty)
    s = s & R("a3exact", 3, -40, 100, 0, 0, Empty)
    ' E: convergence limits (iteration count, error 5)
    For i = 1 To 12
        s = s & R("far_n1_" & i, 1, 0, -1, 2 ^ i, 0, Empty)
    Next
    For Each g In Array(1, 2, 5, 10, 50)
        s = s & R("a360_far_g" & CStr(g), 360, -1, 1000, 0, 0, g)
    Next
    s = s & R("noroot", 10, 100, 100, 100, 0, Empty)
    s = s & R("zero_rate", 10, -10, 100, 0, 0, Empty)
    Run = s
End Function
