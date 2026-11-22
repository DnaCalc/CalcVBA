Attribute VB_Name = "P_MathExtra"
Option Explicit
' Exact-word witnesses captured in Excel VBA 16.0.20527.
Private Type DoubleValue
    value As Double
End Type

Private Type Words64
    lo As Long
    hi As Long
End Type

Private Function FromWords(ByVal hi As Long, ByVal lo As Long) As Double
    Dim words As Words64, number As DoubleValue
    words.hi = hi
    words.lo = lo
    LSet number = words
    FromWords = number.value
End Function

Public Function Run() As String
    Dim angle As Double
    angle = FromWords(&HC031475D, &HF8CD1620)
    Run = "MathExtra.1=" & Hex8(Sin(FromWords(&H40446B9A, &HBBF3BD55))) & vbLf & _
          "MathExtra.2=" & Hex8(Cos(angle)) & vbLf & _
          "MathExtra.3=" & Hex8(Tan(angle)) & vbLf & _
          "MathExtra.4=" & Hex8(Exp(FromWords(&H40823A91, &H0A596DD1))) & vbLf & _
          "MathExtra.5=" & Hex8(Exp(FromWords(&HC0855EF8, &H34E2FB8F))) & vbLf
End Function
