Attribute VB_Name = "Scalars"
Option Explicit

Public Function AddLong(ByVal a As Long, ByVal b As Long) As Long
    AddLong = a + b
End Function

Private Function TwiceLong(ByVal a As Long) As Long
    TwiceLong = a * 2&
End Function

Public Function DirectLong(ByVal a As Long, ByVal b As Long) As Long
    DirectLong = TwiceLong(a) + b
End Function

Public Function AddDouble(ByVal a As Double, ByVal b As Double) As Double
    AddDouble = a + b
End Function

Public Function Greater(ByVal a As Long, ByVal b As Long) As Boolean
    Greater = a > b
End Function

Public Function EchoBoolean(ByVal value As Boolean) As Boolean
    EchoBoolean = value
End Function

Public Function Mixed5(ByVal a As Double, ByVal b As Long, ByVal flag As Boolean, ByVal d As Double, ByVal e As Long) As Double
    If flag Then
        Mixed5 = a + b + d + e
    Else
        Mixed5 = a - b + d - e
    End If
End Function

Public Function Divide(ByVal a As Double, ByVal b As Double) As Double
    Divide = a / b
End Function

Public Function LibraryAbs(ByVal value As Double) As Double
    LibraryAbs = Abs(value)
End Function

Public Function LibrarySqr(ByVal value As Double) As Double
    LibrarySqr = Sqr(value)
End Function

Public Function RoundLong(ByVal value As Double) As Long
    RoundLong = CLng(value)
End Function

Public Function RoundThenAdd(ByVal value As Double) As Long
    RoundThenAdd = RoundLong(value) + 1&
End Function

Public Function BooleanDouble(ByVal value As Boolean) As Double
    BooleanDouble = CDbl(value)
End Function

Public Function BooleanLong(ByVal value As Boolean) As Long
    BooleanLong = CLng(value)
End Function

Public Function DoubleBoolean(ByVal value As Double) As Boolean
    DoubleBoolean = CBool(value)
End Function

Public Function LongBoolean(ByVal value As Long) As Boolean
    LongBoolean = CBool(value)
End Function

Public Function SumTo(ByVal limit As Long) As Long
    Dim index As Long, total As Long
    index = 1&
    Do While index <= limit
        total = total + index
        index = index + 1&
    Loop
    SumTo = total
End Function

Public Function Depth(ByVal remaining As Long) As Long
    If remaining <= 0& Then
        Depth = 0&
    Else
        Depth = 1& + Depth(remaining - 1&)
    End If
End Function

' This procedure is outside the export closure and need not be AOT-admitted.
Public Function UnusedString(ByVal text As String) As String
    UnusedString = Left$(text, 1)
End Function

Public Function Weighted32(ByVal a01 As Long, ByVal a02 As Long, ByVal a03 As Long, ByVal a04 As Long, ByVal a05 As Long, ByVal a06 As Long, ByVal a07 As Long, ByVal a08 As Long, ByVal a09 As Long, ByVal a10 As Long, ByVal a11 As Long, ByVal a12 As Long, ByVal a13 As Long, ByVal a14 As Long, ByVal a15 As Long, ByVal a16 As Long, ByVal a17 As Long, ByVal a18 As Long, ByVal a19 As Long, ByVal a20 As Long, ByVal a21 As Long, ByVal a22 As Long, ByVal a23 As Long, ByVal a24 As Long, ByVal a25 As Long, ByVal a26 As Long, ByVal a27 As Long, ByVal a28 As Long, ByVal a29 As Long, ByVal a30 As Long, ByVal a31 As Long, ByVal a32 As Long) As Long
    Weighted32 = NativeInternals.WeightedBody32(a01, a02, a03, a04, a05, a06, a07, a08, a09, a10, a11, a12, a13, a14, a15, a16, a17, a18, a19, a20, a21, a22, a23, a24, a25, a26, a27, a28, a29, a30, a31, a32)
End Function
