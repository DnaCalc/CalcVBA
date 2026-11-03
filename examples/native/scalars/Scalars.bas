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

' This procedure is outside the export closure and need not be AOT-admitted.
Public Function UnusedString(ByVal text As String) As String
    UnusedString = Left$(text, 1)
End Function
