Option Explicit
Public WithEvents Source As __LIB__.Calculator
Public Seen As Long
Public ReadBack As Long
Private Sub Source_Changed(ByVal value As Variant)
    Seen = Seen + CLng(value)
    ReadBack = Source.Value
    Source.Value = ReadBack + 1
    If CLng(value) = 99 Then Source.Fire 1
End Sub
