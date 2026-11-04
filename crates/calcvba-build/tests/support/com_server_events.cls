Option Explicit
Public WithEvents Source As __LIB__.Calculator
Public Seen As Long
Private Sub Source_Changed(ByVal value As Variant)
    Seen = CLng(value)
End Sub
