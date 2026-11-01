Option Explicit

Private Function Advance(Optional ByRef n As Long = 40) As Long
    n = n + 2
    Advance = n
End Function
Private Function MissingFlag(Optional ByRef value As Variant) As Boolean
    MissingFlag = IsMissing(value)
End Function
Private Function ZeroDefault(Optional ByVal value As Long) As Long
    ZeroDefault = value
End Function
Private Function Mix(Optional ByVal left As Long = 4, Optional ByVal right As Double = 0.5) As Double
    Mix = left + right
End Function
Public Function RunProbe() As String
    On Error GoTo Failed
    Dim n As Long, result As Long, s As String
    s = "default=" & Advance()
    n = 40
    result = Advance(n)
    s = s & "|alias=" & result & ":" & n
    result = Advance((n))
    s = s & "|private=" & result & ":" & n
    s = s & "|missing=" & MissingFlag()
    s = s & "|present=" & MissingFlag(Empty)
    s = s & "|zero=" & ZeroDefault()
    s = s & "|named=" & Mix(right:=2.5)
    RunProbe = s
    Exit Function
Failed:
    RunProbe = "error=" & Err.Number & ":" & Err.Description
End Function
