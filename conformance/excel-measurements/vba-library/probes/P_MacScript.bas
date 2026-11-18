Attribute VB_Name = "P_MacScript"
Option Explicit
' MacScript (Interaction; Macintosh only, on Windows it raises an error).
Private Const N As Long = 5

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = MacScript("return 1")                                                 ' Windows: error
        Case 2: v = MacScript("")                                                         ' empty script
        Case 3: v = MacScript(Script:="return 1")                                         ' named argument
        Case 4: v = MacScript(Null)                                                       ' Null script: error 94
        Case 5: v = TypeName(MacScript("return 1"))                                       ' still an error on Windows
    End Select
    T = "MacScript." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "MacScript." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
