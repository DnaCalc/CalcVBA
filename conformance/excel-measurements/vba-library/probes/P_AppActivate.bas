Attribute VB_Name = "P_AppActivate"
Option Explicit
' AppActivate (Interaction; only a window title that does not exist, never a real window).
Private Const N As Long = 3

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: AppActivate "NoSuchWindowTitle_CalcVbaProbe": v = "activated"              ' no such window: error 5
        Case 2: AppActivate "NoSuchWindowTitle_CalcVbaProbe", True: v = "activated"        ' with Wait
        Case 3: AppActivate Title:="NoSuchWindowTitle_CalcVbaProbe", Wait:=False: v = "activated"   ' named arguments
    End Select
    T = "AppActivate." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "AppActivate." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
