Attribute VB_Name = "P_ShellReturnType"
Option Explicit
' UNMEASURED successful return type. Runner records OwnedTask() for child-process ownership.
Private Const N As Long = 1
Private mOwnedTask As Double
Public Function OwnedTask() As Double
    OwnedTask = mOwnedTask
End Function
Private Function T(ByVal k As Long) As String
    Dim task As Variant, v As Variant
    On Error GoTo fail
    Select Case k
        Case 1
            mOwnedTask = 0
            task = Shell("""" & Environ$("COMSPEC") & """ /d /c exit 0", vbHide)
            mOwnedTask = task
            v = TypeName(task)
    End Select
    T = "ShellReturnType." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "ShellReturnType." & k & "=ERR" & Err.Number
End Function
Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
