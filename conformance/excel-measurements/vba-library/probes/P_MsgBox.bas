Attribute VB_Name = "P_MsgBox"
Option Explicit
' MsgBox (Interaction; never called, only the constants used with it).
Private Const N As Long = 10

Private Function T(ByVal k As Long) As String
    Dim v As Variant
    On Error GoTo fail
    Select Case k
        Case 1: v = vbOKOnly & "," & vbOKCancel & "," & vbAbortRetryIgnore & "," & vbYesNoCancel & "," & vbYesNo & "," & vbRetryCancel   ' button styles
        Case 2: v = vbCritical & "," & vbQuestion & "," & vbExclamation & "," & vbInformation                                         ' icons
        Case 3: v = vbDefaultButton1 & "," & vbDefaultButton2 & "," & vbDefaultButton3 & "," & vbDefaultButton4                       ' default buttons
        Case 4: v = vbApplicationModal & "," & vbSystemModal                                                                           ' modality
        Case 5: v = vbMsgBoxHelpButton & "," & vbMsgBoxSetForeground & "," & vbMsgBoxRight & "," & vbMsgBoxRtlReading                 ' extra flags
        Case 6: v = vbOK & "," & vbCancel & "," & vbAbort & "," & vbRetry & "," & vbIgnore & "," & vbYes & "," & vbNo                  ' results
        Case 7: v = TypeName(vbYes)                                                       ' enum constant type
        Case 8: v = vbYesNo + vbQuestion + vbDefaultButton2                               ' typical combination
        Case 9: v = (vbYesNo Or vbExclamation) And 15                                     ' button bits
        Case 10: v = VarType(vbOKOnly)
    End Select
    T = "MsgBox." & k & "=" & Sh(v)
    Exit Function
fail:
    T = "MsgBox." & k & "=ERR" & Err.Number
End Function

Public Function Run() As String
    Dim k As Long, s As String
    For k = 1 To N
        s = s & T(k) & vbLf
    Next
    Run = s
End Function
